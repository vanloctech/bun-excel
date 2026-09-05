import {
  createCipheriv,
  createHash,
  createHmac,
  randomBytes,
} from 'node:crypto';
import { utils, write } from 'cfb';

// MS-OFFCRYPTO 2.3.4.10–15: Agile AES-256-CBC / SHA-512.
// https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/74d60145-a0f0-44be-99ce-c65d211b4eb7
const SPIN_COUNT = 100_000;
const BLOCK_KEYS = {
  verifier: Buffer.from('fea7d2763b4b9e79', 'hex'),
  verifierHash: Buffer.from('d7aa0f6d3061344e', 'hex'),
  key: Buffer.from('146e0be7abacd0d6', 'hex'),
  hmacKey: Buffer.from('5fb2ad010cb9e1f6', 'hex'),
  hmacValue: Buffer.from('a0677f02b22c8433', 'hex'),
};

export function validateExcelPassword(
  password: unknown,
): asserts password is string | undefined {
  if (password === undefined) return;
  if (
    typeof password !== 'string' ||
    password.length === 0 ||
    password.length > 255 ||
    password.includes('\0')
  )
    throw new Error(
      'password must contain 1–255 UTF-16 code units and no NUL characters',
    );
}

/** Streaming writers must never silently ignore a password. */
export function rejectStreamingPassword(options?: { password?: string }): void {
  if (options?.password !== undefined)
    throw new Error(
      'Password encryption is not supported by streaming writers; use writeExcel() or buildExcelBuffer()',
    );
}

function hash(...parts: Uint8Array[]): Buffer {
  const digest = createHash('sha512');
  for (const part of parts) digest.update(part);
  return digest.digest();
}

function encrypt(data: Uint8Array, key: Uint8Array, iv: Uint8Array): Buffer {
  const cipher = createCipheriv('aes-256-cbc', key, iv);
  cipher.setAutoPadding(false);
  return Buffer.concat([cipher.update(data), cipher.final()]);
}

/** Per-file key material, shared by buffered and disk-backed encryption. */
export function createPackageEncryption(password: string) {
  validateExcelPassword(password);
  const packageKey = randomBytes(32);
  const packageSalt = randomBytes(16);
  const passwordSalt = randomBytes(16);
  const passwordBytes = Buffer.from(password, 'utf16le');
  const hmacKey = randomBytes(64);
  let passwordHash = hash(passwordSalt, passwordBytes);
  const counter = Buffer.alloc(4);
  const dispose = () => {
    packageKey.fill(0);
    passwordBytes.fill(0);
    passwordHash.fill(0);
    hmacKey.fill(0);
  };
  try {
    for (let i = 0; i < SPIN_COUNT; i++) {
      counter.writeUInt32LE(i);
      const next = hash(counter, passwordHash);
      passwordHash.fill(0);
      passwordHash = next;
    }
  } catch (error) {
    dispose();
    throw error;
  }
  passwordBytes.fill(0);
  return {
    dispose,
    hmac: () => createHmac('sha512', hmacKey),
    segment(chunk: Uint8Array, index: number): Buffer {
      counter.writeUInt32LE(index);
      let padded: Uint8Array = chunk;
      if (chunk.byteLength % 16) {
        padded = Buffer.alloc(Math.ceil(chunk.byteLength / 16) * 16);
        padded.set(chunk);
      }
      return encrypt(
        padded,
        packageKey,
        hash(packageSalt, counter).subarray(0, 16),
      );
    },
    info(hmacDigest: Uint8Array): Buffer {
      const wrapKey = (block: Uint8Array, data: Uint8Array) => {
        const key = hash(passwordHash, block).subarray(0, 32);
        try {
          return encrypt(data, key, passwordSalt).toString('base64');
        } finally {
          key.fill(0);
        }
      };
      const integrity = (block: Uint8Array, data: Uint8Array) =>
        encrypt(
          data,
          packageKey,
          hash(packageSalt, block).subarray(0, 16),
        ).toString('base64');
      const verifier = randomBytes(16);
      const params =
        'saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512"';
      const xml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">' +
        `<keyData ${params} saltValue="${packageSalt.toString('base64')}"/>` +
        `<dataIntegrity encryptedHmacKey="${integrity(BLOCK_KEYS.hmacKey, hmacKey)}" encryptedHmacValue="${integrity(BLOCK_KEYS.hmacValue, hmacDigest)}"/>` +
        '<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">' +
        `<p:encryptedKey ${params} spinCount="${SPIN_COUNT}" saltValue="${passwordSalt.toString('base64')}" ` +
        `encryptedVerifierHashInput="${wrapKey(BLOCK_KEYS.verifier, verifier)}" encryptedVerifierHashValue="${wrapKey(BLOCK_KEYS.verifierHash, hash(verifier))}" encryptedKeyValue="${wrapKey(BLOCK_KEYS.key, packageKey)}"/>` +
        '</keyEncryptor></keyEncryptors></encryption>';
      return Buffer.concat([
        Buffer.from('0400040040000000', 'hex'),
        Buffer.from(xml),
      ]);
    },
  };
}

/** Wrap an XLSX ZIP in the Office encrypted compound-file format. */
export function encryptExcelPackage(
  input: Uint8Array,
  password: string,
): Uint8Array {
  const encryption = createPackageEncryption(password);
  try {
    const payload = Buffer.alloc(8 + Math.ceil(input.byteLength / 16) * 16);
    payload.writeBigUInt64LE(BigInt(input.byteLength));
    for (
      let offset = 0, segment = 0;
      offset < input.byteLength;
      offset += 4096, segment++
    )
      payload.set(
        encryption.segment(input.subarray(offset, offset + 4096), segment),
        offset + 8,
      );
    const container = utils.cfb_new();
    utils.cfb_add(
      container,
      'EncryptionInfo',
      encryption.info(encryption.hmac().update(payload).digest()),
    );
    utils.cfb_add(container, 'EncryptedPackage', payload);
    utils.cfb_del(container, '\u0001Sh33tJ5');
    return write(container, { type: 'buffer', fileType: 'cfb' }) as Uint8Array;
  } finally {
    encryption.dispose();
  }
}
