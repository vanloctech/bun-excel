import { afterAll, beforeAll, expect, spyOn, test } from 'bun:test';
import {
  createDecipheriv,
  createHash,
  createHmac,
  randomBytes,
} from 'node:crypto';
import { mkdirSync, readdirSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { find, read } from 'cfb';
import { unzipSync } from 'fflate';
import {
  buildExcelBuffer,
  createChunkedExcelStream,
  createExcelStream,
  createMultiSheetExcelStream,
  ExcelTemplate,
  type ExcelWriteOptions,
  exportExcelRows,
  exportExcelRowsToResponse,
  exportMultiSheetExcel,
  type Workbook,
  writeExcel,
} from '../src';
import { encryptedCfbMetadata } from '../src/excel/encrypted-cfb';
import { writeEncryptedExcelChunks } from '../src/excel/encrypted-file';
import { encryptExcelPackage } from '../src/excel/encryption';
import { findChild, parseXML } from '../src/excel/native-xml';

const TMP = './tests/.tmp-encryption';
const password = 'Mật khẩu 🔐';
const book: Workbook = {
  creator: 'Encryption test',
  created: new Date('2026-01-01T00:00:00Z'),
  modified: new Date('2026-01-01T00:00:00Z'),
  worksheets: [
    {
      name: 'Dữ liệu',
      rows: [
        {
          cells: [
            { value: 'private Việt 🔐' },
            { value: 42 },
            { value: false },
          ],
        },
        {
          cells: [
            { value: new Date('2026-01-02T00:00:00Z') },
            { value: 84, formula: 'B1*2', style: { font: { bold: true } } },
          ],
        },
      ],
    },
    {
      name: 'Second',
      rows: [
        {
          cells: [
            { value: 'Linked', hyperlink: { target: 'https://example.com' } },
          ],
        },
      ],
    },
  ],
};
beforeAll(() => mkdirSync(TMP, { recursive: true }));
afterAll(() => rmSync(TMP, { recursive: true, force: true }));

// Test-side decoder deliberately does not import encryption helpers/constants.
// Interoperability is also checked by the optional independent Python test below.
function decryptPackage(bytes: Uint8Array, secret: string): Buffer {
  const cfb = read(bytes, { type: 'buffer' });
  const info = find(cfb, 'EncryptionInfo');
  const payloadEntry = find(cfb, 'EncryptedPackage');
  if (!info || !payloadEntry) throw new Error('Missing encrypted streams');
  const descriptor = Buffer.from(info.content);
  expect(descriptor.subarray(0, 8).toString('hex')).toBe('0400040040000000');
  const encryption = parseXML(descriptor.subarray(8).toString());
  expect(encryption.name).toBe('encryption');
  const params = findChild(encryption, 'keyData')?.attributes;
  const integrity = findChild(encryption, 'dataIntegrity')?.attributes;
  const encryptors = findChild(encryption, 'keyEncryptors');
  const encryptor = encryptors && findChild(encryptors, 'keyEncryptor');
  const key = encryptor && findChild(encryptor, 'encryptedKey')?.attributes;
  if (!params || !integrity || !key) throw new Error('Incomplete descriptor');
  expect(params).toMatchObject({
    keyBits: '256',
    hashAlgorithm: 'SHA512',
    cipherAlgorithm: 'AES',
    cipherChaining: 'ChainingModeCBC',
    blockSize: '16',
  });
  expect(key.spinCount).toBe('100000');
  const digest = (data: Uint8Array) =>
    createHash('sha512').update(data).digest();
  const decode = (value: string) => Buffer.from(value, 'base64');
  const salt = decode(key.saltValue);
  let derived = digest(Buffer.concat([salt, Buffer.from(secret, 'utf16le')]));
  for (let i = 0; i < Number(key.spinCount); i++) {
    const round = Buffer.alloc(4);
    round.writeUInt32LE(i);
    derived = digest(Buffer.concat([round, derived]));
  }
  const aes = (data: Uint8Array, aesKey: Uint8Array, iv: Uint8Array) => {
    const decipher = createDecipheriv('aes-256-cbc', aesKey, iv);
    decipher.setAutoPadding(false);
    return Buffer.concat([decipher.update(data), decipher.final()]);
  };
  const unwrap = (value: string, block: string) =>
    aes(
      decode(value),
      digest(Buffer.concat([derived, Buffer.from(block, 'hex')])).subarray(
        0,
        32,
      ),
      salt,
    );
  const verifier = unwrap(key.encryptedVerifierHashInput, 'fea7d2763b4b9e79');
  if (
    !digest(verifier).equals(
      unwrap(key.encryptedVerifierHashValue, 'd7aa0f6d3061344e'),
    )
  )
    throw new Error('Wrong password');
  const packageKey = unwrap(key.encryptedKeyValue, '146e0be7abacd0d6');
  const payload = Buffer.from(payloadEntry.content);
  const iv = (block: Uint8Array) =>
    digest(Buffer.concat([decode(params.saltValue), block])).subarray(0, 16);
  const hmacKey = aes(
    decode(integrity.encryptedHmacKey),
    packageKey,
    iv(Buffer.from('5fb2ad010cb9e1f6', 'hex')),
  );
  const hmac = aes(
    decode(integrity.encryptedHmacValue),
    packageKey,
    iv(Buffer.from('a0677f02b22c8433', 'hex')),
  );
  if (!createHmac('sha512', hmacKey).update(payload).digest().equals(hmac))
    throw new Error('Integrity failed');
  const chunks: Buffer[] = [];
  for (
    let offset = 8, index = 0;
    offset < payload.length;
    offset += 4096, index++
  ) {
    const block = Buffer.alloc(4);
    block.writeUInt32LE(index);
    chunks.push(
      aes(payload.subarray(offset, offset + 4096), packageKey, iv(block)),
    );
  }
  return Buffer.concat(chunks).subarray(0, Number(payload.readBigUInt64LE(0)));
}

for (const size of [0, 1, 15, 16, 17, 4095, 4096, 4097, 8192]) {
  test(`encryption preserves ${size} bytes across AES/segment boundaries and sliced views`, () => {
    const backing = randomBytes(size + 20);
    const input = backing.subarray(7, 7 + size);
    const before = Buffer.from(backing);
    expect(
      decryptPackage(encryptExcelPackage(input, password), password),
    ).toEqual(input);
    expect(backing).toEqual(before);
  });
}
for (const compress of [true, false]) {
  test(`encrypted buffer preserves every workbook ZIP part (compress=${compress})`, () => {
    const plain = buildExcelBuffer(book, { compress });
    const encrypted = buildExcelBuffer(book, { compress, password });
    expect(Buffer.from(encrypted.subarray(0, 8)).toString('hex')).toBe(
      'd0cf11e0a1b11ae1',
    );
    expect(
      Buffer.from(encrypted).includes(Buffer.from('private Việt 🔐')),
    ).toBe(false);
    expect(unzipSync(decryptPackage(encrypted, password))).toEqual(
      unzipSync(plain),
    );
  });
}
test('same workbook/password uses independent randomness and rejects a wrong password', () => {
  const first = buildExcelBuffer(book, { password });
  const second = buildExcelBuffer(book, { password });
  expect(first).not.toEqual(second);
  expect(() => decryptPackage(first, 'wrong')).toThrow('Wrong password');
});
test('file, BunFile and template exports encrypt; undefined keeps ordinary ZIP output', async () => {
  for (const target of [`${TMP}/file.xlsx`, Bun.file(`${TMP}/bunfile.xlsx`)]) {
    await writeExcel(target, book, { password });
    const data =
      typeof target === 'string'
        ? await Bun.file(target).bytes()
        : await target.bytes();
    expect(unzipSync(decryptPackage(data, password))).toEqual(
      unzipSync(buildExcelBuffer(book)),
    );
  }
  const template = new ExcelTemplate(book);
  expect(
    unzipSync(decryptPackage(template.build({ password }), password)),
  ).toEqual(unzipSync(buildExcelBuffer(book)));
  await template.write(`${TMP}/template.xlsx`, { password });
  expect(
    unzipSync(
      decryptPackage(await Bun.file(`${TMP}/template.xlsx`).bytes(), password),
    ),
  ).toEqual(unzipSync(buildExcelBuffer(book)));
  expect(
    buildExcelBuffer(book, { password: undefined }).subarray(0, 2),
  ).toEqual(new Uint8Array([0x50, 0x4b]));
});
for (const invalid of ['', null, 123, true, 'x'.repeat(256), 'nul\0value']) {
  test(`invalid password is rejected before replacing a target (${typeof invalid})`, async () => {
    const path = `${TMP}/preserved.xlsx`;
    await Bun.write(path, 'keep existing');
    const options = { password: invalid } as ExcelWriteOptions;
    expect(() => buildExcelBuffer(book, options)).toThrow('password must');
    await expect(writeExcel(path, book, options)).rejects.toThrow(
      'password must',
    );
    expect(await Bun.file(path).text()).toBe('keep existing');
  });
}
for (const secret of [' ', 'x'.repeat(255)]) {
  test(`valid boundary password length ${secret.length} is preserved exactly`, () => {
    const input = buildExcelBuffer(book, { password: secret });
    expect(unzipSync(decryptPackage(input, secret))).toEqual(
      unzipSync(buildExcelBuffer(book)),
    );
  });
}
test('serialization failure preserves existing target and creates no plaintext staging files', async () => {
  const path = `${TMP}/failure.xlsx`;
  await Bun.write(path, 'keep existing');
  const before = readdirSync(tmpdir())
    .filter((p) => p.startsWith('bun-excel-write-'))
    .sort();
  const bad: Workbook = {
    worksheets: [
      {
        name: 'Bad',
        rows: [
          {
            cells: [
              {
                get value(): string {
                  throw new Error('serialization failed');
                },
              },
            ],
          },
        ],
      },
    ],
  };
  await expect(writeExcel(path, bad, { password })).rejects.toThrow(
    'serialization failed',
  );
  expect(await Bun.file(path).text()).toBe('keep existing');
  expect(
    readdirSync(tmpdir())
      .filter((p) => p.startsWith('bun-excel-write-'))
      .sort(),
  ).toEqual(before);
});
test('streaming APIs reject passwords before writing files or consuming rows', async () => {
  const target = `${TMP}/stream.xlsx`;
  for (const create of [
    createExcelStream,
    createChunkedExcelStream,
    createMultiSheetExcelStream,
  ])
    expect(() => create(target, { password })).toThrow(
      'not supported by streaming',
    );
  let consumed = false;
  function* rows() {
    consumed = true;
    yield [];
  }
  for (const mode of ['stream', 'chunked'] as const)
    await expect(
      exportExcelRows({ target, rows: rows(), password, mode }),
    ).rejects.toThrow('not supported by streaming');
  await expect(
    exportExcelRowsToResponse({ rows: rows(), password }),
  ).rejects.toThrow('not supported by streaming');
  await expect(
    exportMultiSheetExcel({
      target,
      sheets: [{ name: 'A', rows: rows(), options: { password } }],
    }),
  ).rejects.toThrow('not supported by streaming');
  expect(await Bun.file(target).exists()).toBe(false);
  expect(consumed).toBe(false);
});

test.skipIf(!process.env.MSOFFCRYPTO_PYTHON)(
  'independent msoffcrypto-tool verifies passwords, integrity and all XLSX parts',
  async () => {
    const plainPath = `${TMP}/independent-plain.xlsx`;
    const encryptedPath = `${TMP}/independent-encrypted.xlsx`;
    const large: Workbook = {
      ...book,
      worksheets: [
        ...book.worksheets,
        {
          name: 'Large',
          rows: Array.from({ length: 4200 }, (_, i) => ({
            cells: [{ value: i }, { value: `Row ${i} ${'x'.repeat(4000)}` }],
          })),
        },
      ],
    };
    for (const compress of [true, false]) {
      await Bun.write(plainPath, buildExcelBuffer(large, { compress }));
      await writeExcel(encryptedPath, large, { password, compress });
      const child = Bun.spawn(
        [
          process.env.MSOFFCRYPTO_PYTHON ?? 'python3',
          'tests/helpers/verify-encryption.py',
          encryptedPath,
          plainPath,
        ],
        { stdin: new Blob([password]), stdout: 'pipe', stderr: 'pipe' },
      );
      const [status, stdout, stderr] = await Promise.all([
        child.exited,
        new Response(child.stdout).text(),
        new Response(child.stderr).text(),
      ]);
      expect({ status, stderr }).toEqual({ status: 0, stderr: '' });
      expect(stdout).toContain('checks passed');
    }
  },
  30_000,
);

function encryptedTemps() {
  return readdirSync(tmpdir())
    .filter((name) => name.startsWith('bun-excel-encrypted-'))
    .sort();
}
for (const size of [1024 * 1024 - 1, 1024 * 1024, 1024 * 1024 + 1]) {
  test(`hybrid threshold and arbitrary ZIP chunk boundaries: ${size}`, async () => {
    const bytes = randomBytes(size);
    const path = `${TMP}/threshold.xlsx`;
    const before = encryptedTemps();
    let staged = false;
    function* chunks() {
      for (let offset = 0; offset < bytes.length; offset += 997)
        yield bytes.subarray(offset, offset + 997);
      staged = encryptedTemps().length > before.length;
    }
    await writeEncryptedExcelChunks(path, chunks(), password);
    expect(staged).toBe(size > 1024 * 1024);
    expect(decryptPackage(await Bun.file(path).bytes(), password)).toEqual(
      bytes,
    );
    expect(encryptedTemps()).toEqual(before);
  });
}
test('large encryption never stages plaintext and cleans up after source failure', async () => {
  const before = encryptedTemps();
  const target = `${TMP}/source-failure.xlsx`;
  await Bun.write(target, 'keep existing');
  let closed = false;
  const reason = new Error('source failed after staging');
  function* chunks() {
    try {
      yield Buffer.alloc(1024 * 1024 + 1, 0x61);
      const directory = encryptedTemps().find((name) => !before.includes(name));
      expect(directory).toBeDefined();
      const staged = readFileSync(`${tmpdir()}/${directory}/package.xlsx`);
      expect(staged.length).toBeGreaterThan(1024 * 1024);
      expect(
        staged.includes(Buffer.from('aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa')),
      ).toBe(false);
      throw reason;
    } finally {
      closed = true;
    }
  }
  await expect(
    writeEncryptedExcelChunks(target, chunks(), password),
  ).rejects.toBe(reason);
  expect(closed).toBe(true);
  expect(await Bun.file(target).text()).toBe('keep existing');
  expect(encryptedTemps()).toEqual(before);
});
test('large encryption destination failure removes its staged container', async () => {
  const before = encryptedTemps();
  await expect(
    writeEncryptedExcelChunks(TMP, [Buffer.alloc(1024 * 1024 + 1)], password),
  ).rejects.toThrow();
  expect(encryptedTemps()).toEqual(before);
});
test('CFB metadata enforces stream and file limits without allocating large payloads', () => {
  for (const size of [
    0,
    4095,
    Number.NaN,
    Number.POSITIVE_INFINITY,
    4096.5,
    0x80000000,
    0x7fffffff,
  ])
    expect(() => encryptedCfbMetadata(size, new Uint8Array())).toThrow();
  expect(() => encryptedCfbMetadata(4096, new Uint8Array(4097))).toThrow(
    'EncryptionInfo',
  );
  const metadata = encryptedCfbMetadata(20 * 1024 * 1024, new Uint8Array());
  expect(metadata.header.readUInt32LE(72)).toBeGreaterThan(1);
});

for (const failure of ['short', 'error'] as const) {
  test(`integrity read ${failure} preserves the target and removes staging files`, async () => {
    const path = `${TMP}/integrity-read-failure.xlsx`;
    await Bun.write(path, 'keep existing');
    const before = encryptedTemps();
    const nativeFile = Bun.file.bind(Bun);
    let intercepted = false;
    const factory = spyOn(Bun, 'file').mockImplementation((source, options) => {
      let file: Bun.BunFile;
      if (typeof source === 'string' || source instanceof URL)
        file = nativeFile(source, options);
      else if (typeof source === 'number') file = nativeFile(source, options);
      else file = nativeFile(source, options);
      if (
        typeof source === 'string' &&
        source.includes('bun-excel-encrypted-')
      ) {
        const slice = file.slice.bind(file);
        spyOn(file, 'slice').mockImplementation(
          (start?: number | string, end?: number | string, type?: string) => {
            intercepted = true;
            if (failure === 'error') throw new Error('integrity read failed');
            if (typeof start === 'string') return slice(start);
            if (typeof end === 'string') return slice(start, end);
            return slice(start, (end ?? file.size) - 16, type);
          },
        );
      }
      return file;
    });
    try {
      await expect(
        writeEncryptedExcelChunks(
          path,
          [Buffer.alloc(1024 * 1024 + 1)],
          password,
        ),
      ).rejects.toThrow(
        failure === 'short'
          ? 'Incomplete encrypted package'
          : 'integrity read failed',
      );
      expect(intercepted).toBe(true);
      expect(await nativeFile(path).text()).toBe('keep existing');
      expect(encryptedTemps()).toEqual(before);
    } finally {
      factory.mockRestore();
    }
  });
}
