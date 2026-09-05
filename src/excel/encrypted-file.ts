import { type FileHandle, mkdtemp, open, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { FileTarget } from '../types';
import { encryptedCfbMetadata } from './encrypted-cfb';
import { createPackageEncryption, encryptExcelPackage } from './encryption';
import { joinZipChunks } from './zip-buffer';

const MEMORY_LIMIT = 1024 * 1024;

async function writeAt(file: FileHandle, bytes: Uint8Array, position: number) {
  let offset = 0;
  while (offset < bytes.length) {
    const { bytesWritten } = await file.write(
      bytes,
      offset,
      bytes.length - offset,
      position + offset,
    );
    if (!bytesWritten) throw new Error('Encrypted file write made no progress');
    offset += bytesWritten;
  }
}

async function writeLargeEncryptedFile(
  target: FileTarget,
  chunks: Iterable<Uint8Array>,
  password: string,
) {
  const directory = await mkdtemp(join(tmpdir(), 'bun-excel-encrypted-'));
  const path = join(directory, 'package.xlsx');
  let sink: Bun.FileSink | undefined;
  let handle: FileHandle | undefined;
  let encryption: ReturnType<typeof createPackageEncryption> | undefined;
  try {
    await (await open(path, 'wx', 0o600)).close();
    const file = Bun.file(path);
    encryption = createPackageEncryption(password);
    sink = file.writer({ highWaterMark: 256 * 1024 });
    sink.write(Buffer.alloc(520)); // CFB header + original ZIP length, patched after serialization.
    const partial = Buffer.alloc(4096);
    let used = 0;
    let index = 0;
    let zipSize = 0;
    let buffered = 520;
    for (const chunk of chunks) {
      zipSize += chunk.length;
      if (zipSize >= 0x80000000)
        throw new Error(
          'Encrypted compound file exceeds the 2 GiB CFB v3 limit',
        );
      let offset = 0;
      while (offset < chunk.length) {
        const count = Math.min(4096 - used, chunk.length - offset);
        partial.set(chunk.subarray(offset, offset + count), used);
        offset += count;
        used += count;
        if (used === 4096) {
          sink.write(encryption.segment(partial, index++));
          used = 0;
          buffered += 4096;
          if (buffered >= 256 * 1024) {
            await sink.flush();
            buffered = 0;
          }
        }
      }
    }
    if (used) sink.write(encryption.segment(partial.subarray(0, used), index));
    partial.fill(0);
    await sink.end();
    sink = undefined;
    handle = await open(path, 'r+');
    const length = Buffer.alloc(8);
    length.writeBigUInt64LE(BigInt(zipSize));
    await writeAt(handle, length, 512);
    const payloadSize = 8 + Math.ceil(zipSize / 16) * 16;
    const hmac = encryption.hmac();
    let integrityBytes = 0;
    // The HMAC covers the final length prefix, so calculate it after patching.
    const reader = file
      .slice(512, 512 + payloadSize)
      .stream()
      .getReader();
    try {
      for (;;) {
        const { value, done } = await reader.read();
        if (done) break;
        integrityBytes += value.byteLength;
        hmac.update(value);
      }
    } finally {
      try {
        await reader.cancel();
      } finally {
        reader.releaseLock();
      }
    }
    if (integrityBytes !== payloadSize)
      throw new Error(
        'Incomplete encrypted package while calculating integrity',
      );
    const metadata = encryptedCfbMetadata(
      payloadSize,
      encryption.info(hmac.digest()),
    );
    await writeAt(handle, metadata.footer, metadata.footerOffset);
    await writeAt(handle, metadata.header, 0);
    await handle.close();
    handle = undefined;
    await Bun.write(target, file);
  } finally {
    encryption?.dispose();
    if (sink) {
      try {
        await sink.end();
      } catch {
        /* Preserve the original error. */
      }
    }
    if (handle) await handle.close().catch(() => {});
    await rm(directory, { recursive: true, force: true });
  }
}

/** Select the path using actual serialized ZIP size, not estimated row counts. */
export async function writeEncryptedExcelChunks(
  target: FileTarget,
  chunks: Iterable<Uint8Array>,
  password: string,
): Promise<void> {
  const iterator = chunks[Symbol.iterator]();
  const pending: Uint8Array[] = [];
  let size = 0;
  try {
    for (;;) {
      const next = iterator.next();
      if (next.done) break;
      pending.push(next.value);
      size += next.value.length;
      if (size > MEMORY_LIMIT) {
        function* replay() {
          yield* pending;
          pending.length = 0;
          yield* { [Symbol.iterator]: () => iterator };
        }
        await writeLargeEncryptedFile(target, replay(), password);
        return;
      }
    }
    await Bun.write(
      target,
      encryptExcelPackage(joinZipChunks(pending), password),
    );
  } finally {
    iterator.return?.();
  }
}
