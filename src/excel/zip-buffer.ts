import { Zip, type ZipInputFile } from 'fflate';

/** Emit each completed ZIP entry before requesting the next XLSX part. */
export function* zipChunks(
  files: Iterable<readonly [string, Uint8Array]>,
  compress: boolean,
): Generator<Uint8Array> {
  const chunks: Uint8Array[] = [];
  const zip = new Zip((error, data) => {
    if (error) throw error;
    chunks.push(data);
  });

  for (const [filename, data] of files) {
    const entry: ZipInputFile = {
      filename,
      size: data.byteLength,
      crc: Bun.hash.crc32(data),
      compression: compress ? 8 : 0,
    };
    zip.add(entry);
    // ZIP requires raw DEFLATE, without a zlib or gzip wrapper.
    const input =
      data.buffer instanceof ArrayBuffer
        ? new Uint8Array(data.buffer, data.byteOffset, data.byteLength)
        : new Uint8Array(data);
    const bytes = compress ? Bun.deflateSync(input, { level: 6 }) : data;
    entry.ondata?.(null, bytes, true);
    yield* chunks;
    chunks.length = 0;
  }
  zip.end();
  yield* chunks;
}

/** Package buffered XLSX parts using native DEFLATE and CRC32. */
export function zipBuffer(
  files: Record<string, Uint8Array> | Iterable<readonly [string, Uint8Array]>,
  compress: boolean,
): Uint8Array {
  const parts = Symbol.iterator in files ? files : Object.entries(files);
  const chunks = [
    ...zipChunks(parts as Iterable<readonly [string, Uint8Array]>, compress),
  ];
  return joinZipChunks(chunks);
}

export function joinZipChunks(chunks: readonly Uint8Array[]): Uint8Array {
  const size = chunks.reduce((sum, chunk) => sum + chunk.byteLength, 0);
  const result = new Uint8Array(size);
  let offset = 0;
  for (const chunk of chunks) {
    result.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return result;
}
