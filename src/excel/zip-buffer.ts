import { Zip, type ZipInputFile } from 'fflate';

/** Package buffered XLSX parts using native DEFLATE and CRC32. */
export function zipBuffer(
  files: Record<string, Uint8Array>,
  compress: boolean,
): Uint8Array {
  const chunks: Uint8Array[] = [];
  let size = 0;
  const zip = new Zip((error, data) => {
    if (error) throw error;
    chunks.push(data);
    size += data.byteLength;
  });

  for (const [filename, data] of Object.entries(files)) {
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
  }
  zip.end();

  const result = new Uint8Array(size);
  let offset = 0;
  for (const chunk of chunks) {
    result.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return result;
}
