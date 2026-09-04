import { expect, test } from 'bun:test';
import { createNativeInflaters } from '../src/excel/native-inflate';

test('native raw DEFLATE handles split input and produces a final event', async () => {
  const source = new TextEncoder().encode('Tiếng Việt 😀 '.repeat(1000));
  const compressed = Bun.deflateSync(source);
  const pool = createNativeInflaters();
  const decoder = new pool.decoder();
  const chunks: Uint8Array[] = [];
  let finalSeen = false;
  decoder.ondata = (error, data, final) => {
    if (error) throw error;
    chunks.push(data);
    finalSeen ||= final;
  };
  for (let offset = 0; offset < compressed.length; offset += 7) {
    decoder.push(
      compressed.subarray(offset, offset + 7),
      offset + 7 >= compressed.length,
    );
    await pool.drain();
  }
  expect(finalSeen).toBe(true);
  expect(await new Blob(chunks as Uint8Array<ArrayBuffer>[]).text()).toBe(
    new TextDecoder().decode(source),
  );
});

test('native decompression errors propagate and can be cleaned up', async () => {
  const pool = createNativeInflaters();
  const decoder = new pool.decoder();
  decoder.ondata = () => {};
  decoder.push(new Uint8Array([255, 255, 255]), true);
  await expect(pool.drain()).rejects.toThrow();
  await pool.abort();
});

test('sink exceptions reject drain instead of leaving a blocked writer', async () => {
  const pool = createNativeInflaters();
  const decoder = new pool.decoder();
  decoder.ondata = () => {
    throw new Error('sink failed');
  };
  decoder.push(Bun.deflateSync(new TextEncoder().encode('data')), true);
  await expect(pool.drain()).rejects.toThrow('sink failed');
  await pool.abort();
});
