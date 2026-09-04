import { afterAll, beforeAll, expect, test } from 'bun:test';
import { mkdirSync, rmSync } from 'node:fs';
import { unzipSync } from 'fflate';
import type { FileTarget } from '../src';
import { StreamingZipWriter } from '../src/excel/zip-stream';

const TMP = './tests/.tmp-native-zip';
beforeAll(() => mkdirSync(TMP, { recursive: true }));
afterAll(() => rmSync(TMP, { recursive: true, force: true }));

for (const level of [0, 1, 6, 9] as const) {
  test(`native streaming ZIP preserves bytes and CRC across parts (level=${level})`, async () => {
    const target = `${TMP}/level-${level}.zip`;
    const data = new TextEncoder().encode('Tiếng Việt 😀 '.repeat(20000));
    const writer = new StreamingZipWriter(target, { level });
    await writer.addFile('empty.xml', []);
    await writer.addFile('文字.xml', [
      data.subarray(0, 13),
      new Blob([data.subarray(13, 150000)]),
      data.subarray(150000),
      'tail',
    ]);
    await writer.close();
    const bytes = await Bun.file(target).bytes();
    const contents = unzipSync(bytes);
    expect(contents['empty.xml']).toEqual(new Uint8Array(0));
    const expected = new Uint8Array(data.length + 4);
    expected.set(data);
    expected.set(new TextEncoder().encode('tail'), data.length);
    expect(contents['文字.xml']).toEqual(expected);
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let offset = view.getUint32(bytes.length - 6, true);
    for (const name of ['empty.xml', '文字.xml']) {
      expect(view.getUint16(offset + 10, true)).toBe(8);
      expect(view.getUint32(offset + 16, true)).toBe(
        Bun.hash.crc32(contents[name]),
      );
      expect(view.getUint32(offset + 24, true)).toBe(contents[name].length);
      offset +=
        46 +
        view.getUint16(offset + 28, true) +
        view.getUint16(offset + 30, true) +
        view.getUint16(offset + 32, true);
    }
  });
}

function mockTarget(
  flush: () => Promise<void> | void,
  write = (_chunk: unknown) => 0,
): FileTarget {
  return {
    writer: () => ({ write, flush, end() {} }),
  } as unknown as FileTarget;
}

function randomChunk() {
  const chunk = new Uint8Array(64 * 1024);
  crypto.getRandomValues(chunk);
  return chunk;
}

test('native compression respects a slow destination and cancels input on sink failure', async () => {
  let release!: () => void;
  let started!: () => void;
  const flushing = new Promise<void>((resolve) => {
    started = resolve;
  });
  const blocked = new Promise<void>((resolve) => {
    release = resolve;
  });
  let pulls = 0;
  let cancelled = false;
  const chunk = randomChunk();
  const blob = {
    stream: () =>
      new ReadableStream<Uint8Array>({
        pull(controller) {
          if (++pulls > 128) controller.close();
          else controller.enqueue(chunk);
        },
        cancel() {
          cancelled = true;
        },
      }),
  } as Blob;
  const writer = new StreamingZipWriter(
    mockTarget(async () => {
      started();
      await blocked;
      throw new Error('sink failed');
    }),
    { flushThreshold: 1 },
  );
  const result = writer.addFile('large.bin', [blob]);
  void result.catch(() => {});
  await flushing;
  await Bun.sleep(20);
  const before = pulls;
  await Bun.sleep(20);
  release();
  expect(pulls).toBe(before);
  expect(pulls).toBeLessThan(20);
  await expect(result).rejects.toThrow('sink failed');
  expect(cancelled).toBe(true);
  await expect(writer.close()).rejects.toThrow('sink failed');
});

test('input errors propagate through native compression without hanging', async () => {
  const blob = {
    stream: () =>
      new ReadableStream<Uint8Array>({
        pull(controller) {
          controller.error(new Error('input failed'));
        },
      }),
  } as Blob;
  const writer = new StreamingZipWriter(mockTarget(() => {}));
  await expect(writer.addFile('broken.bin', [blob])).rejects.toThrow(
    'input failed',
  );
  await expect(writer.close()).rejects.toThrow('input failed');
});

test('synchronous destination errors cancel a pending source read', async () => {
  let cancelled = false;
  let pulls = 0;
  const chunk = randomChunk();
  const blob = {
    stream: () =>
      new ReadableStream<Uint8Array>({
        pull(controller) {
          if (++pulls <= 2) controller.enqueue(chunk);
        },
        cancel() {
          cancelled = true;
        },
      }),
  } as Blob;
  const writer = new StreamingZipWriter(
    mockTarget(
      () => {},
      () => {
        throw new Error('write failed');
      },
    ),
  );
  await expect(writer.addFile('broken.bin', [blob])).rejects.toThrow(
    'write failed',
  );
  expect(cancelled).toBe(true);
});

for (const compress of [true, false]) {
  test(`invalid ZIP entry names fail before opening input (compress=${compress})`, async () => {
    let opened = false;
    const blob = {
      stream() {
        opened = true;
        return new ReadableStream<Uint8Array>({
          start(controller) {
            controller.close();
          },
        });
      },
    } as Blob;
    const writer = new StreamingZipWriter(
      mockTarget(() => {}),
      { compress },
    );
    await expect(writer.addFile('x'.repeat(65536), [blob])).rejects.toThrow();
    expect(opened).toBe(false);
    await expect(writer.close()).rejects.toThrow();
  });
}
