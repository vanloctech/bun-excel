import { afterAll, beforeAll, expect, spyOn, test } from 'bun:test';
import { mkdirSync, readdirSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import {
  buildExcelBuffer,
  type ExcelReadProgress,
  type ExcelReadStreamOptions,
  type FileSource,
  readExcelStream,
} from '../src';

const TMP = './tests/.tmp-read-progress';
const PATH = `${TMP}/data.xlsx`;
beforeAll(async () => {
  mkdirSync(TMP, { recursive: true });
  await Bun.write(
    PATH,
    buildExcelBuffer({
      worksheets: ['First', 'Second'].map((name) => ({
        name,
        rows: Array.from({ length: 25 }, (_, index) => ({
          cells: [{ value: `${name}-${index}` }],
        })),
      })),
    }),
  );
});
afterAll(() => rmSync(TMP, { recursive: true, force: true }));
async function consume(
  options?: ExcelReadStreamOptions,
  source: FileSource = PATH,
) {
  const rows = [];
  for await (const row of readExcelStream(source, options)) rows.push(row);
  return rows;
}
function temps() {
  return readdirSync(tmpdir())
    .filter((name) => name.startsWith('bun-excel-stream-'))
    .sort();
}

test('progress matches selected rows, sheet limits, bytes, strings and completion after cleanup', async () => {
  const events: ExcelReadProgress[] = [];
  const before = temps();
  const rows = await consume({
    sheets: [0, 1],
    startRow: 2,
    maxRows: 5,
    columns: [0],
    progressIntervalRows: 2,
    onProgress(event) {
      events.push(event);
      if (event.stage === 'completed') expect(temps()).toEqual(before);
    },
  });
  expect(rows).toHaveLength(10);
  expect(events[0].stage).toBe('metadata');
  expect(events.at(-1)).toMatchObject({
    stage: 'completed',
    rowsRead: 10,
    sheetRowsRead: 5,
    sheetIndex: 1,
    sharedStringsRead: 50,
  });
  expect(events.at(-1)?.bytesRead).toBeGreaterThanOrEqual(
    (await Bun.file(PATH).stat()).size,
  );
  for (let i = 1; i < events.length; i++) {
    expect(events[i].rowsRead).toBeGreaterThanOrEqual(events[i - 1].rowsRead);
    expect(events[i].bytesRead).toBeGreaterThanOrEqual(events[i - 1].bytesRead);
    expect(events[i].elapsedMs).toBeGreaterThanOrEqual(events[i - 1].elapsedMs);
  }
  expect(
    events
      .filter((e) => e.stage === 'reading' && e.sheetIndex === 0)
      .map((e) => e.sheetRowsRead),
  ).toEqual([0, 2, 4, 5]);
  expect(new Set(events.map((e) => e.stage))).toEqual(
    new Set([
      'metadata',
      'extracting',
      'sharedStrings',
      'reading',
      'completed',
    ]),
  );
});

test('pre-aborted signals reject before I/O, including maxRows=0', async () => {
  const controller = new AbortController();
  const reason = new Error('stop before reading');
  controller.abort(reason);
  const source = Bun.file(PATH);
  const exists = spyOn(source, 'exists');
  try {
    for (const maxRows of [undefined, 0])
      await expect(
        consume({ signal: controller.signal, maxRows }, source),
      ).rejects.toBe(reason);
    expect(exists).not.toHaveBeenCalled();
  } finally {
    exists.mockRestore();
  }
});

for (const stage of [
  'metadata',
  'extracting',
  'sharedStrings',
  'reading',
] as const) {
  test(`abort during ${stage} propagates the reason and cleans temporary files`, async () => {
    const before = temps();
    const controller = new AbortController();
    const reason = new Error(`stop ${stage}`);
    const events: ExcelReadProgress[] = [];
    await expect(
      consume({
        signal: controller.signal,
        onProgress(event) {
          events.push(event);
          if (
            event.stage === stage &&
            (stage !== 'extracting' || event.bytesRead > 0)
          )
            controller.abort(reason);
        },
      }),
    ).rejects.toBe(reason);
    expect(events.some((e) => e.stage === 'completed')).toBe(false);
    expect(temps()).toEqual(before);
  });
}

test('abort after a yielded row stops subsequent rows and cleans up', async () => {
  const before = temps();
  const controller = new AbortController();
  const iterator = readExcelStream(PATH, { signal: controller.signal });
  expect((await iterator.next()).done).toBe(false);
  controller.abort();
  await expect(iterator.next()).rejects.toMatchObject({ name: 'AbortError' });
  expect(temps()).toEqual(before);
});

for (const phase of ['exists', 'stat', 'stream'] as const) {
  test(`abort interrupts a pending ${phase} operation`, async () => {
    const controller = new AbortController();
    const source = Bun.file(PATH);
    const reason = new Error('pending cancelled');
    let entered!: () => void;
    const ready = new Promise<void>((resolve) => {
      entered = resolve;
    });
    let input: ReadableStream<Uint8Array<ArrayBuffer>> | undefined;
    const mock =
      phase === 'stream'
        ? spyOn(source, 'stream').mockImplementation(() => {
            input = new ReadableStream({
              pull() {
                entered();
              },
            });
            return input;
          })
        : spyOn(source, phase).mockImplementation(() => {
            entered();
            return new Promise<never>(() => {});
          });
    const work = consume({ signal: controller.signal }, source);
    const outcome = work.then(
      () => undefined,
      (error) => error,
    );
    try {
      await ready;
      controller.abort(reason);
      expect(await outcome).toBe(reason);
      if (input) expect(input.locked).toBe(false);
    } finally {
      mock.mockRestore();
    }
  });
}

for (const stage of [
  'extracting',
  'sharedStrings',
  'reading',
  'completed',
] as const) {
  test(`callback rejection in ${stage} preserves its error and removes temporary files`, async () => {
    const before = temps();
    const reason = new Error('callback failed');
    await expect(
      consume({
        onProgress: async (event) => {
          if (
            event.stage === stage &&
            (stage !== 'extracting' || event.bytesRead > 0)
          )
            throw reason;
        },
      }),
    ).rejects.toBe(reason);
    expect(temps()).toEqual(before);
  });
}

test('abort interrupts an unresolved async callback without waiting for user code', async () => {
  const before = temps();
  const controller = new AbortController();
  let entered!: () => void;
  const ready = new Promise<void>((resolve) => {
    entered = resolve;
  });
  const work = consume({
    signal: controller.signal,
    onProgress(event) {
      if (event.stage === 'sharedStrings') {
        entered();
        return new Promise<void>(() => {});
      }
      return undefined;
    },
  });
  const outcome = work.then(
    () => undefined,
    (error) => error,
  );
  await ready;
  controller.abort();
  expect(await outcome).toMatchObject({ name: 'AbortError' });
  expect(temps()).toEqual(before);
});

test('callbacks are awaited and progress snapshots cannot mutate internal counters', async () => {
  const counts: number[] = [];
  let active = false;
  await consume({
    maxRows: 2,
    progressIntervalRows: 1,
    async onProgress(event) {
      expect(active).toBe(false);
      active = true;
      counts.push(event.rowsRead);
      event.rowsRead = -999;
      await Promise.resolve();
      active = false;
    },
  });
  expect(counts.at(-1)).toBe(4);
  expect(counts.every((count) => count >= 0)).toBe(true);
});

test('consumer break cleans up without reporting successful completion', async () => {
  const before = temps();
  const events: ExcelReadProgress[] = [];
  for await (const _ of readExcelStream(PATH, {
    onProgress: (event) => {
      events.push(event);
    },
  }))
    break;
  expect(events.some((e) => e.stage === 'completed')).toBe(false);
  expect(temps()).toEqual(before);
});

test('zero rows and empty sheet selections complete with zero row counters', async () => {
  for (const options of [
    { maxRows: 0 },
    { sheets: [] },
    { sheets: ['Missing'] },
  ]) {
    const events: ExcelReadProgress[] = [];
    expect(
      await consume({
        ...options,
        onProgress: (event) => {
          events.push(event);
        },
      }),
    ).toEqual([]);
    expect(events.at(-1)).toMatchObject({ stage: 'completed', rowsRead: 0 });
  }
});

for (const interval of [
  0,
  -1,
  1.5,
  Number.NaN,
  Number.POSITIVE_INFINITY,
  null,
]) {
  test(`rejects invalid progress interval ${String(interval)}`, async () => {
    await expect(
      consume({ progressIntervalRows: interval } as ExcelReadStreamOptions),
    ).rejects.toThrow('progressIntervalRows');
  });
}

test('abort while ZIP extraction has created a worksheet sink cleans up active decoders and files', async () => {
  const before = temps();
  const controller = new AbortController();
  const reason = new Error('abort active extraction');
  const originalFile = Bun.file.bind(Bun);
  let triggered = false;
  const file = spyOn(Bun, 'file').mockImplementation((path, options) => {
    let result: Bun.BunFile;
    if (typeof path === 'string' || path instanceof URL)
      result = originalFile(path, options);
    else if (typeof path === 'number') result = originalFile(path, options);
    else result = originalFile(path, options);
    if (
      !triggered &&
      typeof path === 'string' &&
      path.includes('bun-excel-stream-')
    ) {
      triggered = true;
      queueMicrotask(() => controller.abort(reason));
    }
    return result;
  });
  try {
    await expect(consume({ signal: controller.signal })).rejects.toBe(reason);
    expect(triggered).toBe(true);
    expect(temps()).toEqual(before);
  } finally {
    file.mockRestore();
  }
});
