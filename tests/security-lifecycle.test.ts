import { describe, expect, spyOn, test } from 'bun:test';
import { readdirSync, statSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import {
  buildExcelBuffer,
  createChunkedExcelStream,
  createMultiSheetExcelStream,
  exportExcelRowsToResponse,
  exportMultiSheetExcelToResponse,
  readCSV,
  readCSVStream,
  readExcelStream,
  writeExcel,
} from '../src';
import {
  createPrivateTempFile,
  removePrivateTempFile,
} from '../src/excel/runtime-utils';
import type { CSVReadOptions } from '../src/types';

function temporaryEntries(prefix: string): string[] {
  return readdirSync(tmpdir())
    .filter((name) => name.startsWith(prefix))
    .sort();
}

function expectPrivate(path: string): void {
  expect(statSync(path).mode % 0o1000).toBe(0o600);
  expect(statSync(dirname(path)).mode % 0o1000).toBe(0o700);
}

function csvSource(bytes: Uint8Array, chunkSize: number): Bun.BunFile {
  return {
    exists: async () => true,
    stat: async () => ({ size: bytes.length }),
    text: async () => new TextDecoder().decode(bytes),
    stream: () =>
      new ReadableStream<Uint8Array>({
        start(controller) {
          for (let i = 0; i < bytes.length; i += chunkSize) {
            controller.enqueue(bytes.subarray(i, i + chunkSize));
          }
          controller.close();
        },
      }),
  } as Bun.BunFile;
}

async function streamedCSV(source: Bun.BunFile, options?: CSVReadOptions) {
  const rows = [];
  for await (const row of readCSVStream(source, options)) rows.push(row);
  return rows;
}

describe('CSV boundary and length regressions', () => {
  for (const quote of [false, true]) {
    for (const length of [1_000_000, 1_000_001]) {
      test(`enforces decoded field length: quoted=${quote}, length=${length}`, async () => {
        const input = quote ? `"${'""'.repeat(length)}"` : 'x'.repeat(length);
        const source = csvSource(new TextEncoder().encode(input), 65536);
        if (length > 1_000_000) {
          await expect(readCSV(source)).rejects.toThrow('maximum length');
          await expect(streamedCSV(source)).rejects.toThrow('maximum length');
        } else {
          const expected = (quote ? '"' : 'x').repeat(length);
          expect(
            (await readCSV(source)).worksheets[0].rows[0].cells[0].value,
          ).toBe(expected);
          expect((await streamedCSV(source))[0].cells[0].value).toBe(expected);
        }
      });
    }
  }

  test('preserves escaped quotes, CRLF, quoted newlines and UTF-8 at every split', async () => {
    const input = 'name,value\r\n"Tiếng Việt 😀","a""b\r\nc"\r\n\r\n"",';
    const bytes = new TextEncoder().encode(input);
    for (const hasHeader of [false, true]) {
      for (const skipEmptyLines of [false, true]) {
        const options = { hasHeader, skipEmptyLines };
        const expected = (
          await readCSV(csvSource(bytes, bytes.length), options)
        ).worksheets[0].rows;
        for (let size = 1; size <= bytes.length; size++) {
          expect(await streamedCSV(csvSource(bytes, size), options)).toEqual(
            expected,
          );
        }
      }
    }
    expect(
      (await streamedCSV(csvSource(bytes, 1)))[1].cells.map(
        (cell) => cell.value,
      ),
    ).toEqual(['Tiếng Việt 😀', 'a"b\r\nc']);
  });

  test('flushes a truncated UTF-8 sequence and retains a quoted empty final row', async () => {
    expect(
      (await streamedCSV(csvSource(new Uint8Array([0x61, 0xe2, 0x82]), 1)))[0]
        .cells[0].value,
    ).toBe('a�');
    const source = csvSource(new TextEncoder().encode('""'), 1);
    expect(await streamedCSV(source, { skipEmptyLines: false })).toEqual([
      { cells: [{ value: null }] },
    ]);
    expect(await streamedCSV(source)).toEqual([]);
  });
});

describe('HTTP temporary output lifecycle', () => {
  for (const multiple of [false, true]) {
    const run = (options = {}) =>
      multiple
        ? exportMultiSheetExcelToResponse({
            sheets: [{ name: 'Sheet', rows: [['hello']] }],
            ...options,
          })
        : exportExcelRowsToResponse({ rows: [['hello']], ...options });
    test(`cleans header and completed callback failures, multi=${multiple}`, async () => {
      const before = temporaryEntries('bun-excel-response-');
      await expect(run({ filename: '\ud800.xlsx' })).rejects.toThrow();
      expect(temporaryEntries('bun-excel-response-')).toEqual(before);
      await expect(
        run({
          onProgress: ({ stage }: { stage: string }) => {
            if (stage === 'completed') throw new Error('callback failure');
          },
        }),
      ).rejects.toThrow('callback failure');
      expect(temporaryEntries('bun-excel-response-')).toEqual(before);
    });
    test(`cleans completed and cancelled response bodies, multi=${multiple}`, async () => {
      const before = temporaryEntries('bun-excel-response-');
      const complete = await run();
      expectPrivate(complete.diagnostics.target);
      expect(
        (await complete.response.arrayBuffer()).byteLength,
      ).toBeGreaterThan(0);
      expect(temporaryEntries('bun-excel-response-')).toEqual(before);
      const cancelled = await run();
      await cancelled.response.body?.cancel();
      expect(temporaryEntries('bun-excel-response-')).toEqual(before);
    });
  }

  test('cleans a failed body reader and releases its lock', async () => {
    const before = temporaryEntries('bun-excel-response-');
    const original = Bun.file;
    let failedStream: ReadableStream<Uint8Array<ArrayBuffer>> | undefined;
    const mock = spyOn(Bun, 'file').mockImplementation(((
      ...args: Parameters<typeof Bun.file>
    ) => {
      const file = original(...args);
      if (String(args[0]).includes('bun-excel-response-')) {
        file.stream = () => {
          failedStream = new ReadableStream({
            start(controller) {
              controller.error(new Error('read failure'));
            },
          });
          return failedStream;
        };
      }
      return file;
    }) as typeof Bun.file);
    try {
      const { response } = await exportExcelRowsToResponse({
        rows: [['hello']],
        validateOutput: false,
      });
      await expect(response.arrayBuffer()).rejects.toThrow('read failure');
      expect(failedStream?.locked).toBe(false);
      expect(temporaryEntries('bun-excel-response-')).toEqual(before);
    } finally {
      mock.mockRestore();
    }
  });
});

test('temporary file permissions remain private with umask 000', async () => {
  const previous = process.umask(0);
  let path: string | undefined;
  try {
    path = createPrivateTempFile('bun-excel-permissions');
    expectPrivate(path);
    for (const writer of [
      createChunkedExcelStream(path),
      createMultiSheetExcelStream(path),
    ]) {
      try {
        const directories = temporaryEntries('bun-xlsx-');
        expect(directories.length).toBeGreaterThan(0);
        for (const directory of directories)
          expectPrivate(join(tmpdir(), directory, 'data.tmp'));
      } finally {
        await writer.cancel();
      }
    }
  } finally {
    process.umask(previous);
    if (path) await removePrivateTempFile(path);
  }
});

test('streaming read spools use private files and clean up on early return', async () => {
  const path = createPrivateTempFile('bun-excel-read-security');
  const before = temporaryEntries('bun-excel-stream-');
  try {
    await Bun.write(
      path,
      buildExcelBuffer({
        worksheets: [
          { name: 'Sheet', rows: [{ cells: [{ value: 'secret' }] }] },
        ],
      }),
    );
    for await (const _row of readExcelStream(path)) {
      const created = temporaryEntries('bun-excel-stream-').filter(
        (entry) => !before.includes(entry),
      );
      expect(created.length).toBeGreaterThan(0);
      for (const entry of created)
        expectPrivate(join(tmpdir(), entry, 'data.tmp'));
      break;
    }
    expect(temporaryEntries('bun-excel-stream-')).toEqual(before);
  } finally {
    await removePrivateTempFile(path);
  }
});

test('large buffered export cleans a failed staging writer', async () => {
  const path = createPrivateTempFile('bun-excel-write-security');
  const before = temporaryEntries('bun-excel-write-');
  const original = Bun.file;
  let observed = false;
  const mock = spyOn(Bun, 'file').mockImplementation(((
    ...args: Parameters<typeof Bun.file>
  ) => {
    const file = original(...args);
    if (
      String(args[0]).includes('bun-excel-write-') &&
      String(args[0]) !== path
    ) {
      expectPrivate(String(args[0]));
      observed = true;
      file.writer = () => {
        throw new Error('staging failure');
      };
    }
    return file;
  }) as typeof Bun.file);
  try {
    await expect(
      writeExcel(
        path,
        {
          worksheets: [
            {
              name: 'Sheet',
              rows: [{ cells: [{ value: 'x'.repeat(2 * 1024 * 1024) }] }],
            },
          ],
        },
        { compress: false },
      ),
    ).rejects.toThrow('staging failure');
    expect(observed).toBe(true);
    expect(temporaryEntries('bun-excel-write-')).toEqual(before);
  } finally {
    mock.mockRestore();
    await removePrivateTempFile(path);
  }
});

test('partial writer initialization closes and removes previously allocated resources', async () => {
  const target = createPrivateTempFile('bun-excel-init-security');
  const before = temporaryEntries('bun-xlsx-');
  const original = Bun.file;
  const mock = spyOn(Bun, 'file').mockImplementation(((
    ...args: Parameters<typeof Bun.file>
  ) => {
    const file = original(...args);
    if (String(args[0]).includes('bun-xlsx-links-')) {
      file.writer = () => {
        throw new Error('writer initialization failed');
      };
    }
    return file;
  }) as typeof Bun.file);
  try {
    expect(() => createChunkedExcelStream(target)).toThrow(
      'writer initialization failed',
    );
    expect(() => createMultiSheetExcelStream(target)).toThrow(
      'writer initialization failed',
    );
    // Constructor cleanup closes asynchronous sinks before removing directories.
    for (let attempt = 0; attempt < 100; attempt++) {
      if (
        JSON.stringify(temporaryEntries('bun-xlsx-')) === JSON.stringify(before)
      )
        break;
      await Bun.sleep(5);
    }
    expect(temporaryEntries('bun-xlsx-')).toEqual(before);
  } finally {
    mock.mockRestore();
    await removePrivateTempFile(target);
  }
});

test('failed read spool initialization removes its private file', async () => {
  const path = createPrivateTempFile('bun-excel-read-security');
  await Bun.write(
    path,
    buildExcelBuffer({ worksheets: [{ name: 'Sheet', rows: [] }] }),
  );
  const before = temporaryEntries('bun-excel-stream-');
  const original = Bun.file;
  const mock = spyOn(Bun, 'file').mockImplementation(((
    ...args: Parameters<typeof Bun.file>
  ) => {
    const file = original(...args);
    if (String(args[0]).includes('bun-excel-stream-')) {
      expectPrivate(String(args[0]));
      file.writer = () => {
        throw new Error('spool initialization failed');
      };
    }
    return file;
  }) as typeof Bun.file);
  try {
    await expect(readExcelStream(path).next()).rejects.toThrow(
      'spool initialization failed',
    );
    expect(temporaryEntries('bun-excel-stream-')).toEqual(before);
  } finally {
    mock.mockRestore();
    await removePrivateTempFile(path);
  }
});

test('encrypted staging is private and cleaned when its writer fails', async () => {
  const path = createPrivateTempFile('bun-excel-encryption-security');
  const before = temporaryEntries('bun-excel-encrypted-');
  const original = Bun.file;
  const mock = spyOn(Bun, 'file').mockImplementation(((
    ...args: Parameters<typeof Bun.file>
  ) => {
    const file = original(...args);
    if (String(args[0]).includes('bun-excel-encrypted-')) {
      expectPrivate(String(args[0]));
      file.writer = () => {
        throw new Error('encrypted staging failed');
      };
    }
    return file;
  }) as typeof Bun.file);
  try {
    await expect(
      writeExcel(
        path,
        {
          worksheets: [
            {
              name: 'Sheet',
              rows: [{ cells: [{ value: 'x'.repeat(2 * 1024 * 1024) }] }],
            },
          ],
        },
        { compress: false, password: 'test-password' },
      ),
    ).rejects.toThrow('encrypted staging failed');
    expect(temporaryEntries('bun-excel-encrypted-')).toEqual(before);
  } finally {
    mock.mockRestore();
    await removePrivateTempFile(path);
  }
});

test('cancelling an HTTP body with a pending read waits for file cleanup', async () => {
  const before = temporaryEntries('bun-excel-response-');
  const original = Bun.file;
  let source: ReadableStream<Uint8Array<ArrayBuffer>> | undefined;
  let sourceCancelled = false;
  const mock = spyOn(Bun, 'file').mockImplementation(((
    ...args: Parameters<typeof Bun.file>
  ) => {
    const file = original(...args);
    if (String(args[0]).includes('bun-excel-response-')) {
      file.stream = () => {
        source = new ReadableStream({
          cancel() {
            sourceCancelled = true;
          },
        });
        return source;
      };
    }
    return file;
  }) as typeof Bun.file);
  try {
    const { response } = await exportExcelRowsToResponse({
      rows: [['hello']],
      validateOutput: false,
    });
    if (!response.body) throw new Error('Missing response body');
    const reader = response.body.getReader();
    const pending = reader.read();
    await reader.cancel();
    expect((await pending).done).toBe(true);
    expect(sourceCancelled).toBe(true);
    expect(source?.locked).toBe(false);
    expect(temporaryEntries('bun-excel-response-')).toEqual(before);
    reader.releaseLock();
  } finally {
    mock.mockRestore();
  }
});
