import { afterAll, beforeAll, expect, test } from 'bun:test';
import { mkdirSync, readdirSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { unzipSync } from 'fflate';
import type { Workbook } from '../src';
import { buildExcelBuffer, readExcel, writeExcel } from '../src';

const TMP = './tests/.tmp-write-buffer';
beforeAll(() => mkdirSync(TMP, { recursive: true }));
afterAll(() => rmSync(TMP, { recursive: true, force: true }));

function temporaryArchives() {
  return readdirSync(tmpdir())
    .filter((name) => name.startsWith('bun-excel-write-'))
    .sort();
}

for (const compress of [true, false]) {
  test(`file and buffer exports preserve all parts and workbook input (compress=${compress})`, async () => {
    const workbook: Workbook = {
      worksheets: [
        {
          name: 'First',
          rows: [
            {
              cells: [
                {
                  value: 'Shared 😀',
                  comment: { text: 'Note' },
                  style: { font: { bold: true } },
                },
                { value: 123 },
              ],
            },
          ],
        },
        {
          name: 'Second',
          rows: [
            {
              cells: [
                { value: 'Shared 😀' },
                { formula: 'First!B1', value: 123 },
              ],
            },
          ],
        },
      ],
    };
    const original = structuredClone(workbook);
    const options = {
      compress,
      created: new Date('2026-09-04T00:00:00Z'),
      modified: new Date('2026-09-04T00:00:00Z'),
    };
    const expected = unzipSync(buildExcelBuffer(workbook, options));
    const before = temporaryArchives();
    const target = Bun.file(`${TMP}/parts-${compress}.xlsx`);
    await writeExcel(target, workbook, options);
    expect(unzipSync(await target.bytes())).toEqual(expected);
    expect(workbook).toEqual(original);
    expect((await readExcel(target)).worksheets[1].rows[0].cells[0].value).toBe(
      'Shared 😀',
    );
    expect(temporaryArchives()).toEqual(before);
  });
}

test('serialization failure on a later sheet preserves the target and removes staging files', async () => {
  const target = `${TMP}/existing.xlsx`;
  await Bun.write(target, 'existing content');
  const before = temporaryArchives();
  const workbook: Workbook = {
    worksheets: [
      {
        name: 'First',
        rows: Array.from({ length: 30000 }, (_, index) => ({
          cells: [{ value: index }],
        })),
      },
      {
        name: 'Invalid',
        rows: [],
        images: [
          {
            data: new Uint8Array([1]),
            format: 'png',
            range: { startRow: -1, startCol: 0, endRow: 0, endCol: 0 },
          },
        ],
      },
    ],
  };
  await expect(
    writeExcel(target, workbook, { compress: false }),
  ).rejects.toThrow('Invalid image startRow');
  expect(await Bun.file(target).text()).toBe('existing content');
  expect(temporaryArchives()).toEqual(before);
});

test('destination failure removes the staged archive', async () => {
  const before = temporaryArchives();
  await expect(
    writeExcel(
      TMP,
      {
        worksheets: [
          {
            name: 'Data',
            rows: [{ cells: [{ value: 'Large'.repeat(300000) }] }],
          },
        ],
      },
      { compress: false },
    ),
  ).rejects.toThrow();
  expect(temporaryArchives()).toEqual(before);
});

test('large uncompressed file export matches the buffer and removes its staged archive', async () => {
  const workbook: Workbook = {
    worksheets: [
      {
        name: 'Large',
        rows: Array.from({ length: 30000 }, (_, index) => ({
          cells: [{ value: index }],
        })),
      },
    ],
  };
  const options = {
    compress: false,
    created: new Date('2026-09-04T00:00:00Z'),
    modified: new Date('2026-09-04T00:00:00Z'),
  };
  const expected = buildExcelBuffer(workbook, options);
  expect(expected.length).toBeGreaterThan(1024 * 1024);
  const before = temporaryArchives();
  const path = `${TMP}/large.xlsx`;
  await writeExcel(path, workbook, options);
  expect(unzipSync(await Bun.file(path).bytes())).toEqual(unzipSync(expected));
  expect(temporaryArchives()).toEqual(before);
});
