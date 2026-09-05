import { afterAll, beforeAll, expect, spyOn, test } from 'bun:test';
import { mkdirSync, readdirSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { unzipSync, zipSync } from 'fflate';
import {
  buildExcelBuffer,
  type ExcelReadStreamOptions,
  type ExcelReadStreamRow,
  type FileSource,
  readExcelStream,
} from '../src';

const TMP = './tests/.tmp-stream-read-selection';
const PATH = `${TMP}/sparse.xlsx`;
const encoder = new TextEncoder();
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
let parts: Record<string, Uint8Array>;
let baseline: ExcelReadStreamRow[];

function worksheet(rows: string): Uint8Array {
  return encoder.encode(
    `<worksheet xmlns="${NS}"><sheetData>${rows}</sheetData></worksheet>`,
  );
}

async function collect(
  source: FileSource = PATH,
  options?: ExcelReadStreamOptions,
) {
  const rows: ExcelReadStreamRow[] = [];
  for await (const row of readExcelStream(source, options)) rows.push(row);
  return rows;
}

function temporaryFiles() {
  return readdirSync(tmpdir())
    .filter((name) => name.startsWith('bun-excel-stream-'))
    .sort();
}

beforeAll(async () => {
  mkdirSync(TMP, { recursive: true });
  parts = unzipSync(
    buildExcelBuffer({
      worksheets: ['First', 'Second'].map((name) => ({
        name,
        rows: [
          {
            cells: [
              { value: 'shared text' },
              {
                value: new Date('2026-09-05T00:00:00Z'),
                style: { numberFormat: 'yyyy-mm-dd' },
              },
            ],
          },
        ],
      })),
    }),
  );
  for (let index = 1; index <= 2; index++) {
    parts[`xl/worksheets/sheet${index}.xml`] = worksheet(`
      <row r="1"><c r="A1" t="inlineStr"><is><t>header ${index}</t></is></c></row>
      <row r="3" ht="24" hidden="1" collapsed="1" outlineLevel="2">
        <c r="A3" t="s"><v>0</v></c>
        <c r="C3" s="1"><v>45000</v></c>
        <c r="F3"><f>SUM(20,22)</f><v>42</v></c>
        <c r="H3" t="b"><v>1</v></c>
        <c r="J3" t="inlineStr"><is><r><rPr><b/></rPr><t>rich</t></r><r><t> text</t></r></is></c>
      </row>
      <row r="7"><c r="C7"><v>${index * 7}</v></c><c r="F7"/></row>
      <row r="9"/>
      <row r="10"><c r="A10" t="inlineStr"><is><t>tail</t></is></c></row>
    `);
  }
  await Bun.write(PATH, zipSync(parts));
  baseline = await collect();
});
afterAll(() => rmSync(TMP, { recursive: true, force: true }));

test('default selection preserves existing rows and dense missing-cell placeholders', async () => {
  expect(baseline.map(({ rowIndex }) => rowIndex)).toEqual([
    0, 2, 6, 8, 9, 0, 2, 6, 8, 9,
  ]);
  expect(baseline[1].row.cells[1]).toEqual({ value: null });
  expect(
    await collect(PATH, { startRow: 0, endRow: 1_048_575, maxRows: 1_048_576 }),
  ).toEqual(baseline);
});

for (const [label, options, expectedIndices] of [
  ['start inclusive', { startRow: 2 }, [2, 6, 8, 9]],
  ['end inclusive', { endRow: 2 }, [0, 2]],
  ['one exact row', { startRow: 6, endRow: 6 }, [6]],
  ['missing row interval', { startRow: 3, endRow: 5 }, []],
  ['past last populated row', { startRow: 100 }, []],
  ['first row', { endRow: 0 }, [0]],
  [
    'limit counts actual rows after filtering',
    { startRow: 1, maxRows: 2 },
    [2, 6],
  ],
  ['empty rows count toward limit', { startRow: 7, maxRows: 1 }, [8]],
  [
    'limit beyond matching rows',
    { startRow: 6, endRow: 8, maxRows: 100 },
    [6, 8],
  ],
] as const) {
  test(`row selection: ${label}, independently per sheet`, async () => {
    const before = temporaryFiles();
    const rows = await collect(PATH, options);
    expect(rows.map(({ rowIndex }) => rowIndex)).toEqual([
      ...expectedIndices,
      ...expectedIndices,
    ]);
    expect(rows).toEqual(
      baseline.filter(({ rowIndex }) =>
        (expectedIndices as readonly number[]).includes(rowIndex),
      ),
    );
    expect(temporaryFiles()).toEqual(before);
  });
}

for (const sheets of [['Second'], [1]]) {
  test(`combines sheet selection ${JSON.stringify(sheets)} with rows and columns`, async () => {
    const rows = await collect(Bun.file(PATH), {
      sheets,
      startRow: 2,
      endRow: 6,
      maxRows: 1,
      columns: [5, 0, 5],
    });
    expect(rows).toHaveLength(1);
    expect(rows[0].sheetIndex).toBe(1);
    expect(rows[0].sheetName).toBe('Second');
    expect(rows[0].rowIndex).toBe(2);
    expect(Object.keys(rows[0].row.cells)).toEqual(['0', '5']);
    expect(rows[0].row.cells[0]).toEqual(baseline[6].row.cells[0]);
    expect(rows[0].row.cells[5]).toEqual(baseline[6].row.cells[5]);
    expect(rows[0].row.cells[1]).toBeUndefined();
  });
}

test('column projection preserves styles, dates, booleans, formula caches, rich text and row metadata', async () => {
  for (const includeStyles of [true, false]) {
    const options = {
      sheets: ['First'],
      startRow: 2,
      endRow: 2,
      includeStyles,
    };
    const [full] = await collect(PATH, options);
    const [selected] = await collect(PATH, {
      ...options,
      columns: [9, 7, 5, 2, 0],
    });
    expect({ ...selected.row, cells: [] }).toEqual({ ...full.row, cells: [] });
    expect(selected.row).toMatchObject({
      height: 24,
      hidden: true,
      collapsed: true,
      outlineLevel: 2,
    });
    for (const col of [0, 2, 5, 7, 9])
      expect(selected.row.cells[col]).toEqual(full.row.cells[col]);
    expect(selected.row.cells[2].value instanceof Date).toBe(includeStyles);
    expect(selected.row.cells[5]).toMatchObject({
      type: 'formula',
      value: 42,
      formula: 'SUM(20,22)',
    });
    expect(selected.row.cells[7].value).toBe(true);
    expect(selected.row.cells[9].richText).toBeDefined();
    expect(Object.keys(selected.row.cells)).toEqual(['0', '2', '5', '7', '9']);
  }
});

for (const columns of [[], [1, 4, 12]]) {
  test(`empty/missing columns ${JSON.stringify(columns)} keep rows without inventing cells`, async () => {
    const rows = await collect(PATH, { columns, maxRows: 2 });
    expect(rows.map(({ rowIndex }) => rowIndex)).toEqual([0, 2, 0, 2]);
    expect(rows.every(({ row }) => row.cells.length === 0)).toBe(true);
  });
}

test('an explicit blank selected cell remains present while missing cells remain holes', async () => {
  const [entry] = await collect(PATH, {
    sheets: ['First'],
    startRow: 6,
    maxRows: 1,
    columns: [0, 5],
  });
  expect(Object.keys(entry.row.cells)).toEqual(['5']);
  expect(entry.row.cells[5].value).toBeNull();
  expect(entry.row.cells[0]).toBeUndefined();
});

test('maxRows=0 returns without opening a source, but still validates other options', async () => {
  const file = Bun.file(`${TMP}/does-not-exist.xlsx`);
  const exists = spyOn(file, 'exists');
  const stream = spyOn(file, 'stream');
  try {
    expect(await collect(file, { maxRows: 0 })).toEqual([]);
    await expect(collect(file, { maxRows: 0, columns: [-1] })).rejects.toThrow(
      'columns entry',
    );
    expect(exists).not.toHaveBeenCalled();
    expect(stream).not.toHaveBeenCalled();
  } finally {
    exists.mockRestore();
    stream.mockRestore();
  }
});

for (const option of ['startRow', 'endRow', 'maxRows'] as const) {
  for (const invalid of [
    -1,
    0.5,
    Number.NaN,
    Number.POSITIVE_INFINITY,
    1_048_577,
    '1',
    null,
  ]) {
    test(`rejects invalid ${option}=${String(invalid)} before source I/O`, async () => {
      await expect(
        collect(`${TMP}/missing.xlsx`, {
          [option]: invalid,
        } as unknown as ExcelReadStreamOptions),
      ).rejects.toThrow(option);
    });
  }
}
for (const invalid of [
  -1,
  1.5,
  Number.NaN,
  Number.POSITIVE_INFINITY,
  16_384,
  'A',
  null,
  undefined,
]) {
  test(`rejects invalid column ${String(invalid)}`, async () => {
    await expect(
      collect(PATH, {
        columns: [invalid],
      } as unknown as ExcelReadStreamOptions),
    ).rejects.toThrow('columns entry');
  });
}
for (const invalid of [null, 'A:C', new Set([0]), 0]) {
  test(`rejects non-array columns ${String(invalid)}`, async () => {
    await expect(
      collect(PATH, { columns: invalid } as unknown as ExcelReadStreamOptions),
    ).rejects.toThrow('columns must be an array');
  });
}

test('rejects reversed ranges and out-of-range row coordinates', async () => {
  await expect(collect(PATH, { startRow: 3, endRow: 2 })).rejects.toThrow(
    'endRow',
  );
  await expect(collect(PATH, { startRow: 1_048_576 })).rejects.toThrow(
    'startRow',
  );
  await expect(collect(PATH, { endRow: 1_048_576 })).rejects.toThrow('endRow');
});

test('accepts Excel boundary coordinates without filling excluded columns', async () => {
  const path = `${TMP}/boundary.xlsx`;
  await Bun.write(
    path,
    zipSync({
      ...parts,
      'xl/worksheets/sheet1.xml': worksheet(
        '<row r="1048576"><c r="XFD1048576"><v>42</v></c></row>',
      ),
    }),
  );
  const [entry] = await collect(path, {
    sheets: ['First'],
    startRow: 1_048_575,
    endRow: 1_048_575,
    columns: [16_383],
    maxRows: 1,
  });
  expect(entry.rowIndex).toBe(1_048_575);
  expect(entry.row.cells.length).toBe(16_384);
  expect(Object.keys(entry.row.cells)).toEqual(['16383']);
  expect(entry.row.cells[16_383].value).toBe(42);
});

test('range filtering preserves XML order even when row indices decrease', async () => {
  const path = `${TMP}/unordered.xlsx`;
  await Bun.write(
    path,
    zipSync({
      ...parts,
      'xl/worksheets/sheet1.xml': worksheet(
        '<row r="10"/><row r="3"/><row r="1"/>',
      ),
    }),
  );
  expect(
    (await collect(path, { sheets: ['First'], endRow: 2, maxRows: 2 })).map(
      ({ rowIndex }) => rowIndex,
    ),
  ).toEqual([2, 0]);
});

test('missing or invalid row indices do not count toward maxRows', async () => {
  const path = `${TMP}/invalid-indices.xlsx`;
  await Bun.write(
    path,
    zipSync({
      ...parts,
      'xl/worksheets/sheet1.xml': worksheet(
        '<row/><row r="bad"/><row r="0"/><row r="1048577"/><row r="3"/>',
      ),
    }),
  );
  expect(
    (await collect(path, { sheets: ['First'], maxRows: 1 })).map(
      ({ rowIndex }) => rowIndex,
    ),
  ).toEqual([2]);
});

test('empty/missing selected sheets still yield no rows', async () => {
  for (const sheets of [[], ['Missing']])
    expect(
      await collect(PATH, { sheets, startRow: 2, columns: [0], maxRows: 1 }),
    ).toEqual([]);
});

test('column selection is snapshotted for the lifetime of the iterator', async () => {
  const columns = [0];
  const iterator = readExcelStream(PATH, { columns });
  try {
    await iterator.next();
    columns[0] = 5;
    const next = await iterator.next();
    expect(next.value?.row.cells[0]?.value).toBe('shared text');
    expect(next.value?.row.cells[5]).toBeUndefined();
  } finally {
    await iterator.return(undefined);
  }
});

test('consumer break and consumer error both remove temporary worksheets and strings', async () => {
  for (const fail of [false, true]) {
    const before = temporaryFiles();
    const consume = async () => {
      for await (const _ of readExcelStream(PATH, {
        columns: [0],
        startRow: 2,
        maxRows: 10,
      })) {
        if (fail) throw new Error('consumer failed');
        break;
      }
    };
    if (fail) await expect(consume()).rejects.toThrow('consumer failed');
    else await consume();
    expect(temporaryFiles()).toEqual(before);
  }
});

test('maxRows closes the parser early; unread XML is not full-document validation', async () => {
  const path = `${TMP}/unread-invalid-tail.xlsx`;
  await Bun.write(
    path,
    zipSync({
      ...parts,
      'xl/worksheets/sheet1.xml': worksheet(
        `<row r="1"><c r="A1" t="inlineStr"><is><t>${'x'.repeat(150_000)}</t></is></c></row><row r="2"><c></row>`,
      ),
    }),
  );
  const before = temporaryFiles();
  expect(
    await collect(path, { sheets: ['First'], columns: [], maxRows: 1 }),
  ).toHaveLength(1);
  expect(temporaryFiles()).toEqual(before);
  await expect(
    collect(path, { sheets: ['First'], columns: [] }),
  ).rejects.toThrow();
  expect(temporaryFiles()).toEqual(before);
});

test('filtered-out rows and columns still undergo native XML validation', async () => {
  const path = `${TMP}/invalid-filtered-row.xlsx`;
  await Bun.write(
    path,
    zipSync({
      ...parts,
      'xl/worksheets/sheet1.xml': worksheet(
        '<row r="1"><c r="B1"><v>bad</c></row><row r="3"/>',
      ),
    }),
  );
  const before = temporaryFiles();
  await expect(
    collect(path, { sheets: ['First'], startRow: 2, columns: [0] }),
  ).rejects.toThrow();
  expect(temporaryFiles()).toEqual(before);
});

test('selected reads work with small source chunks and reordered ZIP metadata', async () => {
  const entries = Object.fromEntries(Object.entries(parts).reverse());
  const bytes = new Uint8Array(zipSync(entries));
  const file = Bun.file(PATH);
  const stream = spyOn(file, 'stream').mockImplementation(() => {
    let offset = 0;
    return new ReadableStream<Uint8Array<ArrayBuffer>>({
      pull(controller) {
        if (offset === bytes.length) {
          controller.close();
          return;
        }
        const end = Math.min(offset + 7, bytes.length);
        controller.enqueue(bytes.subarray(offset, end));
        offset = end;
      },
    });
  });
  try {
    const rows = await collect(file, {
      sheets: ['Second'],
      startRow: 2,
      maxRows: 1,
      columns: [0],
    });
    expect(rows).toHaveLength(1);
    expect(rows[0].sheetIndex).toBe(1);
    expect(rows[0].rowIndex).toBe(2);
    expect(rows[0].row.cells[0].value).toBe('shared text');
    expect(stream).toHaveBeenCalledTimes(2);
  } finally {
    stream.mockRestore();
  }
});

test('range and limit stay correct across multiple native XML batches', async () => {
  const path = `${TMP}/multiple-batches.xlsx`;
  await Bun.write(
    path,
    buildExcelBuffer({
      worksheets: [
        {
          name: 'Wide',
          rows: Array.from({ length: 3000 }, (_, index) => ({
            cells: Array.from({ length: 8 }, (_, col) => ({
              value: index * 8 + col,
            })),
          })),
        },
      ],
    }),
  );
  const before = temporaryFiles();
  const rows = await collect(path, {
    startRow: 500,
    endRow: 2500,
    maxRows: 1600,
    columns: [1, 7],
  });
  expect(rows).toHaveLength(1600);
  for (let index = 0; index < rows.length; index++) {
    expect(rows[index].rowIndex).toBe(index + 500);
    expect(Object.keys(rows[index].row.cells)).toEqual(['1', '7']);
    expect(rows[index].row.cells[7].value).toBe((index + 500) * 8 + 7);
  }
  expect(temporaryFiles()).toEqual(before);
});
