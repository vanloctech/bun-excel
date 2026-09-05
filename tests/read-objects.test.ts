import { afterAll, beforeAll, expect, expectTypeOf, test } from 'bun:test';
import { mkdirSync, readdirSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { strToU8, unzipSync, zipSync } from 'fflate';
import {
  buildExcelBuffer,
  type CellValue,
  ExcelHeaderError,
  type ExcelObjectData,
  type ExcelObjectReadOptions,
  type ExcelObjectSchema,
  type ExcelReadProgress,
  readExcelObjectsStream,
  type Workbook,
} from '../src';

const TMP = './tests/.tmp-read-objects';
let id = 0;
beforeAll(() => mkdirSync(TMP, { recursive: true }));
afterAll(() => rmSync(TMP, { recursive: true, force: true }));
async function workbook(book: Workbook) {
  const path = `${TMP}/${id++}.xlsx`;
  await Bun.write(path, buildExcelBuffer(book));
  return path;
}
async function fixture(rows: CellValue[][]) {
  return workbook({
    worksheets: [
      {
        name: 'Products',
        rows: rows.map((cells) => ({
          cells: cells.map((value) => ({ value })),
        })),
      },
    ],
  });
}
async function collect<const S extends ExcelObjectSchema>(
  path: string,
  options: ExcelObjectReadOptions<S>,
) {
  return Array.fromAsync(readExcelObjectsStream(path, options));
}
const schema = {
  name: { header: 'Tên sản phẩm', type: 'string', required: true },
  price: { header: 'Giá', type: 'number', required: true },
  quantity: { header: 'Số lượng', type: 'number' },
} as const;
function temps() {
  return readdirSync(tmpdir())
    .filter((p) => p.startsWith('bun-excel-stream-'))
    .sort();
}
test('maps Vietnamese headers, infers types, reports invalid cells and continues', async () => {
  const path = await fixture([
    ['Tên sản phẩm', 'Giá', 'Số lượng'],
    ['Bàn phím', 450000, 10],
    ['Chuột', 'abc', 5],
    ['Cáp', 0],
  ]);
  const rows = await collect(path, { schema });
  expect(rows).toHaveLength(3);
  expect(rows[0]).toEqual({
    ok: true,
    sheetIndex: 0,
    sheetName: 'Products',
    rowIndex: 1,
    data: { name: 'Bàn phím', price: 450000, quantity: 10 },
  });
  expect(rows[1]).toMatchObject({
    ok: false,
    rowIndex: 2,
    errors: [
      {
        key: 'price',
        cell: 'B3',
        code: 'invalid_type',
        expected: 'number',
        value: 'abc',
      },
    ],
  });
  expect(rows[2]).toMatchObject({
    ok: true,
    data: { name: 'Cáp', price: 0, quantity: null },
  });
  if (rows[0].ok) {
    expectTypeOf(rows[0].data.price).toEqualTypeOf<number>();
    expectTypeOf(rows[0].data.name).toEqualTypeOf<string>();
    expectTypeOf(rows[0].data.quantity).toEqualTypeOf<number | null>();
    // @ts-expect-error no undeclared fields
    void rows[0].data.unknown;
  }
  expectTypeOf<
    ExcelObjectData<{
      active: { header: 'Active'; type: 'boolean'; required: true };
    }>['active']
  >().toEqualTypeOf<boolean>();
});
test('binds headers separately per sheet, supports selection and a nonfirst header', async () => {
  const path = await workbook({
    worksheets: [
      {
        name: 'A',
        rows: [
          { cells: [] },
          { cells: [{ value: 'Tên sản phẩm' }, { value: 'Giá' }] },
          { cells: [{ value: 'A' }, { value: 1 }] },
        ],
      },
      {
        name: 'B',
        rows: [
          { cells: [] },
          { cells: [{ value: 'Giá' }, { value: 'Tên sản phẩm' }] },
          { cells: [{ value: 2 }, { value: 'B' }] },
        ],
      },
    ],
  });
  const rows = await collect(path, { schema, headerRow: 1 });
  expect(rows.map((r) => r.ok && r.data)).toEqual([
    { name: 'A', price: 1, quantity: null },
    { name: 'B', price: 2, quantity: null },
  ]);
  expect(await collect(path, { schema, headerRow: 1, sheets: ['B'] })).toEqual([
    rows[1],
  ]);
  expect(await collect(path, { schema, sheets: [] })).toEqual([]);
});
for (const [value, expected] of [
  ['12', 12],
  [' -1.5e2 ', -150],
  ['.5', 0.5],
  ['0x10', undefined],
  ['Infinity', undefined],
  ['1e999', undefined],
  ['1,000', undefined],
  [' ', undefined],
  [true, undefined],
] as const) {
  test(`explicit number coercion: ${String(value)}`, async () => {
    const path = await fixture([['Value'], [value]]);
    const [row] = await collect(path, {
      schema: { value: { header: 'Value', type: 'number', coerce: true } },
    });
    if (expected === undefined) expect(row.ok).toBe(false);
    else expect(row).toMatchObject({ ok: true, data: { value: expected } });
  });
}
test('strict defaults; explicit boolean/string conversions do not guess dates or truthiness', async () => {
  const path = await fixture([
    ['N', 'B', 'S'],
    ['12', 'false', 42],
    [0, true, false],
    ['2026-01-01', 'FALSE', new Date('2026-01-01')],
  ]);
  const strict = await collect(path, {
    schema: {
      n: { header: 'N', type: 'number' },
      b: { header: 'B', type: 'boolean' },
      s: { header: 'S', type: 'string' },
    },
  });
  expect(strict[0]).toMatchObject({
    ok: false,
    errors: [{ cell: 'A2' }, { cell: 'B2' }, { cell: 'C2' }],
  });
  const coercing = await collect(path, {
    schema: {
      n: { header: 'N', type: 'number', coerce: true },
      b: { header: 'B', type: 'boolean', coerce: true },
      s: { header: 'S', type: 'string', coerce: true },
    },
  });
  expect(coercing[0]).toMatchObject({
    ok: true,
    data: { n: 12, b: false, s: '42' },
  });
  expect(coercing[1]).toMatchObject({
    ok: true,
    data: { n: 0, b: true, s: 'false' },
  });
  expect(coercing[2].ok).toBe(false);
});
test('empty values, optional absent headers and whitespace retain explicit semantics', async () => {
  const path = await fixture([['Value'], [''], [null], ['  ']]);
  const required = await collect(path, {
    schema: { v: { header: 'Value', type: 'string', required: true } },
  });
  expect(required.slice(0, 2).map((r) => !r.ok && r.errors[0].code)).toEqual([
    'required',
    'required',
  ]);
  expect(required[2]).toMatchObject({ ok: true, data: { v: '  ' } });
  const optional = await collect(path, {
    schema: {
      v: { header: 'Value', type: 'string' },
      absent: { header: 'Other', type: 'boolean' },
    },
  });
  expect(optional[0]).toMatchObject({
    ok: true,
    data: { v: null, absent: null },
  });
});
for (const [rows, code] of [
  [[], 'missing_header_row'],
  [[['Other']], 'missing_header'],
  [[['Value', 'Value']], 'duplicate_header'],
  [[[' Value ']], 'missing_header'],
] as const) {
  test(`header failure ${code} cleans temporary files`, async () => {
    const path = await fixture(rows.map((row) => [...row]));
    const before = temps();
    let error: unknown;
    try {
      await collect(path, {
        schema: { v: { header: 'Value', type: 'string', required: true } },
      });
    } catch (caught) {
      error = caught;
    }
    expect(error).toBeInstanceOf(ExcelHeaderError);
    expect(error).toMatchObject({
      code,
      sheetIndex: 0,
      sheetName: 'Products',
      rowIndex: 0,
    });
    expect(temps()).toEqual(before);
  });
}
test('header-only sheets yield no objects; an empty later sheet still fails', async () => {
  const path = await fixture([['Tên sản phẩm', 'Giá']]);
  expect(await collect(path, { schema })).toEqual([]);
  const multiple = await workbook({
    worksheets: [
      {
        name: 'Good',
        rows: [{ cells: [{ value: 'Tên sản phẩm' }, { value: 'Giá' }] }],
      },
      { name: 'Empty', rows: [] },
    ],
  });
  await expect(collect(multiple, { schema })).rejects.toMatchObject({
    code: 'missing_header_row',
    sheetName: 'Empty',
  });
});
test('missing physical header, wide cell addresses and blank rows', async () => {
  const path = await fixture([['Value'], ['x']]);
  const zip = unzipSync(await Bun.file(path).bytes());
  zip['xl/worksheets/sheet1.xml'] = strToU8(
    '<worksheet><sheetData><row r="2"><c r="AA2" t="inlineStr"><is><t>Value</t></is></c></row><row r="4"><c r="AA4" t="inlineStr"><is><t>bad</t></is></c></row><row r="5"/></sheetData></worksheet>',
  );
  await Bun.write(path, zipSync(zip));
  await expect(
    collect(path, {
      schema: { v: { header: 'Value', type: 'number', required: true } },
    }),
  ).rejects.toMatchObject({ code: 'missing_header_row' });
  const rows = await collect(path, {
    headerRow: 1,
    schema: { v: { header: 'Value', type: 'number', required: true } },
  });
  expect(rows[0]).toMatchObject({
    rowIndex: 3,
    errors: [{ cell: 'AA4', code: 'invalid_type' }],
  });
  expect(rows[1]).toMatchObject({
    rowIndex: 4,
    errors: [{ cell: 'AA5', code: 'required' }],
  });
});
test('schema keys cannot mutate object prototypes; schema is snapshotted', async () => {
  const path = await fixture([['Value'], ['a'], ['b']]);
  const fields = JSON.parse(
    '{"__proto__":{"header":"Value","type":"string","required":true},"constructor":{"header":"Value","type":"string"}}',
  );
  const iterator = readExcelObjectsStream(path, { schema: fields });
  const first = (await iterator.next()).value;
  fields.__proto__.header = 'Changed';
  const second = (await iterator.next()).value;
  expect(first?.ok && Object.getPrototypeOf(first.data)).toBe(Object.prototype);
  expect(
    second?.ok &&
      Object.getOwnPropertyDescriptor(second.data, '__proto__')?.value,
  ).toBe('b');
  await iterator.return(undefined);
});
for (const options of [
  { schema: {} },
  { schema: [] },
  { schema: null },
  { schema: { v: { header: '', type: 'string' } } },
  { schema: { v: { header: 'V', type: 'date' } } },
  { schema: { v: { header: 'V', type: 'number', coerce: 1 } } },
  { schema, headerRow: null },
  { schema, headerRow: -1 },
  { schema, headerRow: 1.5 },
  { schema, headerRow: 1048576 },
]) {
  test(`invalid options fail before file access: ${JSON.stringify(options)}`, async () => {
    await expect(
      collect(
        '/nonexistent.xlsx',
        options as ExcelObjectReadOptions<ExcelObjectSchema>,
      ),
    ).rejects.not.toThrow('File not found');
  });
}
test('consumer break, abort and progress callback failure clean up', async () => {
  const path = await fixture([
    ['Tên sản phẩm', 'Giá'],
    ['a', 1],
    ['b', 2],
  ]);
  const before = temps();
  const events: ExcelReadProgress[] = [];
  for await (const _ of readExcelObjectsStream(path, {
    schema,
    onProgress: (e) => {
      events.push(e);
    },
  }))
    break;
  expect(temps()).toEqual(before);
  expect(events.some((e) => e.stage === 'completed')).toBe(false);
  const controller = new AbortController();
  const iterator = readExcelObjectsStream(path, {
    schema,
    signal: controller.signal,
  });
  await iterator.next();
  controller.abort();
  await expect(iterator.next()).rejects.toMatchObject({ name: 'AbortError' });
  expect(temps()).toEqual(before);
  await expect(
    collect(path, { schema, signal: controller.signal }),
  ).rejects.toMatchObject({ name: 'AbortError' });
  const reason = new Error('callback');
  await expect(
    collect(path, {
      schema,
      onProgress(e) {
        if (e.stage === 'reading') throw reason;
      },
    }),
  ).rejects.toBe(reason);
  expect(temps()).toEqual(before);
});
test('multiple XML batches stay ordered; progress counts source rows including headers', async () => {
  const path = await fixture([
    ['Tên sản phẩm', 'Giá'],
    ...Array.from({ length: 1800 }, (_, i) => [`${i}-${'x'.repeat(100)}`, i]),
  ]);
  let last: ExcelReadProgress | undefined;
  const rows = await collect(path, {
    schema,
    onProgress: async (e) => {
      await Promise.resolve();
      last = e;
    },
  });
  expect(rows).toHaveLength(1800);
  expect(rows[1799]).toMatchObject({ rowIndex: 1800, data: { price: 1799 } });
  expect(last).toMatchObject({ stage: 'completed', rowsRead: 1801 });
});

test('optional duplicate headers fail, unreferenced duplicates are ignored', async () => {
  const path = await fixture([
    ['V', 'V', 'Name'],
    [1, 2, 'item'],
  ]);
  await expect(
    collect(path, { schema: { v: { header: 'V', type: 'number' } } }),
  ).rejects.toMatchObject({ code: 'duplicate_header', header: 'V' });
  expect(
    await collect(path, {
      schema: { name: { header: 'Name', type: 'string' } },
    }),
  ).toMatchObject([{ ok: true, data: { name: 'item' } }]);
});

test('duplicate physical header rows throw and malformed XML propagates', async () => {
  const path = await fixture([['V']]);
  const zip = unzipSync(await Bun.file(path).bytes());
  for (const [xml, expected] of [
    [
      '<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>V</t></is></c></row><row r="1"/></sheetData></worksheet>',
      'duplicate_header',
    ],
    ['<worksheet><sheetData><row></wrong></sheetData></worksheet>', undefined],
  ] as const) {
    zip['xl/worksheets/sheet1.xml'] = strToU8(xml);
    await Bun.write(path, zipSync(zip));
    const before = temps();
    const work = collect(path, {
      schema: { v: { header: 'V', type: 'string' } },
    });
    if (expected) await expect(work).rejects.toMatchObject({ code: expected });
    else await expect(work).rejects.toThrow();
    expect(temps()).toEqual(before);
  }
});

test('schema fields defined through class getters retain validation and coercion', async () => {
  class PriceField {
    get header() {
      return 'Price';
    }
    get type() {
      return 'number' as const;
    }
    get required() {
      return true as const;
    }
    get coerce() {
      return true;
    }
  }
  const path = await fixture([['Price'], ['12'], ['']]);
  const rows = await collect(path, { schema: { price: new PriceField() } });
  expect(rows[0]).toMatchObject({ ok: true, data: { price: 12 } });
  expect(rows[1]).toMatchObject({
    ok: false,
    errors: [{ cell: 'A3', code: 'required' }],
  });
});

test('schema getters are read once and the validated snapshot is used throughout', async () => {
  let reads = 0;
  const field = {
    header: 'Value',
    get type() {
      reads++;
      return reads === 1 ? ('number' as const) : ('string' as const);
    },
    required: true,
  };
  const path = await fixture([['Value'], [12], [13]]);
  const rows = await collect(path, { schema: { value: field } });
  expect(reads).toBe(1);
  expect(rows.map((row) => row.ok && row.data.value)).toEqual([12, 13]);
});
