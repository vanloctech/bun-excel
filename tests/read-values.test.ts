import { afterAll, beforeAll, expect, test } from 'bun:test';
import { mkdirSync, readdirSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { unzipSync, zipSync } from 'fflate';
import {
  buildExcelBuffer,
  type ExcelReadValuesBatch,
  type ExcelReadValuesOptions,
  type FileSource,
  readExcelStream,
  readExcelValuesStream,
} from '../src';

const TMP = './tests/.tmp-read-values';
const PATH = `${TMP}/values.xlsx`;
const encoder = new TextEncoder();
let parts: Record<string, Uint8Array>;
const wrap = (rows: string) =>
  encoder.encode(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x="urn:test"><sheetData>${rows}</sheetData></worksheet>`,
  );
function temps() {
  return readdirSync(tmpdir())
    .filter((name) => name.startsWith('bun-excel-stream-'))
    .sort();
}
async function collect(
  source: FileSource = PATH,
  options?: ExcelReadValuesOptions,
) {
  const batches: ExcelReadValuesBatch[] = [];
  for await (const batch of readExcelValuesStream(source, options))
    batches.push(batch);
  return batches;
}
function flatten(batches: ExcelReadValuesBatch[]) {
  return batches.flatMap((batch) =>
    batch.rows.map((values, i) => ({
      sheetIndex: batch.sheetIndex,
      sheetName: batch.sheetName,
      rowIndex: batch.rowIndices[i],
      values,
    })),
  );
}
async function baseline(options?: ExcelReadValuesOptions) {
  const rows = [];
  for await (const item of readExcelStream(PATH, options))
    rows.push({
      sheetIndex: item.sheetIndex,
      sheetName: item.sheetName,
      rowIndex: item.rowIndex,
      values: Array.from(item.row.cells, (cell) => cell?.value ?? null),
    });
  return rows;
}
beforeAll(async () => {
  mkdirSync(TMP, { recursive: true });
  parts = unzipSync(
    buildExcelBuffer({
      worksheets: ['First', 'Second', 'Empty'].map((name) => ({
        name,
        rows: [
          {
            cells: [
              { value: 'shared string' },
              { value: 45000, style: { numberFormat: 'yyyy-mm-dd' } },
            ],
          },
        ],
      })),
    }),
  );
  for (let i = 1; i <= 2; i++)
    parts[`xl/worksheets/sheet${i}.xml`] = wrap(`
    <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="b"><v>1</v></c><c r="C1" t="b"><v>0</v></c><c r="D1" s="1"><v>45000</v></c><c r="E1"><v>-1.25e2</v></c><c r="F1" t="e"><v>#DIV/0!</v></c></row>
    <row r="3" hidden="1" ht="30"><c r="A3"><f>1+2</f><v>3</v></c><c r="C3"><f>1+2</f></c><c r="F3" t="str"><f>"text"</f><v>text</v></c><c r="G3" s="1"><f>TODAY()</f><v>45000</v></c></row>
    <row r="5"><x:c r="B5" t="inlineStr"><x:is><x:r><x:rPr><x:b/></x:rPr><x:t> rich </x:t></x:r><x:r><x:t><![CDATA[😀<&]]></x:t></x:r><x:r><x:rPr/></x:r></x:is></x:c><c r="D5" t="inlineStr"><is><t xml:space="preserve">  spaced  </t></is></c><c r="F5" t="inlineStr"><is><t/></is></c></row>
    <row r="7"><c r="C7" t="s"><v>999999</v></c><c r="D7"><v/></c><c r="F7"/><c r="A7"><v>1</v></c><c r="A7"><v>2</v></c></row>
    <row r="9"/>
    <row r="0"><c r="A0"><v>99</v></c></row>
    <row r="1048577"/><row/>
  `);
  parts['xl/worksheets/sheet3.xml'] = wrap('');
  await Bun.write(PATH, zipSync(parts));
});
afterAll(() => rmSync(TMP, { recursive: true, force: true }));

for (const options of [
  undefined,
  { batchSize: 1 },
  { batchSize: 2 },
  { batchSize: 4096 },
  { includeStyles: false },
  { columns: [5, 1, 5] },
  { columns: [] },
  { startRow: 2, endRow: 6, maxRows: 2 },
  { sheets: ['Second'] },
  { sheets: [1, 0], maxRows: 1 },
  { startRow: 99 },
  { sheets: ['Empty'] },
]) {
  test(`values match row reader: ${JSON.stringify(options)}`, async () => {
    const batches = await collect(PATH, options);
    expect(flatten(batches)).toEqual(await baseline(options));
    for (const batch of batches) {
      expect(batch.rows.length).toBeGreaterThan(0);
      expect(batch.rows.length).toBeLessThanOrEqual(options?.batchSize ?? 256);
      expect(batch.rowIndices.length).toBe(batch.rows.length);
    }
  });
}

test('returns cached formula values, Date, flattened text and null gaps without metadata', async () => {
  const rows = flatten(await collect());
  expect(rows[0].values).toEqual([
    'shared string',
    true,
    false,
    new Date('2023-03-15T00:00:00.000Z'),
    -125,
    '#DIV/0!',
  ]);
  expect(rows[1].values).toEqual([
    3,
    null,
    null,
    null,
    null,
    'text',
    new Date('2023-03-15T00:00:00.000Z'),
  ]);
  expect(rows[2].values).toEqual([
    null,
    ' rich 😀<&',
    null,
    '  spaced  ',
    null,
    '',
  ]);
  expect(rows[3].values).toEqual([2, null, '999999', '', null, null]);
  expect(rows[4].values).toEqual([]);
});

test('returned batches remain independent when retained or mutated', async () => {
  const iterator = readExcelValuesStream(PATH, { batchSize: 1 });
  try {
    const first = (await iterator.next()).value;
    if (!first) throw new Error('Missing batch');
    first.rows[0][0] = 'changed';
    first.rowIndices[0] = 123;
    const second = (await iterator.next()).value;
    if (!second) throw new Error('Missing batch');
    expect(second.rowIndices).toEqual([2]);
    expect(second.rows[0][0]).toBe(3);
    expect(first.rows[0][0]).toBe('changed');
  } finally {
    await iterator.return(undefined);
  }
});

test('bounds wide batches by cell slots', async () => {
  const path = `${TMP}/wide.xlsx`;
  await Bun.write(
    path,
    zipSync({
      ...parts,
      'xl/worksheets/sheet1.xml': wrap(
        Array.from(
          { length: 10 },
          (_, i) =>
            `<row r="${i + 1}"><c r="XFD${i + 1}"><v>${i}</v></c></row>`,
        ).join(''),
      ),
    }),
  );
  const batches = await collect(path, { sheets: [0], batchSize: 4096 });
  expect(batches.map((batch) => batch.rows.length)).toEqual([4, 4, 2]);
  for (const batch of batches)
    expect(
      batch.rows.reduce((n, row) => n + row.length, 0),
    ).toBeLessThanOrEqual(65536);
});

test('flushes large text at native batch boundaries', async () => {
  const path = `${TMP}/text.xlsx`;
  const text = 'x'.repeat(140000);
  await Bun.write(
    path,
    zipSync({
      ...parts,
      'xl/worksheets/sheet1.xml': wrap(
        Array.from(
          { length: 3 },
          (_, i) =>
            `<row r="${i + 1}"><c r="A${i + 1}" t="inlineStr"><is><t>${text}</t></is></c></row>`,
        ).join(''),
      ),
    }),
  );
  const batches = await collect(path, { sheets: [0], batchSize: 4096 });
  expect(batches.map((batch) => batch.rows.length)).toEqual([1, 1, 1]);
  expect(batches[2].rows[0][0]).toBe(text);
});

for (const batchSize of [
  0,
  -1,
  1.5,
  4097,
  Number.NaN,
  Number.POSITIVE_INFINITY,
]) {
  test(`rejects invalid batchSize ${batchSize} before I/O`, async () => {
    await expect(
      readExcelValuesStream('missing.xlsx', { batchSize }).next(),
    ).rejects.toThrow('batchSize');
  });
}

test('maxRows zero skips I/O, and inherited options remain validated', async () => {
  expect(await collect('missing.xlsx', { maxRows: 0 })).toEqual([]);
  await expect(
    readExcelValuesStream(PATH, { columns: [-1] }).next(),
  ).rejects.toThrow();
  await expect(
    readExcelValuesStream(PATH, { progressIntervalRows: 0 }).next(),
  ).rejects.toThrow();
});

test('cleans early return, abort and progress callback failure', async () => {
  const before = temps();
  for await (const _batch of readExcelValuesStream(PATH, { batchSize: 1 }))
    break;
  expect(temps()).toEqual(before);
  const controller = new AbortController();
  const iterator = readExcelValuesStream(PATH, {
    batchSize: 1,
    signal: controller.signal,
  });
  await iterator.next();
  controller.abort(new Error('stop values'));
  await expect(iterator.next()).rejects.toThrow('stop values');
  expect(temps()).toEqual(before);
  await expect(
    collect(PATH, {
      onProgress(progress) {
        if (progress.rowsRead) throw new Error('progress failed');
      },
    }),
  ).rejects.toThrow('progress failed');
  expect(temps()).toEqual(before);
});

test('progress counts rows and finishes after all sheets, not after each batch', async () => {
  const events: { rowsRead: number; stage: string }[] = [];
  await collect(PATH, {
    batchSize: 2,
    progressIntervalRows: 3,
    onProgress: async (event) => {
      events.push(event);
    },
  });
  expect(
    events.some((event) => event.stage === 'reading' && event.rowsRead === 4),
  ).toBe(true);
  expect(events.at(-1)).toMatchObject({ stage: 'completed', rowsRead: 10 });
  expect(events.filter((event) => event.stage === 'completed')).toHaveLength(1);
});

test('malformed XML and DTDs reject and remove temporary resources', async () => {
  const before = temps();
  for (const xml of [
    '<row r="1"><c></row>',
    '<!DOCTYPE worksheet [<!ENTITY x "bad">]><row r="1"/>',
  ]) {
    const path = `${TMP}/invalid.xlsx`;
    await Bun.write(
      path,
      zipSync({ ...parts, 'xl/worksheets/sheet1.xml': wrap(xml) }),
    );
    await expect(collect(path)).rejects.toThrow();
    expect(temps()).toEqual(before);
  }
});

test('supports Bun files with tiny input chunks and releases the source', async () => {
  const bytes = new Uint8Array(await Bun.file(PATH).arrayBuffer());
  const source = Bun.file(PATH);
  let closed = 0;
  source.stream = () => {
    let offset = 0;
    return new ReadableStream<Uint8Array<ArrayBuffer>>({
      pull(controller) {
        if (offset >= bytes.length) {
          closed++;
          controller.close();
          return;
        }
        controller.enqueue(bytes.subarray(offset, offset + 7));
        offset += 7;
      },
    });
  };
  expect(flatten(await collect(source))).toEqual(await baseline());
  expect(closed).toBe(1);
});

test('pre-aborted requests reject even when maxRows is zero', async () => {
  const controller = new AbortController();
  const reason = new Error('pre-aborted');
  controller.abort(reason);
  await expect(
    collect('missing.xlsx', { maxRows: 0, signal: controller.signal }),
  ).rejects.toBe(reason);
});

test('awaits progress callbacks and aborts an unresolved callback', async () => {
  const controller = new AbortController();
  const reason = new Error('cancel callback');
  const before = temps();
  let entered: (() => void) | undefined;
  const started = new Promise<void>((resolve) => {
    entered = resolve;
  });
  const iterator = readExcelValuesStream(PATH, {
    batchSize: 1,
    progressIntervalRows: 1,
    signal: controller.signal,
    onProgress(progress) {
      if (progress.stage === 'reading') {
        entered?.();
        return new Promise<void>(() => {});
      }
      return undefined;
    },
  });
  const pending = iterator.next();
  const outcome = pending.then(
    () => undefined,
    (error: unknown) => error,
  );
  await started;
  controller.abort(reason);
  expect(await outcome).toBe(reason);
  expect(temps()).toEqual(before);
});
