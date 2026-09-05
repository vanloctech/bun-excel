import { afterAll, beforeAll, expect, spyOn, test } from 'bun:test';
import { mkdirSync, rmSync } from 'node:fs';
import { unzipSync, zipSync } from 'fflate';
import {
  buildExcelBuffer,
  readExcel,
  readExcelInfo,
  readExcelStream,
  type WorksheetState,
} from '../src';

const TMP = './tests/.tmp-read-info';
const PATH = `${TMP}/workbook.xlsx`;
const encoder = new TextEncoder();
const decoder = new TextDecoder();
let parts: Record<string, Uint8Array>;

beforeAll(async () => {
  mkdirSync(TMP, { recursive: true });
  parts = unzipSync(
    buildExcelBuffer({
      creator: 'Metadata author',
      created: new Date('2026-09-05T00:00:00Z'),
      modified: new Date('2026-09-05T12:00:00Z'),
      views: { activeTab: 1, firstSheet: 0 },
      definedNames: [
        { name: 'Data', refersTo: "'Orders & 📦'!$A$1", localSheetId: 0 },
      ],
      worksheets: ['visible', 'hidden', 'veryHidden'].map((state, index) => ({
        name: ['Orders & 📦', 'Hidden', 'Internal'][index],
        state: state as WorksheetState,
        rows: [
          {
            cells: [
              {
                value: 'Shared',
                comment: { text: 'Note' },
                style: { font: { bold: true } },
              },
            ],
          },
        ],
        images: [
          {
            data: new Uint8Array([1, 2, 3]),
            format: 'png',
            range: { startRow: 1, startCol: 0, endRow: 2, endCol: 1 },
          },
        ],
        tables: [
          {
            name: `Table${index}`,
            range: { startRow: 0, startCol: 0, endRow: 1, endCol: 0 },
          },
        ],
      })),
    }),
  );
  await Bun.write(PATH, zipSync(parts));
});
afterAll(() => rmSync(TMP, { recursive: true, force: true }));

async function save(name: string, entries: Record<string, Uint8Array>) {
  const path = `${TMP}/${name}.xlsx`;
  await Bun.write(path, zipSync(entries));
  return path;
}

function corrupt(bytes: Uint8Array, filename: string): Uint8Array {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let offset = 0;
  while (view.getUint32(offset, true) === 0x04034b50) {
    const nameLength = view.getUint16(offset + 26, true);
    const extraLength = view.getUint16(offset + 28, true);
    const size = view.getUint32(offset + 18, true);
    const start = offset + 30;
    const payload = start + nameLength + extraLength;
    if (
      decoder.decode(bytes.subarray(start, start + nameLength)) === filename
    ) {
      bytes[payload] = 7;
      return bytes;
    }
    offset = payload + size;
  }
  throw new Error(`Missing ZIP entry: ${filename}`);
}

test('returns sheet order, visibility, byte size and the same workbook metadata as readExcel', async () => {
  const workbook = await readExcel(PATH);
  const info = await readExcelInfo(PATH);
  expect(info).toEqual({
    fileSize: (await Bun.file(PATH).stat()).size,
    sheets: [
      { index: 0, name: 'Orders & 📦', state: 'visible' },
      { index: 1, name: 'Hidden', state: 'hidden' },
      { index: 2, name: 'Internal', state: 'veryHidden' },
    ],
    creator: workbook.creator,
    created: workbook.created,
    modified: workbook.modified,
    definedNames: workbook.definedNames,
    views: workbook.views,
  });
  expect('worksheets' in info).toBe(false);
  expect('rows' in info.sheets[0]).toBe(false);
  const rows = [];
  for await (const row of readExcelStream(PATH, {
    sheets: [info.sheets[1].index],
    maxRows: 1,
  }))
    rows.push(row);
  expect(rows).toHaveLength(1);
  expect(rows[0].sheetName).toBe('Hidden');
});

test('streams Bun.file input without buffered reads or creating temporary files', async () => {
  const source = Bun.file(PATH);
  const bytes = spyOn(source, 'bytes');
  const arrayBuffer = spyOn(source, 'arrayBuffer');
  const stream = spyOn(source, 'stream');
  const fileFactory = spyOn(Bun, 'file');
  try {
    expect((await readExcelInfo(source)).sheets).toHaveLength(3);
    expect(bytes).not.toHaveBeenCalled();
    expect(arrayBuffer).not.toHaveBeenCalled();
    expect(stream).toHaveBeenCalledTimes(1);
    expect(fileFactory).not.toHaveBeenCalled();
  } finally {
    bytes.mockRestore();
    arrayBuffer.mockRestore();
    stream.mockRestore();
    fileFactory.mockRestore();
  }
});

for (const compress of [true, false]) {
  test(`works regardless of metadata order and ZIP compression (compress=${compress})`, async () => {
    const path = `${TMP}/order-${compress}.xlsx`;
    await Bun.write(
      path,
      zipSync(Object.fromEntries(Object.entries(parts).reverse()), {
        level: compress ? 6 : 0,
      }),
    );
    const expected = await readExcelInfo(PATH);
    expect(await readExcelInfo(path)).toEqual({
      ...expected,
      fileSize: (await Bun.file(path).stat()).size,
    });
  });
}

test('skips corrupt worksheet, styles, strings, image, comment and table payloads', async () => {
  const metadata = new Set([
    'xl/workbook.xml',
    'xl/_rels/workbook.xml.rels',
    'docProps/core.xml',
  ]);
  let bytes = zipSync(parts);
  for (const path of Object.keys(parts))
    if (!metadata.has(path)) bytes = corrupt(bytes, path);
  const path = `${TMP}/corrupt-resources.xlsx`;
  await Bun.write(path, bytes);
  const expected = await readExcelInfo(PATH);
  expect(await readExcelInfo(path)).toEqual({
    ...expected,
    fileSize: bytes.length,
  });
  await expect(readExcel(path)).rejects.toThrow();
});

for (const metadata of [
  'xl/workbook.xml',
  'xl/_rels/workbook.xml.rels',
  'docProps/core.xml',
]) {
  test(`rejects corrupted or malformed ${metadata}`, async () => {
    const path = `${TMP}/corrupt-${metadata.split('/').pop()}.xlsx`;
    await Bun.write(path, corrupt(zipSync(parts), metadata));
    await expect(readExcelInfo(path)).rejects.toThrow();
    const invalid = await save(`invalid-${metadata.split('/').pop()}`, {
      ...parts,
      [metadata]: encoder.encode('<broken>'),
    });
    await expect(readExcelInfo(invalid)).rejects.toThrow();
  });
}

for (const metadata of ['xl/workbook.xml', 'xl/_rels/workbook.xml.rels']) {
  test(`rejects missing ${metadata}`, async () => {
    const entries = { ...parts };
    delete entries[metadata];
    await expect(
      readExcelInfo(
        await save(`missing-${metadata.split('/').pop()}`, entries),
      ),
    ).rejects.toThrow('workbook metadata is missing');
  });
}

test('optional core properties and dates may be absent or invalid', async () => {
  const entries = { ...parts };
  delete entries['docProps/core.xml'];
  const absent = await readExcelInfo(await save('no-core', entries));
  expect(absent.creator).toBeUndefined();
  expect(absent.created).toBeUndefined();
  expect(absent.modified).toBeUndefined();
  entries['docProps/core.xml'] = encoder.encode(
    '<coreProperties><creator>Creator &amp; Co.</creator><created>invalid</created><modified>invalid</modified></coreProperties>',
  );
  const invalid = await readExcelInfo(await save('invalid-dates', entries));
  expect(invalid.creator).toBe('Creator & Co.');
  expect(invalid.created).toBeUndefined();
  expect(invalid.modified).toBeUndefined();
});

test('supports absent visibility, empty workbooks and prefixed workbook XML', async () => {
  const path = `${TMP}/defaults.xlsx`;
  await Bun.write(
    path,
    buildExcelBuffer({ worksheets: [{ name: 'Default', rows: [] }] }),
  );
  expect((await readExcelInfo(path)).sheets).toEqual([
    { index: 0, name: 'Default', state: 'visible' },
  ]);
  const empty = `${TMP}/empty.xlsx`;
  await Bun.write(empty, buildExcelBuffer({ worksheets: [] }));
  expect((await readExcelInfo(empty)).sheets).toEqual([]);
  const prefixed = await save('prefixed', {
    ...parts,
    'xl/workbook.xml': encoder.encode(
      '<s:workbook xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><s:sheets><s:sheet name="Prefixed"/></s:sheets></s:workbook>',
    ),
  });
  expect((await readExcelInfo(prefixed)).sheets).toEqual([
    { index: 0, name: 'Prefixed', state: 'visible' },
  ]);
});

for (const [name, xml] of [
  ['wrong-root', '<wrong/>'],
  ['missing-name', '<workbook><sheets><sheet/></sheets></workbook>'],
  [
    'invalid-state',
    '<workbook><sheets><sheet name="Sheet" state="invalid"/></sheets></workbook>',
  ],
  ['DTD', '<!DOCTYPE workbook><workbook/>'],
]) {
  test(`rejects invalid workbook metadata: ${name}`, async () => {
    await expect(
      readExcelInfo(
        await save(name, { ...parts, 'xl/workbook.xml': encoder.encode(xml) }),
      ),
    ).rejects.toThrow();
  });
}

test('source errors propagate and the stream is released', async () => {
  const source = Bun.file(PATH);
  const bytes = new Uint8Array(zipSync(parts));
  let input: ReadableStream<Uint8Array<ArrayBuffer>> | undefined;
  const stream = spyOn(source, 'stream').mockImplementation(() => {
    let sent = false;
    input = new ReadableStream<Uint8Array<ArrayBuffer>>({
      pull(controller) {
        if (sent) {
          controller.error(new Error('source failure'));
          return;
        }
        controller.enqueue(bytes.subarray(0, 100));
        sent = true;
      },
    });
    return input;
  });
  try {
    await expect(readExcelInfo(source)).rejects.toThrow('source failure');
    expect(input?.locked).toBe(false);
  } finally {
    stream.mockRestore();
  }
});

test('missing, empty and oversized sources fail before metadata processing', async () => {
  await expect(readExcelInfo(`${TMP}/missing.xlsx`)).rejects.toThrow(
    'File not found',
  );
  const empty = `${TMP}/zero.xlsx`;
  await Bun.write(empty, new Uint8Array());
  await expect(readExcelInfo(empty)).rejects.toThrow();
  const source = Bun.file(PATH);
  const actual = await source.stat();
  const stat = spyOn(source, 'stat').mockResolvedValue({
    ...actual,
    size: 201 * 1024 * 1024,
  });
  const stream = spyOn(source, 'stream');
  try {
    await expect(readExcelInfo(source)).rejects.toThrow('File too large');
    expect(stream).not.toHaveBeenCalled();
  } finally {
    stat.mockRestore();
    stream.mockRestore();
  }
});

test('scans beyond early metadata and validates excluded ZIP paths and sizes', async () => {
  const unsafe = await save('unsafe', {
    ...parts,
    '../ignored.bin': new Uint8Array([1]),
  });
  await expect(readExcelInfo(unsafe)).rejects.toThrow('Malicious');
  const bytes = zipSync({ 'ignored.bin': new Uint8Array([1]), ...parts });
  new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).setUint32(
    22,
    1024 * 1024 * 1024 + 1,
    true,
  );
  const oversized = `${TMP}/oversized-entry.xlsx`;
  await Bun.write(oversized, bytes);
  await expect(readExcelInfo(oversized)).rejects.toThrow(
    'Declared decompressed size exceeds limit',
  );
});

test('rejects a truncated metadata payload', async () => {
  const bytes = zipSync({ 'xl/workbook.xml': parts['xl/workbook.xml'] });
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const payload = 30 + view.getUint16(26, true) + view.getUint16(28, true);
  const path = `${TMP}/truncated.xlsx`;
  await Bun.write(
    path,
    bytes.subarray(0, payload + Math.floor(view.getUint32(18, true) / 2)),
  );
  await expect(readExcelInfo(path)).rejects.toThrow();
});

test('reads metadata from small source chunks and releases the reader on success', async () => {
  const source = Bun.file(PATH);
  const bytes = new Uint8Array(zipSync(parts));
  let input: ReadableStream<Uint8Array<ArrayBuffer>> | undefined;
  const stream = spyOn(source, 'stream').mockImplementation(() => {
    let offset = 0;
    input = new ReadableStream<Uint8Array<ArrayBuffer>>({
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
    return input;
  });
  try {
    const info = await readExcelInfo(source);
    expect(info.sheets).toHaveLength(3);
    expect(info.creator).toBe('Metadata author');
    expect(input?.locked).toBe(false);
  } finally {
    stream.mockRestore();
  }
});

test('rejects invalid relationship roots and empty sheet names', async () => {
  await expect(
    readExcelInfo(
      await save('wrong-rels-root', {
        ...parts,
        'xl/_rels/workbook.xml.rels': encoder.encode('<wrong/>'),
      }),
    ),
  ).rejects.toThrow('invalid workbook metadata root');
  await expect(
    readExcelInfo(
      await save('empty-sheet-name', {
        ...parts,
        'xl/workbook.xml': encoder.encode(
          '<workbook><sheets><sheet name=""/></sheets></workbook>',
        ),
      }),
    ),
  ).rejects.toThrow('invalid sheet name');
});

test('the ZIP entry count limit also applies to excluded entries', async () => {
  const entries = { ...parts };
  for (let index = 0; index < 10_000; index++)
    entries[`ignored/${index}.bin`] = new Uint8Array();
  await expect(
    readExcelInfo(await save('too-many-entries', entries)),
  ).rejects.toThrow('ZIP has too many entries');
});
