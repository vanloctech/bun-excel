import { afterAll, beforeAll, expect, test } from 'bun:test';
import { mkdirSync, rmSync } from 'node:fs';
import { unzipSync, zipSync } from 'fflate';
import {
  buildExcelBuffer,
  type ExcelReadOptions,
  readExcel,
  type Workbook,
} from '../src';

const TMP = './tests/.tmp-read-feature-options';
const PATH = `${TMP}/features.xlsx`;
const encoder = new TextEncoder();
const decoder = new TextDecoder();
let parts: Record<string, Uint8Array>;
let baseline: Workbook;
let media: string[];

beforeAll(async () => {
  mkdirSync(TMP, { recursive: true });
  parts = unzipSync(
    buildExcelBuffer({
      creator: 'Feature options',
      created: new Date('2026-09-05T00:00:00Z'),
      worksheets: ['First', 'Second'].map((name, index) => ({
        name,
        rows: [
          {
            cells: [
              {
                value: 'Name',
                comment: { text: `Comment ${index}`, author: 'Author' },
              },
              { value: 'Date' },
              { value: 'Total' },
            ],
          },
          {
            cells: [
              { value: name, hyperlink: { target: 'https://example.com' } },
              {
                value: new Date('2026-09-05T00:00:00Z'),
                style: { numberFormat: 'yyyy-mm-dd' },
              },
              { formula: 'SUM(20,22)', value: 42 },
            ],
          },
        ],
        images: [
          {
            data: new Uint8Array([index + 1, 2, 3]),
            format: 'png',
            range: { startRow: 3, startCol: 0, endRow: 4, endCol: 1 },
            name: `Logo ${index}`,
          },
        ],
        tables: [
          {
            name: `Table${index}`,
            range: { startRow: 0, startCol: 0, endRow: 1, endCol: 2 },
            headerRow: true,
            style: { name: 'TableStyleMedium2' },
          },
        ],
        autoFilter: { startRow: 0, startCol: 0, endRow: 1, endCol: 2 },
        printArea: { startRow: 0, startCol: 0, endRow: 4, endCol: 2 },
      })),
    }),
  );
  media = Object.keys(parts).filter((path) => path.startsWith('xl/media/'));
  await Bun.write(PATH, zipSync(parts));
  baseline = await readExcel(PATH);
});
afterAll(() => rmSync(TMP, { recursive: true, force: true }));

function withoutFeatures(
  workbook: Workbook,
  options: ExcelReadOptions,
): Workbook {
  const expected = structuredClone(workbook);
  for (const sheet of expected.worksheets) {
    if (options.includeImages === false) delete sheet.images;
    if (options.includeTables === false) delete sheet.tables;
    if (options.includeComments === false)
      for (const row of sheet.rows)
        for (const cell of row.cells) if (cell) delete cell.comment;
  }
  return expected;
}

// Reserved DEFLATE block type: successful reads prove a part was not inflated.
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
      expect(view.getUint16(offset + 8, true)).toBe(8);
      bytes[payload] = 7;
      return bytes;
    }
    offset = payload + size;
  }
  throw new Error(`Missing ZIP entry: ${filename}`);
}

const featureOptions = [false, true].flatMap((includeImages) =>
  [false, true].flatMap((includeComments) =>
    [false, true].map((includeTables) => ({
      includeImages,
      includeComments,
      includeTables,
    })),
  ),
);
for (const options of featureOptions) {
  for (const sheets of [undefined, ['Second'], [1]]) {
    test(`resource flags ${JSON.stringify(options)}, sheets=${JSON.stringify(sheets)}`, async () => {
      const result = await readExcel(Bun.file(PATH), { ...options, sheets });
      const expected = withoutFeatures(baseline, options);
      if (sheets) expected.worksheets = [expected.worksheets[1]];
      expect(result).toEqual(expected);
    });
  }
}

test('omitting flags preserves the default behavior', async () => {
  expect(
    await readExcel(PATH, {
      includeImages: undefined,
      includeComments: undefined,
      includeTables: undefined,
    }),
  ).toEqual(baseline);
  expect(baseline.worksheets[0].images).toHaveLength(1);
  expect(baseline.worksheets[0].tables).toHaveLength(1);
  expect(baseline.worksheets[0].rows[0].cells[0].comment?.text).toBe(
    'Comment 0',
  );
});

for (const kind of [
  'includeImages',
  'includeComments',
  'includeTables',
] as const) {
  test(`${kind}=false skips corrupt compressed resources, including renamed parts`, async () => {
    const renamed = { ...parts };
    const relPath = 'xl/worksheets/_rels/sheet1.xml.rels';
    renamed['xl/notes/custom.xml'] = renamed['xl/comments1.xml'];
    delete renamed['xl/comments1.xml'];
    renamed['xl/customTables/custom.xml'] = renamed['xl/tables/table1.xml'];
    delete renamed['xl/tables/table1.xml'];
    renamed[relPath] = encoder.encode(
      decoder
        .decode(renamed[relPath])
        .replace('../comments1.xml', '../notes/custom.xml')
        .replace('../tables/table1.xml', '../customTables/custom.xml'),
    );
    const targets =
      kind === 'includeImages'
        ? [
            'xl/drawings/drawing1.xml',
            'xl/drawings/_rels/drawing1.xml.rels',
            media[0],
          ]
        : [
            kind === 'includeComments'
              ? 'xl/notes/custom.xml'
              : 'xl/customTables/custom.xml',
          ];
    for (const [index, target] of targets.entries()) {
      const path = `${TMP}/corrupt-${kind}-${index}.xlsx`;
      await Bun.write(path, corrupt(zipSync(renamed), target));
      for (const sheets of [undefined, ['First'], [0]]) {
        const options = { [kind]: false, sheets };
        const result = await readExcel(path, options);
        const expected = withoutFeatures(baseline, options);
        if (sheets) expected.worksheets = [expected.worksheets[0]];
        expect(result).toEqual(expected);
      }
      await expect(readExcel(path)).rejects.toThrow();
      await expect(
        readExcel(path, { [kind]: true, sheets: ['First'] }),
      ).rejects.toThrow();
    }
  });
}

test('disabled resources with malformed XML are not parsed', async () => {
  const path = `${TMP}/malformed.xml.xlsx`;
  const broken = { ...parts };
  for (const key of [
    'xl/comments1.xml',
    'xl/tables/table1.xml',
    'xl/drawings/drawing1.xml',
    'xl/drawings/_rels/drawing1.xml.rels',
  ])
    broken[key] = encoder.encode('<broken>');
  await Bun.write(path, zipSync(broken));
  const options = {
    includeImages: false,
    includeComments: false,
    includeTables: false,
  };
  expect(await readExcel(path, options)).toEqual(
    withoutFeatures(baseline, options),
  );
  await expect(readExcel(path)).rejects.toThrow();
});

test('feature flags and disabled styles leave values, formulas and hyperlinks intact', async () => {
  const result = await readExcel(PATH, {
    includeStyles: false,
    includeImages: false,
    includeComments: false,
    includeTables: false,
  });
  for (const sheet of result.worksheets) {
    expect(typeof sheet.rows[1].cells[1].value).toBe('number');
    expect(sheet.rows[1].cells[1].style).toBeUndefined();
    expect(sheet.rows[1].cells[2]).toMatchObject({
      formula: 'SUM(20,22)',
      value: 42,
    });
    expect(sheet.rows[1].cells[0].hyperlink?.target).toBe(
      'https://example.com',
    );
    expect(sheet.autoFilter).toEqual(baseline.worksheets[0].autoFilter);
    expect(sheet.printArea).toEqual(baseline.worksheets[0].printArea);
  }
});

test('resources shared between worksheets are retained when enabled', async () => {
  const shared = { ...parts };
  shared['xl/drawings/_rels/drawing2.xml.rels'] = encoder.encode(
    decoder
      .decode(shared['xl/drawings/_rels/drawing2.xml.rels'])
      .replace(media[1].slice(3), media[0].slice(3)),
  );
  shared['xl/worksheets/_rels/sheet2.xml.rels'] = encoder.encode(
    decoder
      .decode(shared['xl/worksheets/_rels/sheet2.xml.rels'])
      .replace('../comments2.xml', '../comments1.xml')
      .replace('../tables/table2.xml', '../tables/table1.xml'),
  );
  const path = `${TMP}/shared.xlsx`;
  await Bun.write(path, zipSync(shared));
  const full = await readExcel(path);
  expect(full.worksheets[1].images?.[0].data).toEqual(
    full.worksheets[0].images?.[0].data,
  );
  expect(full.worksheets[1].rows[0].cells[0].comment).toEqual(
    full.worksheets[0].rows[0].cells[0].comment,
  );
  expect(full.worksheets[1].tables).toEqual(full.worksheets[0].tables);
  for (const options of [
    { includeComments: false },
    { includeImages: false },
    { includeTables: false },
  ]) {
    const expected = withoutFeatures(full, options);
    expect(await readExcel(path, options)).toEqual(expected);
    expected.worksheets = [expected.worksheets[1]];
    expect(await readExcel(path, { ...options, sheets: ['Second'] })).toEqual(
      expected,
    );
  }
});

test('disabled comments do not create cells outside worksheet data', async () => {
  const path = `${TMP}/comment-only-cell.xlsx`;
  await Bun.write(
    path,
    zipSync({
      ...parts,
      'xl/comments1.xml': encoder.encode(
        decoder
          .decode(parts['xl/comments1.xml'])
          .replace('ref="A1"', 'ref="D20"'),
      ),
    }),
  );
  expect(
    (await readExcel(path)).worksheets[0].rows[19].cells[3].comment,
  ).toBeDefined();
  const result = await readExcel(path, { includeComments: false });
  expect(result.worksheets[0].rows).toHaveLength(2);
});

test('disabled features work on sheets without relationship parts and on empty selections', async () => {
  const path = `${TMP}/plain.xlsx`;
  await Bun.write(
    path,
    buildExcelBuffer({
      worksheets: [
        { name: 'Plain', rows: [{ cells: [{ value: 'data' }] }] },
        { name: 'Empty', rows: [] },
      ],
    }),
  );
  const flags = {
    includeImages: false,
    includeComments: false,
    includeTables: false,
  };
  expect(await readExcel(path, flags)).toEqual(await readExcel(path));
  expect((await readExcel(PATH, { ...flags, sheets: [] })).worksheets).toEqual(
    [],
  );
  expect(
    (await readExcel(PATH, { ...flags, sheets: ['Missing'] })).worksheets,
  ).toEqual([]);
});

test('ZIP safety checks still cover excluded entries', async () => {
  const path = `${TMP}/unsafe.xlsx`;
  await Bun.write(
    path,
    zipSync({ ...parts, '../ignored-image.png': new Uint8Array([1]) }),
  );
  await expect(readExcel(path, { includeImages: false })).rejects.toThrow(
    'Malicious',
  );
});

test('worksheet relationships and enabled resources still report errors', async () => {
  const path = `${TMP}/corrupt-rels.xlsx`;
  await Bun.write(
    path,
    corrupt(zipSync(parts), 'xl/worksheets/_rels/sheet1.xml.rels'),
  );
  await expect(
    readExcel(path, {
      includeImages: false,
      includeComments: false,
      includeTables: false,
    }),
  ).rejects.toThrow();
  const enabled = `${TMP}/enabled-table.xlsx`;
  await Bun.write(enabled, corrupt(zipSync(parts), 'xl/tables/table1.xml'));
  await expect(
    readExcel(enabled, { includeImages: false, includeComments: false }),
  ).rejects.toThrow();
});
