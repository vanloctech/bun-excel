import { afterAll, beforeAll, expect, spyOn, test } from 'bun:test';
import { mkdirSync, readdirSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { unzipSync, zipSync } from 'fflate';
import type { ExcelReadOptions, ExcelReadStreamRow } from '../src';
import { buildExcelBuffer, readExcel, readExcelStream } from '../src';

const TMP = './tests/.tmp-selective';
beforeAll(() => mkdirSync(TMP, { recursive: true }));
afterAll(() => rmSync(TMP, { recursive: true, force: true }));

test('buffered: selected sheets retain their own print areas by original workbook index', async () => {
  const path = `${TMP}/print-areas.xlsx`;
  const worksheets = ['First', 'Second', 'Third', 'Fourth'].map(
    (name, index) => ({
      name,
      rows: [{ cells: [{ value: name }] }],
      printArea:
        index === 2
          ? undefined
          : { startRow: 0, startCol: 0, endRow: index + 1, endCol: index + 1 },
    }),
  );
  await Bun.write(path, buildExcelBuffer({ worksheets }));
  for (const sheets of [
    ['Second'],
    [1],
    ['Third'],
    [2],
    ['Second', 'Fourth'],
    [1, 3],
    undefined,
  ]) {
    const result = await readExcel(path, { sheets });
    const expected = worksheets.filter(
      (sheet, index) =>
        !sheets ||
        (sheets as (string | number)[]).some(
          (selected) => selected === sheet.name || selected === index,
        ),
    );
    expect(
      result.worksheets.map(({ name, printArea }) => ({ name, printArea })),
    ).toEqual(expected.map(({ name, printArea }) => ({ name, printArea })));
  }
});

async function fixture(metadataLast = false) {
  const zip = unzipSync(
    await buildExcelBuffer({
      worksheets: [
        { name: 'First', rows: [{ cells: [{ value: 'one' }] }] },
        {
          name: 'Second',
          rows: [
            { cells: [{ value: 'two', style: { font: { bold: true } } }] },
          ],
        },
      ],
    }),
  );
  if (!metadataLast) return zip;
  const reordered: Record<string, Uint8Array> = {};
  for (const name of Object.keys(zip).reverse()) reordered[name] = zip[name];
  const workbook = reordered['xl/workbook.xml'];
  const rels = reordered['xl/_rels/workbook.xml.rels'];
  delete reordered['xl/workbook.xml'];
  delete reordered['xl/_rels/workbook.xml.rels'];
  reordered['xl/workbook.xml'] = workbook;
  reordered['xl/_rels/workbook.xml.rels'] = rels;
  return reordered;
}

// Corrupt the compressed payload, not the XML, so successful reads demonstrate
// that the excluded ZIP entry was never decompressed.
function corruptEntry(bytes: Uint8Array, filename: string): Uint8Array {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const decoder = new TextDecoder();
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
      bytes[payload] = 7; // DEFLATE reserved block type.
      return bytes;
    }
    offset = payload + size;
  }
  throw new Error(`Missing ZIP entry: ${filename}`);
}

async function readValues(
  path: string,
  streaming: boolean,
  options?: ExcelReadOptions,
) {
  if (!streaming) {
    const workbook = await readExcel(path, options);
    return workbook.worksheets.flatMap((sheet) =>
      sheet.rows.map((row) => row.cells[0]?.value),
    );
  }
  const values = [];
  for await (const entry of readExcelStream(Bun.file(path), options))
    values.push(entry.row.cells[0]?.value);
  return values;
}

for (const streaming of [false, true]) {
  const mode = streaming ? 'streaming' : 'buffered';
  test(`${mode}: includeStyles=false skips decompression of styles`, async () => {
    const path = `${TMP}/${mode}-styles.xlsx`;
    await Bun.write(
      path,
      corruptEntry(zipSync(await fixture()), 'xl/styles.xml'),
    );
    expect(await readValues(path, streaming, { includeStyles: false })).toEqual(
      ['one', 'two'],
    );
    await expect(readValues(path, streaming)).rejects.toThrow();
  });

  for (const metadataLast of [false, true]) {
    test(`${mode}: selected sheet skips other compressed worksheets (metadataLast=${metadataLast})`, async () => {
      const path = `${TMP}/${mode}-${metadataLast}.xlsx`;
      await Bun.write(
        path,
        corruptEntry(
          zipSync(await fixture(metadataLast)),
          'xl/worksheets/sheet1.xml',
        ),
      );
      for (const sheets of [['Second'], [1]] as const) {
        expect(
          await readValues(path, streaming, {
            sheets: [...sheets],
          } as ExcelReadOptions),
        ).toEqual(['two']);
      }
      expect(await readValues(path, streaming, { sheets: [] })).toEqual([]);
      expect(
        await readValues(path, streaming, { sheets: ['Missing'] }),
      ).toEqual([]);
      await expect(
        readValues(path, streaming, { sheets: ['First'] }),
      ).rejects.toThrow();
    });
  }

  test(`${mode}: skipped entries still receive ZIP path validation`, async () => {
    const zip = await fixture();
    zip['../ignored.xml'] = new TextEncoder().encode('<ignored/>');
    const path = `${TMP}/${mode}-unsafe.xlsx`;
    await Bun.write(path, zipSync(zip));
    await expect(
      readValues(path, streaming, { sheets: ['Second'], includeStyles: false }),
    ).rejects.toThrow('Malicious');
  });
}

for (const scenario of ['moved', 'removed', 'renamed', 'retargeted'] as const) {
  test(`streaming: rejects a ${scenario} selected sheet between read passes and cleans up`, async () => {
    const original = await fixture();
    const changed = { ...original };
    const decoder = new TextDecoder();
    const encoder = new TextEncoder();
    if (scenario === 'retargeted') {
      changed['xl/_rels/workbook.xml.rels'] = encoder.encode(
        decoder
          .decode(changed['xl/_rels/workbook.xml.rels'])
          .replace('worksheets/sheet2.xml', 'worksheets/replaced.xml'),
      );
      changed['xl/worksheets/replaced.xml'] =
        changed['xl/worksheets/sheet2.xml'];
      delete changed['xl/worksheets/sheet2.xml'];
    } else {
      const names = {
        moved: ['Second', 'First'],
        removed: ['First'],
        renamed: ['First', 'Renamed'],
      }[scenario];
      Object.assign(
        changed,
        unzipSync(
          buildExcelBuffer({
            worksheets: names.map((name) => ({
              name,
              rows: [{ cells: [{ value: name }] }],
            })),
          }),
        ),
      );
    }
    const originalPath = `${TMP}/${scenario}-original.xlsx`;
    const changedPath = `${TMP}/${scenario}-changed.xlsx`;
    await Bun.write(originalPath, zipSync(original));
    await Bun.write(changedPath, zipSync(changed));
    const source = Bun.file(originalPath);
    const stream = spyOn(source, 'stream')
      .mockImplementationOnce(() => Bun.file(originalPath).stream())
      .mockImplementationOnce(() => Bun.file(changedPath).stream());
    const temporaryFiles = () =>
      readdirSync(tmpdir())
        .filter((name) => name.startsWith('bun-excel-stream-'))
        .sort();
    const before = temporaryFiles();
    const rows: ExcelReadStreamRow[] = [];
    try {
      const consume = async () => {
        for await (const row of readExcelStream(source, {
          sheets: scenario === 'renamed' ? [1] : ['Second'],
        }))
          rows.push(row);
      };
      await expect(consume()).rejects.toThrow(
        'XLSX sheet selection changed between read passes',
      );
      expect(stream).toHaveBeenCalledTimes(2);
      expect(rows).toEqual([]);
      expect(temporaryFiles()).toEqual(before);
    } finally {
      stream.mockRestore();
    }
  });
}
