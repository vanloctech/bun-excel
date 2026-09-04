import { afterAll, beforeAll, expect, test } from 'bun:test';
import { mkdirSync, rmSync } from 'node:fs';
import { unzipSync, zipSync } from 'fflate';
import type { ExcelReadOptions } from '../src';
import { buildExcelBuffer, readExcel, readExcelStream } from '../src';

const TMP = './tests/.tmp-selective';
beforeAll(() => mkdirSync(TMP, { recursive: true }));
afterAll(() => rmSync(TMP, { recursive: true, force: true }));

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
