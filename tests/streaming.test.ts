import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { mkdirSync, rmSync } from 'node:fs';
import {
  buildExcelResponse,
  createChunkedExcelStream,
  createCSVStream,
  createExcelStream,
  createMultiSheetExcelStream,
  exportExcelRows,
  exportExcelRowsToResponse,
  exportMultiSheetExcel,
  readCSV,
  readExcel,
  readExcelStream,
  writeExcel,
} from '../src';
import { ManagedFileSink } from '../src/excel/file-sink';

const TMP = './tests/.tmp';
const PNG_1X1 = Uint8Array.from(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADElEQVR42mP8/5+hHgAHggJ/PF6edAAAAABJRU5ErkJggg==',
    'base64',
  ),
);

beforeAll(() => {
  mkdirSync(TMP, { recursive: true });
});

afterAll(() => {
  rmSync(TMP, { recursive: true, force: true });
});

describe('CSV Stream Writer', () => {
  test('writes CSV via stream', async () => {
    const path = `${TMP}/csv-stream.csv`;
    const stream = createCSVStream(path, {
      headers: ['ID', 'Name'],
      includeHeader: true,
    });

    stream.writeRow([1, 'Alice']);
    stream.writeRow([2, 'Bob']);
    await stream.end();

    const wb = await readCSV(path);
    const rows = wb.worksheets[0].rows;
    expect(rows).toHaveLength(3); // header + 2 data
    expect(rows[0].cells[0].value).toBe('ID');
    expect(rows[1].cells[1].value).toBe('Alice');
  });

  test('writes CSV via stream to Bun.file target', async () => {
    const path = `${TMP}/csv-stream-bun-file.csv`;
    const stream = createCSVStream(Bun.file(path), {
      headers: ['ID', 'Name'],
      includeHeader: true,
    });

    stream.writeRow([1, 'Alice']);
    stream.writeRow([2, 'Bob']);
    await stream.end();

    const wb = await readCSV(path);
    expect(wb.worksheets[0].rows[2].cells[1].value).toBe('Bob');
  });

  test('handles large CSV stream', async () => {
    const path = `${TMP}/csv-large-stream.csv`;
    const stream = createCSVStream(path, {
      headers: ['ID', 'Value'],
      includeHeader: true,
    });

    const count = 5000;
    for (let i = 0; i < count; i++) {
      stream.writeRow([i, Math.random()]);
    }
    await stream.end();

    const wb = await readCSV(path);
    expect(wb.worksheets[0].rows).toHaveLength(count + 1);
  });
});

describe('Excel Stream Writer', () => {
  test('writes Excel via stream', async () => {
    const path = `${TMP}/excel-stream.xlsx`;
    const stream = createExcelStream(path, {
      sheetName: 'Data',
      columns: [{ width: 10 }, { width: 20 }],
    });

    stream.writeRow({
      cells: [
        { value: 'ID', style: { font: { bold: true } } },
        { value: 'Name', style: { font: { bold: true } } },
      ],
    });

    for (let i = 0; i < 100; i++) {
      stream.writeRow([i + 1, `Item_${i + 1}`]);
    }

    await stream.end();

    const wb = await readExcel(path);
    expect(wb.worksheets).toHaveLength(1);
    expect(wb.worksheets[0].rows).toHaveLength(101);
    expect(wb.worksheets[0].rows[0].cells[0].value).toBe('ID');
    expect(wb.worksheets[0].rows[1].cells[0].value).toBe(1);
  });

  test('writes Excel via stream to Bun.file target', async () => {
    const path = `${TMP}/excel-stream-bun-file.xlsx`;
    const stream = createExcelStream(Bun.file(path), {
      sheetName: 'Data',
    });

    stream.writeRow(['ID', 'Name']);
    stream.writeRow([1, 'Alice']);
    await stream.end();

    const wb = await readExcel(path);
    expect(wb.worksheets[0].rows[1].cells[1].value).toBe('Alice');
  });

  test('writes stream with freeze pane', async () => {
    const path = `${TMP}/stream-freeze.xlsx`;
    const stream = createExcelStream(path, {
      sheetName: 'Frozen',
      freezePane: { row: 1, col: 0 },
    });

    stream.writeRow(['Header']);
    stream.writeRow(['Data']);
    await stream.end();

    const wb = await readExcel(path);
    expect(wb.worksheets[0].freezePane).toEqual({ row: 1, col: 0 });
  });

  test('writes stream with split pane and workbook properties', async () => {
    const path = `${TMP}/stream-split.xlsx`;
    const created = new Date('2026-03-01T00:00:00.000Z');
    const modified = new Date('2026-03-01T05:00:00.000Z');
    const stream = createExcelStream(path, {
      sheetName: 'Split',
      creator: 'stream-writer',
      created,
      modified,
      splitPane: {
        x: 900,
        y: 1600,
        topLeftCell: { row: 1, col: 1 },
      },
    });

    stream.writeRow(['Header']);
    stream.writeRow(['Data']);
    await stream.end();

    const wb = await readExcel(path);
    expect(wb.creator).toBe('stream-writer');
    expect(wb.created?.toISOString()).toBe(created.toISOString());
    expect(wb.modified?.toISOString()).toBe(modified.toISOString());
    expect(wb.worksheets[0].splitPane).toEqual({
      x: 900,
      y: 1600,
      topLeftCell: { row: 1, col: 1 },
    });
  });

  test('writes stream with merge cells', async () => {
    const path = `${TMP}/stream-merge.xlsx`;
    const stream = createExcelStream(path, {
      sheetName: 'Merged',
      mergeCells: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 2 }],
    });

    stream.writeRow(['Title', null, null]);
    stream.writeRow(['A', 'B', 'C']);
    await stream.end();

    const wb = await readExcel(path);
    expect(wb.worksheets[0].mergeCells).toHaveLength(1);
  });

  test('writes stream with auto filter', async () => {
    const path = `${TMP}/stream-autofilter.xlsx`;
    const stream = createExcelStream(path, {
      sheetName: 'Filtered',
      autoFilter: { startRow: 0, startCol: 0, endRow: 10, endCol: 1 },
    });

    stream.writeRow(['Name', 'Score']);
    stream.writeRow(['Alice', 95]);
    await stream.end();

    const wb = await readExcel(path);
    expect(wb.worksheets[0].autoFilter).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: 10,
      endCol: 1,
    });
  });

  test('writes stream with data validation', async () => {
    const path = `${TMP}/stream-validation.xlsx`;
    const stream = createExcelStream(path, {
      sheetName: 'Validated',
      dataValidations: [
        {
          range: { startRow: 1, startCol: 0, endRow: 20, endCol: 0 },
          type: 'list',
          formula1: ['Low', 'Medium', 'High'],
        },
      ],
    });

    stream.writeRow(['Priority']);
    stream.writeRow([null]);
    await stream.end();

    const wb = await readExcel(path);
    expect(wb.worksheets[0].dataValidations?.[0].formula1).toEqual([
      'Low',
      'Medium',
      'High',
    ]);
  });

  test('writes stream with conditional formatting', async () => {
    const path = `${TMP}/stream-conditional.xlsx`;
    const stream = createExcelStream(path, {
      sheetName: 'Conditional',
      conditionalFormattings: [
        {
          range: { startRow: 1, startCol: 0, endRow: 100, endCol: 0 },
          rules: [
            {
              type: 'expression',
              formula: 'A2>10',
              style: {
                fill: {
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: 'C6EFCE',
                },
              },
            },
          ],
        },
      ],
    });

    stream.writeRow(['Value']);
    stream.writeRow([12]);
    await stream.end();

    const rule = readConditionalRule(await readExcel(path));
    expect(rule?.type).toBe('expression');
    if (rule?.type === 'expression') {
      expect(rule.formula).toBe('A2>10');
      expect(rule.style?.fill?.fgColor).toBe('C6EFCE');
    }
  });

  test('writes stream with rich text and worksheet settings', async () => {
    const path = `${TMP}/stream-rich-settings.xlsx`;
    const stream = createExcelStream(path, {
      sheetName: 'Report',
      state: 'hidden',
      pageSetup: { orientation: 'landscape' },
      headerFooter: {
        oddHeader: { center: 'Quarterly Report' },
      },
      printArea: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
    });

    stream.writeRow({
      cells: [
        {
          value: 'Rich text',
          richText: [
            { text: 'Rich ', font: { bold: true } },
            { text: 'text', font: { italic: true, color: 'FF0000' } },
          ],
        },
        { value: 'B1' },
      ],
    });
    stream.writeRow(['A2', 'B2']);
    await stream.end();

    const wb = await readExcel(path);
    expect(wb.worksheets[0].state).toBe('hidden');
    expect(wb.worksheets[0].pageSetup?.orientation).toBe('landscape');
    expect(wb.worksheets[0].headerFooter?.oddHeader?.center).toBe(
      'Quarterly Report',
    );
    expect(wb.worksheets[0].printArea).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: 1,
      endCol: 1,
    });
    expect(wb.worksheets[0].rows[0].cells[0].richText).toEqual([
      { text: 'Rich ', font: { bold: true } },
      { text: 'text', font: { italic: true, color: 'FF0000' } },
    ]);
  });

  test('writes stream with comments, images, and tables', async () => {
    const path = `${TMP}/stream-comments-images-tables.xlsx`;
    const stream = createExcelStream(path, {
      sheetName: 'Assets',
      images: [
        {
          data: PNG_1X1,
          format: 'png',
          range: { startRow: 4, startCol: 0, endRow: 5, endCol: 1 },
          name: 'Logo',
        },
      ],
      tables: [
        {
          name: 'AssetTable',
          range: { startRow: 0, startCol: 0, endRow: 2, endCol: 2 },
          headerRow: true,
          totalsRow: false,
          columns: [{ name: 'Name' }, { name: 'Value' }, { name: 'Link' }],
        },
      ],
    });

    stream.writeRow({
      cells: [
        { value: 'Name', comment: { text: 'Header comment', author: 'Loc' } },
        { value: 'Value' },
        { value: 'Link' },
      ],
    });
    stream.writeRow({
      cells: [
        { value: 'A' },
        { value: 1 },
        { value: 'Docs', hyperlink: { target: 'https://bun.sh' } },
      ],
    });
    stream.writeRow({
      cells: [
        { value: 'B', comment: { text: 'Body comment' } },
        { value: 2 },
        { value: 'Guide' },
      ],
    });
    await stream.end();

    const sheet = (await readExcel(path)).worksheets[0];
    expect(sheet.rows[0].cells[0].comment?.author).toBe('Loc');
    expect(sheet.rows[2].cells[0].comment?.text).toBe('Body comment');
    expect(sheet.images?.[0].range).toEqual({
      startRow: 4,
      startCol: 0,
      endRow: 5,
      endCol: 1,
    });
    expect(sheet.tables?.[0].name).toBe('AssetTable');
    expect(sheet.tables?.[0].columns?.map((column) => column.name)).toEqual([
      'Name',
      'Value',
      'Link',
    ]);
  });

  test('handles large stream (10K rows)', async () => {
    const path = `${TMP}/stream-large.xlsx`;
    const stream = createExcelStream(path, { sheetName: 'Large' });

    for (let i = 0; i < 10000; i++) {
      stream.writeRow([i, `Row_${i}`, Math.random()]);
    }
    await stream.end();

    const wb = await readExcel(path);
    expect(wb.worksheets[0].rows).toHaveLength(10000);
  });
});

describe('Multi-Sheet Stream Writer', () => {
  test('writes multiple sheets via stream', async () => {
    const path = `${TMP}/multi-stream.xlsx`;
    const stream = createMultiSheetExcelStream(path);

    stream.addSheet('Sheet1', { columns: [{ width: 15 }] });
    stream.writeRow(['First sheet']);
    stream.writeRow(['Data 1']);

    stream.addSheet('Sheet2', { columns: [{ width: 15 }] });
    stream.writeRow(['Second sheet']);

    await stream.end();

    const wb = await readExcel(path);
    expect(wb.worksheets).toHaveLength(2);
    expect(wb.worksheets[0].name).toBe('Sheet1');
    expect(wb.worksheets[1].name).toBe('Sheet2');
    expect(wb.worksheets[0].rows).toHaveLength(2);
    expect(wb.worksheets[1].rows).toHaveLength(1);
  });

  test('writes multiple sheets via stream to Bun.file target', async () => {
    const path = `${TMP}/multi-stream-bun-file.xlsx`;
    const stream = createMultiSheetExcelStream(Bun.file(path));

    stream.addSheet('Sheet1');
    stream.writeRow(['First']);
    stream.addSheet('Sheet2');
    stream.writeRow(['Second']);
    await stream.end();

    const wb = await readExcel(path);
    expect(wb.worksheets).toHaveLength(2);
    expect(wb.worksheets[1].rows[0].cells[0].value).toBe('Second');
  });

  test('writes multiple sheets with per-sheet config and hyperlinks', async () => {
    const path = `${TMP}/multi-stream-configured.xlsx`;
    const stream = createMultiSheetExcelStream(path, {
      creator: 'multi-stream',
    });

    stream.addSheet('Sheet1', {
      freezePane: { row: 1, col: 0 },
      autoFilter: { startRow: 0, startCol: 0, endRow: 10, endCol: 1 },
    });
    stream.writeRow({
      cells: [
        {
          value: 'Docs',
          hyperlink: {
            target: 'https://bun.sh',
            tooltip: 'Bun docs',
          },
        },
        { value: 'Status' },
      ],
    });
    stream.writeRow(['Bun', 'Ready']);

    stream.addSheet('Sheet2', {
      splitPane: {
        x: 900,
        y: 1600,
        topLeftCell: { row: 1, col: 1 },
      },
      conditionalFormattings: [
        {
          range: { startRow: 1, startCol: 0, endRow: 10, endCol: 0 },
          rules: [
            {
              type: 'expression',
              formula: 'A2>10',
              style: {
                fill: {
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: 'C6EFCE',
                },
              },
            },
          ],
        },
      ],
    });
    stream.writeRow(['Value']);
    stream.writeRow([12]);

    await stream.end();

    const wb = await readExcel(path);
    expect(wb.creator).toBe('multi-stream');
    expect(wb.worksheets).toHaveLength(2);
    expect(wb.worksheets[0].freezePane).toEqual({ row: 1, col: 0 });
    expect(wb.worksheets[0].autoFilter).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: 10,
      endCol: 1,
    });
    expect(wb.worksheets[0].rows[0].cells[0].hyperlink).toEqual({
      target: 'https://bun.sh',
      tooltip: 'Bun docs',
    });
    expect(wb.worksheets[1].splitPane).toEqual({
      x: 900,
      y: 1600,
      topLeftCell: { row: 1, col: 1 },
    });

    const rule = wb.worksheets[1].conditionalFormattings?.[0].rules[0];
    expect(rule?.type).toBe('expression');
    if (rule?.type === 'expression') {
      expect(rule.style?.fill?.fgColor).toBe('C6EFCE');
    }
  });

  test('writes multiple sheets with comments, images, and tables', async () => {
    const path = `${TMP}/multi-stream-comments-images-tables.xlsx`;
    const stream = createMultiSheetExcelStream(path);

    stream.addSheet('Sheet1', {
      images: [
        {
          data: PNG_1X1,
          format: 'png',
          range: { startRow: 4, startCol: 0, endRow: 5, endCol: 1 },
          name: 'Logo',
        },
      ],
      tables: [
        {
          name: 'AssetsTable',
          range: { startRow: 0, startCol: 0, endRow: 2, endCol: 2 },
          columns: [{ name: 'Name' }, { name: 'Value' }, { name: 'Link' }],
        },
      ],
    });
    stream.writeRow({
      cells: [
        { value: 'Name', comment: { text: 'Header comment', author: 'Loc' } },
        { value: 'Value' },
        { value: 'Link' },
      ],
    });
    stream.writeRow(['A', 1, 'Docs']);
    stream.writeRow(['B', 2, 'Guide']);

    stream.addSheet('Notes');
    stream.writeRow({
      cells: [{ value: 'Reminder', comment: { text: 'Second sheet note' } }],
    });

    await stream.end();

    const wb = await readExcel(path);
    expect(wb.worksheets).toHaveLength(2);
    expect(wb.worksheets[0].rows[0].cells[0].comment?.author).toBe('Loc');
    expect(wb.worksheets[0].images?.[0].range).toEqual({
      startRow: 4,
      startCol: 0,
      endRow: 5,
      endCol: 1,
    });
    expect(wb.worksheets[0].tables?.[0].name).toBe('AssetsTable');
    expect(wb.worksheets[1].rows[0].cells[0].comment?.text).toBe(
      'Second sheet note',
    );
  });
});

describe('Chunked Stream Writer', () => {
  test('writes Excel via chunked stream', async () => {
    const path = `${TMP}/chunked.xlsx`;
    const stream = createChunkedExcelStream(path, {
      sheetName: 'Chunked',
      columns: [{ width: 10 }, { width: 20 }],
    });

    stream.writeRow({
      cells: [
        { value: 'ID', style: { font: { bold: true } } },
        { value: 'Name', style: { font: { bold: true } } },
      ],
    });

    for (let i = 0; i < 100; i++) {
      stream.writeRow([i + 1, `Item_${i + 1}`]);
    }

    await stream.end();

    const wb = await readExcel(path);
    expect(wb.worksheets).toHaveLength(1);
    expect(wb.worksheets[0].rows).toHaveLength(101);
  });

  test('writes Excel via chunked stream to Bun.file target', async () => {
    const path = `${TMP}/chunked-bun-file.xlsx`;
    const stream = createChunkedExcelStream(Bun.file(path), {
      sheetName: 'Chunked',
    });

    stream.writeRow(['ID', 'Name']);
    stream.writeRow([1, 'Alice']);
    await stream.end();

    const wb = await readExcel(path);
    expect(wb.worksheets[0].rows[1].cells[1].value).toBe('Alice');
  });

  test('chunked stream with styles', async () => {
    const path = `${TMP}/chunked-styles.xlsx`;
    const stream = createChunkedExcelStream(path, {
      sheetName: 'Styled',
    });

    stream.writeRow({
      cells: [
        {
          value: 'Bold',
          style: {
            font: { bold: true, color: 'FF0000' },
            fill: { type: 'pattern', pattern: 'solid', fgColor: 'FFFF00' },
          },
        },
      ],
    });

    stream.writeRow([1234.5]);
    await stream.end();

    const wb = await readExcel(path);
    expect(wb.worksheets[0].rows[0].cells[0].value).toBe('Bold');
  });

  test('chunked stream writes hyperlinks without buffering them in memory', async () => {
    const path = `${TMP}/chunked-hyperlinks.xlsx`;
    const stream = createChunkedExcelStream(path, {
      sheetName: 'Links',
    });

    stream.writeRow({
      cells: [
        {
          value: 'External',
          hyperlink: {
            target: 'https://bun.sh',
            tooltip: 'Bun website',
          },
        },
        {
          value: 'Internal',
          hyperlink: {
            target: 'Sheet1!A10',
            tooltip: 'Jump inside sheet',
          },
        },
      ],
    });

    await stream.end();

    const wb = await readExcel(path);
    expect(wb.worksheets[0].rows[0].cells[0].hyperlink).toEqual({
      target: 'https://bun.sh',
      tooltip: 'Bun website',
    });
    expect(wb.worksheets[0].rows[0].cells[1].hyperlink).toEqual({
      target: 'Sheet1!A10',
      tooltip: 'Jump inside sheet',
    });
  });

  test('chunked stream with freeze pane', async () => {
    const path = `${TMP}/chunked-freeze.xlsx`;
    const stream = createChunkedExcelStream(path, {
      sheetName: 'Frozen',
      freezePane: { row: 1, col: 0 },
    });

    stream.writeRow(['Header']);
    stream.writeRow(['Data']);
    await stream.end();

    const wb = await readExcel(path);
    expect(wb.worksheets[0].freezePane).toEqual({ row: 1, col: 0 });
  });

  test('chunked stream with auto filter', async () => {
    const path = `${TMP}/chunked-autofilter.xlsx`;
    const stream = createChunkedExcelStream(path, {
      sheetName: 'Filtered',
      autoFilter: { startRow: 0, startCol: 0, endRow: 25, endCol: 2 },
    });

    stream.writeRow(['Name', 'Score', 'Team']);
    stream.writeRow(['Alice', 95, 'A']);
    await stream.end();

    const wb = await readExcel(path);
    expect(wb.worksheets[0].autoFilter).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: 25,
      endCol: 2,
    });
  });

  test('chunked stream with split pane', async () => {
    const path = `${TMP}/chunked-split.xlsx`;
    const stream = createChunkedExcelStream(path, {
      sheetName: 'Split',
      splitPane: {
        x: 1440,
        y: 2200,
        topLeftCell: { row: 2, col: 1 },
      },
    });

    stream.writeRow(['Header']);
    stream.writeRow(['Data']);
    await stream.end();

    const wb = await readExcel(path);
    expect(wb.worksheets[0].splitPane).toEqual({
      x: 1440,
      y: 2200,
      topLeftCell: { row: 2, col: 1 },
    });
  });

  test('chunked stream with data validation', async () => {
    const path = `${TMP}/chunked-validation.xlsx`;
    const stream = createChunkedExcelStream(path, {
      sheetName: 'Validated',
      dataValidations: [
        {
          range: { startRow: 1, startCol: 1, endRow: 10, endCol: 1 },
          type: 'whole',
          operator: 'between',
          formula1: 1,
          formula2: 5,
        },
      ],
    });

    stream.writeRow(['Task', 'Score']);
    stream.writeRow(['One', 3]);
    await stream.end();

    const wb = await readExcel(path);
    const validation = wb.worksheets[0].dataValidations?.[0];
    expect(validation?.type).toBe('whole');
    expect(validation?.formula1).toBe(1);
    expect(validation?.formula2).toBe(5);
  });

  test('chunked stream with conditional formatting', async () => {
    const path = `${TMP}/chunked-conditional.xlsx`;
    const stream = createChunkedExcelStream(path, {
      sheetName: 'Conditional',
      conditionalFormattings: [
        {
          range: { startRow: 1, startCol: 1, endRow: 20, endCol: 1 },
          rules: [
            {
              type: 'dataBar',
              color: '5B9BD5',
              min: { type: 'min' },
              max: { type: 'max' },
              showValue: false,
            },
          ],
        },
      ],
    });

    stream.writeRow(['Item', 'Score']);
    stream.writeRow(['One', 5]);
    await stream.end();

    const rule = readConditionalRule(await readExcel(path));
    expect(rule?.type).toBe('dataBar');
    if (rule?.type === 'dataBar') {
      expect(rule.color).toBe('5B9BD5');
      expect(rule.showValue).toBe(false);
    }
  });

  test('chunked stream with comments, images, and tables', async () => {
    const path = `${TMP}/chunked-comments-images-tables.xlsx`;
    const stream = createChunkedExcelStream(path, {
      sheetName: 'Assets',
      images: [
        {
          data: PNG_1X1,
          format: 'png',
          range: { startRow: 4, startCol: 0, endRow: 5, endCol: 1 },
          name: 'Logo',
        },
      ],
      tables: [
        {
          name: 'ChunkedAssetTable',
          range: { startRow: 0, startCol: 0, endRow: 2, endCol: 2 },
          columns: [{ name: 'Name' }, { name: 'Value' }, { name: 'Link' }],
        },
      ],
    });

    stream.writeRow({
      cells: [
        { value: 'Name', comment: { text: 'Header comment', author: 'Loc' } },
        { value: 'Value' },
        { value: 'Link' },
      ],
    });
    stream.writeRow(['A', 1, 'Docs']);
    stream.writeRow({
      cells: [
        { value: 'B', comment: { text: 'Body comment' } },
        { value: 2 },
        { value: 'Guide' },
      ],
    });
    await stream.end();

    const sheet = (await readExcel(path)).worksheets[0];
    expect(sheet.rows[0].cells[0].comment?.author).toBe('Loc');
    expect(sheet.rows[2].cells[0].comment?.text).toBe('Body comment');
    expect(sheet.images?.[0].range).toEqual({
      startRow: 4,
      startCol: 0,
      endRow: 5,
      endCol: 1,
    });
    expect(sheet.tables?.[0].name).toBe('ChunkedAssetTable');
    expect(sheet.tables?.[0].columns?.map((column) => column.name)).toEqual([
      'Name',
      'Value',
      'Link',
    ]);
  });

  test('chunked stream supports uncompressed ZIP output', async () => {
    const path = `${TMP}/chunked-store.xlsx`;
    const stream = createChunkedExcelStream(path, {
      sheetName: 'Stored',
      compress: false,
    });

    for (let i = 0; i < 250; i++) {
      stream.writeRow([i, `Row_${i}`]);
    }

    await stream.end();

    const wb = await readExcel(path);
    expect(wb.worksheets[0].rows).toHaveLength(250);
    expect(wb.worksheets[0].rows[249].cells[1].value).toBe('Row_249');
  });

  test('chunked stream handles large data (5K rows)', async () => {
    const path = `${TMP}/chunked-large.xlsx`;
    const stream = createChunkedExcelStream(path, { sheetName: 'Large' });

    for (let i = 0; i < 5000; i++) {
      stream.writeRow([i, `Row_${i}`, Math.random() * 10000]);
    }
    await stream.end();

    const wb = await readExcel(path);
    expect(wb.worksheets[0].rows).toHaveLength(5000);
    expect(wb.worksheets[0].rows[0].cells[0].value).toBe(0);
    expect(wb.worksheets[0].rows[4999].cells[0].value).toBe(4999);
  });
});

describe('Excel Stream Reader', () => {
  test('streams rows from a workbook path with dates and formulas', async () => {
    const path = `${TMP}/excel-read-stream.xlsx`;
    await writeExcel(path, {
      worksheets: [
        {
          name: 'Data',
          rows: [
            {
              cells: [{ value: 'Date' }, { value: 'Name' }, { value: 'Total' }],
            },
            {
              cells: [
                {
                  value: new Date('2026-03-16T00:00:00.000Z'),
                  style: { numberFormat: 'yyyy-mm-dd' },
                },
                { value: 'Alice' },
                {
                  value: 42,
                  type: 'formula',
                  formula: 'SUM(40,2)',
                  formulaResult: 42,
                },
              ],
            },
          ],
        },
      ],
    });

    const streamedRows = [];
    for await (const row of readExcelStream(path)) {
      streamedRows.push(row);
    }

    expect(streamedRows).toHaveLength(2);
    expect(streamedRows[0]?.sheetName).toBe('Data');
    expect(streamedRows[0]?.rowIndex).toBe(0);
    expect(streamedRows[1]?.row.cells[0]?.value).toBeInstanceOf(Date);
    expect(streamedRows[1]?.row.cells[1]?.value).toBe('Alice');
    expect(streamedRows[1]?.row.cells[2]?.formula).toBe('SUM(40,2)');
    expect(streamedRows[1]?.row.cells[2]?.value).toBe(42);
  });

  test('streams selected sheets from Bun.file source', async () => {
    const path = `${TMP}/excel-read-stream-sheets.xlsx`;
    await writeExcel(path, {
      worksheets: [
        {
          name: 'First',
          rows: [{ cells: [{ value: 'A' }] }],
        },
        {
          name: 'Second',
          rows: [{ cells: [{ value: 'B' }] }, { cells: [{ value: 'C' }] }],
        },
      ],
    });

    const streamedRows = [];
    for await (const row of readExcelStream(Bun.file(path), {
      sheets: ['Second'],
    })) {
      streamedRows.push(row);
    }

    expect(streamedRows).toHaveLength(2);
    expect(new Set(streamedRows.map((row) => row.sheetName))).toEqual(
      new Set(['Second']),
    );
    expect(streamedRows[0]?.sheetIndex).toBe(1);
    expect(streamedRows[1]?.row.cells[0]?.value).toBe('C');
  });
});

describe('Production Excel Export API', () => {
  test('passes S3 writer options through ManagedFileSink', async () => {
    const calls: unknown[] = [];
    const fakeS3File = {
      name: 's3://bucket/report.xlsx',
      presign() {
        return Promise.resolve('');
      },
      writer(options?: unknown) {
        calls.push(options);
        return {
          write() {},
          flush() {},
          end() {},
        };
      },
    } as unknown as Bun.S3File;

    const sink = new ManagedFileSink(fakeS3File, {
      s3WriterOptions: {
        partSize: 8 * 1024 * 1024,
        queueSize: 4,
        retry: 5,
      },
    });
    sink.write('hello');
    await sink.end();

    expect(calls).toEqual([
      {
        partSize: 8 * 1024 * 1024,
        queueSize: 4,
        retry: 5,
      },
    ]);
  });

  test('exports rows with progress and diagnostics', async () => {
    const path = `${TMP}/export-rows.xlsx`;
    const progressStages: string[] = [];

    const diagnostics = await exportExcelRows({
      target: path,
      sheetName: 'Report',
      mode: 'chunked',
      progressIntervalRows: 2,
      rows: [
        ['ID', 'Name'],
        [1, 'Alice'],
        [2, 'Bob'],
        [3, 'Carol'],
      ],
      onProgress(progress) {
        progressStages.push(progress.stage);
      },
    });

    const workbook = await readExcel(path);
    expect(workbook.worksheets[0].rows).toHaveLength(4);
    expect(diagnostics.rowsWritten).toBe(4);
    expect(diagnostics.outputSizeBytes).toBeGreaterThan(0);
    expect(diagnostics.memory.peak.rssBytes).toBeGreaterThan(0);
    expect(progressStages).toContain('writing');
    expect(progressStages.at(-1)).toBe('completed');
  });

  test('aborts export jobs with AbortSignal', async () => {
    const path = `${TMP}/export-aborted.xlsx`;
    const controller = new AbortController();

    async function* rows() {
      yield ['ID', 'Value'];
      yield [1, 'A'];
      controller.abort(new Error('stop export'));
      yield [2, 'B'];
    }

    await expect(
      exportExcelRows({
        target: path,
        sheetName: 'Abort',
        rows: rows(),
        signal: controller.signal,
      }),
    ).rejects.toMatchObject({
      name: 'AbortError',
    });

    expect(await Bun.file(path).exists()).toBe(false);
  });

  test('builds Bun Response for Excel downloads', async () => {
    const response = await buildExcelResponse(
      {
        worksheets: [
          {
            name: 'Sheet1',
            rows: [{ cells: [{ value: 'Hello' }, { value: 123 }] }],
          },
        ],
      },
      {
        filename: 'report.xlsx',
      },
    );

    expect(response.headers.get('content-type')).toBe(
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );
    expect(response.headers.get('content-disposition')).toContain(
      'report.xlsx',
    );

    const responsePath = `${TMP}/response-export.xlsx`;
    await Bun.write(responsePath, new Uint8Array(await response.arrayBuffer()));
    const workbook = await readExcel(responsePath);
    expect(workbook.worksheets[0].rows[0].cells[0].value).toBe('Hello');
  });

  test('streams exported Excel as Response without building full buffer in user code', async () => {
    const { response, diagnostics } = await exportExcelRowsToResponse({
      filename: 'streamed.xlsx',
      sheetName: 'Orders',
      mode: 'chunked',
      rows: [
        ['ID', 'Name'],
        [1, 'Alice'],
        [2, 'Bob'],
      ],
    });

    expect(response.headers.get('content-disposition')).toContain(
      'streamed.xlsx',
    );
    expect(diagnostics.rowsWritten).toBe(3);

    const responsePath = `${TMP}/response-streamed.xlsx`;
    await Bun.write(responsePath, new Uint8Array(await response.arrayBuffer()));
    const workbook = await readExcel(responsePath);
    expect(workbook.worksheets[0].rows[2].cells[1].value).toBe('Bob');
  });

  test('exports multiple sheets with progress and diagnostics', async () => {
    const path = `${TMP}/export-multi-sheet.xlsx`;
    const seenSheets = new Set<string>();

    const diagnostics = await exportMultiSheetExcel({
      target: path,
      creator: 'export-engine',
      progressIntervalRows: 2,
      sheets: [
        {
          name: 'Orders',
          options: { columns: [{ width: 18 }, { width: 22 }] },
          rows: [
            ['ID', 'Customer'],
            [1, 'Alice'],
            [2, 'Bob'],
          ],
        },
        {
          name: 'Summary',
          rows: [
            ['Metric', 'Value'],
            ['Orders', 2],
          ],
        },
      ],
      onProgress(progress) {
        if (progress.sheetName) {
          seenSheets.add(progress.sheetName);
        }
      },
    });

    const workbook = await readExcel(path);
    expect(workbook.creator).toBe('export-engine');
    expect(workbook.worksheets).toHaveLength(2);
    expect(workbook.worksheets[0].rows[2].cells[1].value).toBe('Bob');
    expect(workbook.worksheets[1].rows[1].cells[0].value).toBe('Orders');
    expect(diagnostics.mode).toBe('multi-sheet');
    expect(diagnostics.sheetCount).toBe(2);
    expect(diagnostics.rowsWritten).toBe(5);
    expect(seenSheets).toEqual(new Set(['Orders', 'Summary']));
  });

  test('aborts multi-sheet export jobs with AbortSignal', async () => {
    const path = `${TMP}/export-multi-aborted.xlsx`;
    const controller = new AbortController();

    async function* secondSheetRows() {
      yield ['Metric', 'Value'];
      controller.abort(new Error('stop multi export'));
      yield ['Orders', 2];
    }

    await expect(
      exportMultiSheetExcel({
        target: path,
        signal: controller.signal,
        sheets: [
          {
            name: 'Orders',
            rows: [
              ['ID', 'Name'],
              [1, 'Alice'],
            ],
          },
          {
            name: 'Summary',
            rows: secondSheetRows(),
          },
        ],
      }),
    ).rejects.toMatchObject({
      name: 'AbortError',
    });

    expect(await Bun.file(path).exists()).toBe(false);
  });
});

function readConditionalRule(workbook: Awaited<ReturnType<typeof readExcel>>) {
  return workbook.worksheets[0].conditionalFormattings?.[0].rules[0];
}
