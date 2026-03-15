import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { mkdirSync, rmSync } from 'node:fs';
import {
  createChunkedExcelStream,
  createCSVStream,
  createExcelStream,
  createMultiSheetExcelStream,
  readCSV,
  readExcel,
} from '../src';

const TMP = './tests/.tmp';

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

function readConditionalRule(workbook: Awaited<ReturnType<typeof readExcel>>) {
  return workbook.worksheets[0].conditionalFormattings?.[0].rules[0];
}
