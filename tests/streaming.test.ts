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
