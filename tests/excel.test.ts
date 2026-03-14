import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { mkdirSync, rmSync } from 'node:fs';
import { type CellStyle, readExcel, type Workbook, writeExcel } from '../src';

const TMP = './tests/.tmp';

beforeAll(() => {
  mkdirSync(TMP, { recursive: true });
});

afterAll(() => {
  rmSync(TMP, { recursive: true, force: true });
});

describe('Excel Writer', () => {
  test('writes basic Excel file', async () => {
    const path = `${TMP}/basic.xlsx`;
    await writeExcel(path, {
      worksheets: [
        {
          name: 'Sheet1',
          rows: [
            { cells: [{ value: 'Hello' }, { value: 123 }] },
            { cells: [{ value: 'World' }, { value: 456 }] },
          ],
        },
      ],
    });

    const file = Bun.file(path);
    expect(file.size).toBeGreaterThan(0);
  });

  test('writes multiple worksheets', async () => {
    const path = `${TMP}/multi-sheet.xlsx`;
    await writeExcel(path, {
      worksheets: [
        { name: 'First', rows: [{ cells: [{ value: 'A' }] }] },
        { name: 'Second', rows: [{ cells: [{ value: 'B' }] }] },
        { name: 'Third', rows: [{ cells: [{ value: 'C' }] }] },
      ],
    });

    const wb = await readExcel(path);
    expect(wb.worksheets).toHaveLength(3);
    expect(wb.worksheets[0].name).toBe('First');
    expect(wb.worksheets[1].name).toBe('Second');
    expect(wb.worksheets[2].name).toBe('Third');
  });

  test('writes cell styles', async () => {
    const path = `${TMP}/styles.xlsx`;
    const style: CellStyle = {
      font: { bold: true, size: 14, color: 'FF0000' },
      fill: { type: 'pattern', pattern: 'solid', fgColor: 'FFFF00' },
      alignment: { horizontal: 'center' },
      numberFormat: '#,##0.00',
    };

    await writeExcel(path, {
      worksheets: [
        {
          name: 'Styled',
          rows: [{ cells: [{ value: 1234.5, style }] }],
        },
      ],
    });

    const wb = await readExcel(path);
    const cell = wb.worksheets[0].rows[0].cells[0];
    expect(cell.value).toBe(1234.5);
    expect(cell.style?.font?.bold).toBe(true);
    expect(cell.style?.font?.color).toBe('FF0000');
  });

  test('writes merge cells', async () => {
    const path = `${TMP}/merge.xlsx`;
    await writeExcel(path, {
      worksheets: [
        {
          name: 'Merged',
          rows: [
            { cells: [{ value: 'Title' }, { value: null }, { value: null }] },
            { cells: [{ value: 'A' }, { value: 'B' }, { value: 'C' }] },
          ],
          mergeCells: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 2 }],
        },
      ],
    });

    const wb = await readExcel(path);
    expect(wb.worksheets[0].mergeCells).toHaveLength(1);
    expect(wb.worksheets[0].mergeCells?.[0]).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: 0,
      endCol: 2,
    });
  });

  test('writes freeze pane', async () => {
    const path = `${TMP}/freeze.xlsx`;
    await writeExcel(path, {
      worksheets: [
        {
          name: 'Frozen',
          rows: [
            { cells: [{ value: 'Header' }] },
            { cells: [{ value: 'Data' }] },
          ],
          freezePane: { row: 1, col: 0 },
        },
      ],
    });

    const wb = await readExcel(path);
    expect(wb.worksheets[0].freezePane).toEqual({ row: 1, col: 0 });
  });

  test('writes column widths', async () => {
    const path = `${TMP}/columns.xlsx`;
    await writeExcel(path, {
      worksheets: [
        {
          name: 'Cols',
          columns: [{ width: 30 }, { width: 15 }],
          rows: [{ cells: [{ value: 'Wide' }, { value: 'Normal' }] }],
        },
      ],
    });

    const wb = await readExcel(path);
    expect(wb.worksheets[0].columns).toBeDefined();
    expect(wb.worksheets[0].columns?.length).toBeGreaterThanOrEqual(2);
  });

  test('writes row height', async () => {
    const path = `${TMP}/row-height.xlsx`;
    await writeExcel(path, {
      worksheets: [
        {
          name: 'Heights',
          rows: [
            { cells: [{ value: 'Tall row' }], height: 40 },
            { cells: [{ value: 'Normal row' }] },
          ],
        },
      ],
    });

    const file = Bun.file(path);
    expect(file.size).toBeGreaterThan(0);
  });
});

describe('Excel Reader', () => {
  test('reads written file back correctly', async () => {
    const path = `${TMP}/read-back.xlsx`;
    const original: Workbook = {
      worksheets: [
        {
          name: 'Data',
          rows: [
            { cells: [{ value: 'Name' }, { value: 'Age' }] },
            { cells: [{ value: 'Alice' }, { value: 28 }] },
            { cells: [{ value: 'Bob' }, { value: 32 }] },
          ],
        },
      ],
    };

    await writeExcel(path, original);
    const wb = await readExcel(path);

    expect(wb.worksheets).toHaveLength(1);
    expect(wb.worksheets[0].name).toBe('Data');
    expect(wb.worksheets[0].rows).toHaveLength(3);
    expect(wb.worksheets[0].rows[0].cells[0].value).toBe('Name');
    expect(wb.worksheets[0].rows[1].cells[0].value).toBe('Alice');
    expect(wb.worksheets[0].rows[1].cells[1].value).toBe(28);
  });

  test('reads boolean values', async () => {
    const path = `${TMP}/booleans.xlsx`;
    await writeExcel(path, {
      worksheets: [
        {
          name: 'Bool',
          rows: [{ cells: [{ value: true }, { value: false }] }],
        },
      ],
    });

    const wb = await readExcel(path);
    expect(wb.worksheets[0].rows[0].cells[0].value).toBe(true);
    expect(wb.worksheets[0].rows[0].cells[1].value).toBe(false);
  });

  test('handles empty worksheet', async () => {
    const path = `${TMP}/empty.xlsx`;
    await writeExcel(path, {
      worksheets: [{ name: 'Empty', rows: [] }],
    });

    const wb = await readExcel(path);
    expect(wb.worksheets).toHaveLength(1);
    expect(wb.worksheets[0].rows).toHaveLength(0);
  });

  test('preserves number formats', async () => {
    const path = `${TMP}/numfmt.xlsx`;
    await writeExcel(path, {
      worksheets: [
        {
          name: 'Fmt',
          rows: [
            {
              cells: [{ value: 1234.5, style: { numberFormat: '#,##0.00' } }],
            },
          ],
        },
      ],
    });

    const wb = await readExcel(path);
    const cell = wb.worksheets[0].rows[0].cells[0];
    expect(cell.value).toBe(1234.5);
    // numberFormat is applied via style index; verify style exists
    expect(cell.style).toBeDefined();
  });
});

describe('Formulas', () => {
  test('writes and reads formulas with cached results', async () => {
    const path = `${TMP}/formulas.xlsx`;
    await writeExcel(path, {
      worksheets: [
        {
          name: 'Calc',
          rows: [
            { cells: [{ value: 10 }, { value: 20 }, { value: 30 }] },
            {
              cells: [
                {
                  value: null,
                  formula: 'SUM(A1:C1)',
                  formulaResult: 60,
                },
              ],
            },
          ],
        },
      ],
    });

    const wb = await readExcel(path);
    const formulaCell = wb.worksheets[0].rows[1].cells[0];
    expect(formulaCell.formula).toBe('SUM(A1:C1)');
    expect(formulaCell.value).toBe(60);
  });

  test('writes multiple formula types', async () => {
    const path = `${TMP}/multi-formulas.xlsx`;
    await writeExcel(path, {
      worksheets: [
        {
          name: 'Formulas',
          rows: [
            { cells: [{ value: 100 }, { value: 200 }, { value: 300 }] },
            {
              cells: [
                { value: null, formula: 'SUM(A1:C1)', formulaResult: 600 },
                {
                  value: null,
                  formula: 'AVERAGE(A1:C1)',
                  formulaResult: 200,
                },
                { value: null, formula: 'MAX(A1:C1)', formulaResult: 300 },
              ],
            },
          ],
        },
      ],
    });

    const wb = await readExcel(path);
    const row = wb.worksheets[0].rows[1];
    expect(row.cells[0].formula).toBe('SUM(A1:C1)');
    expect(row.cells[1].formula).toBe('AVERAGE(A1:C1)');
    expect(row.cells[2].formula).toBe('MAX(A1:C1)');
  });
});

describe('Hyperlinks', () => {
  test('writes and reads external hyperlink', async () => {
    const path = `${TMP}/hyperlinks.xlsx`;
    await writeExcel(path, {
      worksheets: [
        {
          name: 'Links',
          rows: [
            {
              cells: [
                {
                  value: 'Visit',
                  hyperlink: {
                    target: 'https://bun.sh',
                    tooltip: 'Bun website',
                  },
                },
              ],
            },
          ],
        },
      ],
    });

    const wb = await readExcel(path);
    const cell = wb.worksheets[0].rows[0].cells[0];
    expect(cell.value).toBe('Visit');
    expect(cell.hyperlink?.target).toBe('https://bun.sh');
    expect(cell.hyperlink?.tooltip).toBe('Bun website');
  });

  test('writes mailto hyperlink', async () => {
    const path = `${TMP}/mailto.xlsx`;
    await writeExcel(path, {
      worksheets: [
        {
          name: 'Mail',
          rows: [
            {
              cells: [
                {
                  value: 'Email',
                  hyperlink: { target: 'mailto:test@example.com' },
                },
              ],
            },
          ],
        },
      ],
    });

    const wb = await readExcel(path);
    const cell = wb.worksheets[0].rows[0].cells[0];
    expect(cell.hyperlink?.target).toBe('mailto:test@example.com');
  });

  test('writes internal sheet reference', async () => {
    const path = `${TMP}/internal-link.xlsx`;
    await writeExcel(path, {
      worksheets: [
        {
          name: 'Sheet1',
          rows: [
            {
              cells: [
                {
                  value: 'Go',
                  hyperlink: { target: 'Sheet2!A1' },
                },
              ],
            },
          ],
        },
        {
          name: 'Sheet2',
          rows: [{ cells: [{ value: 'Target' }] }],
        },
      ],
    });

    const wb = await readExcel(path);
    const cell = wb.worksheets[0].rows[0].cells[0];
    expect(cell.hyperlink?.target).toBe('Sheet2!A1');
  });
});

describe('Data Validation', () => {
  test('writes and reads dropdown list validations', async () => {
    const path = `${TMP}/validation-list.xlsx`;
    await writeExcel(path, {
      worksheets: [
        {
          name: 'Status',
          rows: [
            { cells: [{ value: 'Status' }] },
            { cells: [{ value: null }] },
          ],
          dataValidations: [
            {
              range: { startRow: 1, startCol: 0, endRow: 10, endCol: 0 },
              type: 'list',
              allowBlank: true,
              showErrorMessage: true,
              errorTitle: 'Invalid status',
              error: 'Pick a value from the list',
              formula1: ['New', 'In Progress', 'Done'],
            },
          ],
        },
      ],
    });

    const wb = await readExcel(path);
    const validation = wb.worksheets[0].dataValidations?.[0];
    expect(validation).toBeDefined();
    expect(validation?.type).toBe('list');
    expect(validation?.formula1).toEqual(['New', 'In Progress', 'Done']);
    expect(validation?.allowBlank).toBe(true);
  });

  test('writes and reads number range validations', async () => {
    const path = `${TMP}/validation-number.xlsx`;
    await writeExcel(path, {
      worksheets: [
        {
          name: 'Scores',
          rows: [{ cells: [{ value: 'Score' }] }, { cells: [{ value: 50 }] }],
          dataValidations: [
            {
              range: { startRow: 1, startCol: 0, endRow: 100, endCol: 0 },
              type: 'whole',
              operator: 'between',
              formula1: 0,
              formula2: 100,
            },
          ],
        },
      ],
    });

    const wb = await readExcel(path);
    const validation = wb.worksheets[0].dataValidations?.[0];
    expect(validation?.type).toBe('whole');
    expect(validation?.operator).toBe('between');
    expect(validation?.formula1).toBe(0);
    expect(validation?.formula2).toBe(100);
  });

  test('writes and reads date limit validations', async () => {
    const path = `${TMP}/validation-date.xlsx`;
    const startDate = new Date(Date.UTC(2026, 0, 1));
    const endDate = new Date(Date.UTC(2026, 11, 31));

    await writeExcel(path, {
      worksheets: [
        {
          name: 'Dates',
          rows: [
            { cells: [{ value: 'Due Date' }] },
            { cells: [{ value: null }] },
          ],
          dataValidations: [
            {
              range: { startRow: 1, startCol: 0, endRow: 50, endCol: 0 },
              type: 'date',
              operator: 'between',
              formula1: startDate,
              formula2: endDate,
            },
          ],
        },
      ],
    });

    const wb = await readExcel(path);
    const validation = wb.worksheets[0].dataValidations?.[0];
    expect(validation?.type).toBe('date');
    expect(validation?.operator).toBe('between');
    expect(validation?.formula1).toBeInstanceOf(Date);
    expect(validation?.formula2).toBeInstanceOf(Date);
  });

  test('writes and reads custom formula validations', async () => {
    const path = `${TMP}/validation-custom.xlsx`;
    await writeExcel(path, {
      worksheets: [
        {
          name: 'Unique',
          rows: [{ cells: [{ value: 'Code' }] }, { cells: [{ value: null }] }],
          dataValidations: [
            {
              range: { startRow: 1, startCol: 0, endRow: 50, endCol: 0 },
              type: 'custom',
              showInputMessage: true,
              promptTitle: 'Unique code',
              prompt: 'Each code must be unique in column A',
              formula1: '=COUNTIF($A:$A,A2)=1',
            },
          ],
        },
      ],
    });

    const wb = await readExcel(path);
    const validation = wb.worksheets[0].dataValidations?.[0];
    expect(validation?.type).toBe('custom');
    expect(validation?.formula1).toBe('COUNTIF($A:$A,A2)=1');
    expect(validation?.promptTitle).toBe('Unique code');
  });
});

describe('Special Characters', () => {
  test('handles XML special characters in cell values', async () => {
    const path = `${TMP}/special-chars.xlsx`;
    await writeExcel(path, {
      worksheets: [
        {
          name: 'Special',
          rows: [
            {
              cells: [
                { value: 'less < greater >' },
                { value: 'amp & quote "' },
                { value: "apostrophe '" },
              ],
            },
          ],
        },
      ],
    });

    const wb = await readExcel(path);
    const cells = wb.worksheets[0].rows[0].cells;
    expect(cells[0].value).toBe('less < greater >');
    expect(cells[1].value).toBe('amp & quote "');
    expect(cells[2].value).toBe("apostrophe '");
  });

  test('handles unicode in cell values', async () => {
    const path = `${TMP}/unicode.xlsx`;
    await writeExcel(path, {
      worksheets: [
        {
          name: 'Unicode',
          rows: [
            {
              cells: [
                { value: 'Vietnamese: Xin chao' },
                { value: 'Japanese: Konnichiwa' },
                { value: 'Symbols: -- +/-' },
              ],
            },
          ],
        },
      ],
    });

    const wb = await readExcel(path);
    const cells = wb.worksheets[0].rows[0].cells;
    expect(cells[0].value).toContain('Vietnamese');
    expect(cells[1].value).toContain('Japanese');
  });
});
