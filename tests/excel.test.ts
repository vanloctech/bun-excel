import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { mkdirSync, rmSync } from 'node:fs';
import {
  type CellStyle,
  duplicateRow,
  insertColumns,
  insertRows,
  loadExcelTemplate,
  readExcel,
  spliceRows,
  type Workbook,
  writeExcel,
} from '../src';

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

  test('writes Excel to Bun.file target', async () => {
    const path = `${TMP}/bun-file-target.xlsx`;
    await writeExcel(Bun.file(path), {
      worksheets: [
        {
          name: 'Sheet1',
          rows: [{ cells: [{ value: 'Hello' }, { value: 123 }] }],
        },
      ],
    });

    const wb = await readExcel(path);
    expect(wb.worksheets[0].rows[0].cells[0].value).toBe('Hello');
    expect(wb.worksheets[0].rows[0].cells[1].value).toBe(123);
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

  test('writes auto filter', async () => {
    const path = `${TMP}/autofilter.xlsx`;
    await writeExcel(path, {
      worksheets: [
        {
          name: 'Filtered',
          rows: [
            { cells: [{ value: 'Name' }, { value: 'Score' }] },
            { cells: [{ value: 'Alice' }, { value: 95 }] },
            { cells: [{ value: 'Bob' }, { value: 87 }] },
          ],
          autoFilter: { startRow: 0, startCol: 0, endRow: 2, endCol: 1 },
        },
      ],
    });

    const wb = await readExcel(path);
    expect(wb.worksheets[0].autoFilter).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: 2,
      endCol: 1,
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

  test('writes split pane', async () => {
    const path = `${TMP}/split.xlsx`;
    await writeExcel(path, {
      worksheets: [
        {
          name: 'Split',
          rows: [
            { cells: [{ value: 'Header' }] },
            { cells: [{ value: 'Data' }] },
          ],
          splitPane: {
            x: 1200,
            y: 1800,
            topLeftCell: { row: 1, col: 1 },
          },
        },
      ],
    });

    const wb = await readExcel(path);
    expect(wb.worksheets[0].splitPane).toEqual({
      x: 1200,
      y: 1800,
      topLeftCell: { row: 1, col: 1 },
    });
    expect(wb.worksheets[0].freezePane).toBeUndefined();
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

  test('reads Excel from Bun.file source', async () => {
    const path = `${TMP}/bun-file-source.xlsx`;
    await writeExcel(path, {
      worksheets: [
        {
          name: 'Data',
          rows: [{ cells: [{ value: 'Alice' }, { value: 28 }] }],
        },
      ],
    });

    const wb = await readExcel(Bun.file(path));
    expect(wb.worksheets[0].rows[0].cells[0].value).toBe('Alice');
    expect(wb.worksheets[0].rows[0].cells[1].value).toBe(28);
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

  test('reads workbook properties', async () => {
    const path = `${TMP}/workbook-props.xlsx`;
    const created = new Date('2026-02-01T10:00:00.000Z');
    const modified = new Date('2026-02-02T12:30:00.000Z');

    await writeExcel(path, {
      worksheets: [{ name: 'Meta', rows: [{ cells: [{ value: 'Hello' }] }] }],
      creator: 'bun-spreadsheet',
      created,
      modified,
    });

    const wb = await readExcel(path);
    expect(wb.creator).toBe('bun-spreadsheet');
    expect(wb.created?.toISOString()).toBe(created.toISOString());
    expect(wb.modified?.toISOString()).toBe(modified.toISOString());
  });

  test('reads date cells as Date values when number format is date-based', async () => {
    const path = `${TMP}/date-cells.xlsx`;
    const input = new Date('2026-01-15T00:00:00.000Z');

    await writeExcel(path, {
      worksheets: [
        {
          name: 'Dates',
          rows: [
            {
              cells: [
                {
                  value: input,
                  style: { numberFormat: 'yyyy-mm-dd' },
                },
              ],
            },
          ],
        },
      ],
    });

    const cell = (await readExcel(path)).worksheets[0].rows[0].cells[0];
    expect(cell.type).toBe('date');
    expect(cell.value).toBeInstanceOf(Date);
    expect((cell.value as Date).toISOString()).toBe(input.toISOString());
  });

  test('writes and reads gradient fills', async () => {
    const path = `${TMP}/gradient-fill.xlsx`;
    await writeExcel(path, {
      worksheets: [
        {
          name: 'Gradient',
          rows: [
            {
              cells: [
                {
                  value: 'Heatmap',
                  style: {
                    fill: {
                      type: 'gradient',
                      fgColor: 'FFF2CC',
                      bgColor: 'F4B183',
                    },
                  },
                },
              ],
            },
          ],
        },
      ],
    });

    const cell = (await readExcel(path)).worksheets[0].rows[0].cells[0];
    expect(cell.style?.fill).toEqual({
      type: 'gradient',
      fgColor: 'FFF2CC',
      bgColor: 'F4B183',
    });
  });

  test('writes and reads workbook and worksheet view settings', async () => {
    const path = `${TMP}/worksheet-settings.xlsx`;
    await writeExcel(path, {
      creator: 'bun-spreadsheet',
      views: {
        activeTab: 1,
        windowWidth: 24000,
        windowHeight: 12000,
      },
      definedNames: [{ name: 'SalesData', refersTo: "'Visible'!$A$2:$B$3" }],
      worksheets: [
        {
          name: 'Visible',
          state: 'visible',
          rows: [
            {
              cells: [{ value: 'Name' }, { value: 'Score' }],
              outlineLevel: 1,
              collapsed: true,
            },
            {
              cells: [
                { value: 'Alice' },
                {
                  value: 95,
                  style: { protection: { locked: false } },
                },
              ],
              hidden: true,
              outlineLevel: 1,
            },
          ],
          columns: [
            { width: 20, hidden: true, outlineLevel: 1 },
            { width: 15, collapsed: true, outlineLevel: 1 },
          ],
          protection: {
            sheet: true,
            autoFilter: true,
            password: 'secret',
          },
          pageMargins: { left: 0.5, right: 0.5, top: 0.75, bottom: 0.75 },
          pageSetup: {
            orientation: 'landscape',
            fitToWidth: 1,
            fitToHeight: 1,
          },
          headerFooter: {
            oddHeader: { left: 'Report', right: 'Page &P' },
            oddFooter: { center: 'Confidential' },
          },
          printArea: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
        },
        {
          name: 'Hidden',
          state: 'hidden',
          rows: [{ cells: [{ value: 'secret' }] }],
        },
      ],
    });

    const wb = await readExcel(path);
    expect(wb.views?.activeTab).toBe(1);
    expect(wb.definedNames?.[0]).toEqual({
      name: 'SalesData',
      refersTo: "'Visible'!$A$2:$B$3",
      comment: undefined,
      hidden: false,
      localSheetId: undefined,
    });
    expect(wb.worksheets[0].columns?.[0].hidden).toBe(true);
    expect(wb.worksheets[0].rows[0].outlineLevel).toBe(1);
    expect(wb.worksheets[0].rows[1].hidden).toBe(true);
    expect(wb.worksheets[0].protection?.sheet).toBe(true);
    expect(wb.worksheets[0].pageSetup?.orientation).toBe('landscape');
    expect(wb.worksheets[0].headerFooter?.oddHeader?.left).toBe('Report');
    expect(wb.worksheets[0].printArea).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: 1,
      endCol: 1,
    });
    expect(wb.worksheets[1].state).toBe('hidden');
  });
});

describe('Rich Text', () => {
  test('writes and reads partial rich text styles', async () => {
    const path = `${TMP}/rich-text.xlsx`;
    await writeExcel(path, {
      worksheets: [
        {
          name: 'Rich',
          rows: [
            {
              cells: [
                {
                  value: 'Hello Bun',
                  richText: [
                    { text: 'Hello ', font: { bold: true, color: 'FF0000' } },
                    { text: 'Bun', font: { italic: true, color: '0000FF' } },
                  ],
                },
              ],
            },
          ],
        },
      ],
    });

    const cell = (await readExcel(path)).worksheets[0].rows[0].cells[0];
    expect(cell.value).toBe('Hello Bun');
    expect(cell.richText).toEqual([
      { text: 'Hello ', font: { bold: true, color: 'FF0000' } },
      { text: 'Bun', font: { italic: true, color: '0000FF' } },
    ]);
  });
});

describe('Comments, Images, and Tables', () => {
  test('writes and reads comments, images, and tables', async () => {
    const path = `${TMP}/comments-images-tables.xlsx`;
    await writeExcel(path, {
      worksheets: [
        {
          name: 'Assets',
          rows: [
            {
              cells: [
                {
                  value: 'Product',
                  comment: { text: 'Main product header', author: 'Loc' },
                },
                { value: 'Price' },
                { value: 'Link' },
              ],
            },
            {
              cells: [
                { value: 'Widget' },
                { value: 99.5 },
                {
                  value: 'Docs',
                  hyperlink: { target: 'https://bun.sh' },
                },
              ],
            },
            {
              cells: [
                { value: 'Gadget', comment: { text: 'Promo item' } },
                { value: 149.25 },
                { value: 'Guide' },
              ],
            },
          ],
          images: [
            {
              data: PNG_1X1,
              format: 'png',
              range: { startRow: 4, startCol: 0, endRow: 5, endCol: 1 },
              name: 'Logo',
              description: 'Tiny logo',
            },
          ],
          tables: [
            {
              name: 'SalesTable',
              range: { startRow: 0, startCol: 0, endRow: 2, endCol: 2 },
              headerRow: true,
              totalsRow: false,
              style: { name: 'TableStyleMedium2', showRowStripes: true },
            },
          ],
        },
      ],
    });

    const wb = await readExcel(path);
    const sheet = wb.worksheets[0];
    expect(sheet.rows[0].cells[0].comment).toEqual({
      text: 'Main product header',
      author: 'Loc',
    });
    expect(sheet.rows[2].cells[0].comment).toEqual({
      text: 'Promo item',
      author: 'Author',
    });
    expect(sheet.images).toHaveLength(1);
    expect(sheet.images?.[0].format).toBe('png');
    expect(sheet.images?.[0].range).toEqual({
      startRow: 4,
      startCol: 0,
      endRow: 5,
      endCol: 1,
    });
    expect(new Uint8Array(sheet.images?.[0].data as Uint8Array)[0]).toBe(137);
    expect(sheet.tables?.[0]).toMatchObject({
      name: 'SalesTable',
      range: { startRow: 0, startCol: 0, endRow: 2, endCol: 2 },
      headerRow: true,
      totalsRow: false,
    });
    expect(sheet.tables?.[0].columns?.map((column) => column.name)).toEqual([
      'Product',
      'Price',
      'Link',
    ]);
  });
});

describe('Worksheet Operations', () => {
  test('inserts, splices, duplicates rows and inserts columns', () => {
    const worksheet: Workbook['worksheets'][number] = {
      name: 'Ops',
      rows: [
        { cells: [{ value: 'A1' }, { value: 'B1' }] },
        { cells: [{ value: 'A2' }, { value: 'B2' }] },
      ],
      columns: [{ width: 10 }, { width: 10 }],
    };

    insertRows(worksheet, 1, [{ cells: [{ value: 'X1' }, { value: 'Y1' }] }]);
    duplicateRow(worksheet, 0, 1);
    spliceRows(worksheet, 3, 1, [{ cells: [{ value: 'R' }, { value: 'S' }] }]);
    insertColumns(worksheet, 1, 1);

    expect(worksheet.rows).toHaveLength(4);
    expect(worksheet.rows[1].cells[0].value).toBe('A1');
    expect(worksheet.rows[2].cells[0].value).toBe('X1');
    expect(worksheet.rows[3].cells[0].value).toBe('R');
    expect(worksheet.rows[0].cells).toHaveLength(3);
    expect(worksheet.columns).toHaveLength(3);
  });
});

describe('Template Mode', () => {
  test('loads a template and fills cells and defined ranges while preserving template features', async () => {
    const templatePath = `${TMP}/invoice-template.xlsx`;
    const outputPath = `${TMP}/invoice-filled.xlsx`;

    await writeExcel(templatePath, {
      creator: 'Template Author',
      definedNames: [
        { name: 'InvoiceNumber', refersTo: "'Invoice'!$B$2" },
        { name: 'InvoiceTotal', refersTo: "'Invoice'!$C$3" },
        { name: 'LineItems', refersTo: "'Invoice'!$A$6:$B$7" },
      ],
      worksheets: [
        {
          name: 'Invoice',
          freezePane: { row: 1, col: 0 },
          images: [
            {
              data: PNG_1X1,
              format: 'png',
              range: { startRow: 0, startCol: 3, endRow: 2, endCol: 4 },
              name: 'Template Logo',
            },
          ],
          tables: [
            {
              name: 'ItemsTable',
              range: { startRow: 4, startCol: 0, endRow: 6, endCol: 1 },
              columns: [{ name: 'Item' }, { name: 'Qty' }],
              style: { name: 'TableStyleMedium2', showRowStripes: true },
            },
          ],
          rows: [
            {
              cells: [
                { value: 'Invoice Template', style: { font: { bold: true } } },
              ],
            },
            {
              cells: [
                { value: 'Invoice #' },
                {
                  value: 'TBD',
                  style: { font: { bold: true, color: '1F4E78' } },
                  comment: { text: 'Template comment', author: 'Loc' },
                },
              ],
            },
            {
              cells: [
                { value: 'Total' },
                { value: null },
                {
                  value: 0,
                  style: { numberFormat: '$#,##0.00' },
                },
              ],
            },
            { cells: [] },
            {
              cells: [
                { value: 'Item', style: { font: { bold: true } } },
                { value: 'Qty', style: { font: { bold: true } } },
              ],
            },
            { cells: [{ value: '' }, { value: 0 }] },
            { cells: [{ value: '' }, { value: 0 }] },
          ],
        },
      ],
    });

    const template = await loadExcelTemplate(Bun.file(templatePath));
    template.setCell('Invoice', 'A1', 'Invoice 2026');
    template.setDefinedName('InvoiceNumber', 'INV-001');
    template.setDefinedName('InvoiceTotal', 1234.5);
    template.setDefinedName('LineItems', [
      ['Apple', 2],
      ['Orange', 5],
    ]);
    await template.write(outputPath);

    const sheet = (await readExcel(outputPath)).worksheets[0];
    expect(sheet.rows[0].cells[0].value).toBe('Invoice 2026');
    expect(sheet.rows[1].cells[1].value).toBe('INV-001');
    expect(sheet.rows[1].cells[1].style?.font?.bold).toBe(true);
    expect(sheet.rows[1].cells[1].style?.font?.color).toBe('1F4E78');
    expect(sheet.rows[1].cells[1].comment?.text).toBe('Template comment');
    expect(sheet.rows[2].cells[2].value).toBe(1234.5);
    expect(sheet.rows[2].cells[2].style?.numberFormat).toBe('$#,##0.00');
    expect(sheet.rows[5].cells[0].value).toBe('Apple');
    expect(sheet.rows[5].cells[1].value).toBe(2);
    expect(sheet.rows[6].cells[0].value).toBe('Orange');
    expect(sheet.rows[6].cells[1].value).toBe(5);
    expect(sheet.images?.[0].name).toBe('Template Logo');
    expect(sheet.tables?.[0].name).toBe('ItemsTable');
    expect(sheet.freezePane).toEqual({ row: 1, col: 0 });
  });

  test('throws when assigning a scalar value to a multi-cell defined name', async () => {
    const templatePath = `${TMP}/template-defined-range.xlsx`;

    await writeExcel(templatePath, {
      definedNames: [{ name: 'LineItems', refersTo: "'Sheet1'!$A$1:$B$2" }],
      worksheets: [
        {
          name: 'Sheet1',
          rows: [
            { cells: [{ value: '' }, { value: '' }] },
            { cells: [{ value: '' }, { value: '' }] },
          ],
        },
      ],
    });

    const template = await loadExcelTemplate(templatePath);
    expect(() => template.setDefinedName('LineItems', 'single value')).toThrow(
      'Provide a 2D array',
    );
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

describe('Conditional Formatting', () => {
  test('writes and reads highlight cell rules', async () => {
    const path = `${TMP}/conditional-highlight.xlsx`;
    await writeExcel(path, {
      worksheets: [
        {
          name: 'Highlight',
          rows: [{ cells: [{ value: 'Score' }] }, { cells: [{ value: 95 }] }],
          conditionalFormattings: [
            {
              range: { startRow: 1, startCol: 0, endRow: 100, endCol: 0 },
              rules: [
                {
                  type: 'cellIs',
                  operator: 'greaterThan',
                  formula1: 80,
                  stopIfTrue: true,
                  style: {
                    font: { bold: true, color: '9C0006' },
                    fill: {
                      type: 'pattern',
                      pattern: 'solid',
                      fgColor: 'FFC7CE',
                    },
                  },
                },
              ],
            },
          ],
        },
      ],
    });

    const wb = await readExcel(path);
    const rule = wb.worksheets[0].conditionalFormattings?.[0].rules[0];
    expect(rule?.type).toBe('cellIs');
    if (rule?.type === 'cellIs') {
      expect(rule.operator).toBe('greaterThan');
      expect(rule.formula1).toBe(80);
      expect(rule.stopIfTrue).toBe(true);
      expect(rule.style?.fill?.fgColor).toBe('FFC7CE');
    }
  });

  test('writes and reads color scales, data bars, and icon sets', async () => {
    const path = `${TMP}/conditional-visuals.xlsx`;
    await writeExcel(path, {
      worksheets: [
        {
          name: 'Visuals',
          rows: [
            { cells: [{ value: 10 }] },
            { cells: [{ value: 50 }] },
            { cells: [{ value: 90 }] },
          ],
          conditionalFormattings: [
            {
              range: { startRow: 0, startCol: 0, endRow: 20, endCol: 0 },
              rules: [
                {
                  type: 'colorScale',
                  thresholds: [
                    { type: 'min' },
                    { type: 'percentile', value: 50 },
                    { type: 'max' },
                  ],
                  colors: ['F8696B', 'FFEB84', '63BE7B'],
                },
                {
                  type: 'dataBar',
                  min: { type: 'min' },
                  max: { type: 'max' },
                  color: '638EC6',
                  showValue: false,
                  minLength: 10,
                  maxLength: 90,
                },
                {
                  type: 'iconSet',
                  iconSet: '3TrafficLights1',
                  thresholds: [
                    { type: 'percent', value: 0 },
                    { type: 'percent', value: 33 },
                    { type: 'percent', value: 67 },
                  ],
                  showValue: false,
                  reverse: true,
                },
              ],
            },
          ],
        },
      ],
    });

    const rules = wbConditionalRules(await readExcel(path));
    expect(rules[0]?.type).toBe('colorScale');
    if (rules[0]?.type === 'colorScale') {
      expect(rules[0].colors).toEqual(['F8696B', 'FFEB84', '63BE7B']);
      expect(rules[0].thresholds[1].value).toBe(50);
    }

    expect(rules[1]?.type).toBe('dataBar');
    if (rules[1]?.type === 'dataBar') {
      expect(rules[1].color).toBe('638EC6');
      expect(rules[1].showValue).toBe(false);
      expect(rules[1].minLength).toBe(10);
      expect(rules[1].maxLength).toBe(90);
    }

    expect(rules[2]?.type).toBe('iconSet');
    if (rules[2]?.type === 'iconSet') {
      expect(rules[2].iconSet).toBe('3TrafficLights1');
      expect(rules[2].reverse).toBe(true);
      expect(rules[2].thresholds).toHaveLength(3);
    }
  });

  test('writes and reads multiple ranges with preserved priorities', async () => {
    const path = `${TMP}/conditional-multi-range.xlsx`;
    await writeExcel(path, {
      worksheets: [
        {
          name: 'Ranges',
          rows: [
            { cells: [{ value: 'A' }, { value: 'B' }, { value: 'C' }] },
            { cells: [{ value: 1 }, { value: 2 }, { value: 3 }] },
          ],
          conditionalFormattings: [
            {
              range: [
                { startRow: 1, startCol: 0, endRow: 10, endCol: 0 },
                { startRow: 1, startCol: 2, endRow: 10, endCol: 2 },
              ],
              rules: [
                {
                  type: 'expression',
                  formula: '=MOD(ROW(),2)=0',
                  priority: 7,
                  stopIfTrue: true,
                  style: {
                    fill: {
                      type: 'pattern',
                      pattern: 'solid',
                      fgColor: 'F2F2F2',
                    },
                  },
                },
                {
                  type: 'iconSet',
                  iconSet: '3Arrows',
                  priority: 8,
                  showValue: false,
                  thresholds: [
                    { type: 'percent', value: 0, gte: false },
                    { type: 'percent', value: 33, gte: true },
                    { type: 'percent', value: 67, gte: true },
                  ],
                },
              ],
            },
          ],
        },
      ],
    });

    const formatting = await readExcel(path).then(
      (workbook) => workbook.worksheets[0].conditionalFormattings?.[0],
    );

    expect(formatting?.range).toEqual([
      { startRow: 1, startCol: 0, endRow: 10, endCol: 0 },
      { startRow: 1, startCol: 2, endRow: 10, endCol: 2 },
    ]);

    const expressionRule = formatting?.rules[0];
    expect(expressionRule?.type).toBe('expression');
    if (expressionRule?.type === 'expression') {
      expect(expressionRule.priority).toBe(7);
      expect(expressionRule.stopIfTrue).toBe(true);
      expect(expressionRule.formula).toBe('MOD(ROW(),2)=0');
    }

    const iconRule = formatting?.rules[1];
    expect(iconRule?.type).toBe('iconSet');
    if (iconRule?.type === 'iconSet') {
      expect(iconRule.priority).toBe(8);
      expect(iconRule.showValue).toBe(false);
      expect(iconRule.thresholds.map((threshold) => threshold.gte)).toEqual([
        false,
        true,
        true,
      ]);
    }
  });
});

function wbConditionalRules(workbook: Workbook) {
  return workbook.worksheets[0].conditionalFormattings?.[0].rules || [];
}

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
