// ============================================
// bun-spreadsheet — Full Feature Demo
// ============================================

import {
  type CellStyle,
  createCSVStream,
  createExcelStream,
  createMultiSheetExcelStream,
  readCSV,
  readExcel,
  type Workbook,
  writeCSV,
  writeExcel,
} from '../src/index';

const OUTPUT = './output';

import { mkdirSync } from 'node:fs';

mkdirSync(OUTPUT, { recursive: true });

console.log('bun-spreadsheet Demo');
console.log('='.repeat(60));

// --- Styles ------------------------------------------------------------------

const headerStyle: CellStyle = {
  font: { bold: true, size: 14, color: 'FFFFFF', name: 'Arial' },
  fill: { type: 'pattern', pattern: 'solid', fgColor: '4472C4' },
  alignment: { horizontal: 'center', vertical: 'center' },
  border: {
    top: { style: 'thin', color: '000000' },
    bottom: { style: 'medium', color: '000000' },
    left: { style: 'thin', color: '000000' },
    right: { style: 'thin', color: '000000' },
  },
};

const numberStyle: CellStyle = {
  numberFormat: '#,##0.00',
  font: { size: 11 },
  alignment: { horizontal: 'right' },
  border: {
    top: { style: 'thin', color: 'D9D9D9' },
    bottom: { style: 'thin', color: 'D9D9D9' },
    left: { style: 'thin', color: 'D9D9D9' },
    right: { style: 'thin', color: 'D9D9D9' },
  },
};

const highlightStyle: CellStyle = {
  font: { bold: true, color: '006100' },
  fill: { type: 'pattern', pattern: 'solid', fgColor: 'C6EFCE' },
  numberFormat: '#,##0.00',
  alignment: { horizontal: 'right' },
};

const dateStyle: CellStyle = {
  numberFormat: 'yyyy-mm-dd',
  alignment: { horizontal: 'center' },
};

const linkStyle: CellStyle = {
  font: { color: '0563C1', underline: true, size: 11 },
};

// =============================================================================
// 1. Write Excel with styles, merge cells, freeze pane
// =============================================================================
console.log('\n[1/9] Writing styled Excel...');

const workbook: Workbook = {
  worksheets: [
    {
      name: 'Sales Report',
      columns: [
        { width: 15, header: 'Date' },
        { width: 25, header: 'Product' },
        { width: 15, header: 'Category' },
        { width: 12, header: 'Quantity' },
        { width: 15, header: 'Unit Price' },
        { width: 18, header: 'Total' },
      ],
      freezePane: { row: 1, col: 0 },
      rows: [
        {
          cells: [
            { value: 'Date', style: headerStyle },
            { value: 'Product', style: headerStyle },
            { value: 'Category', style: headerStyle },
            { value: 'Quantity', style: headerStyle },
            { value: 'Unit Price', style: headerStyle },
            { value: 'Total', style: headerStyle },
          ],
          height: 30,
        },
        {
          cells: [
            { value: '2024-01-15', style: dateStyle },
            { value: 'Laptop Pro 16' },
            { value: 'Electronics' },
            { value: 5, style: numberStyle },
            { value: 1299.99, style: numberStyle },
            { value: 6499.95, style: highlightStyle },
          ],
        },
        {
          cells: [
            { value: '2024-01-16', style: dateStyle },
            { value: 'Wireless Mouse' },
            { value: 'Accessories' },
            { value: 50, style: numberStyle },
            { value: 29.99, style: numberStyle },
            { value: 1499.5, style: numberStyle },
          ],
        },
        {
          cells: [
            { value: '2024-01-17', style: dateStyle },
            { value: 'USB-C Hub 7-in-1' },
            { value: 'Accessories' },
            { value: 30, style: numberStyle },
            { value: 49.99, style: numberStyle },
            { value: 1499.7, style: numberStyle },
          ],
        },
        {
          cells: [
            { value: '2024-01-18', style: dateStyle },
            { value: 'Monitor 27" 4K' },
            { value: 'Electronics' },
            { value: 10, style: numberStyle },
            { value: 599.99, style: numberStyle },
            { value: 5999.9, style: highlightStyle },
          ],
        },
        {
          cells: [
            { value: '2024-01-19', style: dateStyle },
            { value: 'Keyboard Mechanical' },
            { value: 'Accessories' },
            { value: 25, style: numberStyle },
            { value: 89.99, style: numberStyle },
            { value: 2249.75, style: numberStyle },
          ],
        },
      ],
      mergeCells: [],
    },
    {
      name: 'Summary',
      columns: [{ width: 20 }, { width: 20 }],
      rows: [
        {
          cells: [
            {
              value: 'Sales Summary',
              style: {
                font: { bold: true, size: 18, color: '1F4E79' },
                alignment: { horizontal: 'center' },
              },
            },
            { value: null },
          ],
        },
        { cells: [] },
        {
          cells: [
            { value: 'Total Revenue', style: { font: { bold: true } } },
            { value: 17748.8, style: highlightStyle },
          ],
        },
        {
          cells: [
            { value: 'Total Items Sold', style: { font: { bold: true } } },
            { value: 120, style: numberStyle },
          ],
        },
        {
          cells: [
            { value: 'Average Order', style: { font: { bold: true } } },
            { value: 3549.76, style: numberStyle },
          ],
        },
      ],
      mergeCells: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }],
    },
  ],
  creator: 'bun-spreadsheet',
};

await writeExcel(`${OUTPUT}/styled-report.xlsx`, workbook);
console.log('      -> output/styled-report.xlsx');

// =============================================================================
// 2. Read Excel
// =============================================================================
console.log('\n[2/9] Reading Excel file back...');

const readBack = await readExcel(`${OUTPUT}/styled-report.xlsx`);
console.log(`      ${readBack.worksheets.length} worksheets found:`);

for (const ws of readBack.worksheets) {
  console.log(`      - "${ws.name}": ${ws.rows.length} rows`);
  for (let r = 0; r < Math.min(3, ws.rows.length); r++) {
    const row = ws.rows[r];
    if (!row) continue;
    const values = row.cells
      .map((c) =>
        c.value !== null && c.value !== undefined ? String(c.value) : '',
      )
      .join(' | ');
    console.log(`        Row ${r + 1}: ${values}`);
  }
  if (ws.rows.length > 3) {
    console.log(`        ... and ${ws.rows.length - 3} more rows`);
  }
}

// =============================================================================
// 3. Write CSV
// =============================================================================
console.log('\n[3/9] Writing CSV...');

const csvData = [
  ['Name', 'Email', 'Age', 'City'],
  ['Alice', 'alice@example.com', 28, 'Hanoi'],
  ['Bob', 'bob@example.com', 32, 'Ho Chi Minh'],
  ['Charlie', 'charlie@example.com', 25, 'Da Nang'],
  ['Diana', 'diana@example.com', 30, 'Hue'],
];

await writeCSV(`${OUTPUT}/contacts.csv`, csvData as (string | number)[][], {
  includeHeader: false,
});
console.log('      -> output/contacts.csv');

// =============================================================================
// 4. Read CSV
// =============================================================================
console.log('\n[4/9] Reading CSV back...');

const csvWorkbook = await readCSV(`${OUTPUT}/contacts.csv`);
const csvSheet = csvWorkbook.worksheets[0];
console.log(`      ${csvSheet.rows.length} rows:`);
for (let r = 0; r < Math.min(5, csvSheet.rows.length); r++) {
  const values = csvSheet.rows[r].cells.map((c) => String(c.value)).join(' | ');
  console.log(`        Row ${r + 1}: ${values}`);
}

// =============================================================================
// 5. CSV Stream (10K rows)
// =============================================================================
console.log('\n[5/9] CSV streaming write (10,000 rows)...');

const csvStream = createCSVStream(`${OUTPUT}/csv-stream-10k.csv`, {
  headers: ['ID', 'Name', 'Value', 'Timestamp'],
  includeHeader: true,
});

const csvStart = performance.now();
for (let i = 0; i < 10000; i++) {
  csvStream.writeRow([
    i + 1,
    `Item_${i + 1}`,
    Math.round(Math.random() * 10000) / 100,
    new Date(2024, 0, 1 + (i % 365)).toISOString(),
  ]);
}
await csvStream.end();
const csvMs = (performance.now() - csvStart).toFixed(2);
console.log(`      -> output/csv-stream-10k.csv (${csvMs}ms)`);

// =============================================================================
// 6. Excel Stream (10K rows with styles)
// =============================================================================
console.log('\n[6/9] Excel streaming write (10,000 rows)...');

const streamHeaderStyle: CellStyle = {
  font: { bold: true, color: 'FFFFFF', size: 12 },
  fill: { type: 'pattern', pattern: 'solid', fgColor: '2E75B6' },
  alignment: { horizontal: 'center' },
};

const excelStream = createExcelStream(`${OUTPUT}/excel-stream-10k.xlsx`, {
  sheetName: 'Data',
  columns: [
    { width: 10 },
    { width: 25 },
    { width: 15 },
    { width: 20 },
    { width: 15 },
  ],
  freezePane: { row: 1, col: 0 },
});

excelStream.writeRow({
  cells: [
    { value: 'ID', style: streamHeaderStyle },
    { value: 'Product Name', style: streamHeaderStyle },
    { value: 'Price', style: streamHeaderStyle },
    { value: 'Created Date', style: streamHeaderStyle },
    { value: 'In Stock', style: streamHeaderStyle },
  ],
  height: 25,
});

const excelStart = performance.now();
for (let i = 0; i < 10000; i++) {
  const price = Math.round(Math.random() * 100000) / 100;
  const rowStyle: CellStyle | undefined =
    price > 800 ? highlightStyle : numberStyle;

  excelStream.writeRow({
    cells: [
      { value: i + 1 },
      { value: `Product_${String(i + 1).padStart(5, '0')}` },
      { value: price, style: rowStyle },
      {
        value: `2024-${String((i % 12) + 1).padStart(2, '0')}-${String((i % 28) + 1).padStart(2, '0')}`,
        style: dateStyle,
      },
      { value: Math.random() > 0.3 },
    ],
  });
}

await excelStream.end();
const excelMs = (performance.now() - excelStart).toFixed(2);
console.log(`      -> output/excel-stream-10k.xlsx (${excelMs}ms)`);

// =============================================================================
// 7. Multi-Sheet Stream
// =============================================================================
console.log('\n[7/9] Multi-sheet streaming write...');

const multiStream = createMultiSheetExcelStream(`${OUTPUT}/multi-sheet.xlsx`);

multiStream.addSheet('Revenue', {
  columns: [{ width: 15 }, { width: 15 }, { width: 15 }],
});

multiStream.writeRow({
  cells: [
    { value: 'Month', style: headerStyle },
    { value: 'Revenue', style: headerStyle },
    { value: 'Growth', style: headerStyle },
  ],
});

const months = [
  'Jan',
  'Feb',
  'Mar',
  'Apr',
  'May',
  'Jun',
  'Jul',
  'Aug',
  'Sep',
  'Oct',
  'Nov',
  'Dec',
];

for (let i = 0; i < 12; i++) {
  const revenue = 50000 + Math.random() * 50000;
  const growth = (Math.random() - 0.3) * 20;
  multiStream.writeRow([
    months[i],
    Math.round(revenue * 100) / 100,
    Math.round(growth * 100) / 100,
  ]);
}

multiStream.addSheet('Expenses', {
  columns: [{ width: 20 }, { width: 15 }],
});

multiStream.writeRow({
  cells: [
    { value: 'Category', style: headerStyle },
    { value: 'Amount', style: headerStyle },
  ],
});

const expenses = [
  ['Salaries', 120000],
  ['Marketing', 35000],
  ['Operations', 28000],
  ['Technology', 45000],
  ['Office', 15000],
];

for (const [cat, amt] of expenses) {
  multiStream.writeRow([cat, amt]);
}

await multiStream.end();
console.log('      -> output/multi-sheet.xlsx (2 sheets)');

// =============================================================================
// 8. Hyperlinks & Formulas
// =============================================================================
console.log('\n[8/9] Writing hyperlinks & formulas...');

const formulaWorkbook: Workbook = {
  worksheets: [
    {
      name: 'Links & Formulas',
      columns: [{ width: 25 }, { width: 30 }, { width: 20 }, { width: 20 }],
      rows: [
        {
          cells: [
            { value: 'Description', style: headerStyle },
            { value: 'Link / Formula', style: headerStyle },
            { value: 'Value A', style: headerStyle },
            { value: 'Value B', style: headerStyle },
          ],
          height: 30,
        },
        {
          cells: [
            { value: 'Website' },
            {
              value: 'Visit Bun.sh',
              style: linkStyle,
              hyperlink: {
                target: 'https://bun.sh',
                tooltip: 'Open Bun website',
              },
            },
            { value: null },
            { value: null },
          ],
        },
        {
          cells: [
            { value: 'Contact' },
            {
              value: 'Send Email',
              style: linkStyle,
              hyperlink: {
                target: 'mailto:hello@example.com',
                tooltip: 'Send email',
              },
            },
            { value: null },
            { value: null },
          ],
        },
        {
          cells: [
            { value: 'Go to Summary' },
            {
              value: 'Click here',
              style: linkStyle,
              hyperlink: {
                target: 'Calculations!A1',
                tooltip: 'Go to Calculations sheet',
              },
            },
            { value: null },
            { value: null },
          ],
        },
        { cells: [] },
        {
          cells: [
            { value: 'Product A' },
            { value: null },
            { value: 150, style: numberStyle },
            { value: 250, style: numberStyle },
          ],
        },
        {
          cells: [
            { value: 'Product B' },
            { value: null },
            { value: 300, style: numberStyle },
            { value: 450, style: numberStyle },
          ],
        },
        {
          cells: [
            { value: 'Product C' },
            { value: null },
            { value: 500, style: numberStyle },
            { value: 120, style: numberStyle },
          ],
        },
        { cells: [] },
        {
          cells: [
            { value: 'SUM of A', style: { font: { bold: true } } },
            {
              value: null,
              formula: 'SUM(C6:C8)',
              formulaResult: 950,
              style: highlightStyle,
            },
            { value: null },
            { value: null },
          ],
        },
        {
          cells: [
            { value: 'SUM of B', style: { font: { bold: true } } },
            {
              value: null,
              formula: 'SUM(D6:D8)',
              formulaResult: 820,
              style: highlightStyle,
            },
            { value: null },
            { value: null },
          ],
        },
        {
          cells: [
            { value: 'AVERAGE', style: { font: { bold: true } } },
            {
              value: null,
              formula: 'AVERAGE(C6:D8)',
              formulaResult: 295,
              style: numberStyle,
            },
            { value: null },
            { value: null },
          ],
        },
        {
          cells: [
            { value: 'A > B ?', style: { font: { bold: true } } },
            {
              value: null,
              formula: 'IF(B10>B11,"A wins","B wins")',
              formulaResult: 'A wins',
              style: { font: { bold: true, color: 'FF6600' } },
            },
            { value: null },
            { value: null },
          ],
        },
      ],
    },
    {
      name: 'Calculations',
      columns: [{ width: 25 }, { width: 20 }],
      rows: [
        {
          cells: [
            {
              value: 'This is the Calculations sheet',
              style: { font: { bold: true, size: 14 } },
            },
            { value: null },
          ],
        },
        {
          cells: [
            {
              value: 'Back to Links',
              style: linkStyle,
              hyperlink: {
                target: "'Links & Formulas'!A1",
                tooltip: 'Go back',
              },
            },
            { value: null },
          ],
        },
      ],
    },
  ],
  creator: 'bun-spreadsheet',
};

await writeExcel(`${OUTPUT}/hyperlinks-formulas.xlsx`, formulaWorkbook);
console.log('      -> output/hyperlinks-formulas.xlsx');

// =============================================================================
// 9. Read back Hyperlinks & Formulas
// =============================================================================
console.log('\n[9/9] Verifying hyperlinks & formulas...');

const hlReadBack = await readExcel(`${OUTPUT}/hyperlinks-formulas.xlsx`);
const hlSheet = hlReadBack.worksheets[0];
console.log(`      ${hlSheet.rows.length} rows in "${hlSheet.name}":`);

for (let r = 0; r < hlSheet.rows.length; r++) {
  const row = hlSheet.rows[r];
  if (!row) continue;
  for (let c = 0; c < row.cells.length; c++) {
    const cell = row.cells[c];
    if (cell?.hyperlink) {
      console.log(
        `        [${r + 1},${c + 1}] "${cell.value}" -> ${cell.hyperlink.target}${cell.hyperlink.tooltip ? ` (${cell.hyperlink.tooltip})` : ''}`,
      );
    }
    if (cell?.formula) {
      console.log(
        `        [${r + 1},${c + 1}] =${cell.formula} -> ${cell.value}`,
      );
    }
  }
}

// =============================================================================
// Summary
// =============================================================================
console.log(`\n${'='.repeat(60)}`);
console.log('Done. Output files:\n');
console.log(
  '  output/styled-report.xlsx      Styled Excel (2 sheets, merge, freeze)',
);
console.log('  output/contacts.csv            CSV file');
console.log('  output/csv-stream-10k.csv      10K rows CSV (streamed)');
console.log('  output/excel-stream-10k.xlsx   10K rows Excel (streamed)');
console.log('  output/multi-sheet.xlsx        Multi-sheet Excel (streamed)');
console.log('  output/hyperlinks-formulas.xlsx Hyperlinks + Formulas');
console.log('\nOpen .xlsx files in Excel or Google Sheets to verify.');
