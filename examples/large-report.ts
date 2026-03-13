// ============================================
// Large Report — 30 columns x 30,000 rows
// with merged SUM/AVG/MAX/MIN footer
// ============================================

import {
  type Cell,
  type CellStyle,
  type ColumnConfig,
  type MergeCell,
  type Row,
  type Workbook,
  writeExcel,
} from '../src/index';

const OUTPUT = './output';

import { mkdirSync } from 'node:fs';

mkdirSync(OUTPUT, { recursive: true });

console.log('Large Report Generator');
console.log('='.repeat(60));

// --- Styles ------------------------------------------------------------------

const headerStyle: CellStyle = {
  font: { bold: true, size: 11, color: 'FFFFFF', name: 'Arial' },
  fill: { type: 'pattern', pattern: 'solid', fgColor: '2F5496' },
  alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
  border: {
    top: { style: 'thin', color: '1F3864' },
    bottom: { style: 'medium', color: '1F3864' },
    left: { style: 'thin', color: '1F3864' },
    right: { style: 'thin', color: '1F3864' },
  },
};

const dataStyle: CellStyle = {
  font: { size: 10 },
  border: {
    top: { style: 'thin', color: 'D6DCE4' },
    bottom: { style: 'thin', color: 'D6DCE4' },
    left: { style: 'thin', color: 'D6DCE4' },
    right: { style: 'thin', color: 'D6DCE4' },
  },
};

const numberDataStyle: CellStyle = {
  ...dataStyle,
  numberFormat: '#,##0',
  alignment: { horizontal: 'right' },
};

const currencyStyle: CellStyle = {
  ...dataStyle,
  numberFormat: '#,##0.00',
  alignment: { horizontal: 'right' },
  font: { size: 10 },
};

const percentStyle: CellStyle = {
  ...dataStyle,
  numberFormat: '0.0%',
  alignment: { horizontal: 'center' },
};

const dateDataStyle: CellStyle = {
  ...dataStyle,
  alignment: { horizontal: 'center' },
};

const evenRowStyle: CellStyle = {
  ...dataStyle,
  fill: { type: 'pattern', pattern: 'solid', fgColor: 'F2F2F2' },
};
const evenNumberStyle: CellStyle = {
  ...numberDataStyle,
  fill: { type: 'pattern', pattern: 'solid', fgColor: 'F2F2F2' },
};
const evenCurrencyStyle: CellStyle = {
  ...currencyStyle,
  fill: { type: 'pattern', pattern: 'solid', fgColor: 'F2F2F2' },
};
const evenPercentStyle: CellStyle = {
  ...percentStyle,
  fill: { type: 'pattern', pattern: 'solid', fgColor: 'F2F2F2' },
};
const evenDateStyle: CellStyle = {
  ...dateDataStyle,
  fill: { type: 'pattern', pattern: 'solid', fgColor: 'F2F2F2' },
};

const footerBase: CellStyle = {
  font: { bold: true, size: 12, color: 'FFFFFF' },
  alignment: { horizontal: 'right', vertical: 'center' },
  border: {
    top: { style: 'medium', color: '000000' },
    bottom: { style: 'medium', color: '000000' },
    left: { style: 'thin', color: '000000' },
    right: { style: 'thin', color: '000000' },
  },
};

const titleStyle: CellStyle = {
  font: { bold: true, size: 16, color: '1F3864' },
  alignment: { horizontal: 'center', vertical: 'center' },
};

const subtitleStyle: CellStyle = {
  font: { size: 11, color: '595959', italic: true },
  alignment: { horizontal: 'center', vertical: 'center' },
};

// --- Config ------------------------------------------------------------------

const COL_COUNT = 30;
const DATA_ROWS = 30_000;

const departments = [
  'Sales',
  'Marketing',
  'Engineering',
  'HR',
  'Finance',
  'Operations',
  'Support',
  'Legal',
  'R&D',
  'Product',
];
const regions = ['North', 'South', 'East', 'West', 'Central'];
const statuses = ['Active', 'Pending', 'Completed', 'Cancelled', 'On Hold'];
const categories = ['A', 'B', 'C', 'D', 'E'];

const columnHeaders = [
  'ID',
  'Date',
  'Department',
  'Region',
  'Employee',
  'Category',
  'Status',
  'Revenue',
  'Cost',
  'Profit',
  'Quantity',
  'Unit Price',
  'Discount %',
  'Tax',
  'Net Amount',
  'Budget',
  'Actual',
  'Variance',
  'Target',
  'Achievement %',
  'Hours',
  'Rate',
  'Labor Cost',
  'Material Cost',
  'Overhead',
  'Total Cost',
  'Margin',
  'Commission',
  'Bonus',
  'Grand Total',
];

const columns: ColumnConfig[] = [
  { width: 8 },
  { width: 12 },
  { width: 14 },
  { width: 10 },
  { width: 18 },
  { width: 10 },
  { width: 10 },
  { width: 14 },
  { width: 14 },
  { width: 14 },
  { width: 10 },
  { width: 12 },
  { width: 11 },
  { width: 12 },
  { width: 14 },
  { width: 14 },
  { width: 14 },
  { width: 14 },
  { width: 14 },
  { width: 13 },
  { width: 10 },
  { width: 10 },
  { width: 14 },
  { width: 14 },
  { width: 12 },
  { width: 14 },
  { width: 12 },
  { width: 12 },
  { width: 12 },
  { width: 14 },
];

function colLetter(i: number): string {
  let s = '';
  let n = i;
  while (n >= 0) {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

// --- Generate data -----------------------------------------------------------

console.log(
  `\nGenerating ${DATA_ROWS.toLocaleString()} rows x ${COL_COUNT} columns...`,
);
const startTime = performance.now();

const rows: Row[] = [];

// Title rows
const titleCells: Cell[] = [
  { value: 'COMPREHENSIVE BUSINESS REPORT FY2024', style: titleStyle },
];
for (let c = 1; c < COL_COUNT; c++) titleCells.push({ value: null });
rows.push({ cells: titleCells, height: 35 });

const subtitleCells: Cell[] = [
  {
    value: `Generated: ${new Date().toISOString().split('T')[0]} — ${DATA_ROWS.toLocaleString()} records across ${departments.length} departments`,
    style: subtitleStyle,
  },
];
for (let c = 1; c < COL_COUNT; c++) subtitleCells.push({ value: null });
rows.push({ cells: subtitleCells, height: 22 });

rows.push({ cells: [] });

// Header
rows.push({
  cells: columnHeaders.map((h) => ({ value: h, style: headerStyle })),
  height: 35,
});

const dataStartExcelRow = 5;

for (let i = 0; i < DATA_ROWS; i++) {
  const isEven = i % 2 === 0;
  const ds = isEven ? evenRowStyle : dataStyle;
  const ns = isEven ? evenNumberStyle : numberDataStyle;
  const cs = isEven ? evenCurrencyStyle : currencyStyle;
  const ps = isEven ? evenPercentStyle : percentStyle;
  const dts = isEven ? evenDateStyle : dateDataStyle;

  const revenue = 1000 + Math.random() * 49000;
  const cost = revenue * (0.3 + Math.random() * 0.4);
  const profit = revenue - cost;
  const quantity = Math.floor(1 + Math.random() * 500);
  const unitPrice = revenue / quantity;
  const discountPct = Math.random() * 0.25;
  const tax = (revenue - revenue * discountPct) * 0.1;
  const netAmount = revenue - revenue * discountPct + tax;
  const budget = 5000 + Math.random() * 45000;
  const actual = budget * (0.7 + Math.random() * 0.6);
  const variance = actual - budget;
  const target = 3000 + Math.random() * 47000;
  const achievementPct = actual / target;
  const hours = 10 + Math.random() * 150;
  const rate = 20 + Math.random() * 80;
  const laborCost = hours * rate;
  const materialCost = cost * 0.4;
  const overhead = cost * 0.15;
  const totalCost = laborCost + materialCost + overhead;
  const margin = (revenue - totalCost) / revenue;
  const commission = revenue * 0.03;
  const bonus = profit > 20000 ? profit * 0.05 : 0;
  const grandTotal = netAmount - totalCost + commission + bonus;

  const month = (i % 12) + 1;
  const day = (i % 28) + 1;
  const dateStr = `2024-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;

  const r = (v: number) => Math.round(v * 100) / 100;
  const r3 = (v: number) => Math.round(v * 1000) / 1000;

  rows.push({
    cells: [
      { value: i + 1, style: ns },
      { value: dateStr, style: dts },
      { value: departments[i % departments.length], style: ds },
      { value: regions[i % regions.length], style: ds },
      { value: `Employee_${String(i + 1).padStart(5, '0')}`, style: ds },
      { value: categories[i % categories.length], style: ds },
      { value: statuses[i % statuses.length], style: ds },
      { value: r(revenue), style: cs },
      { value: r(cost), style: cs },
      { value: r(profit), style: cs },
      { value: quantity, style: ns },
      { value: r(unitPrice), style: cs },
      { value: r3(discountPct), style: ps },
      { value: r(tax), style: cs },
      { value: r(netAmount), style: cs },
      { value: r(budget), style: cs },
      { value: r(actual), style: cs },
      { value: r(variance), style: cs },
      { value: r(target), style: cs },
      { value: r3(achievementPct), style: ps },
      { value: Math.round(hours * 10) / 10, style: ns },
      { value: r(rate), style: cs },
      { value: r(laborCost), style: cs },
      { value: r(materialCost), style: cs },
      { value: r(overhead), style: cs },
      { value: r(totalCost), style: cs },
      { value: r3(margin), style: ps },
      { value: r(commission), style: cs },
      { value: r(bonus), style: cs },
      { value: r(grandTotal), style: cs },
    ],
  });
}

const genTime = (performance.now() - startTime).toFixed(0);
console.log(`      Data generated in ${genTime}ms`);

// --- Footer rows (SUM, AVERAGE, MAX, MIN) ------------------------------------

const lastDataExcelRow = dataStartExcelRow + DATA_ROWS - 1;
const numericCols = [
  7, 8, 9, 10, 13, 14, 15, 16, 17, 18, 20, 22, 23, 24, 25, 27, 28, 29,
];

rows.push({ cells: [] });

const footerConfigs = [
  { label: 'TOTAL (SUM)', fn: 'SUM', color: '1F3864' },
  { label: 'AVERAGE', fn: 'AVERAGE', color: '2E75B6' },
  { label: 'MAX', fn: 'MAX', color: '548235' },
  { label: 'MIN', fn: 'MIN', color: 'BF8F00' },
];

for (const { label, fn, color } of footerConfigs) {
  const labelStyle: CellStyle = {
    ...footerBase,
    fill: { type: 'pattern', pattern: 'solid', fgColor: color },
  };
  const valueStyle: CellStyle = {
    ...labelStyle,
    numberFormat: '#,##0.00',
  };

  const cells: Cell[] = [{ value: label, style: labelStyle }];
  for (let c = 1; c < 7; c++) cells.push({ value: null, style: labelStyle });

  for (let c = 7; c < COL_COUNT; c++) {
    if (numericCols.includes(c)) {
      const letter = colLetter(c);
      cells.push({
        value: null,
        formula: `${fn}(${letter}${dataStartExcelRow}:${letter}${lastDataExcelRow})`,
        formulaResult: 0,
        style: valueStyle,
      });
    } else {
      cells.push({ value: null, style: labelStyle });
    }
  }
  rows.push({ cells, height: 30 });
}

// --- Merge cells -------------------------------------------------------------

const mergeCells: MergeCell[] = [
  { startRow: 0, startCol: 0, endRow: 0, endCol: COL_COUNT - 1 },
  { startRow: 1, startCol: 0, endRow: 1, endCol: COL_COUNT - 1 },
  {
    startRow: rows.length - 4,
    startCol: 0,
    endRow: rows.length - 4,
    endCol: 6,
  },
  {
    startRow: rows.length - 3,
    startCol: 0,
    endRow: rows.length - 3,
    endCol: 6,
  },
  {
    startRow: rows.length - 2,
    startCol: 0,
    endRow: rows.length - 2,
    endCol: 6,
  },
  {
    startRow: rows.length - 1,
    startCol: 0,
    endRow: rows.length - 1,
    endCol: 6,
  },
];

// --- Write -------------------------------------------------------------------

console.log('\nWriting XLSX file...');
const writeStart = performance.now();

const workbook: Workbook = {
  worksheets: [
    {
      name: 'Business Report',
      columns,
      rows,
      mergeCells,
      freezePane: { row: 4, col: 1 },
    },
  ],
  creator: 'bun-spreadsheet',
};

await writeExcel(`${OUTPUT}/large-report-30x30k.xlsx`, workbook);
const writeTime = (performance.now() - writeStart).toFixed(0);
const totalTime = (performance.now() - startTime).toFixed(0);

const fileInfo = Bun.file(`${OUTPUT}/large-report-30x30k.xlsx`);
const fileSizeMB = (fileInfo.size / (1024 * 1024)).toFixed(2);

console.log('      -> output/large-report-30x30k.xlsx');
console.log(`\n${'='.repeat(60)}`);
console.log('Summary:\n');
console.log(
  `  Dimensions:  ${COL_COUNT} columns x ${DATA_ROWS.toLocaleString()} data rows`,
);
console.log(
  `  Total rows:  ${rows.length.toLocaleString()} (incl. title, header, footer)`,
);
console.log(
  `  Merge cells: ${mergeCells.length} (title, subtitle, 4 footer labels)`,
);
console.log(
  `  Formulas:    ${numericCols.length * 4} (SUM + AVG + MAX + MIN x ${numericCols.length} columns)`,
);
console.log(`  File size:   ${fileSizeMB} MB`);
console.log(`  Gen time:    ${genTime}ms`);
console.log(`  Write time:  ${writeTime}ms`);
console.log(`  Total time:  ${totalTime}ms`);
