import type {
  Cell,
  CellStyle,
  ColumnConfig,
  MergeCell,
  Row,
  Workbook,
} from '../src';

export const LARGE_REPORT_COL_COUNT = 30;
export const LARGE_REPORT_DATA_ROWS = 30_000;
export const LARGE_REPORT_SHEET_NAME = 'Business Report';
export const LARGE_REPORT_OUTPUT = './output/large-report-30x30k.xlsx';
export const LARGE_REPORT_FREEZE_PANE = { row: 4, col: 1 } as const;

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

export const largeReportColumns: ColumnConfig[] = [
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

export const largeReportNumericCols = [
  7, 8, 9, 10, 13, 14, 15, 16, 17, 18, 20, 22, 23, 24, 25, 27, 28, 29,
];

const footerConfigs = [
  { label: 'TOTAL (SUM)', fn: 'SUM', color: '1F3864' },
  { label: 'AVERAGE', fn: 'AVERAGE', color: '2E75B6' },
  { label: 'MAX', fn: 'MAX', color: '548235' },
  { label: 'MIN', fn: 'MIN', color: 'BF8F00' },
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

function round2(v: number): number {
  return Math.round(v * 100) / 100;
}

function round3(v: number): number {
  return Math.round(v * 1000) / 1000;
}

export function createLargeReportPreludeRows(): Row[] {
  const rows: Row[] = [];

  const titleCells: Cell[] = [
    { value: 'COMPREHENSIVE BUSINESS REPORT FY2024', style: titleStyle },
  ];
  for (let c = 1; c < LARGE_REPORT_COL_COUNT; c++) {
    titleCells.push({ value: null });
  }
  rows.push({ cells: titleCells, height: 35 });

  const subtitleCells: Cell[] = [
    {
      value: `Generated: ${new Date().toISOString().split('T')[0]} - ${LARGE_REPORT_DATA_ROWS.toLocaleString()} records across ${departments.length} departments`,
      style: subtitleStyle,
    },
  ];
  for (let c = 1; c < LARGE_REPORT_COL_COUNT; c++) {
    subtitleCells.push({ value: null });
  }
  rows.push({ cells: subtitleCells, height: 22 });

  rows.push({ cells: [] });
  rows.push({
    cells: columnHeaders.map((h) => ({ value: h, style: headerStyle })),
    height: 35,
  });

  return rows;
}

export function createLargeReportDataRow(i: number): Row {
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

  return {
    cells: [
      { value: i + 1, style: ns },
      { value: dateStr, style: dts },
      { value: departments[i % departments.length], style: ds },
      { value: regions[i % regions.length], style: ds },
      { value: `Employee_${String(i + 1).padStart(5, '0')}`, style: ds },
      { value: categories[i % categories.length], style: ds },
      { value: statuses[i % statuses.length], style: ds },
      { value: round2(revenue), style: cs },
      { value: round2(cost), style: cs },
      { value: round2(profit), style: cs },
      { value: quantity, style: ns },
      { value: round2(unitPrice), style: cs },
      { value: round3(discountPct), style: ps },
      { value: round2(tax), style: cs },
      { value: round2(netAmount), style: cs },
      { value: round2(budget), style: cs },
      { value: round2(actual), style: cs },
      { value: round2(variance), style: cs },
      { value: round2(target), style: cs },
      { value: round3(achievementPct), style: ps },
      { value: Math.round(hours * 10) / 10, style: ns },
      { value: round2(rate), style: cs },
      { value: round2(laborCost), style: cs },
      { value: round2(materialCost), style: cs },
      { value: round2(overhead), style: cs },
      { value: round2(totalCost), style: cs },
      { value: round3(margin), style: ps },
      { value: round2(commission), style: cs },
      { value: round2(bonus), style: cs },
      { value: round2(grandTotal), style: cs },
    ],
  };
}

export function createLargeReportFooterRows(): Row[] {
  const lastDataExcelRow = 5 + LARGE_REPORT_DATA_ROWS - 1;
  const rows: Row[] = [{ cells: [] }];

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

    for (let c = 7; c < LARGE_REPORT_COL_COUNT; c++) {
      if (largeReportNumericCols.includes(c)) {
        const letter = colLetter(c);
        cells.push({
          value: null,
          formula: `${fn}(${letter}5:${letter}${lastDataExcelRow})`,
          formulaResult: 0,
          style: valueStyle,
        });
      } else {
        cells.push({ value: null, style: labelStyle });
      }
    }

    rows.push({ cells, height: 30 });
  }

  return rows;
}

export function createLargeReportMergeCells(): MergeCell[] {
  const footerStartRow = 4 + LARGE_REPORT_DATA_ROWS + 1;
  return [
    {
      startRow: 0,
      startCol: 0,
      endRow: 0,
      endCol: LARGE_REPORT_COL_COUNT - 1,
    },
    {
      startRow: 1,
      startCol: 0,
      endRow: 1,
      endCol: LARGE_REPORT_COL_COUNT - 1,
    },
    {
      startRow: footerStartRow,
      startCol: 0,
      endRow: footerStartRow,
      endCol: 6,
    },
    {
      startRow: footerStartRow + 1,
      startCol: 0,
      endRow: footerStartRow + 1,
      endCol: 6,
    },
    {
      startRow: footerStartRow + 2,
      startCol: 0,
      endRow: footerStartRow + 2,
      endCol: 6,
    },
    {
      startRow: footerStartRow + 3,
      startCol: 0,
      endRow: footerStartRow + 3,
      endCol: 6,
    },
  ];
}

export function buildLargeReportRows(): Row[] {
  const rows = createLargeReportPreludeRows();
  for (let i = 0; i < LARGE_REPORT_DATA_ROWS; i++) {
    rows.push(createLargeReportDataRow(i));
  }
  rows.push(...createLargeReportFooterRows());
  return rows;
}

export function buildLargeReportWorkbook(): Workbook {
  return {
    worksheets: [
      {
        name: LARGE_REPORT_SHEET_NAME,
        columns: largeReportColumns,
        rows: buildLargeReportRows(),
        mergeCells: createLargeReportMergeCells(),
        freezePane: LARGE_REPORT_FREEZE_PANE,
      },
    ],
    creator: 'bun-spreadsheet',
  };
}
