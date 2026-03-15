// ============================================
// Benchmark: Normal vs Stream vs Chunked Stream
// 30 columns x 30,000 rows
// ============================================

import {
  type Cell,
  type CellStyle,
  type ColumnConfig,
  createChunkedExcelStream,
  createExcelStream,
  type Row,
  writeExcel,
} from '../src/index';

const DATA_ROWS = 30_000;
const COL_COUNT = 30;

const numStyle: CellStyle = {
  numberFormat: '#,##0.00',
  alignment: { horizontal: 'right' },
};
const hdrStyle: CellStyle = {
  font: { bold: true, color: 'FFFFFF' },
  fill: { type: 'pattern', pattern: 'solid', fgColor: '2F5496' },
  alignment: { horizontal: 'center' },
};

const cols: ColumnConfig[] = Array.from({ length: COL_COUNT }, () => ({
  width: 14,
}));

function nowMs(): number {
  return Bun.nanoseconds() / 1_000_000;
}

function colLetter(i: number) {
  let s = '';
  let n = i;
  while (n >= 0) {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

function makeCell(c: number, i: number): Cell {
  if (c === 0) return { value: i + 1 };
  if (c < 7) return { value: `val_${i}_${c}` };
  return {
    value: Math.round(Math.random() * 50000 * 100) / 100,
    style: numStyle,
  };
}

const headerCells: Cell[] = Array.from({ length: COL_COUNT }, (_, i) => ({
  value: `Col${i}`,
  style: hdrStyle,
}));

function makeFormulaRow(
  fn: string,
  lastRow: number,
): { cells: Cell[]; height: number } {
  const fc: Cell[] = [{ value: fn, style: hdrStyle }];
  for (let c = 1; c < 7; c++) fc.push({ value: null });
  for (let c = 7; c < COL_COUNT; c++)
    fc.push({
      value: null,
      formula: `${fn}(${colLetter(c)}5:${colLetter(c)}${lastRow})`,
    });
  return { cells: fc, height: 30 };
}

interface Result {
  genMs: number;
  writeMs: number;
  totalMs: number;
  rss: number;
  heap: number;
  fileSize: number;
}

const OUTPUT = './output';

import { mkdirSync } from 'node:fs';

mkdirSync(OUTPUT, { recursive: true });

console.log(
  `Benchmark: ${COL_COUNT} columns x ${DATA_ROWS.toLocaleString()} rows`,
);
console.log('='.repeat(60));

// --- 1. Normal Write ---------------------------------------------------------
console.log('\n[1/3] Normal write (writeExcel)');

Bun.gc(true);
await Bun.sleep(200);

const m1b = process.memoryUsage();
const t1s = nowMs();

const rows: Row[] = [];
rows.push({ cells: [{ value: 'REPORT' }] });
rows.push({ cells: [{ value: 'Subtitle' }] });
rows.push({ cells: [] });
rows.push({ cells: headerCells });

const t1g = nowMs();
for (let i = 0; i < DATA_ROWS; i++) {
  const cells: Cell[] = [];
  for (let c = 0; c < COL_COUNT; c++) cells.push(makeCell(c, i));
  rows.push({ cells });
}
const t1gd = nowMs();
for (const fn of ['SUM', 'AVERAGE', 'MAX', 'MIN'])
  rows.push(makeFormulaRow(fn, DATA_ROWS + 4));

const t1w = nowMs();
await writeExcel(`${OUTPUT}/bench-normal.xlsx`, {
  worksheets: [
    {
      name: 'Report',
      rows,
      columns: cols,
      freezePane: { row: 4, col: 1 },
    },
  ],
});
const t1d = nowMs();
const m1a = process.memoryUsage();
const f1 = Bun.file(`${OUTPUT}/bench-normal.xlsx`);

const r1: Result = {
  genMs: t1gd - t1g,
  writeMs: t1d - t1w,
  totalMs: t1d - t1s,
  rss: (m1a.rss - m1b.rss) / 1024 / 1024,
  heap: (m1a.heapUsed - m1b.heapUsed) / 1024 / 1024,
  fileSize: f1.size / 1024 / 1024,
};
console.log(
  `      Gen: ${r1.genMs.toFixed(0)}ms | Write: ${r1.writeMs.toFixed(0)}ms | Total: ${r1.totalMs.toFixed(0)}ms`,
);
console.log(
  `      RSS: +${r1.rss.toFixed(1)}MB | Heap: +${r1.heap.toFixed(1)}MB | File: ${r1.fileSize.toFixed(2)}MB`,
);

// --- 2. Stream Write ---------------------------------------------------------
rows.length = 0;
Bun.gc(true);
await Bun.sleep(500);

console.log('\n[2/3] Stream write (createExcelStream)');

const m2b = process.memoryUsage();
const t2s = nowMs();

const stream = createExcelStream(`${OUTPUT}/bench-stream.xlsx`, {
  sheetName: 'Report',
  columns: cols,
  freezePane: { row: 1, col: 1 },
});
stream.writeRow({ cells: headerCells, height: 30 });

const t2g = nowMs();
for (let i = 0; i < DATA_ROWS; i++) {
  const cells: Cell[] = [];
  for (let c = 0; c < COL_COUNT; c++) cells.push(makeCell(c, i));
  stream.writeRow({ cells });
}
const t2gd = nowMs();
for (const fn of ['SUM', 'AVERAGE', 'MAX', 'MIN'])
  stream.writeRow(makeFormulaRow(fn, DATA_ROWS + 1));

const t2w = nowMs();
await stream.end();
const t2d = nowMs();
const m2a = process.memoryUsage();
const f2 = Bun.file(`${OUTPUT}/bench-stream.xlsx`);

const r2: Result = {
  genMs: t2gd - t2g,
  writeMs: t2d - t2w,
  totalMs: t2d - t2s,
  rss: (m2a.rss - m2b.rss) / 1024 / 1024,
  heap: (m2a.heapUsed - m2b.heapUsed) / 1024 / 1024,
  fileSize: f2.size / 1024 / 1024,
};
console.log(
  `      Gen: ${r2.genMs.toFixed(0)}ms | Write: ${r2.writeMs.toFixed(0)}ms | Total: ${r2.totalMs.toFixed(0)}ms`,
);
console.log(
  `      RSS: +${r2.rss.toFixed(1)}MB | Heap: +${r2.heap.toFixed(1)}MB | File: ${r2.fileSize.toFixed(2)}MB`,
);

// --- 3. Chunked Stream Write -------------------------------------------------
Bun.gc(true);
await Bun.sleep(500);

console.log('\n[3/3] Chunked stream write (createChunkedExcelStream)');

const m3b = process.memoryUsage();
const t3s = nowMs();

const chunked = createChunkedExcelStream(`${OUTPUT}/bench-chunked.xlsx`, {
  sheetName: 'Report',
  columns: cols,
  freezePane: { row: 1, col: 1 },
});
chunked.writeRow({ cells: headerCells, height: 30 });

const t3g = nowMs();
for (let i = 0; i < DATA_ROWS; i++) {
  const cells: Cell[] = [];
  for (let c = 0; c < COL_COUNT; c++) cells.push(makeCell(c, i));
  chunked.writeRow({ cells });
}
const t3gd = nowMs();
for (const fn of ['SUM', 'AVERAGE', 'MAX', 'MIN'])
  chunked.writeRow(makeFormulaRow(fn, DATA_ROWS + 1));

const t3w = nowMs();
await chunked.end();
const t3d = nowMs();
const m3a = process.memoryUsage();
const f3 = Bun.file(`${OUTPUT}/bench-chunked.xlsx`);

const r3: Result = {
  genMs: t3gd - t3g,
  writeMs: t3d - t3w,
  totalMs: t3d - t3s,
  rss: (m3a.rss - m3b.rss) / 1024 / 1024,
  heap: (m3a.heapUsed - m3b.heapUsed) / 1024 / 1024,
  fileSize: f3.size / 1024 / 1024,
};
console.log(
  `      Gen: ${r3.genMs.toFixed(0)}ms | Write: ${r3.writeMs.toFixed(0)}ms | Total: ${r3.totalMs.toFixed(0)}ms`,
);
console.log(
  `      RSS: +${r3.rss.toFixed(1)}MB | Heap: +${r3.heap.toFixed(1)}MB | File: ${r3.fileSize.toFixed(2)}MB`,
);

// --- Comparison Table --------------------------------------------------------
console.log(`\n${'='.repeat(60)}`);
console.log('Comparison\n');
console.log('                  Normal       Stream      Chunked');
console.log(
  `  Gen:        ${String(r1.genMs.toFixed(0)).padStart(8)}ms  ${String(r2.genMs.toFixed(0)).padStart(8)}ms  ${String(r3.genMs.toFixed(0)).padStart(8)}ms`,
);
console.log(
  `  Write:      ${String(r1.writeMs.toFixed(0)).padStart(8)}ms  ${String(r2.writeMs.toFixed(0)).padStart(8)}ms  ${String(r3.writeMs.toFixed(0)).padStart(8)}ms`,
);
console.log(
  `  Total:      ${String(r1.totalMs.toFixed(0)).padStart(8)}ms  ${String(r2.totalMs.toFixed(0)).padStart(8)}ms  ${String(r3.totalMs.toFixed(0)).padStart(8)}ms`,
);
console.log(
  `  RSS delta:  ${String(r1.rss.toFixed(1)).padStart(7)}MB   ${String(r2.rss.toFixed(1)).padStart(7)}MB   ${String(r3.rss.toFixed(1)).padStart(7)}MB`,
);
console.log(
  `  Heap delta: ${String(r1.heap.toFixed(1)).padStart(7)}MB   ${String(r2.heap.toFixed(1)).padStart(7)}MB   ${String(r3.heap.toFixed(1)).padStart(7)}MB`,
);
console.log(
  `  File size:  ${String(r1.fileSize.toFixed(2)).padStart(7)}MB   ${String(r2.fileSize.toFixed(2)).padStart(7)}MB   ${String(r3.fileSize.toFixed(2)).padStart(7)}MB`,
);
console.log('');
