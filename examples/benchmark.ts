// ============================================
// Benchmark: realistic large-report workload
// 30 columns x 30,000 rows
// ============================================

import { mkdirSync } from 'node:fs';
import {
  createChunkedExcelStream,
  createExcelStream,
  writeExcel,
} from '../src';
import {
  buildLargeReportWorkbook,
  createLargeReportDataRow,
  createLargeReportFooterRows,
  createLargeReportMergeCells,
  createLargeReportPreludeRows,
  LARGE_REPORT_COL_COUNT,
  LARGE_REPORT_DATA_ROWS,
  LARGE_REPORT_FREEZE_PANE,
  LARGE_REPORT_SHEET_NAME,
  largeReportColumns,
} from './large-report-workload';

interface Result {
  genMs: number;
  writeMs: number;
  totalMs: number;
  peakRss: number;
  peakHeapUsed: number;
  fileSize: number;
}

interface PeakTracker {
  baselineRss: number;
  baselineHeapUsed: number;
  peakRss: number;
  peakHeapUsed: number;
}

function nowMs(): number {
  return Bun.nanoseconds() / 1_000_000;
}

function createPeakTracker(): PeakTracker {
  const mem = process.memoryUsage();
  return {
    baselineRss: mem.rss,
    baselineHeapUsed: mem.heapUsed,
    peakRss: mem.rss,
    peakHeapUsed: mem.heapUsed,
  };
}

function samplePeak(tracker: PeakTracker) {
  const mem = process.memoryUsage();
  if (mem.rss > tracker.peakRss) tracker.peakRss = mem.rss;
  if (mem.heapUsed > tracker.peakHeapUsed) tracker.peakHeapUsed = mem.heapUsed;
}

function peakDeltaMb(peak: number, baseline: number): number {
  return (peak - baseline) / 1024 / 1024;
}

const OUTPUT = './output';

mkdirSync(OUTPUT, { recursive: true });

console.log(
  `Benchmark: large-report workload (${LARGE_REPORT_COL_COUNT} columns x ${LARGE_REPORT_DATA_ROWS.toLocaleString()} rows)`,
);
console.log('='.repeat(60));

// --- 1. Normal Write ---------------------------------------------------------
console.log('\n[1/3] Normal write (writeExcel)');

Bun.gc(true);
await Bun.sleep(200);

const p1 = createPeakTracker();
const t1s = nowMs();
const t1g = nowMs();
const normalWorkbook = buildLargeReportWorkbook();
samplePeak(p1);
const t1gd = nowMs();

const t1w = nowMs();
await writeExcel(`${OUTPUT}/bench-normal.xlsx`, normalWorkbook);
const t1d = nowMs();
samplePeak(p1);
const f1 = Bun.file(`${OUTPUT}/bench-normal.xlsx`);

const r1: Result = {
  genMs: t1gd - t1g,
  writeMs: t1d - t1w,
  totalMs: t1d - t1s,
  peakRss: peakDeltaMb(p1.peakRss, p1.baselineRss),
  peakHeapUsed: peakDeltaMb(p1.peakHeapUsed, p1.baselineHeapUsed),
  fileSize: f1.size / 1024 / 1024,
};
console.log(
  `      Gen: ${r1.genMs.toFixed(0)}ms | Write: ${r1.writeMs.toFixed(0)}ms | Total: ${r1.totalMs.toFixed(0)}ms`,
);
console.log(
  `      Peak RSS delta: +${r1.peakRss.toFixed(1)}MB | Peak heapUsed delta: +${r1.peakHeapUsed.toFixed(1)}MB | File: ${r1.fileSize.toFixed(2)}MB`,
);

// --- 2. Stream Write ---------------------------------------------------------
Bun.gc(true);
await Bun.sleep(500);

console.log('\n[2/3] Stream write (createExcelStream)');

const p2 = createPeakTracker();
const t2s = nowMs();

const stream = createExcelStream(`${OUTPUT}/bench-stream.xlsx`, {
  sheetName: LARGE_REPORT_SHEET_NAME,
  columns: largeReportColumns,
  freezePane: LARGE_REPORT_FREEZE_PANE,
  mergeCells: createLargeReportMergeCells(),
});

const t2g = nowMs();
for (const row of createLargeReportPreludeRows()) stream.writeRow(row);
for (let i = 0; i < LARGE_REPORT_DATA_ROWS; i++) {
  stream.writeRow(createLargeReportDataRow(i));
  if ((i + 1) % 250 === 0) samplePeak(p2);
}
for (const row of createLargeReportFooterRows()) stream.writeRow(row);
samplePeak(p2);
const t2gd = nowMs();

const t2w = nowMs();
await stream.end();
const t2d = nowMs();
samplePeak(p2);
const f2 = Bun.file(`${OUTPUT}/bench-stream.xlsx`);

const r2: Result = {
  genMs: t2gd - t2g,
  writeMs: t2d - t2w,
  totalMs: t2d - t2s,
  peakRss: peakDeltaMb(p2.peakRss, p2.baselineRss),
  peakHeapUsed: peakDeltaMb(p2.peakHeapUsed, p2.baselineHeapUsed),
  fileSize: f2.size / 1024 / 1024,
};
console.log(
  `      Gen: ${r2.genMs.toFixed(0)}ms | Write: ${r2.writeMs.toFixed(0)}ms | Total: ${r2.totalMs.toFixed(0)}ms`,
);
console.log(
  `      Peak RSS delta: +${r2.peakRss.toFixed(1)}MB | Peak heapUsed delta: +${r2.peakHeapUsed.toFixed(1)}MB | File: ${r2.fileSize.toFixed(2)}MB`,
);

// --- 3. Chunked Stream Write -------------------------------------------------
Bun.gc(true);
await Bun.sleep(500);

console.log('\n[3/3] Chunked stream write (createChunkedExcelStream)');

const p3 = createPeakTracker();
const t3s = nowMs();

const chunked = createChunkedExcelStream(`${OUTPUT}/bench-chunked.xlsx`, {
  sheetName: LARGE_REPORT_SHEET_NAME,
  columns: largeReportColumns,
  freezePane: LARGE_REPORT_FREEZE_PANE,
  mergeCells: createLargeReportMergeCells(),
});

const t3g = nowMs();
for (const row of createLargeReportPreludeRows()) chunked.writeRow(row);
for (let i = 0; i < LARGE_REPORT_DATA_ROWS; i++) {
  chunked.writeRow(createLargeReportDataRow(i));
  if ((i + 1) % 250 === 0) samplePeak(p3);
}
for (const row of createLargeReportFooterRows()) chunked.writeRow(row);
samplePeak(p3);
const t3gd = nowMs();

const t3w = nowMs();
await chunked.end();
const t3d = nowMs();
samplePeak(p3);
const f3 = Bun.file(`${OUTPUT}/bench-chunked.xlsx`);

const r3: Result = {
  genMs: t3gd - t3g,
  writeMs: t3d - t3w,
  totalMs: t3d - t3s,
  peakRss: peakDeltaMb(p3.peakRss, p3.baselineRss),
  peakHeapUsed: peakDeltaMb(p3.peakHeapUsed, p3.baselineHeapUsed),
  fileSize: f3.size / 1024 / 1024,
};
console.log(
  `      Gen: ${r3.genMs.toFixed(0)}ms | Write: ${r3.writeMs.toFixed(0)}ms | Total: ${r3.totalMs.toFixed(0)}ms`,
);
console.log(
  `      Peak RSS delta: +${r3.peakRss.toFixed(1)}MB | Peak heapUsed delta: +${r3.peakHeapUsed.toFixed(1)}MB | File: ${r3.fileSize.toFixed(2)}MB`,
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
  `  Peak RSS Δ: ${String(r1.peakRss.toFixed(1)).padStart(7)}MB   ${String(r2.peakRss.toFixed(1)).padStart(7)}MB   ${String(r3.peakRss.toFixed(1)).padStart(7)}MB`,
);
console.log(
  `  Peak heap Δ:${String(r1.peakHeapUsed.toFixed(1)).padStart(8)}MB   ${String(r2.peakHeapUsed.toFixed(1)).padStart(7)}MB   ${String(r3.peakHeapUsed.toFixed(1)).padStart(7)}MB`,
);
console.log(
  `  File size:  ${String(r1.fileSize.toFixed(2)).padStart(7)}MB   ${String(r2.fileSize.toFixed(2)).padStart(7)}MB   ${String(r3.fileSize.toFixed(2)).padStart(7)}MB`,
);
console.log('');
