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
  peakHeapUsed: number;
}

function nowMs(): number {
  return Bun.nanoseconds() / 1_000_000;
}

function createPeakTracker(): PeakTracker {
  const mem = process.memoryUsage();
  return {
    peakHeapUsed: mem.heapUsed,
  };
}

function samplePeak(tracker: PeakTracker) {
  const mem = process.memoryUsage();
  if (mem.heapUsed > tracker.peakHeapUsed) tracker.peakHeapUsed = mem.heapUsed;
}

function peakRssMb(): number {
  return process.resourceUsage().maxRSS / 1024;
}

const OUTPUT = './output';

mkdirSync(OUTPUT, { recursive: true });

console.log(
  `Benchmark: large-report workload (${LARGE_REPORT_COL_COUNT} columns x ${LARGE_REPORT_DATA_ROWS.toLocaleString()} rows)`,
);
console.log('='.repeat(60));

const mode = process.argv[2];
if (!mode) {
  console.log('| Mode | Total | Peak RSS | Sampled peak heapUsed | File |');
  console.log('| --- | ---: | ---: | ---: | ---: |');
  for (const childMode of ['normal', 'stream', 'chunked']) {
    const child = Bun.spawn({
      cmd: [process.execPath, import.meta.path, childMode],
      stdout: 'pipe',
      stderr: 'inherit',
    });
    const stdout = await new Response(child.stdout).text();
    if ((await child.exited) !== 0)
      throw new Error(`${childMode} benchmark failed`);
    const line = stdout
      .split('\n')
      .find((value) => value.startsWith('__RESULT__'));
    if (!line) throw new Error(`Missing result for ${childMode}`);
    const result: Result = JSON.parse(line.slice('__RESULT__'.length));
    console.log(
      `| ${childMode} | ${(result.totalMs / 1000).toFixed(2)}s | ${result.peakRss.toFixed(1)}MiB | ${result.peakHeapUsed.toFixed(1)}MiB | ${result.fileSize.toFixed(2)}MiB |`,
    );
    console.log(`__RESULT__${JSON.stringify({ mode: childMode, ...result })}`);
  }
} else if (mode === 'normal') {
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
    peakRss: peakRssMb(),
    peakHeapUsed: p1.peakHeapUsed / 1024 / 1024,
    fileSize: f1.size / 1024 / 1024,
  };
  console.log(`__RESULT__${JSON.stringify(r1)}`);
} else if (mode === 'stream') {
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
    peakRss: peakRssMb(),
    peakHeapUsed: p2.peakHeapUsed / 1024 / 1024,
    fileSize: f2.size / 1024 / 1024,
  };
  console.log(`__RESULT__${JSON.stringify(r2)}`);
} else if (mode === 'chunked') {
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
    peakRss: peakRssMb(),
    peakHeapUsed: p3.peakHeapUsed / 1024 / 1024,
    fileSize: f3.size / 1024 / 1024,
  };
  console.log(`__RESULT__${JSON.stringify(r3)}`);
} else {
  throw new Error(`Unknown benchmark mode: ${mode}`);
}
