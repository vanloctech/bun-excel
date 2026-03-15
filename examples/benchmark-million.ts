// ============================================
// Benchmark: 1,000,000-row Excel export
// Stream vs Chunked Stream
// ============================================

import { mkdirSync } from 'node:fs';
import {
  type Cell,
  type CellStyle,
  type ColumnConfig,
  createChunkedExcelStream,
  createExcelStream,
  type StreamWriter,
} from '../src/index';

type BenchmarkMode = 'stream' | 'chunked';

interface BenchmarkResult {
  mode: BenchmarkMode;
  rows: number;
  columns: number;
  rowWriteMs: number;
  finalizeMs: number;
  totalMs: number;
  rowsPerSecond: number;
  peakRssMb: number;
  peakHeapUsedMb: number;
  endRssMb: number;
  endHeapUsedMb: number;
  fileSizeMb: number;
  outputPath: string;
}

interface BenchmarkSummary {
  generatedAt: string;
  bunVersion: string;
  platform: string;
  arch: string;
  cpuModel?: string;
  totalMemoryGb?: number;
  rows: number;
  columns: number;
  results: BenchmarkResult[];
}

const DATA_ROWS = Number(Bun.env.BENCH_ROWS ?? '1000000');
const COL_COUNT = Number(Bun.env.BENCH_COLS ?? '10');
const SAMPLE_INTERVAL_MS = Number(Bun.env.BENCH_SAMPLE_MS ?? '100');
const LOG_EVERY_ROWS = Number(Bun.env.BENCH_LOG_EVERY ?? '100000');
const OUTPUT = './output';
const RESULT_MARKER = '__BENCHMARK_RESULT__';

const hdrStyle: CellStyle = {
  font: { bold: true, color: 'FFFFFF' },
  fill: { type: 'pattern', pattern: 'solid', fgColor: '2F5496' },
  alignment: { horizontal: 'center' },
};

const amountStyle: CellStyle = {
  numberFormat: '#,##0.00',
  alignment: { horizontal: 'right' },
};

const integerStyle: CellStyle = {
  numberFormat: '0',
  alignment: { horizontal: 'right' },
};

const columns: ColumnConfig[] = [
  { width: 10 },
  { width: 18 },
  { width: 14 },
  { width: 12 },
  { width: 12 },
  { width: 10 },
  { width: 14 },
  { width: 12 },
  { width: 14 },
  { width: 10 },
];

const headerCells: Cell[] = [
  'Row ID',
  'Order Code',
  'Store',
  'Channel',
  'Status',
  'Qty',
  'Subtotal',
  'Tax',
  'Total',
  'Paid',
].map((value) => ({ value, style: hdrStyle }));

const channels = ['POS', 'APP', 'WEB', 'KIOSK'] as const;
const statuses = ['PAID', 'PENDING', 'REFUND', 'VOID', 'FAILED'] as const;

function formatMb(bytes: number): number {
  return bytes / 1024 / 1024;
}

function formatGb(bytes: number): number {
  return bytes / 1024 / 1024 / 1024;
}

function takeMemorySnapshot() {
  const usage = process.memoryUsage();
  return {
    rssMb: formatMb(usage.rss),
    heapUsedMb: formatMb(usage.heapUsed),
  };
}

function createWriter(mode: BenchmarkMode, outputPath: string): StreamWriter {
  const options = {
    sheetName: 'Benchmark',
    columns,
    freezePane: { row: 1, col: 1 },
  };

  if (mode === 'chunked') {
    return createChunkedExcelStream(outputPath, options);
  }

  return createExcelStream(outputPath, options);
}

function makeCells(rowIndex: number): Cell[] {
  const displayIndex = rowIndex + 1;
  const qty = (displayIndex % 7) + 1;
  const subtotal = Number((((displayIndex % 1000) + 25) * 1.37).toFixed(2));
  const tax = Number((subtotal * 0.08).toFixed(2));
  const total = Number((subtotal + tax).toFixed(2));

  return [
    { value: displayIndex },
    { value: `ORD-${displayIndex}` },
    { value: `Store ${(displayIndex % 48) + 1}` },
    { value: channels[displayIndex % channels.length] },
    { value: statuses[displayIndex % statuses.length] },
    { value: qty, style: integerStyle },
    { value: subtotal, style: amountStyle },
    { value: tax, style: amountStyle },
    { value: total, style: amountStyle },
    { value: displayIndex % 2 === 0 },
  ];
}

function getCpuModel(): string | undefined {
  if (process.platform === 'darwin') {
    const result = Bun.spawnSync({
      cmd: ['sysctl', '-n', 'machdep.cpu.brand_string'],
      stdout: 'pipe',
      stderr: 'ignore',
    });
    if (result.exitCode === 0) {
      return result.stdout.toString().trim();
    }
  }

  return undefined;
}

function getTotalMemoryGb(): number | undefined {
  if (process.platform === 'darwin') {
    const result = Bun.spawnSync({
      cmd: ['sysctl', '-n', 'hw.memsize'],
      stdout: 'pipe',
      stderr: 'ignore',
    });
    if (result.exitCode === 0) {
      const bytes = Number(result.stdout.toString().trim());
      if (Number.isFinite(bytes) && bytes > 0) {
        return Number(formatGb(bytes).toFixed(1));
      }
    }
  }

  return undefined;
}

async function runMode(mode: BenchmarkMode): Promise<BenchmarkResult> {
  const outputPath = `${OUTPUT}/benchmark-${mode}-1m.xlsx`;
  const writer = createWriter(mode, outputPath);
  writer.writeRow({ cells: headerCells, height: 24 });

  Bun.gc(true);
  await Bun.sleep(200);

  let peak = takeMemorySnapshot();
  const samplePeak = () => {
    const current = takeMemorySnapshot();
    peak = {
      rssMb: Math.max(peak.rssMb, current.rssMb),
      heapUsedMb: Math.max(peak.heapUsedMb, current.heapUsedMb),
    };
    return current;
  };

  const interval = setInterval(samplePeak, SAMPLE_INTERVAL_MS);

  const startedAt = performance.now();
  for (let rowIndex = 0; rowIndex < DATA_ROWS; rowIndex++) {
    writer.writeRow({ cells: makeCells(rowIndex) });

    const rowsWritten = rowIndex + 1;
    if (rowsWritten % LOG_EVERY_ROWS === 0) {
      const current = samplePeak();
      console.log(
        `[${mode}] ${rowsWritten.toLocaleString()} / ${DATA_ROWS.toLocaleString()} rows | rss=${current.rssMb.toFixed(1)}MB heapUsed=${current.heapUsedMb.toFixed(1)}MB`,
      );
    }
  }

  const rowWriteDoneAt = performance.now();
  console.log(`[${mode}] Finalizing workbook...`);
  await writer.end();
  const finishedAt = performance.now();

  clearInterval(interval);
  const ended = samplePeak();
  const file = Bun.file(outputPath);

  return {
    mode,
    rows: DATA_ROWS,
    columns: COL_COUNT,
    rowWriteMs: rowWriteDoneAt - startedAt,
    finalizeMs: finishedAt - rowWriteDoneAt,
    totalMs: finishedAt - startedAt,
    rowsPerSecond: (DATA_ROWS / (finishedAt - startedAt)) * 1000,
    peakRssMb: peak.rssMb,
    peakHeapUsedMb: peak.heapUsedMb,
    endRssMb: ended.rssMb,
    endHeapUsedMb: ended.heapUsedMb,
    fileSizeMb: Number((file.size / 1024 / 1024).toFixed(2)),
    outputPath,
  };
}

async function runChild(): Promise<void> {
  const mode = Bun.env.BENCH_MODE as BenchmarkMode | undefined;
  if (mode !== 'stream' && mode !== 'chunked') {
    throw new Error('BENCH_MODE must be "stream" or "chunked"');
  }

  console.log(
    `[${mode}] Starting benchmark for ${DATA_ROWS.toLocaleString()} rows x ${COL_COUNT} columns`,
  );
  const result = await runMode(mode);
  console.log(
    `[${mode}] Done in ${(result.totalMs / 1000).toFixed(1)}s | peak RSS ${result.peakRssMb.toFixed(1)}MB | peak heapUsed ${result.peakHeapUsedMb.toFixed(1)}MB | file ${result.fileSizeMb.toFixed(2)}MB`,
  );
  console.log(`${RESULT_MARKER}${JSON.stringify(result)}`);
}

async function runParent(): Promise<void> {
  mkdirSync(OUTPUT, { recursive: true });

  const summary: BenchmarkSummary = {
    generatedAt: new Date().toISOString(),
    bunVersion: Bun.version,
    platform: process.platform,
    arch: process.arch,
    cpuModel: getCpuModel(),
    totalMemoryGb: getTotalMemoryGb(),
    rows: DATA_ROWS,
    columns: COL_COUNT,
    results: [],
  };

  console.log(
    `Benchmark: ${DATA_ROWS.toLocaleString()} rows x ${COL_COUNT} columns`,
  );
  console.log(
    `Bun ${summary.bunVersion} | ${summary.platform} ${summary.arch}`,
  );
  if (summary.cpuModel) {
    console.log(summary.cpuModel);
  }
  if (summary.totalMemoryGb) {
    console.log(`${summary.totalMemoryGb.toFixed(1)} GB RAM`);
  }
  console.log('='.repeat(72));

  for (const mode of ['stream', 'chunked'] as const) {
    console.log(`\nRunning ${mode} benchmark in a fresh Bun process...\n`);
    const child = Bun.spawn({
      cmd: ['bun', 'run', 'examples/benchmark-million.ts'],
      cwd: process.cwd(),
      env: {
        ...Bun.env,
        BENCH_CHILD: '1',
        BENCH_MODE: mode,
        BENCH_ROWS: String(DATA_ROWS),
        BENCH_COLS: String(COL_COUNT),
        BENCH_SAMPLE_MS: String(SAMPLE_INTERVAL_MS),
        BENCH_LOG_EVERY: String(LOG_EVERY_ROWS),
      },
      stdout: 'pipe',
      stderr: 'inherit',
    });

    const decoder = new TextDecoder();
    const reader = child.stdout.getReader();
    let stdout = '';

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;

      const text = decoder.decode(value, { stream: true });
      stdout += text;

      for (const line of text.split('\n')) {
        if (!line || line.startsWith(RESULT_MARKER)) continue;
        console.log(line);
      }
    }

    const exitCode = await child.exited;
    if (exitCode !== 0) {
      throw new Error(`${mode} benchmark failed with exit code ${exitCode}`);
    }

    const resultLine = stdout
      .split('\n')
      .find((line) => line.startsWith(RESULT_MARKER));
    if (!resultLine) {
      throw new Error(`Missing benchmark result for mode "${mode}"`);
    }

    summary.results.push(
      JSON.parse(resultLine.slice(RESULT_MARKER.length)) as BenchmarkResult,
    );
  }

  const resultPath = `${OUTPUT}/benchmark-1m-results.json`;
  await Bun.write(resultPath, JSON.stringify(summary, null, 2));

  console.log(`\n${'='.repeat(72)}`);
  console.log('Summary\n');
  console.log(
    '| Mode | Total | Finalize | Rows/sec | Peak RSS | Peak heapUsed | File |',
  );
  console.log('| --- | ---: | ---: | ---: | ---: | ---: | ---: |');
  for (const result of summary.results) {
    console.log(
      `| ${result.mode} | ${(result.totalMs / 1000).toFixed(1)}s | ${(result.finalizeMs / 1000).toFixed(1)}s | ${Math.round(result.rowsPerSecond).toLocaleString()} | ${result.peakRssMb.toFixed(1)}MB | ${result.peakHeapUsedMb.toFixed(1)}MB | ${result.fileSizeMb.toFixed(2)}MB |`,
    );
  }
  console.log(`\nSaved raw results to ${resultPath}`);
}

if (Bun.env.BENCH_CHILD === '1') {
  await runChild();
} else {
  await runParent();
}
