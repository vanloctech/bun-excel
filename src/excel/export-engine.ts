import {
  describeFileSource,
  getRuntimeFileSize,
  toReadableFile,
} from '../runtime-io';
import type { CellValue, FileTarget, Row, Workbook } from '../types';
import {
  type ChunkedExcelStreamOptions,
  createChunkedExcelStream,
} from './xlsx-chunked-stream-writer';
import {
  createExcelStream,
  createMultiSheetExcelStream,
  type ExcelStreamOptions,
} from './xlsx-stream-writer';
import { buildExcelBuffer, writeExcel } from './xlsx-writer';

export interface ExcelExportMemorySnapshot {
  rssBytes: number;
  heapUsedBytes: number;
  heapTotalBytes: number;
  externalBytes: number;
  arrayBuffersBytes: number;
}

export interface ExcelExportProgress {
  stage: 'writing' | 'finalizing' | 'completed' | 'aborted';
  mode: 'stream' | 'chunked' | 'multi-sheet';
  rowsWritten: number;
  elapsedMs: number;
  target: string;
  memory: ExcelExportMemorySnapshot;
  sheetName?: string;
  sheetIndex?: number;
}

export interface ExcelExportDiagnostics {
  mode: 'stream' | 'chunked' | 'multi-sheet';
  target: string;
  rowsWritten: number;
  startedAt: Date;
  finishedAt: Date;
  durationMs: number;
  outputSizeBytes: number;
  sheetCount?: number;
  memory: {
    baseline: ExcelExportMemorySnapshot;
    peak: ExcelExportMemorySnapshot;
    end: ExcelExportMemorySnapshot;
  };
}

export interface ExportExcelRowsOptions extends ExcelStreamOptions {
  target: FileTarget;
  rows: Iterable<Row | CellValue[]> | AsyncIterable<Row | CellValue[]>;
  mode?: 'stream' | 'chunked';
  signal?: AbortSignal;
  progressIntervalRows?: number;
  onProgress?: (progress: ExcelExportProgress) => void | Promise<void>;
  validateOutput?: boolean;
}

export interface ExportExcelSheetRows {
  name: string;
  rows: Iterable<Row | CellValue[]> | AsyncIterable<Row | CellValue[]>;
  options?: ExcelStreamOptions;
}

export interface ExportMultiSheetExcelOptions {
  target: FileTarget;
  sheets: ExportExcelSheetRows[];
  creator?: Workbook['creator'];
  created?: Workbook['created'];
  modified?: Workbook['modified'];
  compress?: boolean;
  definedNames?: Workbook['definedNames'];
  views?: Workbook['views'];
  signal?: AbortSignal;
  progressIntervalRows?: number;
  onProgress?: (progress: ExcelExportProgress) => void | Promise<void>;
  validateOutput?: boolean;
}

export interface ExcelResponseOptions {
  filename?: string;
  headers?: HeadersInit;
}

function takeMemorySnapshot(): ExcelExportMemorySnapshot {
  const usage = process.memoryUsage();
  return {
    rssBytes: usage.rss,
    heapUsedBytes: usage.heapUsed,
    heapTotalBytes: usage.heapTotal,
    externalBytes: usage.external,
    arrayBuffersBytes: usage.arrayBuffers,
  };
}

function updatePeakSnapshot(
  peak: ExcelExportMemorySnapshot,
  current: ExcelExportMemorySnapshot,
): void {
  peak.rssBytes = Math.max(peak.rssBytes, current.rssBytes);
  peak.heapUsedBytes = Math.max(peak.heapUsedBytes, current.heapUsedBytes);
  peak.heapTotalBytes = Math.max(peak.heapTotalBytes, current.heapTotalBytes);
  peak.externalBytes = Math.max(peak.externalBytes, current.externalBytes);
  peak.arrayBuffersBytes = Math.max(
    peak.arrayBuffersBytes,
    current.arrayBuffersBytes,
  );
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (!signal?.aborted) return;

  const reason =
    signal.reason instanceof Error
      ? signal.reason
      : new Error(String(signal.reason ?? 'The operation was aborted'));
  reason.name = 'AbortError';
  throw reason;
}

async function* toAsyncRowIterable(
  rows: Iterable<Row | CellValue[]> | AsyncIterable<Row | CellValue[]>,
): AsyncGenerator<Row | CellValue[]> {
  if (Symbol.asyncIterator in rows) {
    for await (const row of rows as AsyncIterable<Row | CellValue[]>) {
      yield row;
    }
    return;
  }

  for (const row of rows as Iterable<Row | CellValue[]>) {
    yield row;
  }
}

async function emitProgress(
  callback: ExportExcelRowsOptions['onProgress'],
  progress: ExcelExportProgress,
): Promise<void> {
  if (!callback) return;
  await callback(progress);
}

async function getWrittenTargetSize(target: FileTarget): Promise<number> {
  const file = toReadableFile(target);
  const exists = await file.exists();
  if (!exists) {
    throw new Error(
      `Export target was not created: ${describeFileSource(target)}`,
    );
  }

  return getRuntimeFileSize(file);
}

export async function exportExcelRows(
  options: ExportExcelRowsOptions,
): Promise<ExcelExportDiagnostics> {
  const {
    target,
    rows,
    mode = 'stream',
    signal,
    progressIntervalRows = 1000,
    onProgress,
    validateOutput = true,
    ...writerOptions
  } = options;

  const writer =
    mode === 'chunked'
      ? createChunkedExcelStream(
          target,
          writerOptions as ChunkedExcelStreamOptions,
        )
      : createExcelStream(target, writerOptions);

  const startedAt = new Date();
  const startedAtNs = Bun.nanoseconds();
  const baselineMemory = takeMemorySnapshot();
  const peakMemory = { ...baselineMemory };
  const targetDescription = describeFileSource(target);
  let rowsWritten = 0;

  try {
    throwIfAborted(signal);

    for await (const row of toAsyncRowIterable(rows)) {
      throwIfAborted(signal);
      writer.writeRow(row);
      rowsWritten++;

      const snapshot = takeMemorySnapshot();
      updatePeakSnapshot(peakMemory, snapshot);

      if (rowsWritten % progressIntervalRows === 0) {
        await writer.flush();
        await emitProgress(onProgress, {
          stage: 'writing',
          mode,
          rowsWritten,
          elapsedMs: (Bun.nanoseconds() - startedAtNs) / 1_000_000,
          target: targetDescription,
          memory: snapshot,
        });
      }
    }

    throwIfAborted(signal);
    await emitProgress(onProgress, {
      stage: 'finalizing',
      mode,
      rowsWritten,
      elapsedMs: (Bun.nanoseconds() - startedAtNs) / 1_000_000,
      target: targetDescription,
      memory: takeMemorySnapshot(),
    });

    await writer.end();

    const endMemory = takeMemorySnapshot();
    updatePeakSnapshot(peakMemory, endMemory);

    const outputSizeBytes = validateOutput
      ? await getWrittenTargetSize(target)
      : await getWrittenTargetSize(target);

    if (validateOutput && outputSizeBytes <= 0) {
      throw new Error(
        `Export output is empty: ${targetDescription} (${outputSizeBytes} bytes)`,
      );
    }

    await emitProgress(onProgress, {
      stage: 'completed',
      mode,
      rowsWritten,
      elapsedMs: (Bun.nanoseconds() - startedAtNs) / 1_000_000,
      target: targetDescription,
      memory: endMemory,
    });

    return {
      mode,
      target: targetDescription,
      rowsWritten,
      startedAt,
      finishedAt: new Date(),
      durationMs: (Bun.nanoseconds() - startedAtNs) / 1_000_000,
      outputSizeBytes,
      memory: {
        baseline: baselineMemory,
        peak: peakMemory,
        end: endMemory,
      },
    };
  } catch (error) {
    if (signal?.aborted) {
      await emitProgress(onProgress, {
        stage: 'aborted',
        mode,
        rowsWritten,
        elapsedMs: (Bun.nanoseconds() - startedAtNs) / 1_000_000,
        target: targetDescription,
        memory: takeMemorySnapshot(),
      });
    }

    if ('cancel' in writer && typeof writer.cancel === 'function') {
      await writer.cancel();
    }
    throw error;
  }
}

export async function exportMultiSheetExcel(
  options: ExportMultiSheetExcelOptions,
): Promise<ExcelExportDiagnostics> {
  const {
    target,
    sheets,
    signal,
    progressIntervalRows = 1000,
    onProgress,
    validateOutput = true,
    ...writerOptions
  } = options;

  const writer = createMultiSheetExcelStream(target, writerOptions);
  const startedAt = new Date();
  const startedAtNs = Bun.nanoseconds();
  const baselineMemory = takeMemorySnapshot();
  const peakMemory = { ...baselineMemory };
  const targetDescription = describeFileSource(target);
  let rowsWritten = 0;

  try {
    throwIfAborted(signal);

    for (let sheetIndex = 0; sheetIndex < sheets.length; sheetIndex++) {
      const sheet = sheets[sheetIndex];
      writer.addSheet(sheet.name, sheet.options);

      for await (const row of toAsyncRowIterable(sheet.rows)) {
        throwIfAborted(signal);
        writer.writeRow(row);
        rowsWritten++;

        const snapshot = takeMemorySnapshot();
        updatePeakSnapshot(peakMemory, snapshot);

        if (rowsWritten % progressIntervalRows === 0) {
          await writer.flush();
          await emitProgress(onProgress, {
            stage: 'writing',
            mode: 'multi-sheet',
            rowsWritten,
            elapsedMs: (Bun.nanoseconds() - startedAtNs) / 1_000_000,
            target: targetDescription,
            memory: snapshot,
            sheetName: sheet.name,
            sheetIndex,
          });
        }
      }
    }

    throwIfAborted(signal);
    await emitProgress(onProgress, {
      stage: 'finalizing',
      mode: 'multi-sheet',
      rowsWritten,
      elapsedMs: (Bun.nanoseconds() - startedAtNs) / 1_000_000,
      target: targetDescription,
      memory: takeMemorySnapshot(),
    });

    await writer.end();

    const endMemory = takeMemorySnapshot();
    updatePeakSnapshot(peakMemory, endMemory);
    const outputSizeBytes = await getWrittenTargetSize(target);

    if (validateOutput && outputSizeBytes <= 0) {
      throw new Error(
        `Export output is empty: ${targetDescription} (${outputSizeBytes} bytes)`,
      );
    }

    await emitProgress(onProgress, {
      stage: 'completed',
      mode: 'multi-sheet',
      rowsWritten,
      elapsedMs: (Bun.nanoseconds() - startedAtNs) / 1_000_000,
      target: targetDescription,
      memory: endMemory,
    });

    return {
      mode: 'multi-sheet',
      target: targetDescription,
      rowsWritten,
      startedAt,
      finishedAt: new Date(),
      durationMs: (Bun.nanoseconds() - startedAtNs) / 1_000_000,
      outputSizeBytes,
      sheetCount: sheets.length,
      memory: {
        baseline: baselineMemory,
        peak: peakMemory,
        end: endMemory,
      },
    };
  } catch (error) {
    if (signal?.aborted) {
      await emitProgress(onProgress, {
        stage: 'aborted',
        mode: 'multi-sheet',
        rowsWritten,
        elapsedMs: (Bun.nanoseconds() - startedAtNs) / 1_000_000,
        target: targetDescription,
        memory: takeMemorySnapshot(),
      });
    }

    await writer.cancel();
    throw error;
  }
}

export async function buildExcelResponse(
  workbook: Workbook,
  options: ExcelResponseOptions = {},
): Promise<Response> {
  const buffer = buildExcelBuffer(workbook);
  const responseBuffer = new ArrayBuffer(buffer.byteLength);
  new Uint8Array(responseBuffer).set(buffer);
  const headers = new Headers(options.headers);
  const contentType =
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  headers.set('content-type', contentType);
  if (options.filename) {
    headers.set(
      'content-disposition',
      `attachment; filename="${options.filename.replace(/"/g, '')}"`,
    );
  }
  return new Response(new Blob([responseBuffer], { type: contentType }), {
    headers,
  });
}

export async function writeExcelWithDiagnostics(
  target: FileTarget,
  workbook: Workbook,
): Promise<ExcelExportDiagnostics> {
  const startedAt = new Date();
  const startedAtNs = Bun.nanoseconds();
  const baselineMemory = takeMemorySnapshot();

  await writeExcel(target, workbook);

  const endMemory = takeMemorySnapshot();
  const outputSizeBytes = await getWrittenTargetSize(target);

  return {
    mode: 'stream',
    target: describeFileSource(target),
    rowsWritten: workbook.worksheets.reduce(
      (sum, worksheet) => sum + worksheet.rows.length,
      0,
    ),
    startedAt,
    finishedAt: new Date(),
    durationMs: (Bun.nanoseconds() - startedAtNs) / 1_000_000,
    outputSizeBytes,
    memory: {
      baseline: baselineMemory,
      peak: endMemory,
      end: endMemory,
    },
  };
}
