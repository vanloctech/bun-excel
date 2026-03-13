// ============================================
// CSV Writer — Bun-optimized CSV writing
// ============================================

import { resolve } from 'node:path';
import type {
  Cell,
  CellValue,
  CSVWriteOptions,
  Row,
  StreamWriter,
  Workbook,
} from '../types';

/** Validate path for security */
function validatePath(filePath: string): string {
  if (filePath.includes('\0')) {
    throw new Error('Invalid file path: contains null bytes');
  }
  return resolve(filePath);
}

const DEFAULT_OPTIONS: Required<CSVWriteOptions> = {
  delimiter: ',',
  quoteChar: '"',
  lineEnding: '\n',
  includeHeader: true,
  headers: [],
  bom: false,
};

/**
 * Escape a cell value for CSV output
 */
function escapeCSVValue(
  value: CellValue,
  delimiter: string,
  quoteChar: string,
): string {
  if (value === null || value === undefined) return '';

  let str: string;
  if (value instanceof Date) {
    str = value.toISOString();
  } else {
    str = String(value);
  }

  // Need quoting if contains delimiter, quote, newline
  const needsQuoting =
    str.includes(delimiter) ||
    str.includes(quoteChar) ||
    str.includes('\n') ||
    str.includes('\r');

  if (needsQuoting) {
    const escaped = str.replace(
      new RegExp(quoteChar.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'),
      quoteChar + quoteChar,
    );
    return `${quoteChar}${escaped}${quoteChar}`;
  }

  return str;
}

/**
 * Get cell value from a Cell or CellValue
 */
function getCellValue(cell: Cell | CellValue): CellValue {
  if (cell !== null && typeof cell === 'object' && 'value' in cell) {
    return (cell as Cell).value;
  }
  return cell as CellValue;
}

/**
 * Convert a row to CSV line
 */
function rowToCSVLine(
  row: Row | CellValue[],
  delimiter: string,
  quoteChar: string,
): string {
  const cells = Array.isArray(row) ? row : row.cells;
  return cells
    .map((cell) => {
      const value = getCellValue(cell as Cell | CellValue);
      return escapeCSVValue(value, delimiter, quoteChar);
    })
    .join(delimiter);
}

/**
 * Write a Workbook to CSV file
 * Uses Bun.write() for optimized file writing
 */
export async function writeCSV(
  path: string,
  data: Workbook | CellValue[][],
  options?: CSVWriteOptions,
): Promise<void> {
  const opts = { ...DEFAULT_OPTIONS, ...options };

  let rows: Row[];
  let headers: string[] = opts.headers;

  if (Array.isArray(data)) {
    // Raw data array
    rows = data.map((rowData) => ({
      cells: rowData.map((v) => ({ value: v })),
    }));
  } else {
    // Workbook
    const worksheet = data.worksheets[0];
    if (!worksheet) throw new Error('Workbook has no worksheets');
    rows = worksheet.rows;
    if (headers.length === 0 && worksheet.columns) {
      headers = worksheet.columns
        .map((c) => c.header || '')
        .filter((h) => h.length > 0);
    }
  }

  const lines: string[] = [];

  // BOM for UTF-8
  let prefix = '';
  if (opts.bom) {
    prefix = '\uFEFF';
  }

  // Headers
  if (opts.includeHeader && headers.length > 0) {
    lines.push(
      headers
        .map((h) => escapeCSVValue(h, opts.delimiter, opts.quoteChar))
        .join(opts.delimiter),
    );
  }

  // Data rows
  for (const row of rows) {
    lines.push(rowToCSVLine(row, opts.delimiter, opts.quoteChar));
  }

  const content = prefix + lines.join(opts.lineEnding) + opts.lineEnding;

  // Use Bun.write() for optimized writing
  await Bun.write(validatePath(path), content);
}

/**
 * CSV Stream Writer — uses Bun's FileSink for incremental writing
 */
export class CSVStreamWriter implements StreamWriter {
  private writer: ReturnType<ReturnType<typeof Bun.file>['writer']>;
  private opts: Required<CSVWriteOptions>;
  private headerWritten = false;

  constructor(path: string, options?: CSVWriteOptions) {
    this.opts = { ...DEFAULT_OPTIONS, ...options };

    // Use Bun.file().writer() — FileSink for incremental writes
    const file = Bun.file(validatePath(path));
    this.writer = file.writer({ highWaterMark: 1024 * 1024 }); // 1MB buffer

    // Write BOM if needed
    if (this.opts.bom) {
      this.writer.write('\uFEFF');
    }
  }

  /**
   * Write headers (auto-called on first writeRow if headers are set)
   */
  private writeHeaders(): void {
    if (this.headerWritten) return;
    this.headerWritten = true;

    if (this.opts.includeHeader && this.opts.headers.length > 0) {
      const headerLine =
        this.opts.headers
          .map((h) =>
            escapeCSVValue(h, this.opts.delimiter, this.opts.quoteChar),
          )
          .join(this.opts.delimiter) + this.opts.lineEnding;
      this.writer.write(headerLine);
    }
  }

  /**
   * Write a single row
   */
  writeRow(row: Row | CellValue[]): void {
    this.writeHeaders();
    const line =
      rowToCSVLine(row, this.opts.delimiter, this.opts.quoteChar) +
      this.opts.lineEnding;
    this.writer.write(line);
  }

  /**
   * Flush the buffer to disk
   */
  flush(): void | Promise<void> {
    const result = this.writer.flush();
    if (result instanceof Promise) {
      return result.then(() => {});
    }
  }

  /**
   * End the stream and close the file
   */
  async end(): Promise<void> {
    await this.writer.end();
  }
}

/**
 * Create a CSV stream writer
 */
export function createCSVStream(
  path: string,
  options?: CSVWriteOptions,
): CSVStreamWriter {
  return new CSVStreamWriter(path, options);
}
