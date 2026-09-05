// ============================================
// CSV Reader — Bun-optimized CSV parsing
// ============================================

// Top-level regex for performance (biome: useTopLevelRegex)
const ISO_DATE_REGEX = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}:\d{2})?/;

import {
  describeFileSource,
  getRuntimeFileSize,
  toReadableFile,
} from '../runtime-io';
import type {
  Cell,
  CellValue,
  CSVReadOptions,
  FileSource,
  Row,
  Workbook,
  Worksheet,
} from '../types';

/** Security limits */
const MAX_CSV_FILE_SIZE = 500 * 1024 * 1024; // 500MB max
const MAX_FIELD_LENGTH = 1_000_000; // UTF-16 code units per decoded field

const DEFAULT_OPTIONS: Required<CSVReadOptions> = {
  delimiter: ',',
  quoteChar: '"',
  escapeChar: '"',
  hasHeader: false,
  encoding: 'utf-8',
  skipEmptyLines: true,
};

/** Shared incremental parser; quote and CRLF state survives chunk boundaries. */
class CSVParser {
  private row: string[] = [];
  private field = '';
  private inQuotes = false;
  private pendingQuote = false;
  private skipLF = false;
  private started = false;

  private readonly options: Required<CSVReadOptions>;

  constructor(options: Required<CSVReadOptions>) {
    this.options = options;
  }

  private append(char: string): void {
    if (this.field.length + char.length > MAX_FIELD_LENGTH) {
      throw new Error(
        `CSV field exceeds maximum length (${MAX_FIELD_LENGTH} chars)`,
      );
    }
    this.field += char;
  }

  private finishRow(): string[] | undefined {
    this.row.push(this.field);
    const row = this.row;
    this.row = [];
    this.field = '';
    this.started = false;
    return !this.options.skipEmptyLines || row.some((field) => field.length > 0)
      ? row
      : undefined;
  }

  *feed(content: string, final = false): Generator<string[]> {
    const { delimiter, quoteChar } = this.options;
    for (let i = 0; i < content.length; i++) {
      const char = content[i];
      if (this.skipLF) {
        this.skipLF = false;
        if (char === '\n') continue;
      }
      this.started = true;
      if (this.pendingQuote) {
        this.pendingQuote = false;
        if (char === quoteChar) {
          this.append(char);
          continue;
        }
        this.inQuotes = false;
      }
      if (this.inQuotes) {
        if (char === quoteChar) this.pendingQuote = true;
        else this.append(char);
      } else if (char === quoteChar) {
        this.inQuotes = true;
      } else if (char === delimiter) {
        this.row.push(this.field);
        this.field = '';
      } else if (char === '\r' || char === '\n') {
        this.skipLF = char === '\r';
        const row = this.finishRow();
        if (row) yield row;
      } else {
        this.append(char);
      }
    }
    if (final && this.started) {
      const row = this.finishRow();
      if (row) yield row;
    }
  }
}

/**
 * Auto-detect cell value type
 */
function detectCellValue(raw: string): CellValue {
  if (raw === '') return null;

  // Boolean
  const lower = raw.toLowerCase();
  if (lower === 'true') return true;
  if (lower === 'false') return false;

  // Number
  const num = Number(raw);
  if (!Number.isNaN(num) && raw.trim() !== '') return num;

  // Date (ISO format)
  if (ISO_DATE_REGEX.test(raw)) {
    const date = new Date(raw);
    if (!Number.isNaN(date.getTime())) return date;
  }

  return raw;
}

/**
 * Read a CSV file and return a Workbook
 * Uses Bun.file().text() for optimized file reading
 */
export async function readCSV(
  source: FileSource,
  options?: CSVReadOptions,
): Promise<Workbook> {
  const opts = { ...DEFAULT_OPTIONS, ...options };
  const file = toReadableFile(source);
  const exists = await file.exists();
  if (!exists) {
    throw new Error(`File not found: ${describeFileSource(source)}`);
  }

  // Check file size before loading into memory
  const fileSize = await getRuntimeFileSize(file);
  if (fileSize > MAX_CSV_FILE_SIZE) {
    throw new Error(
      `CSV file too large: ${fileSize} bytes (max: ${MAX_CSV_FILE_SIZE}). Use readCSVStream() for large files.`,
    );
  }

  const content = await file.text();
  const rawRows = [...new CSVParser(opts).feed(content, true)];

  let headers: string[] | undefined;
  let dataStartIndex = 0;

  if (opts.hasHeader && rawRows.length > 0) {
    headers = rawRows[0];
    dataStartIndex = 1;
  }

  const rows: Row[] = [];
  for (let r = dataStartIndex; r < rawRows.length; r++) {
    const cells: Cell[] = rawRows[r].map((value) => ({
      value: detectCellValue(value),
    }));
    rows.push({ cells });
  }

  const worksheet: Worksheet = {
    name: 'Sheet1',
    rows,
    columns: headers?.map((h) => ({ header: h })),
  };

  return { worksheets: [worksheet] };
}

/**
 * Read a large CSV file as a stream using Bun.file().stream()
 * Returns an AsyncGenerator that yields rows one at a time
 */
export async function* readCSVStream(
  source: FileSource,
  options?: CSVReadOptions,
): AsyncGenerator<Row, void, unknown> {
  const opts = { ...DEFAULT_OPTIONS, ...options };
  const file = toReadableFile(source);
  const exists = await file.exists();
  if (!exists) {
    throw new Error(`File not found: ${describeFileSource(source)}`);
  }

  const stream = file.stream();
  const decoder = new TextDecoder(opts.encoding);

  const parser = new CSVParser(opts);
  let rowIndex = 0;
  for await (const chunk of stream) {
    for (const row of parser.feed(decoder.decode(chunk, { stream: true }))) {
      if (!(opts.hasHeader && rowIndex++ === 0)) {
        yield {
          cells: row.map((value) => ({ value: detectCellValue(value) })),
        };
      }
    }
  }
  for (const row of parser.feed(decoder.decode(), true)) {
    if (!(opts.hasHeader && rowIndex++ === 0)) {
      yield { cells: row.map((value) => ({ value: detectCellValue(value) })) };
    }
  }
}
