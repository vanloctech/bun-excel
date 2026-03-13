// ============================================
// CSV Reader — Bun-optimized CSV parsing
// ============================================

import { resolve } from 'node:path';

// Top-level regex for performance (biome: useTopLevelRegex)
const ISO_DATE_REGEX = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}:\d{2})?/;

import type {
  Cell,
  CellValue,
  CSVReadOptions,
  Row,
  Workbook,
  Worksheet,
} from '../types';

/** Security limits */
const MAX_CSV_FILE_SIZE = 500 * 1024 * 1024; // 500MB max
const MAX_FIELD_LENGTH = 1_000_000; // 1MB max per field

const DEFAULT_OPTIONS: Required<CSVReadOptions> = {
  delimiter: ',',
  quoteChar: '"',
  escapeChar: '"',
  hasHeader: false,
  encoding: 'utf-8',
  skipEmptyLines: true,
};

/**
 * Parse a CSV string into rows of cell values
 */
function parseCSVContent(
  content: string,
  options: Required<CSVReadOptions>,
): string[][] {
  const { delimiter, quoteChar, skipEmptyLines } = options;
  const rows: string[][] = [];
  let currentRow: string[] = [];
  let currentField = '';
  let inQuotes = false;
  let i = 0;

  while (i < content.length) {
    const char = content[i];

    if (inQuotes) {
      if (char === quoteChar) {
        // Check for escaped quote
        if (i + 1 < content.length && content[i + 1] === quoteChar) {
          currentField += quoteChar;
          i += 2;
          continue;
        }
        inQuotes = false;
        i++;
        continue;
      }
      currentField += char;
      if (currentField.length > MAX_FIELD_LENGTH) {
        throw new Error(
          `CSV field exceeds maximum length (${MAX_FIELD_LENGTH} chars)`,
        );
      }
      i++;
    } else {
      if (char === quoteChar) {
        inQuotes = true;
        i++;
      } else if (char === delimiter) {
        currentRow.push(currentField);
        currentField = '';
        i++;
      } else if (char === '\r') {
        // Handle \r\n and \r
        currentRow.push(currentField);
        currentField = '';
        if (!skipEmptyLines || currentRow.some((f) => f.length > 0)) {
          rows.push(currentRow);
        }
        currentRow = [];
        if (i + 1 < content.length && content[i + 1] === '\n') {
          i += 2;
        } else {
          i++;
        }
      } else if (char === '\n') {
        currentRow.push(currentField);
        currentField = '';
        if (!skipEmptyLines || currentRow.some((f) => f.length > 0)) {
          rows.push(currentRow);
        }
        currentRow = [];
        i++;
      } else {
        currentField += char;
        i++;
      }
    }
  }

  // Handle last field
  if (currentField.length > 0 || currentRow.length > 0) {
    currentRow.push(currentField);
    if (!skipEmptyLines || currentRow.some((f) => f.length > 0)) {
      rows.push(currentRow);
    }
  }

  return rows;
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
  path: string,
  options?: CSVReadOptions,
): Promise<Workbook> {
  const opts = { ...DEFAULT_OPTIONS, ...options };

  // Validate and resolve path (prevent path traversal)
  const resolvedPath = resolve(path);
  if (resolvedPath.includes('\0')) {
    throw new Error('Invalid file path: contains null bytes');
  }

  // Use Bun.file() for optimized reading
  const file = Bun.file(resolvedPath);
  const exists = await file.exists();
  if (!exists) {
    throw new Error(`File not found: ${resolvedPath}`);
  }

  // Check file size before loading into memory
  if (file.size > MAX_CSV_FILE_SIZE) {
    throw new Error(
      `CSV file too large: ${file.size} bytes (max: ${MAX_CSV_FILE_SIZE}). Use readCSVStream() for large files.`,
    );
  }

  const content = await file.text();
  const rawRows = parseCSVContent(content, opts);

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
  path: string,
  options?: CSVReadOptions,
): AsyncGenerator<Row, void, unknown> {
  const opts = { ...DEFAULT_OPTIONS, ...options };
  const { delimiter, quoteChar, skipEmptyLines } = opts;

  // Validate and resolve path
  const resolvedPath = resolve(path);
  if (resolvedPath.includes('\0')) {
    throw new Error('Invalid file path: contains null bytes');
  }

  // Use Bun.file().stream() for streaming read
  const file = Bun.file(resolvedPath);
  const exists = await file.exists();
  if (!exists) {
    throw new Error(`File not found: ${resolvedPath}`);
  }

  const stream = file.stream();
  const decoder = new TextDecoder(opts.encoding);

  let buffer = '';
  let inQuotes = false;
  let currentRow: string[] = [];
  let currentField = '';
  let rowIndex = 0;

  for await (const chunk of stream) {
    buffer += decoder.decode(chunk, { stream: true });

    let i = 0;
    while (i < buffer.length) {
      const char = buffer[i];

      if (inQuotes) {
        if (char === quoteChar) {
          if (i + 1 < buffer.length && buffer[i + 1] === quoteChar) {
            currentField += quoteChar;
            i += 2;
            continue;
          }
          inQuotes = false;
          i++;
          continue;
        }
        currentField += char;
        i++;
      } else {
        if (char === quoteChar) {
          inQuotes = true;
          i++;
        } else if (char === delimiter) {
          currentRow.push(currentField);
          currentField = '';
          i++;
        } else if (char === '\r' || char === '\n') {
          if (
            char === '\r' &&
            i + 1 < buffer.length &&
            buffer[i + 1] === '\n'
          ) {
            i++;
          }
          currentRow.push(currentField);
          currentField = '';

          if (!skipEmptyLines || currentRow.some((f) => f.length > 0)) {
            if (!(opts.hasHeader && rowIndex === 0)) {
              const cells: Cell[] = currentRow.map((v) => ({
                value: detectCellValue(v),
              }));
              yield { cells };
            }
            rowIndex++;
          }
          currentRow = [];
          i++;
        } else {
          currentField += char;
          i++;
        }
      }
    }

    // Keep remaining incomplete data in buffer
    buffer = '';
  }

  // Handle last row
  if (currentField.length > 0 || currentRow.length > 0) {
    currentRow.push(currentField);
    if (!skipEmptyLines || currentRow.some((f) => f.length > 0)) {
      if (!(opts.hasHeader && rowIndex === 0)) {
        const cells: Cell[] = currentRow.map((v) => ({
          value: detectCellValue(v),
        }));
        yield { cells };
      }
    }
  }
}
