import type {
  CellValue,
  ExcelObjectData,
  ExcelObjectField,
  ExcelObjectReadOptions,
  ExcelObjectResult,
  ExcelObjectSchema,
  ExcelObjectValidationError,
  FileSource,
  Row,
} from '../types';
import { readExcelStream } from './xlsx-reader';

/** A worksheet cannot be mapped unambiguously to the supplied schema. */
export class ExcelHeaderError extends Error {
  override readonly name = 'ExcelHeaderError';
  readonly code: 'missing_header_row' | 'missing_header' | 'duplicate_header';
  readonly sheetIndex: number;
  readonly sheetName: string;
  readonly rowIndex: number;
  readonly header?: string;
  constructor(
    code: ExcelHeaderError['code'],
    sheetIndex: number,
    sheetName: string,
    rowIndex: number,
    header?: string,
  ) {
    super(
      `${code} in worksheet "${sheetName}" at row ${rowIndex + 1}${header === undefined ? '' : `: ${header}`}`,
    );
    this.code = code;
    this.sheetIndex = sheetIndex;
    this.sheetName = sheetName;
    this.rowIndex = rowIndex;
    this.header = header;
  }
}

function cellAddress(column: number, row: number): string {
  let letters = '';
  for (let n = column + 1; n > 0; n = Math.floor((n - 1) / 26))
    letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
  return `${letters}${row + 1}`;
}

const DECIMAL_NUMBER = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;

function convert(value: CellValue, field: ExcelObjectField): CellValue {
  if (!field.coerce) return value;
  if (
    field.type === 'string' &&
    (typeof value === 'number' || typeof value === 'boolean')
  )
    return String(value);
  if (typeof value !== 'string') return value;
  if (field.type === 'number' && DECIMAL_NUMBER.test(value.trim()))
    return Number(value.trim());
  if (field.type === 'boolean') {
    if (value === 'true') return true;
    if (value === 'false') return false;
  }
  return value;
}

/** Stream typed objects; schema/header failures throw, invalid data yields ok:false. */
export async function* readExcelObjectsStream<
  const S extends ExcelObjectSchema,
>(
  source: FileSource,
  options: ExcelObjectReadOptions<S>,
): AsyncGenerator<ExcelObjectResult<S>> {
  const headerRow = options.headerRow === undefined ? 0 : options.headerRow;
  if (
    !Number.isSafeInteger(headerRow) ||
    headerRow < 0 ||
    headerRow > 1_048_575
  )
    throw new Error('headerRow must be an integer between 0 and 1048575');
  if (
    !options.schema ||
    typeof options.schema !== 'object' ||
    Array.isArray(options.schema)
  )
    throw new Error('schema must be a non-empty object');
  const fields = Object.entries(options.schema).map(([key, input]) => {
    const field = {
      header: input?.header,
      type: input?.type,
      required: input?.required,
      coerce: input?.coerce,
    };
    if (
      typeof field.header !== 'string' ||
      field.header.length === 0 ||
      !['string', 'number', 'boolean'].includes(field.type) ||
      (field.required !== undefined && typeof field.required !== 'boolean') ||
      (field.coerce !== undefined && typeof field.coerce !== 'boolean')
    )
      throw new Error(`Invalid schema field: ${key}`);
    return {
      key,
      field,
      column: undefined as number | undefined,
    };
  });
  if (fields.length === 0) throw new Error('schema must be a non-empty object');
  let sheetIndex: number | undefined;
  let sheetName = '';
  let hasHeader = false;
  const checkHeader = () => {
    if (sheetIndex !== undefined && !hasHeader)
      throw new ExcelHeaderError(
        'missing_header_row',
        sheetIndex,
        sheetName,
        headerRow,
      );
  };
  const bindHeader = (row: Row, index: number) => {
    const headers = new Map<string, number[]>();
    for (let i = 0; i < row.cells.length; i++) {
      const value = row.cells[i]?.value;
      if (typeof value !== 'string') continue;
      const columns = headers.get(value);
      if (columns) columns.push(i);
      else headers.set(value, [i]);
    }
    for (const entry of fields) {
      const columns = headers.get(entry.field.header);
      if (columns && columns.length > 1)
        throw new ExcelHeaderError(
          'duplicate_header',
          index,
          sheetName,
          headerRow,
          entry.field.header,
        );
      if (!columns && entry.field.required)
        throw new ExcelHeaderError(
          'missing_header',
          index,
          sheetName,
          headerRow,
          entry.field.header,
        );
      entry.column = columns?.[0];
    }
    hasHeader = true;
  };
  for await (const entry of readExcelStream(source, {
    sheets: options.sheets,
    includeStyles: options.includeStyles,
    signal: options.signal,
    progressIntervalRows: options.progressIntervalRows,
    async onProgress(progress) {
      if (
        progress.stage === 'reading' &&
        progress.sheetIndex !== undefined &&
        progress.sheetName !== undefined &&
        progress.sheetIndex !== sheetIndex
      ) {
        checkHeader();
        sheetIndex = progress.sheetIndex;
        sheetName = progress.sheetName;
        hasHeader = false;
      }
      if (progress.stage === 'completed') checkHeader();
      await options.onProgress?.(progress);
    },
  })) {
    if (entry.rowIndex < headerRow) continue;
    if (entry.rowIndex === headerRow) {
      if (hasHeader)
        throw new ExcelHeaderError(
          'duplicate_header',
          entry.sheetIndex,
          entry.sheetName,
          headerRow,
        );
      bindHeader(entry.row, entry.sheetIndex);
      continue;
    }
    checkHeader();
    const data: Record<string, CellValue> = {};
    const errors: ExcelObjectValidationError[] = [];
    for (const { key, field, column } of fields) {
      const raw =
        column === undefined ? undefined : entry.row.cells[column]?.value;
      const empty = raw === null || raw === undefined || raw === '';
      const value = empty ? null : convert(raw, field);
      let code: ExcelObjectValidationError['code'] | undefined;
      if (empty) {
        if (field.required) code = 'required';
      } else if (
        typeof value !== field.type ||
        (typeof value === 'number' && !Number.isFinite(value))
      ) {
        code = 'invalid_type';
      }
      if (code && column !== undefined) {
        errors.push({
          key,
          cell: cellAddress(column, entry.rowIndex),
          code,
          expected: field.type,
          value: raw,
          message:
            code === 'required'
              ? `${field.header} is required`
              : `${field.header} must be ${field.type}`,
        });
      } else {
        if (key === '__proto__')
          Object.defineProperty(data, key, {
            value,
            enumerable: true,
            configurable: true,
            writable: true,
          });
        else data[key] = value;
      }
    }
    const location = {
      sheetIndex: entry.sheetIndex,
      sheetName: entry.sheetName,
      rowIndex: entry.rowIndex,
    };
    yield errors.length
      ? { ...location, ok: false, errors }
      : { ...location, ok: true, data: data as ExcelObjectData<S> };
  }
}
