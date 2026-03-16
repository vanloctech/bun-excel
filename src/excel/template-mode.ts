import type {
  Cell,
  CellRange,
  CellValue,
  DefinedName,
  ExcelReadOptions,
  ExcelWriteOptions,
  FileSource,
  FileTarget,
  Row,
  Workbook,
  Worksheet,
} from '../types';
import { readExcel } from './xlsx-reader';
import { buildExcelBuffer, writeExcel } from './xlsx-writer';
import { parseCellRef } from './xml-builder';

export type TemplateSheetRef = string | number;
export type TemplateCellRef = string | { row: number; col: number };
export type TemplateCellInput = CellValue | Partial<Cell>;
export type TemplateCellMatrix = TemplateCellInput[][];

const LEADING_EQUALS_REGEX = /^=+/;
const LEADING_SINGLE_QUOTE_REGEX = /^'/;
const TRAILING_SINGLE_QUOTE_REGEX = /'$/;
const CELL_KEYS = new Set<keyof Cell>([
  'value',
  'style',
  'type',
  'richText',
  'formula',
  'formulaResult',
  'hyperlink',
  'comment',
]);

function toCellCoordinates(ref: TemplateCellRef): { row: number; col: number } {
  if (typeof ref === 'string') {
    return parseCellRef(ref.replace(/\$/g, ''));
  }
  return ref;
}

function parseRangeRef(rangeRef: string): CellRange {
  const [startRef, endRef = startRef] = rangeRef.split(':');
  const start = parseCellRef(startRef.replace(/\$/g, ''));
  const end = parseCellRef(endRef.replace(/\$/g, ''));
  return {
    startRow: start.row,
    startCol: start.col,
    endRow: end.row,
    endCol: end.col,
  };
}

function normalizeSheetName(sheetName: string): string {
  return sheetName
    .replace(LEADING_EQUALS_REGEX, '')
    .replace(LEADING_SINGLE_QUOTE_REGEX, '')
    .replace(TRAILING_SINGLE_QUOTE_REGEX, '')
    .replace(/''/g, "'");
}

function resolveWorksheet(
  workbook: Workbook,
  sheetRef: TemplateSheetRef,
): Worksheet {
  const worksheet =
    typeof sheetRef === 'number'
      ? workbook.worksheets[sheetRef]
      : workbook.worksheets.find((sheet) => sheet.name === sheetRef);
  if (!worksheet) {
    throw new Error(`Worksheet not found: ${String(sheetRef)}`);
  }
  return worksheet;
}

function ensureCell(
  worksheet: Worksheet,
  rowIndex: number,
  colIndex: number,
): Cell {
  while (worksheet.rows.length <= rowIndex) {
    worksheet.rows.push({ cells: [] });
  }

  const existingRow = worksheet.rows[rowIndex];
  if (!existingRow) {
    const newRow: Row = { cells: [] };
    worksheet.rows[rowIndex] = newRow;
    return ensureCell(worksheet, rowIndex, colIndex);
  }

  while (existingRow.cells.length <= colIndex) {
    existingRow.cells.push({ value: null });
  }

  const existingCell = existingRow.cells[colIndex];
  if (existingCell) {
    return existingCell;
  }

  const newCell: Cell = { value: null };
  existingRow.cells[colIndex] = newCell;
  return newCell;
}

function isTemplateCellPatch(value: TemplateCellInput): value is Partial<Cell> {
  if (
    value === null ||
    value === undefined ||
    typeof value !== 'object' ||
    value instanceof Date ||
    Array.isArray(value)
  ) {
    return false;
  }

  return Object.keys(value).some((key) => CELL_KEYS.has(key as keyof Cell));
}

function applyCellInput(cell: Cell, input: TemplateCellInput): void {
  if (!isTemplateCellPatch(input)) {
    cell.value = input;
    return;
  }

  const patch: Partial<Cell> = {};
  for (const [key, value] of Object.entries(input) as [
    keyof Cell,
    Cell[keyof Cell],
  ][]) {
    if (value !== undefined) {
      (patch as Record<string, unknown>)[key] = value;
    }
  }
  Object.assign(cell, patch);
}

function resolveDefinedNameRange(
  workbook: Workbook,
  definedName: DefinedName,
): { sheetIndex: number; range: CellRange } | undefined {
  let sheetIndex = definedName.localSheetId;
  let refersTo = definedName.refersTo.replace(LEADING_EQUALS_REGEX, '');
  const bangIndex = refersTo.lastIndexOf('!');

  if (bangIndex !== -1) {
    const rawSheetName = refersTo.slice(0, bangIndex);
    refersTo = refersTo.slice(bangIndex + 1);

    if (sheetIndex === undefined) {
      sheetIndex = workbook.worksheets.findIndex(
        (worksheet) => worksheet.name === normalizeSheetName(rawSheetName),
      );
    }
  }

  if (sheetIndex === undefined || sheetIndex < 0) {
    return undefined;
  }

  return {
    sheetIndex,
    range: parseRangeRef(refersTo),
  };
}

function findDefinedName(
  workbook: Workbook,
  name: string,
  scope?: TemplateSheetRef,
): DefinedName | undefined {
  const definedNames = workbook.definedNames ?? [];
  let scopedSheetIndex: number | undefined;
  if (scope !== undefined) {
    scopedSheetIndex =
      typeof scope === 'number'
        ? scope
        : workbook.worksheets.findIndex(
            (worksheet) => worksheet.name === scope,
          );
  }

  return definedNames.find((definedName) => {
    if (definedName.name !== name) return false;
    if (scopedSheetIndex === undefined) return true;
    return definedName.localSheetId === undefined
      ? resolveDefinedNameRange(workbook, definedName)?.sheetIndex ===
          scopedSheetIndex
      : definedName.localSheetId === scopedSheetIndex;
  });
}

export class ExcelTemplate {
  readonly workbook: Workbook;

  constructor(workbook: Workbook) {
    this.workbook = workbook;
  }

  getWorksheet(sheetRef: TemplateSheetRef): Worksheet {
    return resolveWorksheet(this.workbook, sheetRef);
  }

  getDefinedName(
    name: string,
    scope?: TemplateSheetRef,
  ): DefinedName | undefined {
    return findDefinedName(this.workbook, name, scope);
  }

  setCell(
    sheetRef: TemplateSheetRef,
    ref: TemplateCellRef,
    input: TemplateCellInput,
  ): this {
    const worksheet = this.getWorksheet(sheetRef);
    const { row, col } = toCellCoordinates(ref);
    const cell = ensureCell(worksheet, row, col);
    applyCellInput(cell, input);
    return this;
  }

  fillRange(
    sheetRef: TemplateSheetRef,
    startRef: TemplateCellRef,
    values: TemplateCellMatrix,
  ): this {
    const start = toCellCoordinates(startRef);
    for (let rowOffset = 0; rowOffset < values.length; rowOffset++) {
      const row = values[rowOffset];
      for (let colOffset = 0; colOffset < row.length; colOffset++) {
        this.setCell(
          sheetRef,
          {
            row: start.row + rowOffset,
            col: start.col + colOffset,
          },
          row[colOffset],
        );
      }
    }
    return this;
  }

  setDefinedName(
    name: string,
    input: TemplateCellInput | TemplateCellMatrix,
    scope?: TemplateSheetRef,
  ): this {
    const definedName = this.getDefinedName(name, scope);
    if (!definedName) {
      throw new Error(`Defined name not found: ${name}`);
    }

    const resolved = resolveDefinedNameRange(this.workbook, definedName);
    if (!resolved) {
      throw new Error(`Defined name is not a worksheet cell range: ${name}`);
    }

    const { sheetIndex, range } = resolved;
    const isSingleCell =
      range.startRow === range.endRow && range.startCol === range.endCol;

    if (isSingleCell) {
      const singleValue = Array.isArray(input)
        ? (input[0]?.[0] ?? null)
        : input;
      return this.setCell(
        sheetIndex,
        { row: range.startRow, col: range.startCol },
        singleValue,
      );
    }

    if (!Array.isArray(input) || !Array.isArray(input[0])) {
      throw new Error(
        `Defined name "${name}" refers to multiple cells. Provide a 2D array.`,
      );
    }

    return this.fillRange(
      sheetIndex,
      { row: range.startRow, col: range.startCol },
      input,
    );
  }

  build(options?: ExcelWriteOptions): Uint8Array {
    return buildExcelBuffer(this.workbook, options);
  }

  async write(target: FileTarget, options?: ExcelWriteOptions): Promise<void> {
    await writeExcel(target, this.workbook, options);
  }
}

export async function loadExcelTemplate(
  source: FileSource,
  options?: ExcelReadOptions,
): Promise<ExcelTemplate> {
  const workbook = await readExcel(source, {
    includeStyles: true,
    ...options,
  });
  return new ExcelTemplate(workbook);
}
