import type { CellRange, DataValidation } from '../types';
import {
  buildRangeRef,
  escapeXML,
  getFiniteNumber,
  getNonNegativeIntegerOr,
  parseCellRef,
} from './xml-builder';
import {
  findChild,
  findChildren,
  getTextContent,
  type XMLNode,
} from './xml-parser';

const WHITESPACE_SPLIT_REGEX = /\s+/;

function dateToExcelSerial(date: Date): number {
  const epoch = new Date(Date.UTC(1899, 11, 30));
  return (date.getTime() - epoch.getTime()) / (24 * 60 * 60 * 1000);
}

function excelSerialToDate(serial: number): Date {
  const epoch = new Date(Date.UTC(1899, 11, 30));
  return new Date(epoch.getTime() + serial * 24 * 60 * 60 * 1000);
}

function normalizeRange(range: CellRange): CellRange {
  const startRow = getNonNegativeIntegerOr(range.startRow, 0);
  const startCol = getNonNegativeIntegerOr(range.startCol, 0);
  const endRow = getNonNegativeIntegerOr(range.endRow, startRow);
  const endCol = getNonNegativeIntegerOr(range.endCol, startCol);
  return {
    startRow: Math.min(startRow, endRow),
    startCol: Math.min(startCol, endCol),
    endRow: Math.max(startRow, endRow),
    endCol: Math.max(startCol, endCol),
  };
}

function normalizeRanges(range: CellRange | CellRange[]): CellRange[] {
  return (Array.isArray(range) ? range : [range]).map(normalizeRange);
}

function buildSqref(range: CellRange | CellRange[]): string {
  return normalizeRanges(range)
    .map((entry) =>
      buildRangeRef(entry.startRow, entry.startCol, entry.endRow, entry.endCol),
    )
    .join(' ');
}

function stripLeadingEquals(value: string): string {
  return value.startsWith('=') ? value.slice(1) : value;
}

function serializeListFormula(values: string[]): string {
  const escaped = values.map((value) => value.replace(/"/g, '""'));
  return `"${escaped.join(',')}"`;
}

function serializeFormulaValue(
  value: DataValidation['formula1'] | DataValidation['formula2'],
): string | undefined {
  if (value === null || value === undefined) return undefined;
  if (Array.isArray(value)) return serializeListFormula(value);
  if (value instanceof Date) return String(dateToExcelSerial(value));
  if (typeof value === 'number') {
    const numeric = getFiniteNumber(value);
    return numeric !== undefined ? String(numeric) : undefined;
  }
  return stripLeadingEquals(String(value));
}

function parseListFormula(raw: string): string[] | string {
  if (raw.startsWith('"') && raw.endsWith('"')) {
    return raw
      .slice(1, -1)
      .split(',')
      .map((value) => value.replace(/""/g, '"'));
  }
  return raw;
}

function parseFormula1Value(
  raw: string,
  type: DataValidation['type'],
): DataValidation['formula1'] {
  if (type === 'list') {
    return parseListFormula(raw);
  }

  const numeric = getFiniteNumber(raw);
  if (numeric === undefined) return raw;

  if (type === 'date') return excelSerialToDate(numeric);
  if (type === 'whole' || type === 'decimal' || type === 'time') return numeric;
  if (type === 'textLength') return Math.trunc(numeric);
  return raw;
}

function parseFormula2Value(
  raw: string,
  type: DataValidation['type'],
): DataValidation['formula2'] {
  const numeric = getFiniteNumber(raw);
  if (numeric === undefined) return raw;

  if (type === 'date') return excelSerialToDate(numeric);
  if (type === 'whole' || type === 'decimal' || type === 'time') return numeric;
  if (type === 'textLength') return Math.trunc(numeric);
  return raw;
}

function parseSqref(sqref: string): CellRange[] {
  const ranges: CellRange[] = [];

  for (const part of sqref.split(WHITESPACE_SPLIT_REGEX)) {
    if (!part) continue;
    const [startRef, endRef] = part.split(':');
    try {
      const start = parseCellRef(startRef);
      const end = parseCellRef(endRef || startRef);
      ranges.push({
        startRow: Math.min(start.row, end.row),
        startCol: Math.min(start.col, end.col),
        endRow: Math.max(start.row, end.row),
        endCol: Math.max(start.col, end.col),
      });
    } catch {
      // Skip malformed ranges
    }
  }

  return ranges;
}

export function buildDataValidationsXML(
  validations: DataValidation[] | undefined,
): string {
  if (!validations || validations.length === 0) return '';

  const entries = validations
    .map((validation) => {
      const sqref = buildSqref(validation.range);
      if (!sqref) return '';

      let xml = `<dataValidation type="${escapeXML(validation.type)}" sqref="${escapeXML(sqref)}"`;
      if (validation.operator) {
        xml += ` operator="${escapeXML(validation.operator)}"`;
      }
      if (validation.allowBlank !== undefined) {
        xml += ` allowBlank="${validation.allowBlank ? 1 : 0}"`;
      }
      if (validation.showInputMessage !== undefined) {
        xml += ` showInputMessage="${validation.showInputMessage ? 1 : 0}"`;
      }
      if (validation.showErrorMessage !== undefined) {
        xml += ` showErrorMessage="${validation.showErrorMessage ? 1 : 0}"`;
      }
      if (validation.errorStyle) {
        xml += ` errorStyle="${escapeXML(validation.errorStyle)}"`;
      }
      if (validation.promptTitle) {
        xml += ` promptTitle="${escapeXML(validation.promptTitle)}"`;
      }
      if (validation.prompt) {
        xml += ` prompt="${escapeXML(validation.prompt)}"`;
      }
      if (validation.errorTitle) {
        xml += ` errorTitle="${escapeXML(validation.errorTitle)}"`;
      }
      if (validation.error) {
        xml += ` error="${escapeXML(validation.error)}"`;
      }
      xml += '>';

      const formula1 = serializeFormulaValue(validation.formula1);
      const formula2 = serializeFormulaValue(validation.formula2);
      if (formula1 !== undefined) {
        xml += `<formula1>${escapeXML(formula1)}</formula1>`;
      }
      if (formula2 !== undefined) {
        xml += `<formula2>${escapeXML(formula2)}</formula2>`;
      }

      xml += '</dataValidation>';
      return xml;
    })
    .filter((entry) => entry.length > 0);

  if (entries.length === 0) return '';
  return `<dataValidations count="${entries.length}">${entries.join('')}</dataValidations>`;
}

export function parseDataValidations(root: XMLNode): DataValidation[] {
  const validationsNode = findChild(root, 'dataValidations');
  if (!validationsNode) return [];

  const validations: DataValidation[] = [];
  for (const node of findChildren(validationsNode, 'dataValidation')) {
    const sqref = node.attributes.sqref;
    if (!sqref || !node.attributes.type) continue;

    const ranges = parseSqref(sqref);
    if (ranges.length === 0) continue;

    const type = node.attributes.type as DataValidation['type'];
    const formula1Node = findChild(node, 'formula1');
    const formula2Node = findChild(node, 'formula2');
    const formula1Text = formula1Node
      ? getTextContent(formula1Node)
      : undefined;
    const formula2Text = formula2Node
      ? getTextContent(formula2Node)
      : undefined;

    validations.push({
      range: ranges.length === 1 ? ranges[0] : ranges,
      type,
      operator: node.attributes.operator as DataValidation['operator'],
      allowBlank: node.attributes.allowBlank === '1',
      showInputMessage: node.attributes.showInputMessage === '1',
      showErrorMessage: node.attributes.showErrorMessage === '1',
      errorStyle: node.attributes.errorStyle as DataValidation['errorStyle'],
      promptTitle: node.attributes.promptTitle,
      prompt: node.attributes.prompt,
      errorTitle: node.attributes.errorTitle,
      error: node.attributes.error,
      formula1:
        formula1Text !== undefined
          ? parseFormula1Value(formula1Text, type)
          : undefined,
      formula2:
        formula2Text !== undefined
          ? parseFormula2Value(formula2Text, type)
          : undefined,
    });
  }

  return validations;
}
