import type { CellRange } from '../types';
import { findChild, type XMLNode } from './native-xml';
import {
  buildRangeRef,
  escapeXML,
  getNonNegativeIntegerOr,
  parseCellRef,
} from './xml-builder';

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

export function buildAutoFilterXML(range: CellRange | undefined): string {
  if (!range) return '';
  const normalized = normalizeRange(range);
  const ref = buildRangeRef(
    normalized.startRow,
    normalized.startCol,
    normalized.endRow,
    normalized.endCol,
  );
  return `<autoFilter ref="${escapeXML(ref)}"/>`;
}

export function parseAutoFilter(root: XMLNode): CellRange | undefined {
  const autoFilter = findChild(root, 'autoFilter');
  const ref = autoFilter?.attributes.ref;
  if (!ref) return undefined;

  const [startRef, endRef] = ref.split(':');
  try {
    const start = parseCellRef(startRef);
    const end = parseCellRef(endRef || startRef);
    return normalizeRange({
      startRow: start.row,
      startCol: start.col,
      endRow: end.row,
      endCol: end.col,
    });
  } catch {
    return undefined;
  }
}
