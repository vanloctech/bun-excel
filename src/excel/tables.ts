import type { Worksheet, WorksheetTable } from '../types';
import { findChild, findChildren, parseXML } from './native-xml';
import { buildRangeRef, escapeXML, parseCellRef } from './xml-builder';

function defaultTableColumnName(
  worksheet: Worksheet,
  table: WorksheetTable,
  index: number,
): string {
  if (!table.headerRow) return `Column${index + 1}`;
  const headerRow = worksheet.rows[table.range.startRow];
  const headerCell = headerRow?.cells[table.range.startCol + index];
  const value = headerCell?.value;
  return value === null || value === undefined
    ? `Column${index + 1}`
    : String(value);
}

export function buildTableXML(
  worksheet: Worksheet,
  table: WorksheetTable,
  tableId: number,
): string {
  const ref = buildRangeRef(
    table.range.startRow,
    table.range.startCol,
    table.range.endRow,
    table.range.endCol,
  );
  const displayName = table.displayName || table.name;
  const columnCount = table.range.endCol - table.range.startCol + 1;
  const headerRow = table.headerRow !== false;
  const totalsRow = table.totalsRow === true;
  const autoFilterEndRow = totalsRow
    ? table.range.endRow - 1
    : table.range.endRow;

  let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  xml +=
    '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
  xml += ` id="${tableId}" name="${escapeXML(table.name)}" displayName="${escapeXML(displayName)}"`;
  xml += ` ref="${escapeXML(ref)}"`;
  xml += ` headerRowCount="${headerRow ? 1 : 0}"`;
  if (totalsRow) {
    xml += ' totalsRowCount="1" totalsRowShown="1"';
  } else {
    xml += ' totalsRowShown="0"';
  }
  xml += '>';

  if (headerRow) {
    xml += `<autoFilter ref="${buildRangeRef(
      table.range.startRow,
      table.range.startCol,
      autoFilterEndRow,
      table.range.endCol,
    )}"/>`;
  }

  xml += `<tableColumns count="${columnCount}">`;
  for (let i = 0; i < columnCount; i++) {
    const column = table.columns?.[i];
    xml += `<tableColumn id="${i + 1}" name="${escapeXML(
      column?.name || defaultTableColumnName(worksheet, table, i),
    )}"`;
    if (column?.totalsRowLabel) {
      xml += ` totalsRowLabel="${escapeXML(column.totalsRowLabel)}"`;
    }
    if (column?.totalsRowFunction) {
      xml += ` totalsRowFunction="${escapeXML(column.totalsRowFunction)}"`;
    }
    xml += '/>';
  }
  xml += '</tableColumns>';

  if (table.style) {
    xml += '<tableStyleInfo';
    xml += ` name="${escapeXML(table.style.name || 'TableStyleMedium2')}"`;
    xml += ` showFirstColumn="${table.style.showFirstColumn ? 1 : 0}"`;
    xml += ` showLastColumn="${table.style.showLastColumn ? 1 : 0}"`;
    xml += ` showRowStripes="${table.style.showRowStripes !== false ? 1 : 0}"`;
    xml += ` showColumnStripes="${table.style.showColumnStripes ? 1 : 0}"`;
    xml += '/>';
  }

  xml += '</table>';
  return xml;
}

export function parseTableXML(xml: string): WorksheetTable | undefined {
  const root = parseXML(xml);
  if (!root) return undefined;

  const ref = root.attributes.ref;
  if (!ref) return undefined;
  const [startRef, endRef = startRef] = ref.split(':');
  const start = parseCellRef(startRef);
  const end = parseCellRef(endRef);

  const table: WorksheetTable = {
    name: root.attributes.name,
    displayName: root.attributes.displayName,
    range: {
      startRow: start.row,
      startCol: start.col,
      endRow: end.row,
      endCol: end.col,
    },
    headerRow: root.attributes.headerRowCount !== '0',
    totalsRow: root.attributes.totalsRowShown === '1',
  };

  const columnsNode = findChild(root, 'tableColumns');
  if (columnsNode) {
    table.columns = findChildren(columnsNode, 'tableColumn').map(
      (columnNode) => ({
        name: columnNode.attributes.name || '',
        totalsRowLabel: columnNode.attributes.totalsRowLabel,
        totalsRowFunction: columnNode.attributes
          .totalsRowFunction as NonNullable<
          WorksheetTable['columns']
        >[number]['totalsRowFunction'],
      }),
    );
  }

  const styleInfo = findChild(root, 'tableStyleInfo');
  if (styleInfo) {
    table.style = {
      name: styleInfo.attributes.name,
      showFirstColumn: styleInfo.attributes.showFirstColumn === '1',
      showLastColumn: styleInfo.attributes.showLastColumn === '1',
      showRowStripes: styleInfo.attributes.showRowStripes !== '0',
      showColumnStripes: styleInfo.attributes.showColumnStripes === '1',
    };
  }

  return table;
}
