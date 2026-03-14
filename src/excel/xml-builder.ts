// ============================================
// XML Builder for XLSX generation
// ============================================

// Top-level regex for performance (biome: useTopLevelRegex)
const CELL_REF_PARSE_REGEX = /^([A-Z]+)(\d+)$/;

/**
 * Encode special characters for XML
 */
export function escapeXML(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Convert runtime input to a finite number.
 * Returns undefined for NaN/Infinity/non-numeric values.
 */
export function getFiniteNumber(value: unknown): number | undefined {
  let num: number | undefined;
  if (typeof value === 'number') {
    num = value;
  } else if (typeof value === 'string') {
    num = Number(value);
  }
  return num !== undefined && Number.isFinite(num) ? num : undefined;
}

/**
 * Convert runtime input to a finite number with a safe fallback.
 */
export function getFiniteNumberOr(value: unknown, fallback: number): number {
  return getFiniteNumber(value) ?? fallback;
}

/**
 * Convert runtime input to a non-negative integer with a safe fallback.
 */
export function getNonNegativeIntegerOr(
  value: unknown,
  fallback: number,
): number {
  return Math.max(0, Math.trunc(getFiniteNumberOr(value, fallback)));
}

/**
 * Build an XML tag string
 */
export function tag(
  name: string,
  attrs?: Record<string, string | number | boolean | undefined>,
  content?: string,
  selfClose?: boolean,
): string {
  let attrStr = '';
  if (attrs) {
    for (const [key, value] of Object.entries(attrs)) {
      if (value !== undefined && value !== false) {
        attrStr += ` ${key}="${escapeXML(String(value))}"`;
      }
    }
  }

  if (selfClose && !content) {
    return `<${name}${attrStr}/>`;
  }
  return `<${name}${attrStr}>${content || ''}</${name}>`;
}

/**
 * Convert column index (0-based) to Excel column letter (A, B, ..., Z, AA, AB...)
 */
export function colIndexToLetter(index: number): string {
  let letter = '';
  let n = index;
  while (n >= 0) {
    letter = String.fromCharCode(65 + (n % 26)) + letter;
    n = Math.floor(n / 26) - 1;
  }
  return letter;
}

/**
 * Convert Excel column letter to 0-based index
 */
export function letterToColIndex(letter: string): number {
  let index = 0;
  for (let i = 0; i < letter.length; i++) {
    index = index * 26 + (letter.charCodeAt(i) - 64);
  }
  return index - 1;
}

/**
 * Parse cell reference (e.g. "A1") to row and column indices
 */
export function parseCellRef(ref: string): { row: number; col: number } {
  const match = ref.match(CELL_REF_PARSE_REGEX);
  if (!match) throw new Error(`Invalid cell reference: ${ref}`);
  return {
    col: letterToColIndex(match[1]),
    row: Number.parseInt(match[2], 10) - 1,
  };
}

/**
 * Build cell reference from row and column indices
 */
export function buildCellRef(row: number, col: number): string {
  return `${colIndexToLetter(col)}${row + 1}`;
}

/**
 * Build a worksheet range reference from 0-based coordinates
 */
export function buildRangeRef(
  startRow: number,
  startCol: number,
  endRow: number,
  endCol: number,
): string {
  const startRef = buildCellRef(startRow, startCol);
  const endRef = buildCellRef(endRow, endCol);
  return startRef === endRef ? startRef : `${startRef}:${endRef}`;
}

/**
 * Build [Content_Types].xml
 */
export function buildContentTypes(sheetCount: number): string {
  let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  xml +=
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
  xml +=
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
  xml += '<Default Extension="xml" ContentType="application/xml"/>';
  xml +=
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
  xml +=
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';
  xml +=
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>';

  for (let i = 1; i <= sheetCount; i++) {
    xml += `<Override PartName="/xl/worksheets/sheet${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
  }

  xml += '</Types>';
  return xml;
}

/**
 * Build _rels/.rels
 */
export function buildRootRels(): string {
  let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  xml +=
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
  xml +=
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>';
  xml += '</Relationships>';
  return xml;
}

/**
 * Build xl/_rels/workbook.xml.rels
 */
export function buildWorkbookRels(sheetCount: number): string {
  let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  xml +=
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';

  for (let i = 1; i <= sheetCount; i++) {
    xml += `<Relationship Id="rId${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i}.xml"/>`;
  }

  xml += `<Relationship Id="rId${sheetCount + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`;
  xml += `<Relationship Id="rId${sheetCount + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`;
  xml += '</Relationships>';
  return xml;
}

/**
 * Build xl/workbook.xml
 */
export function buildWorkbookXML(sheetNames: string[]): string {
  let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  xml +=
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
  xml += '<sheets>';

  for (let i = 0; i < sheetNames.length; i++) {
    xml += `<sheet name="${escapeXML(sheetNames[i])}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`;
  }

  xml += '</sheets>';
  xml += '</workbook>';
  return xml;
}

/**
 * Build xl/sharedStrings.xml
 */
export function buildSharedStrings(strings: string[]): string {
  let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  xml += `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.length}" uniqueCount="${strings.length}">`;

  for (const str of strings) {
    xml += `<si><t>${escapeXML(str)}</t></si>`;
  }

  xml += '</sst>';
  return xml;
}
