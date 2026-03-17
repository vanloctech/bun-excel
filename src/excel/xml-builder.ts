// ============================================
// XML Builder for XLSX generation
// ============================================

import type {
  DefinedName,
  HeaderFooter,
  PageMargins,
  PageSetup,
  RichTextRun,
  SplitPane,
  WorkbookView,
  WorksheetProtection,
  WorksheetState,
} from '../types';

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
export function buildContentTypes(
  sheetCount: number,
  options?: {
    commentsCount?: number;
    drawingsCount?: number;
    tablesCount?: number;
    mediaExtensions?: Iterable<string>;
    includeVml?: boolean;
  },
): string {
  let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  xml +=
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
  xml +=
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
  xml += '<Default Extension="xml" ContentType="application/xml"/>';
  if (options?.includeVml) {
    xml +=
      '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>';
  }
  const mediaExtensions = new Set(options?.mediaExtensions || []);
  for (const ext of mediaExtensions) {
    const normalized = ext === 'jpg' ? 'jpeg' : ext;
    const contentType =
      normalized === 'jpeg' ? 'image/jpeg' : `image/${normalized}`;
    xml += `<Default Extension="${escapeXML(ext)}" ContentType="${contentType}"/>`;
  }
  xml +=
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
  xml +=
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>';
  xml +=
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
  xml +=
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';
  xml +=
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>';

  for (let i = 1; i <= sheetCount; i++) {
    xml += `<Override PartName="/xl/worksheets/sheet${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
  }
  for (let i = 1; i <= (options?.commentsCount ?? 0); i++) {
    xml += `<Override PartName="/xl/comments${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>`;
  }
  for (let i = 1; i <= (options?.drawingsCount ?? 0); i++) {
    xml += `<Override PartName="/xl/drawings/drawing${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`;
  }
  for (let i = 1; i <= (options?.tablesCount ?? 0); i++) {
    xml += `<Override PartName="/xl/tables/table${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`;
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
  xml +=
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>';
  xml +=
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>';
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

function normalizeSheetEntry(
  sheet: string | { name: string; state?: WorksheetState },
): { name: string; state?: WorksheetState } {
  return typeof sheet === 'string' ? { name: sheet } : sheet;
}

function buildWorkbookViewXML(view?: WorkbookView): string {
  if (!view) return '<bookViews><workbookView activeTab="0"/></bookViews>';

  let xml = '<bookViews><workbookView';
  const activeTab = getNonNegativeIntegerOr(view.activeTab, 0);
  xml += ` activeTab="${activeTab}"`;
  if (view.firstSheet !== undefined) {
    xml += ` firstSheet="${getNonNegativeIntegerOr(view.firstSheet, 0)}"`;
  }
  if (view.visibility && view.visibility !== 'visible') {
    xml += ` visibility="${escapeXML(view.visibility)}"`;
  }
  if (view.xWindow !== undefined) {
    xml += ` xWindow="${getNonNegativeIntegerOr(view.xWindow, 0)}"`;
  }
  if (view.yWindow !== undefined) {
    xml += ` yWindow="${getNonNegativeIntegerOr(view.yWindow, 0)}"`;
  }
  if (view.windowWidth !== undefined) {
    xml += ` windowWidth="${getNonNegativeIntegerOr(view.windowWidth, 0)}"`;
  }
  if (view.windowHeight !== undefined) {
    xml += ` windowHeight="${getNonNegativeIntegerOr(view.windowHeight, 0)}"`;
  }
  xml += '/></bookViews>';
  return xml;
}

function buildDefinedNamesXML(definedNames?: DefinedName[]): string {
  if (!definedNames || definedNames.length === 0) return '';

  let xml = '<definedNames>';
  for (const definedName of definedNames) {
    let attrs = ` name="${escapeXML(definedName.name)}"`;
    if (definedName.localSheetId !== undefined) {
      attrs += ` localSheetId="${getNonNegativeIntegerOr(
        definedName.localSheetId,
        0,
      )}"`;
    }
    if (definedName.hidden) attrs += ' hidden="1"';
    if (definedName.comment) {
      attrs += ` comment="${escapeXML(definedName.comment)}"`;
    }
    xml += `<definedName${attrs}>${escapeXML(
      definedName.refersTo,
    )}</definedName>`;
  }
  xml += '</definedNames>';
  return xml;
}

/**
 * Build xl/workbook.xml
 */
export function buildWorkbookXML(
  sheets: Array<string | { name: string; state?: WorksheetState }>,
  options?: {
    definedNames?: DefinedName[];
    view?: WorkbookView;
  },
): string {
  let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  xml +=
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
  xml += buildWorkbookViewXML(options?.view);
  xml += '<sheets>';

  for (let i = 0; i < sheets.length; i++) {
    const sheet = normalizeSheetEntry(sheets[i]);
    xml += `<sheet name="${escapeXML(sheet.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"`;
    if (sheet.state && sheet.state !== 'visible') {
      xml += ` state="${escapeXML(sheet.state)}"`;
    }
    xml += '/>';
  }

  xml += '</sheets>';
  xml += buildDefinedNamesXML(options?.definedNames);
  xml += '</workbook>';
  return xml;
}

function formatW3CDate(value: Date): string {
  return value.toISOString();
}

export function buildCorePropsXML(metadata: {
  creator?: string;
  created?: Date;
  modified?: Date;
}): string {
  let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  xml +=
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">';

  if (metadata.creator) {
    xml += `<dc:creator>${escapeXML(metadata.creator)}</dc:creator>`;
    xml += `<cp:lastModifiedBy>${escapeXML(metadata.creator)}</cp:lastModifiedBy>`;
  }
  if (metadata.created) {
    xml += `<dcterms:created xsi:type="dcterms:W3CDTF">${formatW3CDate(metadata.created)}</dcterms:created>`;
  }
  if (metadata.modified) {
    xml += `<dcterms:modified xsi:type="dcterms:W3CDTF">${formatW3CDate(metadata.modified)}</dcterms:modified>`;
  }

  xml += '</cp:coreProperties>';
  return xml;
}

export function buildAppPropsXML(sheetNames: string[]): string {
  let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  xml +=
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">';
  xml += '<Application>bun-excel</Application>';
  xml += '<HeadingPairs>';
  xml += '<vt:vector size="2" baseType="variant">';
  xml += '<vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>';
  xml += `<vt:variant><vt:i4>${sheetNames.length}</vt:i4></vt:variant>`;
  xml += '</vt:vector>';
  xml += '</HeadingPairs>';
  xml += `<TitlesOfParts><vt:vector size="${sheetNames.length}" baseType="lpstr">`;
  for (const sheetName of sheetNames) {
    xml += `<vt:lpstr>${escapeXML(sheetName)}</vt:lpstr>`;
  }
  xml += '</vt:vector></TitlesOfParts>';
  xml += '</Properties>';
  return xml;
}

function getActivePane(xSplit: number, ySplit: number): string {
  if (xSplit > 0 && ySplit > 0) return 'bottomRight';
  if (xSplit > 0) return 'topRight';
  if (ySplit > 0) return 'bottomLeft';
  return 'topLeft';
}

export function buildSheetViewsXML(config: {
  freezePane?: { row: number; col: number };
  splitPane?: SplitPane;
}): string {
  if (config.freezePane) {
    const row = getNonNegativeIntegerOr(config.freezePane.row, 0);
    const col = getNonNegativeIntegerOr(config.freezePane.col, 0);
    const topLeftCell = buildCellRef(row, col);
    return `<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="${col}" ySplit="${row}" topLeftCell="${topLeftCell}" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>`;
  }

  if (config.splitPane) {
    const x = Math.max(0, getFiniteNumberOr(config.splitPane.x, 0));
    const y = Math.max(0, getFiniteNumberOr(config.splitPane.y, 0));
    if (x === 0 && y === 0) return '';

    let xml = '<sheetViews><sheetView tabSelected="1" workbookViewId="0">';
    xml += `<pane xSplit="${x}" ySplit="${y}" activePane="${getActivePane(x, y)}" state="split"`;
    if (config.splitPane.topLeftCell) {
      xml += ` topLeftCell="${buildCellRef(
        getNonNegativeIntegerOr(config.splitPane.topLeftCell.row, 0),
        getNonNegativeIntegerOr(config.splitPane.topLeftCell.col, 0),
      )}"`;
    }
    xml += '/></sheetView></sheetViews>';
    return xml;
  }

  return '';
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

function buildRichTextRunFontXML(font: RichTextRun['font']): string {
  if (!font) return '';

  let xml = '<rPr>';
  if (font.bold) xml += '<b/>';
  if (font.italic) xml += '<i/>';
  if (font.underline) xml += '<u/>';
  if (font.strike) xml += '<strike/>';
  if (font.size !== undefined) {
    xml += `<sz val="${getFiniteNumberOr(font.size, 11)}"/>`;
  }
  if (font.color) {
    xml += `<color rgb="FF${escapeXML(font.color)}"/>`;
  }
  if (font.name) {
    xml += `<rFont val="${escapeXML(font.name)}"/>`;
  }
  xml += '</rPr>';
  return xml;
}

export function buildRichTextXML(runs: RichTextRun[]): string {
  let xml = '<is>';
  for (const run of runs) {
    xml += '<r>';
    xml += buildRichTextRunFontXML(run.font);
    xml += `<t xml:space="preserve">${escapeXML(run.text)}</t>`;
    xml += '</r>';
  }
  xml += '</is>';
  return xml;
}

export function buildSheetPropertiesXML(hasOutlines: boolean): string {
  if (!hasOutlines) return '';
  return '<sheetPr><outlinePr summaryBelow="1" summaryRight="1"/></sheetPr>';
}

function hashSheetProtectionPassword(password: string): string {
  let hash = 0;
  for (let i = 0; i < password.length; i++) {
    let value = password.charCodeAt(i);
    // biome-ignore lint/suspicious/noBitwiseOperators: Excel sheet protection uses the legacy XOR/rotate hash algorithm.
    value = ((value << (i + 1)) | (value >> (15 - i))) & 0x7fff;
    // biome-ignore lint/suspicious/noBitwiseOperators: Excel sheet protection uses the legacy XOR/rotate hash algorithm.
    hash ^= value;
  }
  // biome-ignore lint/suspicious/noBitwiseOperators: Excel sheet protection uses the legacy XOR/rotate hash algorithm.
  hash ^= password.length;
  // biome-ignore lint/suspicious/noBitwiseOperators: Excel sheet protection uses the legacy XOR/rotate hash algorithm.
  hash ^= 0xce4b;
  return hash.toString(16).toUpperCase();
}

export function buildSheetProtectionXML(
  protection?: WorksheetProtection,
): string {
  if (!protection) return '';

  let xml = '<sheetProtection';
  if (protection.password) {
    xml += ` password="${hashSheetProtectionPassword(protection.password)}"`;
  }

  const attrs: Record<string, boolean | undefined> = {
    sheet: protection.sheet,
    objects: protection.objects,
    scenarios: protection.scenarios,
    formatCells: protection.formatCells,
    formatColumns: protection.formatColumns,
    formatRows: protection.formatRows,
    insertColumns: protection.insertColumns,
    insertRows: protection.insertRows,
    insertHyperlinks: protection.insertHyperlinks,
    deleteColumns: protection.deleteColumns,
    deleteRows: protection.deleteRows,
    selectLockedCells: protection.selectLockedCells,
    sort: protection.sort,
    autoFilter: protection.autoFilter,
    pivotTables: protection.pivotTables,
    selectUnlockedCells: protection.selectUnlockedCells,
  };

  for (const [key, value] of Object.entries(attrs)) {
    if (value !== undefined) {
      xml += ` ${key}="${value ? '1' : '0'}"`;
    }
  }

  xml += '/>';
  return xml;
}

export function buildPageMarginsXML(pageMargins?: PageMargins): string {
  if (!pageMargins) return '';
  return tag(
    'pageMargins',
    {
      left: getFiniteNumberOr(pageMargins.left, 0.7),
      right: getFiniteNumberOr(pageMargins.right, 0.7),
      top: getFiniteNumberOr(pageMargins.top, 0.75),
      bottom: getFiniteNumberOr(pageMargins.bottom, 0.75),
      header: getFiniteNumberOr(pageMargins.header, 0.3),
      footer: getFiniteNumberOr(pageMargins.footer, 0.3),
    },
    undefined,
    true,
  );
}

export function buildPageSetupXML(pageSetup?: PageSetup): string {
  if (!pageSetup) return '';

  const attrs: Record<string, string | number | boolean | undefined> = {};
  if (pageSetup.orientation) attrs.orientation = pageSetup.orientation;
  if (pageSetup.paperSize !== undefined) {
    attrs.paperSize = getNonNegativeIntegerOr(pageSetup.paperSize, 0);
  }
  if (pageSetup.scale !== undefined) {
    attrs.scale = getNonNegativeIntegerOr(pageSetup.scale, 100);
  }
  if (pageSetup.fitToWidth !== undefined) {
    attrs.fitToWidth = getNonNegativeIntegerOr(pageSetup.fitToWidth, 1);
  }
  if (pageSetup.fitToHeight !== undefined) {
    attrs.fitToHeight = getNonNegativeIntegerOr(pageSetup.fitToHeight, 1);
  }
  if (pageSetup.firstPageNumber !== undefined) {
    attrs.firstPageNumber = getNonNegativeIntegerOr(
      pageSetup.firstPageNumber,
      1,
    );
  }
  if (pageSetup.useFirstPageNumber !== undefined) {
    attrs.useFirstPageNumber = pageSetup.useFirstPageNumber ? 1 : 0;
  }
  return tag('pageSetup', attrs, undefined, true);
}

function buildHeaderFooterSectionXML(
  prefix: string,
  section?: { left?: string; center?: string; right?: string },
): string {
  if (!section) return '';
  const text = `${section.left ? `&L${section.left}` : ''}${
    section.center ? `&C${section.center}` : ''
  }${section.right ? `&R${section.right}` : ''}`;
  if (!text) return '';
  return `<${prefix}>${escapeXML(text)}</${prefix}>`;
}

export function buildHeaderFooterXML(headerFooter?: HeaderFooter): string {
  if (!headerFooter) return '';

  let xml = '<headerFooter';
  if (headerFooter.differentFirst) xml += ' differentFirst="1"';
  if (headerFooter.differentOddEven) xml += ' differentOddEven="1"';
  xml += '>';
  xml += buildHeaderFooterSectionXML('oddHeader', headerFooter.oddHeader);
  xml += buildHeaderFooterSectionXML('oddFooter', headerFooter.oddFooter);
  xml += buildHeaderFooterSectionXML('evenHeader', headerFooter.evenHeader);
  xml += buildHeaderFooterSectionXML('evenFooter', headerFooter.evenFooter);
  xml += buildHeaderFooterSectionXML('firstHeader', headerFooter.firstHeader);
  xml += buildHeaderFooterSectionXML('firstFooter', headerFooter.firstFooter);
  xml += '</headerFooter>';
  return xml;
}
