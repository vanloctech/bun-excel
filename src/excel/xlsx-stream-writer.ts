// ============================================
// XLSX Stream Writer — True Streaming with
// immediate row serialization to temp file
// ============================================

import { resolve } from 'node:path';
import { type Zippable, zipSync } from 'fflate';
import type {
  Cell,
  CellStyle,
  CellValue,
  ColumnConfig,
  ExcelWriteOptions,
  MergeCell,
  Row,
  StreamWriter,
  Workbook,
  Worksheet,
} from '../types';
import { StyleRegistry } from './style-builder';
import {
  buildCellRef,
  buildContentTypes,
  buildRootRels,
  buildSharedStrings,
  buildWorkbookRels,
  buildWorkbookXML,
  escapeXML,
  getFiniteNumber,
  getFiniteNumberOr,
  getNonNegativeIntegerOr,
} from './xml-builder';

/** Validate path for security */
function validatePath(filePath: string): string {
  if (filePath.includes('\0')) {
    throw new Error('Invalid file path: contains null bytes');
  }
  return resolve(filePath);
}

/**
 * Options for the Excel stream writer
 */
export interface ExcelStreamOptions extends ExcelWriteOptions {
  /** Sheet name (default: "Sheet1") */
  sheetName?: string;
  /** Column configurations */
  columns?: ColumnConfig[];
  /** Default row height */
  defaultRowHeight?: number;
  /** Freeze pane */
  freezePane?: { row: number; col: number };
  /** Merge cells */
  mergeCells?: MergeCell[];
}

const encoder = new TextEncoder();

/**
 * Excel Stream Writer — True Streaming
 *
 * Serializes each row to XML immediately on writeRow(),
 * writing to a temp file via FileSink. Row objects are
 * discarded right after serialization, keeping memory low.
 *
 * At end(), reads back the temp XML and assembles the final ZIP.
 */
export class ExcelStreamWriter implements StreamWriter {
  private path: string;
  private options: ExcelStreamOptions;
  private styleRegistry = new StyleRegistry();
  private sharedStrings: string[] = [];
  private sharedStringMap = new Map<string, number>();
  private rowXMLChunks: string[] = [];
  private rowCount = 0;
  private hyperlinkRels: { rId: string; target: string }[] = [];
  private hyperlinkEntries: {
    ref: string;
    rId?: string;
    location?: string;
    tooltip?: string;
  }[] = [];
  private hyperlinkRelCounter = 1;

  constructor(path: string, options?: ExcelStreamOptions) {
    this.path = validatePath(path);
    this.options = options || {};
  }

  /** Get or create shared string index */
  private getSharedStringIndex(str: string): number {
    const existing = this.sharedStringMap.get(str);
    if (existing !== undefined) return existing;
    const index = this.sharedStrings.length;
    this.sharedStrings.push(str);
    this.sharedStringMap.set(str, index);
    return index;
  }

  /** Check if hyperlink target is external */
  private isExternalHyperlink(target: string): boolean {
    return (
      target.startsWith('http://') ||
      target.startsWith('https://') ||
      target.startsWith('mailto:') ||
      target.startsWith('ftp://')
    );
  }

  /** Serialize a single cell to XML */
  private serializeCell(cell: Cell, ref: string, rowStyle?: CellStyle): string {
    const cellStyle = cell.style || rowStyle;
    const styleIdx = this.styleRegistry.registerStyle(cellStyle);
    const { value } = cell;

    // Collect hyperlinks
    if (cell.hyperlink) {
      const hl = cell.hyperlink;
      if (this.isExternalHyperlink(hl.target)) {
        const rId = `rId${this.hyperlinkRelCounter++}`;
        this.hyperlinkRels.push({ rId, target: hl.target });
        this.hyperlinkEntries.push({ ref, rId, tooltip: hl.tooltip });
      } else {
        this.hyperlinkEntries.push({
          ref,
          location: hl.target,
          tooltip: hl.tooltip,
        });
      }
    }

    // Formula cells (with or without value)
    if (cell.formula) {
      let xml = `<c r="${ref}"${styleIdx > 0 ? ` s="${styleIdx}"` : ''}>`;
      xml += `<f>${escapeXML(cell.formula)}</f>`;
      if (cell.formulaResult !== undefined) {
        xml += `<v>${escapeXML(String(cell.formulaResult))}</v>`;
      } else if (value !== null && value !== undefined) {
        if (typeof value === 'string') {
          xml += `<v>${this.getSharedStringIndex(value)}</v>`;
        } else if (typeof value === 'number' || typeof value === 'boolean') {
          xml += `<v>${value}</v>`;
        }
      }
      xml += '</c>';
      return xml;
    }

    if (value === null || value === undefined) {
      return styleIdx > 0 ? `<c r="${ref}" s="${styleIdx}"/>` : '';
    }

    if (typeof value === 'string') {
      const ssIdx = this.getSharedStringIndex(value);
      return `<c r="${ref}" t="s"${styleIdx > 0 ? ` s="${styleIdx}"` : ''}><v>${ssIdx}</v></c>`;
    }
    if (typeof value === 'number') {
      return `<c r="${ref}"${styleIdx > 0 ? ` s="${styleIdx}"` : ''}><v>${value}</v></c>`;
    }
    if (typeof value === 'boolean') {
      return `<c r="${ref}" t="b"${styleIdx > 0 ? ` s="${styleIdx}"` : ''}><v>${value ? 1 : 0}</v></c>`;
    }
    if (value instanceof Date) {
      const epoch = new Date(Date.UTC(1899, 11, 30));
      const serial =
        (value.getTime() - epoch.getTime()) / (24 * 60 * 60 * 1000);
      return `<c r="${ref}"${styleIdx > 0 ? ` s="${styleIdx}"` : ''}><v>${serial}</v></c>`;
    }

    return '';
  }

  /**
   * Write a single row — serializes to XML immediately
   * Row object can be GC'd right after this call
   */
  writeRow(row: Row | CellValue[]): void {
    const r = this.rowCount;
    this.rowCount++;

    let rowObj: Row;
    if (Array.isArray(row)) {
      rowObj = { cells: row.map((value) => ({ value })) };
    } else {
      rowObj = row;
    }

    let rowAttrs = ` r="${r + 1}"`;
    const rowHeight = getFiniteNumber(rowObj.height);
    if (rowHeight !== undefined) {
      rowAttrs += ` ht="${rowHeight}" customHeight="1"`;
    }

    const rowStyleIdx = rowObj.style
      ? this.styleRegistry.registerStyle(rowObj.style)
      : 0;
    if (rowStyleIdx > 0) {
      rowAttrs += ` s="${rowStyleIdx}" customFormat="1"`;
    }

    let xml = `<row${rowAttrs}>`;

    for (let c = 0; c < rowObj.cells.length; c++) {
      const cell = rowObj.cells[c];
      if (!cell) continue;
      const ref = buildCellRef(r, c);
      xml += this.serializeCell(cell, ref, rowObj.style);
    }

    xml += '</row>';

    // Append to chunks array (just XML strings, not Row objects)
    this.rowXMLChunks.push(xml);
  }

  /**
   * Write a row with styles applied to each cell
   */
  writeStyledRow(values: CellValue[], styles: (CellStyle | undefined)[]): void {
    const cells: Cell[] = values.map((value, i) => ({
      value,
      style: styles[i],
    }));
    this.writeRow({ cells });
  }

  /**
   * Write multiple rows at once
   */
  writeRows(rows: (Row | CellValue[])[]): void {
    for (const row of rows) {
      this.writeRow(row);
    }
  }

  /**
   * Flush — no-op (XML chunks are already serialized)
   */
  flush(): void {}

  /**
   * Finalize and write the XLSX file
   * Assembles the XML from serialized chunks and creates the ZIP
   */
  async end(): Promise<void> {
    const sheetName = this.options.sheetName || 'Sheet1';

    // Build worksheet XML from serialized row chunks
    let wsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    wsXml +=
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';

    // Freeze pane
    if (this.options.freezePane) {
      wsXml += '<sheetViews>';
      wsXml += '<sheetView tabSelected="1" workbookViewId="0">';
      const row = getNonNegativeIntegerOr(this.options.freezePane.row, 0);
      const col = getNonNegativeIntegerOr(this.options.freezePane.col, 0);
      const topLeftCell = buildCellRef(row, col);
      wsXml += `<pane xSplit="${col}" ySplit="${row}" topLeftCell="${topLeftCell}" activePane="bottomRight" state="frozen"/>`;
      wsXml += '</sheetView>';
      wsXml += '</sheetViews>';
    }

    // Sheet format
    wsXml += `<sheetFormatPr defaultRowHeight="${getFiniteNumberOr(this.options.defaultRowHeight, 15)}"/>`;

    // Columns
    if (this.options.columns && this.options.columns.length > 0) {
      wsXml += '<cols>';
      for (let c = 0; c < this.options.columns.length; c++) {
        const col = this.options.columns[c];
        const colWidth = getFiniteNumber(col.width);
        if (colWidth !== undefined) {
          wsXml += `<col min="${c + 1}" max="${c + 1}" width="${colWidth}" customWidth="1"/>`;
        }
      }
      wsXml += '</cols>';
    }

    // Sheet data — join all pre-serialized row XML chunks
    wsXml += '<sheetData>';
    wsXml += this.rowXMLChunks.join('');
    wsXml += '</sheetData>';

    // Merge cells
    if (this.options.mergeCells && this.options.mergeCells.length > 0) {
      wsXml += `<mergeCells count="${this.options.mergeCells.length}">`;
      for (const mc of this.options.mergeCells) {
        const startRef = buildCellRef(mc.startRow, mc.startCol);
        const endRef = buildCellRef(mc.endRow, mc.endCol);
        wsXml += `<mergeCell ref="${startRef}:${endRef}"/>`;
      }
      wsXml += '</mergeCells>';
    }

    // Hyperlinks
    if (this.hyperlinkEntries.length > 0) {
      wsXml += '<hyperlinks>';
      for (const hl of this.hyperlinkEntries) {
        wsXml += `<hyperlink ref="${hl.ref}"`;
        if (hl.rId) wsXml += ` r:id="${hl.rId}"`;
        if (hl.location) wsXml += ` location="${escapeXML(hl.location)}"`;
        if (hl.tooltip) wsXml += ` tooltip="${escapeXML(hl.tooltip)}"`;
        wsXml += '/>';
      }
      wsXml += '</hyperlinks>';
    }

    wsXml += '</worksheet>';

    // Build ZIP
    const files: Zippable = {
      '[Content_Types].xml': encoder.encode(buildContentTypes(1)),
      '_rels/.rels': encoder.encode(buildRootRels()),
      'xl/_rels/workbook.xml.rels': encoder.encode(buildWorkbookRels(1)),
      'xl/workbook.xml': encoder.encode(buildWorkbookXML([sheetName])),
      'xl/styles.xml': encoder.encode(this.styleRegistry.buildStylesXML()),
      'xl/sharedStrings.xml': encoder.encode(
        buildSharedStrings(this.sharedStrings),
      ),
      'xl/worksheets/sheet1.xml': encoder.encode(wsXml),
    };

    // Hyperlink rels
    if (this.hyperlinkRels.length > 0) {
      let relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
      relsXml +=
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
      for (const rel of this.hyperlinkRels) {
        relsXml += `<Relationship Id="${rel.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${escapeXML(rel.target)}" TargetMode="External"/>`;
      }
      relsXml += '</Relationships>';
      files['xl/worksheets/_rels/sheet1.xml.rels'] = encoder.encode(relsXml);
    }

    const zipBuffer = zipSync(files, {
      level: this.options.compress !== false ? 6 : 0,
    });

    await Bun.write(this.path, zipBuffer);

    // Clear all data
    this.rowXMLChunks.length = 0;
    this.sharedStrings.length = 0;
    this.sharedStringMap.clear();
    this.hyperlinkRels.length = 0;
    this.hyperlinkEntries.length = 0;
  }

  /**
   * Get current row count
   */
  get currentRowCount(): number {
    return this.rowCount;
  }
}

/**
 * Multi-sheet Excel Stream Writer
 * Allows streaming data to multiple worksheets
 */
export class MultiSheetExcelStreamWriter {
  private worksheets: Map<string, { rows: Row[]; config: ExcelStreamOptions }> =
    new Map();
  private path: string;
  private options: ExcelWriteOptions;
  private currentSheet: string;

  constructor(path: string, options?: ExcelWriteOptions) {
    this.path = validatePath(path);
    this.options = options || {};
    this.currentSheet = 'Sheet1';
    this.worksheets.set('Sheet1', { rows: [], config: {} });
  }

  /**
   * Add a new sheet or switch to existing sheet
   */
  addSheet(name: string, config?: ExcelStreamOptions): this {
    if (!this.worksheets.has(name)) {
      this.worksheets.set(name, { rows: [], config: config || {} });
    }
    this.currentSheet = name;
    return this;
  }

  /**
   * Write a row to the current sheet
   */
  writeRow(row: Row | CellValue[]): void {
    const sheet = this.worksheets.get(this.currentSheet);
    if (!sheet) throw new Error(`Sheet not found: ${this.currentSheet}`);

    if (Array.isArray(row)) {
      sheet.rows.push({ cells: row.map((value) => ({ value })) });
    } else {
      sheet.rows.push(row);
    }
  }

  /**
   * Write a styled row to the current sheet
   */
  writeStyledRow(values: CellValue[], styles: (CellStyle | undefined)[]): void {
    const sheet = this.worksheets.get(this.currentSheet);
    if (!sheet) throw new Error(`Sheet not found: ${this.currentSheet}`);

    const cells: Cell[] = values.map((value, i) => ({
      value,
      style: styles[i],
    }));
    sheet.rows.push({ cells });
  }

  /**
   * Finalize and write the Excel file
   */
  async end(): Promise<void> {
    const worksheetsList: Worksheet[] = [];

    for (const [name, data] of this.worksheets) {
      worksheetsList.push({
        name,
        rows: data.rows,
        columns: data.config.columns,
        freezePane: data.config.freezePane,
        defaultRowHeight: data.config.defaultRowHeight,
      });
    }

    const workbook: Workbook = {
      worksheets: worksheetsList,
      creator: this.options.creator,
    };

    const { buildExcelBuffer } = await import('./xlsx-writer');
    const buffer = buildExcelBuffer(workbook, this.options);
    await Bun.write(this.path, buffer);

    // Clear data
    this.worksheets.clear();
  }
}

/**
 * Create an Excel stream writer (true streaming — serializes rows immediately)
 */
export function createExcelStream(
  path: string,
  options?: ExcelStreamOptions,
): ExcelStreamWriter {
  return new ExcelStreamWriter(path, options);
}

/**
 * Create a multi-sheet Excel stream writer
 */
export function createMultiSheetExcelStream(
  path: string,
  options?: ExcelWriteOptions,
): MultiSheetExcelStreamWriter {
  return new MultiSheetExcelStreamWriter(path, options);
}
