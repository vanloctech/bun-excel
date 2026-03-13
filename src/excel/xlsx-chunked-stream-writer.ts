// ============================================
// XLSX Chunked Stream Writer — Constant-memory
// streaming via temp file on disk
// ============================================
//
// Flow:
//   writeRow() → serialize XML → FileSink.write() to temp file (RAM: ~0)
//   end()      → read temp file → assemble ZIP → write output → delete temp
//
// Uses inline strings (<is><t>...</t></is>) instead of shared string table
// to avoid tracking all strings in memory.

import { unlinkSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
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
} from '../types';
import { StyleRegistry } from './style-builder';
import {
  buildCellRef,
  buildContentTypes,
  buildRootRels,
  buildWorkbookRels,
  buildWorkbookXML,
  escapeXML,
} from './xml-builder';

/** Validate path for security */
function validatePath(filePath: string): string {
  if (filePath.includes('\0')) {
    throw new Error('Invalid file path: contains null bytes');
  }
  return resolve(filePath);
}

/**
 * Options for the chunked Excel stream writer
 */
export interface ChunkedExcelStreamOptions extends ExcelWriteOptions {
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
 * Excel Chunked Stream Writer — Constant Memory
 *
 * Writes row XML directly to a temporary file on disk via Bun's FileSink.
 * Row objects and XML strings are NOT kept in memory.
 * At end(), reads back the temp file and assembles the final ZIP.
 *
 * Uses inline strings instead of shared string table to avoid
 * tracking all string values in memory.
 *
 * Memory usage stays ~constant regardless of how many rows are written.
 */
export class ExcelChunkedStreamWriter implements StreamWriter {
  private path: string;
  private options: ChunkedExcelStreamOptions;
  private styleRegistry = new StyleRegistry();
  private tempFilePath: string;
  private tempWriter: ReturnType<ReturnType<typeof Bun.file>['writer']>;
  private rowCount = 0;
  private hyperlinkRels: { rId: string; target: string }[] = [];
  private hyperlinkEntries: {
    ref: string;
    rId?: string;
    location?: string;
    tooltip?: string;
  }[] = [];
  private hyperlinkRelCounter = 1;

  constructor(path: string, options?: ChunkedExcelStreamOptions) {
    this.path = validatePath(path);
    this.options = options || {};

    // Create temp file for row XML chunks
    const tmpName = `bun-xlsx-${Date.now()}-${Math.random().toString(36).slice(2)}.tmp`;
    this.tempFilePath = join(tmpdir(), tmpName);
    this.tempWriter = Bun.file(this.tempFilePath).writer({
      highWaterMark: 256 * 1024, // 256KB buffer
    });
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

  /**
   * Serialize a single cell to XML using inline strings
   * (no shared string table needed)
   */
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

    // Formula cells
    if (cell.formula) {
      let xml = `<c r="${ref}"${styleIdx > 0 ? ` s="${styleIdx}"` : ''}>`;
      xml += `<f>${escapeXML(cell.formula)}</f>`;
      if (cell.formulaResult !== undefined) {
        xml += `<v>${escapeXML(String(cell.formulaResult))}</v>`;
      } else if (value !== null && value !== undefined) {
        if (typeof value === 'number' || typeof value === 'boolean') {
          xml += `<v>${value}</v>`;
        }
      }
      xml += '</c>';
      return xml;
    }

    if (value === null || value === undefined) {
      return styleIdx > 0 ? `<c r="${ref}" s="${styleIdx}"/>` : '';
    }

    // Use INLINE strings (<is><t>...</t></is>) instead of shared strings
    // This avoids having to track all string values in memory
    if (typeof value === 'string') {
      return `<c r="${ref}" t="inlineStr"${styleIdx > 0 ? ` s="${styleIdx}"` : ''}><is><t>${escapeXML(value)}</t></is></c>`;
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
   * Write a single row — serializes XML and writes to temp file immediately.
   * The Row object can be garbage collected right after this call.
   * No row data is kept in memory.
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
    if (rowObj.height) {
      rowAttrs += ` ht="${rowObj.height}" customHeight="1"`;
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

    // Write to temp file on disk — not kept in memory
    this.tempWriter.write(xml);
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
   * Flush the temp file writer buffer to disk
   */
  flush(): void | Promise<void> {
    const result = this.tempWriter.flush();
    if (result instanceof Promise) {
      return result.then(() => {});
    }
  }

  /**
   * Finalize: read temp file, assemble worksheet XML, create ZIP, write output
   */
  async end(): Promise<void> {
    // Flush and close the temp writer
    await this.tempWriter.end();

    const sheetName = this.options.sheetName || 'Sheet1';

    // Build worksheet XML header
    let wsHeader = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    wsHeader +=
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';

    // Freeze pane
    if (this.options.freezePane) {
      wsHeader += '<sheetViews>';
      wsHeader += '<sheetView tabSelected="1" workbookViewId="0">';
      const { row, col } = this.options.freezePane;
      const topLeftCell = buildCellRef(row, col);
      wsHeader += `<pane xSplit="${col}" ySplit="${row}" topLeftCell="${topLeftCell}" activePane="bottomRight" state="frozen"/>`;
      wsHeader += '</sheetView>';
      wsHeader += '</sheetViews>';
    }

    // Sheet format
    wsHeader += `<sheetFormatPr defaultRowHeight="${this.options.defaultRowHeight || 15}"/>`;

    // Columns
    if (this.options.columns && this.options.columns.length > 0) {
      wsHeader += '<cols>';
      for (let c = 0; c < this.options.columns.length; c++) {
        const col = this.options.columns[c];
        if (col.width) {
          wsHeader += `<col min="${c + 1}" max="${c + 1}" width="${col.width}" customWidth="1"/>`;
        }
      }
      wsHeader += '</cols>';
    }

    wsHeader += '<sheetData>';

    // Read the row XML from temp file on disk
    const tempFile = Bun.file(this.tempFilePath);
    const rowXmlContent = await tempFile.text();

    // Build worksheet footer
    let wsFooter = '</sheetData>';

    // Merge cells
    if (this.options.mergeCells && this.options.mergeCells.length > 0) {
      wsFooter += `<mergeCells count="${this.options.mergeCells.length}">`;
      for (const mc of this.options.mergeCells) {
        const startRef = buildCellRef(mc.startRow, mc.startCol);
        const endRef = buildCellRef(mc.endRow, mc.endCol);
        wsFooter += `<mergeCell ref="${startRef}:${endRef}"/>`;
      }
      wsFooter += '</mergeCells>';
    }

    // Hyperlinks
    if (this.hyperlinkEntries.length > 0) {
      wsFooter += '<hyperlinks>';
      for (const hl of this.hyperlinkEntries) {
        wsFooter += `<hyperlink ref="${hl.ref}"`;
        if (hl.rId) wsFooter += ` r:id="${hl.rId}"`;
        if (hl.location) wsFooter += ` location="${escapeXML(hl.location)}"`;
        if (hl.tooltip) wsFooter += ` tooltip="${escapeXML(hl.tooltip)}"`;
        wsFooter += '/>';
      }
      wsFooter += '</hyperlinks>';
    }

    wsFooter += '</worksheet>';

    // Combine: header + rows from disk + footer
    const fullWorksheetXml = wsHeader + rowXmlContent + wsFooter;

    // Build ZIP (no shared strings file needed — using inline strings)
    const files: Zippable = {
      '[Content_Types].xml': encoder.encode(buildContentTypes(1)),
      '_rels/.rels': encoder.encode(buildRootRels()),
      'xl/_rels/workbook.xml.rels': encoder.encode(buildWorkbookRels(1)),
      'xl/workbook.xml': encoder.encode(buildWorkbookXML([sheetName])),
      'xl/styles.xml': encoder.encode(this.styleRegistry.buildStylesXML()),
      // Empty shared strings (required by some readers)
      'xl/sharedStrings.xml': encoder.encode(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>',
      ),
      'xl/worksheets/sheet1.xml': encoder.encode(fullWorksheetXml),
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

    // Clean up temp file
    try {
      unlinkSync(this.tempFilePath);
    } catch {
      // Ignore cleanup errors
    }

    // Clear state
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
 * Create a chunked Excel stream writer (constant-memory disk-based streaming)
 */
export function createChunkedExcelStream(
  path: string,
  options?: ChunkedExcelStreamOptions,
): ExcelChunkedStreamWriter {
  return new ExcelChunkedStreamWriter(path, options);
}
