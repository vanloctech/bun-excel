// ============================================
// XLSX Stream Writer — Bun-native disk-backed
// streaming via FileSink/temp files
// ============================================

import { renameSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { toWriteTarget } from '../runtime-io';
import type {
  Cell,
  CellRange,
  CellStyle,
  CellValue,
  ColumnConfig,
  ConditionalFormatting,
  DataValidation,
  ExcelWriteOptions,
  FileTarget,
  MergeCell,
  Row,
  StreamWriter,
  Worksheet,
} from '../types';
import { buildAutoFilterXML } from './auto-filter';
import { buildConditionalFormattingsXML } from './conditional-formatting';
import { buildDataValidationsXML } from './data-validation';
import { ManagedFileSink } from './file-sink';
import { createTempRuntimeId } from './runtime-utils';
import { StyleRegistry } from './style-builder';
import { ExcelChunkedStreamWriter } from './xlsx-chunked-stream-writer';
import {
  buildAppPropsXML,
  buildCellRef,
  buildContentTypes,
  buildCorePropsXML,
  buildRootRels,
  buildSheetViewsXML,
  buildWorkbookRels,
  buildWorkbookXML,
  escapeXML,
  getFiniteNumber,
  getFiniteNumberOr,
} from './xml-builder';
import { StreamingZipWriter } from './zip-stream';

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
  /** Split pane */
  splitPane?: Worksheet['splitPane'];
  /** Merge cells */
  mergeCells?: MergeCell[];
  /** Auto filter range */
  autoFilter?: CellRange;
  /** Conditional formatting rules */
  conditionalFormattings?: ConditionalFormatting[];
  /** Data validation rules */
  dataValidations?: DataValidation[];
}

function createTempFilePath(prefix: string): string {
  return join(tmpdir(), `${prefix}-${createTempRuntimeId()}.tmp`);
}

function createOutputTempPath(outputPath: string): string {
  return join(
    dirname(outputPath),
    `.bun-spreadsheet-${createTempRuntimeId()}.tmp`,
  );
}

class DiskBackedWorksheetWriter {
  private options: ExcelStreamOptions;
  private readonly styleRegistry: StyleRegistry;
  private readonly rowTempFilePath: string;
  private readonly rowTempWriter: ManagedFileSink;
  private readonly hyperlinkTempFilePath: string;
  private readonly hyperlinkTempWriter: ManagedFileSink;
  private readonly hyperlinkRelTempFilePath: string;
  private readonly hyperlinkRelTempWriter: ManagedFileSink;
  private rowCount = 0;
  private hyperlinkCount = 0;
  private externalHyperlinkCount = 0;
  private hyperlinkRelCounter = 1;
  private closed = false;

  constructor(options: ExcelStreamOptions, styleRegistry: StyleRegistry) {
    this.options = { ...options };
    this.styleRegistry = styleRegistry;
    this.rowTempFilePath = createTempFilePath('bun-xlsx-rows');
    this.rowTempWriter = new ManagedFileSink(this.rowTempFilePath);
    this.hyperlinkTempFilePath = createTempFilePath('bun-xlsx-links');
    this.hyperlinkTempWriter = new ManagedFileSink(this.hyperlinkTempFilePath);
    this.hyperlinkRelTempFilePath = createTempFilePath('bun-xlsx-link-rels');
    this.hyperlinkRelTempWriter = new ManagedFileSink(
      this.hyperlinkRelTempFilePath,
    );
  }

  updateOptions(options?: ExcelStreamOptions): void {
    if (!options) return;
    this.options = { ...this.options, ...options };
  }

  private isExternalHyperlink(target: string): boolean {
    return (
      target.startsWith('http://') ||
      target.startsWith('https://') ||
      target.startsWith('mailto:') ||
      target.startsWith('ftp://')
    );
  }

  private writeHyperlink(ref: string, target: string, tooltip?: string): void {
    this.hyperlinkCount++;

    let hyperlinkXml = `<hyperlink ref="${ref}"`;
    if (tooltip) {
      hyperlinkXml += ` tooltip="${escapeXML(tooltip)}"`;
    }

    if (this.isExternalHyperlink(target)) {
      const rId = `rId${this.hyperlinkRelCounter++}`;
      hyperlinkXml += ` r:id="${rId}"/>`;
      this.hyperlinkTempWriter.write(hyperlinkXml);
      this.hyperlinkRelTempWriter.write(
        `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${escapeXML(target)}" TargetMode="External"/>`,
      );
      this.externalHyperlinkCount++;
      return;
    }

    hyperlinkXml += ` location="${escapeXML(target)}"/>`;
    this.hyperlinkTempWriter.write(hyperlinkXml);
  }

  private serializeCell(cell: Cell, ref: string, rowStyle?: CellStyle): string {
    const cellStyle = cell.style || rowStyle;
    const styleIdx = this.styleRegistry.registerStyle(cellStyle);
    const { value } = cell;

    if (cell.hyperlink) {
      this.writeHyperlink(ref, cell.hyperlink.target, cell.hyperlink.tooltip);
    }

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

  private buildWorksheetPrefix(): string {
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml +=
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
    xml += buildSheetViewsXML({
      freezePane: this.options.freezePane,
      splitPane: this.options.splitPane,
    });
    xml += `<sheetFormatPr defaultRowHeight="${getFiniteNumberOr(this.options.defaultRowHeight, 15)}"/>`;

    if (this.options.columns && this.options.columns.length > 0) {
      xml += '<cols>';
      for (let c = 0; c < this.options.columns.length; c++) {
        const col = this.options.columns[c];
        const colWidth = getFiniteNumber(col.width);
        if (colWidth !== undefined) {
          xml += `<col min="${c + 1}" max="${c + 1}" width="${colWidth}" customWidth="1"/>`;
        }
      }
      xml += '</cols>';
    }

    xml += '<sheetData>';
    return xml;
  }

  private buildWorksheetSuffix(): string[] {
    const parts: string[] = ['</sheetData>'];

    if (this.options.mergeCells && this.options.mergeCells.length > 0) {
      parts.push(`<mergeCells count="${this.options.mergeCells.length}">`);
      for (const mc of this.options.mergeCells) {
        const startRef = buildCellRef(mc.startRow, mc.startCol);
        const endRef = buildCellRef(mc.endRow, mc.endCol);
        parts.push(`<mergeCell ref="${startRef}:${endRef}"/>`);
      }
      parts.push('</mergeCells>');
    }

    const autoFilterXml = buildAutoFilterXML(this.options.autoFilter);
    if (autoFilterXml) {
      parts.push(autoFilterXml);
    }

    const conditionalFormattingXml = buildConditionalFormattingsXML(
      this.options.conditionalFormattings,
      this.styleRegistry,
    );
    if (conditionalFormattingXml) {
      parts.push(conditionalFormattingXml);
    }

    const dataValidationsXml = buildDataValidationsXML(
      this.options.dataValidations,
    );
    if (dataValidationsXml) {
      parts.push(dataValidationsXml);
    }

    parts.push('</worksheet>');
    return parts;
  }

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
    this.rowTempWriter.write(xml);
  }

  writeStyledRow(values: CellValue[], styles: (CellStyle | undefined)[]): void {
    const cells: Cell[] = values.map((value, i) => ({
      value,
      style: styles[i],
    }));
    this.writeRow({ cells });
  }

  writeRows(rows: (Row | CellValue[])[]): void {
    for (const row of rows) {
      this.writeRow(row);
    }
  }

  flush(): Promise<void> {
    return Promise.all([
      this.rowTempWriter.flush(),
      this.hyperlinkTempWriter.flush(),
      this.hyperlinkRelTempWriter.flush(),
    ]).then(() => {});
  }

  async close(): Promise<void> {
    if (this.closed) {
      return;
    }
    this.closed = true;
    await Promise.all([
      this.rowTempWriter.end(),
      this.hyperlinkTempWriter.end(),
      this.hyperlinkRelTempWriter.end(),
    ]);
  }

  buildWorksheetParts(): (string | Blob)[] {
    const parts: (string | Blob)[] = [
      this.buildWorksheetPrefix(),
      Bun.file(this.rowTempFilePath),
    ];
    const suffixParts = this.buildWorksheetSuffix();

    if (this.hyperlinkCount > 0) {
      const closingTag = suffixParts.pop();
      if (closingTag) {
        parts.push(...suffixParts, '<hyperlinks>');
        parts.push(Bun.file(this.hyperlinkTempFilePath));
        parts.push('</hyperlinks>', closingTag);
        return parts;
      }
    }

    parts.push(...suffixParts);
    return parts;
  }

  buildWorksheetRelParts(): (string | Blob)[] | undefined {
    if (this.externalHyperlinkCount === 0) {
      return undefined;
    }

    return [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      Bun.file(this.hyperlinkRelTempFilePath),
      '</Relationships>',
    ];
  }

  async cleanup(): Promise<void> {
    await Promise.all(
      [
        this.rowTempFilePath,
        this.hyperlinkTempFilePath,
        this.hyperlinkRelTempFilePath,
      ].map(async (filePath) => {
        try {
          await Bun.file(filePath).delete();
        } catch {
          // Ignore cleanup errors
        }
      }),
    );
  }
}

/**
 * Excel Stream Writer — Bun-native disk-backed streaming
 *
 * Delegates to the disk-backed chunked writer so the public
 * createExcelStream() API also uses Bun FileSink/temp files
 * instead of keeping row XML in memory.
 */
export class ExcelStreamWriter implements StreamWriter {
  private readonly writer: ExcelChunkedStreamWriter;

  constructor(target: FileTarget, options?: ExcelStreamOptions) {
    this.writer = new ExcelChunkedStreamWriter(target, options);
  }

  /**
   * Write a single row
   */
  writeRow(row: Row | CellValue[]): void {
    this.writer.writeRow(row);
  }

  /**
   * Write a row with styles applied to each cell
   */
  writeStyledRow(values: CellValue[], styles: (CellStyle | undefined)[]): void {
    this.writer.writeStyledRow(values, styles);
  }

  /**
   * Write multiple rows at once
   */
  writeRows(rows: (Row | CellValue[])[]): void {
    this.writer.writeRows(rows);
  }

  /**
   * Flush buffered temp-file writes.
   */
  flush(): Promise<void> {
    return this.writer.flush();
  }

  /**
   * Finalize and write the XLSX file.
   */
  async end(): Promise<void> {
    await this.writer.end();
  }

  /**
   * Get current row count
   */
  get currentRowCount(): number {
    return this.writer.currentRowCount;
  }
}

/**
 * Multi-sheet Excel Stream Writer
 * Allows streaming data to multiple worksheets
 */
export class MultiSheetExcelStreamWriter {
  private readonly worksheets = new Map<
    string,
    { writer: DiskBackedWorksheetWriter; config: ExcelStreamOptions }
  >();
  private readonly target: string | Bun.BunFile | Bun.S3File;
  private readonly options: ExcelWriteOptions;
  private readonly styleRegistry = new StyleRegistry();
  private currentSheet: string;
  private ended = false;

  constructor(target: FileTarget, options?: ExcelWriteOptions) {
    this.target = toWriteTarget(target);
    this.options = options || {};
    this.currentSheet = 'Sheet1';
    this.worksheets.set('Sheet1', {
      writer: new DiskBackedWorksheetWriter({}, this.styleRegistry),
      config: {},
    });
  }

  /**
   * Add a new sheet or switch to existing sheet
   */
  addSheet(name: string, config?: ExcelStreamOptions): this {
    const existing = this.worksheets.get(name);
    if (existing) {
      existing.config = { ...existing.config, ...(config || {}) };
      existing.writer.updateOptions(config);
    } else {
      this.worksheets.set(name, {
        writer: new DiskBackedWorksheetWriter(config || {}, this.styleRegistry),
        config: config || {},
      });
    }
    this.currentSheet = name;
    return this;
  }

  private getCurrentWorksheet() {
    const sheet = this.worksheets.get(this.currentSheet);
    if (!sheet) throw new Error(`Sheet not found: ${this.currentSheet}`);
    return sheet;
  }

  /**
   * Write a row to the current sheet
   */
  writeRow(row: Row | CellValue[]): void {
    this.getCurrentWorksheet().writer.writeRow(row);
  }

  /**
   * Write a styled row to the current sheet
   */
  writeStyledRow(values: CellValue[], styles: (CellStyle | undefined)[]): void {
    this.getCurrentWorksheet().writer.writeStyledRow(values, styles);
  }

  /**
   * Flush buffered temp-file writes for all sheets.
   */
  flush(): Promise<void> {
    return Promise.all(
      [...this.worksheets.values()].map((sheet) => sheet.writer.flush()),
    ).then(() => {});
  }

  /**
   * Finalize and write the Excel file
   */
  async end(): Promise<void> {
    if (this.ended) {
      return;
    }
    this.ended = true;

    const tempOutputPath =
      typeof this.target === 'string'
        ? createOutputTempPath(this.target)
        : undefined;
    const sheets = [...this.worksheets.entries()];
    const sheetNames = sheets.map(([name]) => name);

    try {
      await Promise.all(sheets.map(([, sheet]) => sheet.writer.close()));

      const zipWriter = new StreamingZipWriter(tempOutputPath ?? this.target, {
        compress: this.options.compress,
      });

      await zipWriter.addFile('[Content_Types].xml', [
        buildContentTypes(sheetNames.length),
      ]);
      await zipWriter.addFile('_rels/.rels', [buildRootRels()]);
      await zipWriter.addFile('docProps/app.xml', [
        buildAppPropsXML(sheetNames),
      ]);
      await zipWriter.addFile('docProps/core.xml', [
        buildCorePropsXML({
          creator: this.options.creator,
          created: this.options.created,
          modified: this.options.modified,
        }),
      ]);
      await zipWriter.addFile('xl/_rels/workbook.xml.rels', [
        buildWorkbookRels(sheetNames.length),
      ]);
      await zipWriter.addFile('xl/workbook.xml', [
        buildWorkbookXML(sheetNames),
      ]);
      await zipWriter.addFile('xl/sharedStrings.xml', [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>',
      ]);

      for (let i = 0; i < sheets.length; i++) {
        const [, sheet] = sheets[i];
        await zipWriter.addFile(
          `xl/worksheets/sheet${i + 1}.xml`,
          sheet.writer.buildWorksheetParts(),
        );
      }

      await zipWriter.addFile('xl/styles.xml', [
        this.styleRegistry.buildStylesXML(),
      ]);

      for (let i = 0; i < sheets.length; i++) {
        const [, sheet] = sheets[i];
        const relParts = sheet.writer.buildWorksheetRelParts();
        if (!relParts) continue;
        await zipWriter.addFile(
          `xl/worksheets/_rels/sheet${i + 1}.xml.rels`,
          relParts,
        );
      }

      await zipWriter.close();
      if (typeof this.target === 'string' && tempOutputPath) {
        renameSync(tempOutputPath, this.target);
      }
    } finally {
      await Promise.all([
        ...sheets.map(([, sheet]) => sheet.writer.cleanup()),
        ...(tempOutputPath
          ? [
              (async () => {
                try {
                  await Bun.file(tempOutputPath).delete();
                } catch {
                  // Ignore cleanup errors
                }
              })(),
            ]
          : []),
      ]);
      this.worksheets.clear();
    }
  }
}

/**
 * Create an Excel stream writer (disk-backed Bun-native streaming)
 */
export function createExcelStream(
  target: FileTarget,
  options?: ExcelStreamOptions,
): ExcelStreamWriter {
  return new ExcelStreamWriter(target, options);
}

/**
 * Create a multi-sheet Excel stream writer
 */
export function createMultiSheetExcelStream(
  target: FileTarget,
  options?: ExcelWriteOptions,
): MultiSheetExcelStreamWriter {
  return new MultiSheetExcelStreamWriter(target, options);
}
