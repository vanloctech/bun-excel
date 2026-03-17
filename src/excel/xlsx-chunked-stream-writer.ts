// ============================================
// XLSX Chunked Stream Writer — Disk-backed
// low-memory streaming
// ============================================
//
// Flow:
//   writeRow() → serialize XML → temp files on disk
//   end()      → stream worksheet XML into ZIP entry → rename temp output
//
// Uses inline strings (<is><t>...</t></is>) instead of shared string table
// to avoid tracking all string values in memory.

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
  DefinedName,
  ExcelWriteOptions,
  FileTarget,
  MergeCell,
  Row,
  StreamWriter,
  WorkbookView,
  Worksheet,
} from '../types';
import { buildAutoFilterXML } from './auto-filter';
import { type CommentEntry, commentRefFromCoords } from './comments';
import { buildConditionalFormattingsXML } from './conditional-formatting';
import { buildDataValidationsXML } from './data-validation';
import { ManagedFileSink } from './file-sink';
import { createTempRuntimeId } from './runtime-utils';
import {
  buildSheetRelsXML,
  buildWorksheetFeatureArtifacts,
} from './sheet-parts';
import { StyleRegistry } from './style-builder';
import {
  buildAppPropsXML,
  buildCellRef,
  buildContentTypes,
  buildCorePropsXML,
  buildHeaderFooterXML,
  buildPageMarginsXML,
  buildPageSetupXML,
  buildRichTextXML,
  buildRootRels,
  buildSheetPropertiesXML,
  buildSheetProtectionXML,
  buildSheetViewsXML,
  buildWorkbookRels,
  buildWorkbookXML,
  escapeXML,
  getFiniteNumber,
  getFiniteNumberOr,
} from './xml-builder';
import { StreamingZipWriter } from './zip-stream';

const CELL_REF_PARTS_REGEX = /^([A-Z]+)(\d+)$/;

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
  /** Worksheet state */
  state?: Worksheet['state'];
  /** Sheet protection */
  protection?: Worksheet['protection'];
  /** Page margins */
  pageMargins?: Worksheet['pageMargins'];
  /** Page setup */
  pageSetup?: Worksheet['pageSetup'];
  /** Header/footer */
  headerFooter?: Worksheet['headerFooter'];
  /** Print area */
  printArea?: Worksheet['printArea'];
  /** Worksheet images */
  images?: Worksheet['images'];
  /** Worksheet tables */
  tables?: Worksheet['tables'];
  /** Workbook defined names */
  definedNames?: DefinedName[];
  /** Workbook views */
  views?: WorkbookView;
}

function createTempFilePath(prefix: string): string {
  return join(tmpdir(), `${prefix}-${createTempRuntimeId()}.tmp`);
}

function buildWorksheetOpenTag(): string {
  return [
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
  ].join(' ');
}

function createOutputTempPath(outputPath: string): string {
  return join(dirname(outputPath), `.bun-excel-${createTempRuntimeId()}.tmp`);
}

function quoteSheetName(name: string): string {
  return `'${name.replace(/'/g, "''")}'`;
}

function absoluteCellRef(row: number, col: number): string {
  const ref = buildCellRef(row, col);
  const match = ref.match(CELL_REF_PARTS_REGEX);
  if (!match) return ref;
  return `$${match[1]}$${match[2]}`;
}

/**
 * Excel Chunked Stream Writer — Disk-backed low memory
 *
 * Writes row XML and hyperlink metadata directly to temporary files on disk.
 * At end(), streams worksheet XML directly into a ZIP entry without
 * materializing the full worksheet or archive in memory.
 *
 * Uses inline strings instead of shared string table to avoid
 * tracking all string values in memory.
 */
export class ExcelChunkedStreamWriter implements StreamWriter {
  private readonly target: string | Bun.BunFile | Bun.S3File;
  private readonly options: ChunkedExcelStreamOptions;
  private readonly styleRegistry = new StyleRegistry();
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
  private hasOutlineLevels = false;
  private readonly commentEntries: CommentEntry[] = [];
  private ended = false;

  constructor(target: FileTarget, options?: ChunkedExcelStreamOptions) {
    this.target = toWriteTarget(target);
    this.options = options || {};
    this.rowTempFilePath = createTempFilePath('bun-xlsx-rows');
    this.rowTempWriter = new ManagedFileSink(this.rowTempFilePath);
    this.hyperlinkTempFilePath = createTempFilePath('bun-xlsx-links');
    this.hyperlinkTempWriter = new ManagedFileSink(this.hyperlinkTempFilePath);
    this.hyperlinkRelTempFilePath = createTempFilePath('bun-xlsx-link-rels');
    this.hyperlinkRelTempWriter = new ManagedFileSink(
      this.hyperlinkRelTempFilePath,
    );
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

    if (cell.richText && cell.richText.length > 0) {
      return `<c r="${ref}" t="inlineStr"${
        styleIdx > 0 ? ` s="${styleIdx}"` : ''
      }>${buildRichTextXML(cell.richText)}</c>`;
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

    const epoch = new Date(Date.UTC(1899, 11, 30));
    const serial = (value.getTime() - epoch.getTime()) / (24 * 60 * 60 * 1000);
    return `<c r="${ref}"${styleIdx > 0 ? ` s="${styleIdx}"` : ''}><v>${serial}</v></c>`;
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

  private ensureWritable(): void {
    if (this.ended) {
      throw new Error('Cannot write rows after stream.end() has been called');
    }
  }

  private buildWorksheetPrefix(): string {
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml += buildWorksheetOpenTag();

    xml += buildSheetPropertiesXML(
      this.hasOutlineLevels ||
        !!this.options.columns?.some((column) => !!column.outlineLevel),
    );

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
        let colAttrs = ` min="${c + 1}" max="${c + 1}"`;
        if (colWidth !== undefined) {
          colAttrs += ` width="${colWidth}" customWidth="1"`;
        }
        if (col.hidden) colAttrs += ' hidden="1"';
        if (col.collapsed) colAttrs += ' collapsed="1"';
        if (col.outlineLevel !== undefined) {
          colAttrs += ` outlineLevel="${Math.max(0, Math.trunc(col.outlineLevel))}"`;
        }
        xml += `<col${colAttrs}/>`;
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

    const protectionXml = buildSheetProtectionXML(this.options.protection);
    if (protectionXml) {
      parts.push(protectionXml);
    }

    const pageMarginsXml = buildPageMarginsXML(this.options.pageMargins);
    if (pageMarginsXml) {
      parts.push(pageMarginsXml);
    }

    const pageSetupXml = buildPageSetupXML(this.options.pageSetup);
    if (pageSetupXml) {
      parts.push(pageSetupXml);
    }

    const headerFooterXml = buildHeaderFooterXML(this.options.headerFooter);
    if (headerFooterXml) {
      parts.push(headerFooterXml);
    }

    parts.push('</worksheet>');
    return parts;
  }

  /**
   * Write a single row — serializes XML and appends it to a temp file.
   */
  writeRow(row: Row | CellValue[]): void {
    this.ensureWritable();

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
    if (rowObj.hidden) rowAttrs += ' hidden="1"';
    if (rowObj.collapsed) rowAttrs += ' collapsed="1"';
    if (rowObj.outlineLevel !== undefined) {
      rowAttrs += ` outlineLevel="${Math.max(0, Math.trunc(rowObj.outlineLevel))}"`;
      this.hasOutlineLevels = true;
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
      if (cell.comment) {
        this.commentEntries.push({
          ref: commentRefFromCoords(r, c),
          comment: cell.comment,
        });
      }
      xml += this.serializeCell(cell, ref, rowObj.style);
    }

    xml += '</row>';

    this.rowTempWriter.write(xml);
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
   * Flush temp file buffers to disk.
   */
  flush(): Promise<void> {
    return Promise.all([
      this.rowTempWriter.flush(),
      this.hyperlinkTempWriter.flush(),
      this.hyperlinkRelTempWriter.flush(),
    ]).then(() => {});
  }

  /**
   * Finalize and write the XLSX file using incremental ZIP output.
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
    const tempPaths = [
      this.rowTempFilePath,
      this.hyperlinkTempFilePath,
      this.hyperlinkRelTempFilePath,
      ...(tempOutputPath ? [tempOutputPath] : []),
    ];

    try {
      await Promise.all([
        this.rowTempWriter.end(),
        this.hyperlinkTempWriter.end(),
        this.hyperlinkRelTempWriter.end(),
      ]);

      const sheetName = this.options.sheetName || 'Sheet1';
      const definedNames = [...(this.options.definedNames ?? [])];
      if (this.options.printArea) {
        definedNames.push({
          name: '_xlnm.Print_Area',
          localSheetId: 0,
          refersTo: `${quoteSheetName(sheetName)}!${absoluteCellRef(
            this.options.printArea.startRow,
            this.options.printArea.startCol,
          )}:${absoluteCellRef(
            this.options.printArea.endRow,
            this.options.printArea.endCol,
          )}`,
        });
      }
      const zipWriter = new StreamingZipWriter(tempOutputPath ?? this.target, {
        compress: this.options.compress,
        s3WriterOptions: this.options.s3WriterOptions,
      });

      const featureArtifacts = buildWorksheetFeatureArtifacts(
        {
          name: sheetName,
          rows: [],
          images: this.options.images,
          tables: this.options.tables,
        },
        {
          nextCommentsIndex: 1,
          nextDrawingIndex: 1,
          nextTableIndex: 1,
        },
        {
          commentEntries: this.commentEntries,
          startingRelIndex: this.hyperlinkRelCounter,
        },
      );

      await zipWriter.addFile('[Content_Types].xml', [
        buildContentTypes(1, {
          commentsCount: featureArtifacts.commentCount,
          drawingsCount: featureArtifacts.drawingCount,
          tablesCount: featureArtifacts.tableCount,
          includeVml: featureArtifacts.commentCount > 0,
          mediaExtensions: [...featureArtifacts.mediaExtensions],
        }),
      ]);
      await zipWriter.addFile('_rels/.rels', [buildRootRels()]);
      await zipWriter.addFile('docProps/app.xml', [
        buildAppPropsXML([sheetName]),
      ]);
      await zipWriter.addFile('docProps/core.xml', [
        buildCorePropsXML({
          creator: this.options.creator,
          created: this.options.created,
          modified: this.options.modified,
        }),
      ]);
      await zipWriter.addFile('xl/_rels/workbook.xml.rels', [
        buildWorkbookRels(1),
      ]);
      await zipWriter.addFile('xl/workbook.xml', [
        buildWorkbookXML([{ name: sheetName, state: this.options.state }], {
          definedNames,
          view: this.options.views,
        }),
      ]);
      await zipWriter.addFile('xl/sharedStrings.xml', [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>',
      ]);

      const worksheetParts: (string | Blob)[] = [
        this.buildWorksheetPrefix(),
        Bun.file(this.rowTempFilePath),
      ];

      const worksheetSuffixParts = this.buildWorksheetSuffix();
      worksheetSuffixParts.splice(
        worksheetSuffixParts.length - 1,
        0,
        ...featureArtifacts.xmlPartsBeforeClose,
      );
      if (this.hyperlinkCount > 0) {
        const worksheetClosingTag = worksheetSuffixParts.pop();
        if (worksheetClosingTag) {
          worksheetParts.push(...worksheetSuffixParts, '<hyperlinks>');
          worksheetParts.push(Bun.file(this.hyperlinkTempFilePath));
          worksheetParts.push('</hyperlinks>', worksheetClosingTag);
        }
      } else {
        worksheetParts.push(...worksheetSuffixParts);
      }

      await zipWriter.addFile('xl/worksheets/sheet1.xml', worksheetParts);
      await zipWriter.addFile('xl/styles.xml', [
        this.styleRegistry.buildStylesXML(),
      ]);

      for (const extraFile of featureArtifacts.extraFiles) {
        await zipWriter.addFile(extraFile.path, [extraFile.content]);
      }

      if (
        this.externalHyperlinkCount > 0 ||
        featureArtifacts.relationships.length > 0
      ) {
        const relXml =
          this.externalHyperlinkCount > 0
            ? (() => {
                const rels: string[] = [];
                for (const relationship of featureArtifacts.relationships) {
                  rels.push(
                    `<Relationship Id="${relationship.id}" Type="${relationship.type}" Target="${relationship.target}"${
                      relationship.targetMode
                        ? ` TargetMode="${relationship.targetMode}"`
                        : ''
                    }/>`,
                  );
                }
                return [
                  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
                  ...rels,
                ];
              })()
            : [buildSheetRelsXML(featureArtifacts.relationships)];

        if (this.externalHyperlinkCount > 0) {
          await zipWriter.addFile('xl/worksheets/_rels/sheet1.xml.rels', [
            ...relXml,
            Bun.file(this.hyperlinkRelTempFilePath),
            '</Relationships>',
          ]);
        } else {
          await zipWriter.addFile(
            'xl/worksheets/_rels/sheet1.xml.rels',
            relXml,
          );
        }
      }

      await zipWriter.close();
      if (typeof this.target === 'string' && tempOutputPath) {
        renameSync(tempOutputPath, this.target);
      }
    } finally {
      await Promise.all(
        tempPaths.map(async (filePath) => {
          try {
            await Bun.file(filePath).delete();
          } catch {
            // Ignore cleanup errors
          }
        }),
      );
    }
  }

  async cancel(): Promise<void> {
    if (this.ended) {
      return;
    }
    this.ended = true;

    await Promise.all([
      this.rowTempWriter.end().catch(() => {}),
      this.hyperlinkTempWriter.end().catch(() => {}),
      this.hyperlinkRelTempWriter.end().catch(() => {}),
    ]);

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

  /**
   * Get current row count
   */
  get currentRowCount(): number {
    return this.rowCount;
  }
}

/**
 * Create a chunked Excel stream writer (disk-backed low-memory streaming)
 */
export function createChunkedExcelStream(
  target: FileTarget,
  options?: ChunkedExcelStreamOptions,
): ExcelChunkedStreamWriter {
  return new ExcelChunkedStreamWriter(target, options);
}
