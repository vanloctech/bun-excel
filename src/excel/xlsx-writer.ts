// ============================================
// XLSX Writer — Bun-optimized Excel writing
// ============================================

import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { toWriteTarget } from '../runtime-io';
import type {
  Cell,
  DefinedName,
  ExcelWriteOptions,
  FileTarget,
  Row,
  Workbook,
  Worksheet,
} from '../types';
import { buildAutoFilterXML } from './auto-filter';
import { buildConditionalFormattingsXML } from './conditional-formatting';
import { buildDataValidationsXML } from './data-validation';
import { writeEncryptedExcelChunks } from './encrypted-file';
import { encryptExcelPackage, validateExcelPassword } from './encryption';
import { createTempRuntimeId } from './runtime-utils';
import {
  buildSheetRelsXML,
  buildWorksheetFeatureArtifacts,
  type SheetRelationship,
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
  buildSharedStrings,
  buildSheetPropertiesXML,
  buildSheetProtectionXML,
  buildSheetViewsXML,
  buildWorkbookRels,
  buildWorkbookXML,
  escapeXML,
  getFiniteNumber,
  getFiniteNumberOr,
} from './xml-builder';
import { joinZipChunks, zipBuffer, zipChunks } from './zip-buffer';

const encoder = new TextEncoder();
const CELL_REF_PARTS_REGEX = /^([A-Z]+)(\d+)$/;

/**
 * Write a Workbook to an Excel (.xlsx) file
 * Uses Bun.write() for optimized file output
 */
export async function writeExcel(
  target: FileTarget,
  workbook: Workbook,
  options?: ExcelWriteOptions,
): Promise<void> {
  const password = options?.password;
  validateExcelPassword(password);
  const output = toWriteTarget(target);
  if (password !== undefined) {
    await writeEncryptedExcelChunks(
      output,
      zipChunks(
        buildExcelParts(workbook, options),
        options?.compress !== false,
      ),
      password,
    );
    return;
  }
  const pending: Uint8Array[] = [];
  let size = 0;
  let temporary: Bun.BunFile | undefined;
  let sink: Bun.FileSink | undefined;
  let closed = false;
  try {
    for (const chunk of zipChunks(
      buildExcelParts(workbook, options),
      options?.compress !== false,
    )) {
      if (!sink) {
        pending.push(chunk);
        size += chunk.byteLength;
        // Small exports retain the original in-memory write path.
        if (size <= 1024 * 1024) continue;
        temporary = Bun.file(
          join(tmpdir(), `bun-excel-write-${createTempRuntimeId()}.zip`),
        );
        sink = temporary.writer({ highWaterMark: 256 * 1024 });
        for (const part of pending) sink.write(part);
        pending.length = 0;
      } else {
        sink.write(chunk);
      }
      await sink.flush();
    }
    if (sink && temporary) {
      await sink.end();
      closed = true;
      // Publish only after serialization succeeds, preserving an existing
      // target when workbook validation or compression fails.
      await Bun.write(output, temporary);
    } else {
      await Bun.write(output, joinZipChunks(pending));
    }
  } finally {
    if (sink && !closed) {
      try {
        await sink.end();
      } catch {
        /* Preserve the original write error. */
      }
    }
    if (temporary) await temporary.delete().catch(() => {});
  }
}

/**
 * Build Excel buffer in memory (returns Uint8Array)
 * Useful for sending as HTTP response or further processing
 */
export function buildExcelBuffer(
  workbook: Workbook,
  options?: ExcelWriteOptions,
): Uint8Array {
  const password = options?.password;
  validateExcelPassword(password);
  const buffer = zipBuffer(
    buildExcelParts(workbook, options),
    options?.compress !== false,
  );
  return password === undefined
    ? buffer
    : encryptExcelPackage(buffer, password);
}

function* buildExcelParts(
  workbook: Workbook,
  options?: ExcelWriteOptions,
): Generator<readonly [string, Uint8Array]> {
  const styleRegistry = new StyleRegistry();
  const sharedStrings: string[] = [];
  const sharedStringMap = new Map<string, number>();

  /**
   * Get or create shared string index
   */
  function getSharedStringIndex(str: string): number {
    const existing = sharedStringMap.get(str);
    if (existing !== undefined) return existing;
    const index = sharedStrings.length;
    sharedStrings.push(str);
    sharedStringMap.set(str, index);
    return index;
  }

  interface HyperlinkEntry {
    ref: string;
    rId?: string;
    location?: string;
    tooltip?: string;
  }

  function collectHyperlink(
    ref: string,
    cell: Cell,
    relationships: SheetRelationship[],
    hyperlinkEntries: HyperlinkEntry[],
    nextRelId: () => string,
  ): void {
    if (!cell.hyperlink) return;

    const hyperlink = cell.hyperlink;
    if (isExternalHyperlink(hyperlink.target)) {
      const relationshipId = nextRelId();
      relationships.push({
        id: relationshipId,
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
        target: hyperlink.target,
        targetMode: 'External',
      });
      hyperlinkEntries.push({
        ref,
        rId: relationshipId,
        tooltip: hyperlink.tooltip,
      });
      return;
    }

    hyperlinkEntries.push({
      ref,
      location: hyperlink.target,
      tooltip: hyperlink.tooltip,
    });
  }

  function buildWorksheetColumnsXML(worksheet: Worksheet): string {
    if (!worksheet.columns || worksheet.columns.length === 0) {
      return '';
    }

    let xml = '<cols>';
    for (let c = 0; c < worksheet.columns.length; c++) {
      const col = worksheet.columns[c];
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
    return xml;
  }

  function buildWorksheetCellXML(
    cell: Cell,
    rowStyle: Row['style'],
    ref: string,
  ): string {
    const cellStyle = cell.style || rowStyle;
    const styleIdx = styleRegistry.registerStyle(cellStyle);
    const { value } = cell;

    if (value === null || value === undefined) {
      if (cell.formula) {
        let xml = `<c r="${ref}"${styleIdx > 0 ? ` s="${styleIdx}"` : ''}>`;
        xml += `<f>${escapeXML(cell.formula)}</f>`;
        if (cell.formulaResult !== undefined) {
          xml += `<v>${escapeXML(String(cell.formulaResult))}</v>`;
        }
        xml += '</c>';
        return xml;
      }
      return styleIdx > 0 ? `<c r="${ref}" s="${styleIdx}"/>` : '';
    }

    if (cell.formula) {
      let xml = `<c r="${ref}"${styleIdx > 0 ? ` s="${styleIdx}"` : ''}>`;
      xml += `<f>${escapeXML(cell.formula)}</f>`;
      if (cell.formulaResult !== undefined) {
        xml += `<v>${escapeXML(String(cell.formulaResult))}</v>`;
      } else if (typeof value === 'string') {
        xml += `<v>${getSharedStringIndex(value)}</v>`;
      } else if (typeof value === 'number' || typeof value === 'boolean') {
        xml += `<v>${value}</v>`;
      }
      xml += '</c>';
      return xml;
    }

    if (cell.richText && cell.richText.length > 0) {
      return `<c r="${ref}" t="inlineStr"${
        styleIdx > 0 ? ` s="${styleIdx}"` : ''
      }>${buildRichTextXML(cell.richText)}</c>`;
    }

    if (typeof value === 'string') {
      return `<c r="${ref}" t="s"${
        styleIdx > 0 ? ` s="${styleIdx}"` : ''
      }><v>${getSharedStringIndex(value)}</v></c>`;
    }

    if (typeof value === 'number') {
      return `<c r="${ref}"${styleIdx > 0 ? ` s="${styleIdx}"` : ''}><v>${value}</v></c>`;
    }

    if (typeof value === 'boolean') {
      return `<c r="${ref}" t="b"${
        styleIdx > 0 ? ` s="${styleIdx}"` : ''
      }><v>${value ? 1 : 0}</v></c>`;
    }

    if (value instanceof Date) {
      return `<c r="${ref}"${
        styleIdx > 0 ? ` s="${styleIdx}"` : ''
      }><v>${dateToExcelSerial(value)}</v></c>`;
    }

    return '';
  }

  function buildWorksheetRowsXML(
    worksheet: Worksheet,
    relationships: SheetRelationship[],
    hyperlinkEntries: HyperlinkEntry[],
    nextRelId: () => string,
  ): string {
    const batches: string[] = [];
    const parts: string[] = ['<sheetData>'];
    let size = 0;
    function append(value: string): void {
      parts.push(value);
      size += value.length;
      if (size >= 128 * 1024) {
        batches.push(parts.join(''));
        parts.length = 0;
        size = 0;
      }
    }

    for (let r = 0; r < worksheet.rows.length; r++) {
      const row = worksheet.rows[r];
      if (!row) continue;

      let rowAttrs = ` r="${r + 1}"`;
      const rowHeight = getFiniteNumber(row.height);
      if (rowHeight !== undefined) {
        rowAttrs += ` ht="${rowHeight}" customHeight="1"`;
      }
      if (row.hidden) rowAttrs += ' hidden="1"';
      if (row.collapsed) rowAttrs += ' collapsed="1"';
      if (row.outlineLevel !== undefined) {
        rowAttrs += ` outlineLevel="${Math.max(0, Math.trunc(row.outlineLevel))}"`;
      }

      const rowStyleIdx = row.style
        ? styleRegistry.registerStyle(row.style)
        : 0;
      if (rowStyleIdx > 0) {
        rowAttrs += ` s="${rowStyleIdx}" customFormat="1"`;
      }

      append(`<row${rowAttrs}>`);
      for (let c = 0; c < row.cells.length; c++) {
        const cell = row.cells[c];
        if (!cell) continue;

        const ref = buildCellRef(r, c);
        append(buildWorksheetCellXML(cell, row.style, ref));
        collectHyperlink(ref, cell, relationships, hyperlinkEntries, nextRelId);
      }
      append('</row>');
    }

    append('</sheetData>');
    batches.push(parts.join(''));
    return batches.join('');
  }

  function buildWorksheetHyperlinksXML(
    hyperlinkEntries: HyperlinkEntry[],
  ): string {
    if (hyperlinkEntries.length === 0) {
      return '';
    }

    let xml = '<hyperlinks>';
    for (const hyperlink of hyperlinkEntries) {
      xml += `<hyperlink ref="${hyperlink.ref}"`;
      if (hyperlink.rId) xml += ` r:id="${hyperlink.rId}"`;
      if (hyperlink.location) {
        xml += ` location="${escapeXML(hyperlink.location)}"`;
      }
      if (hyperlink.tooltip) {
        xml += ` tooltip="${escapeXML(hyperlink.tooltip)}"`;
      }
      xml += '/>';
    }
    xml += '</hyperlinks>';
    return xml;
  }

  /**
   * Build worksheet XML + collect hyperlink relationships
   */
  function buildWorksheetXML(worksheet: Worksheet): {
    xml: string;
    relationships: SheetRelationship[];
    extraFiles: { path: string; content: Uint8Array }[];
    mediaExtensions: Set<string>;
    commentCount: number;
    drawingCount: number;
    tableCount: number;
  } {
    const relationships: SheetRelationship[] = [];
    const hyperlinkEntries: HyperlinkEntry[] = [];
    let hyperlinkRelCounter = 1;
    const nextHyperlinkRelId = () => `rId${hyperlinkRelCounter++}`;

    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml +=
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';

    const hasOutlines =
      worksheet.rows.some((row) => !!row.outlineLevel) ||
      !!worksheet.columns?.some((column) => !!column.outlineLevel);
    xml += buildSheetPropertiesXML(hasOutlines);

    xml += buildSheetViewsXML({
      freezePane: worksheet.freezePane,
      splitPane: worksheet.splitPane,
    });

    // Sheet format properties
    xml += `<sheetFormatPr defaultRowHeight="${getFiniteNumberOr(worksheet.defaultRowHeight, 15)}"`;
    const defaultColWidth = getFiniteNumber(worksheet.defaultColWidth);
    if (defaultColWidth !== undefined) {
      xml += ` defaultColWidth="${defaultColWidth}"`;
    }
    xml += '/>';
    xml += buildWorksheetColumnsXML(worksheet);
    xml += buildWorksheetRowsXML(
      worksheet,
      relationships,
      hyperlinkEntries,
      nextHyperlinkRelId,
    );

    const featureArtifacts = buildWorksheetFeatureArtifacts(
      worksheet,
      sheetCounters,
      {
        startingRelIndex: hyperlinkRelCounter,
      },
    );
    relationships.push(...featureArtifacts.relationships);

    if (worksheet.mergeCells && worksheet.mergeCells.length > 0) {
      xml += `<mergeCells count="${worksheet.mergeCells.length}">`;
      for (const mc of worksheet.mergeCells) {
        const startRef = buildCellRef(mc.startRow, mc.startCol);
        const endRef = buildCellRef(mc.endRow, mc.endCol);
        xml += `<mergeCell ref="${startRef}:${endRef}"/>`;
      }
      xml += '</mergeCells>';
    }

    const autoFilterXml = buildAutoFilterXML(worksheet.autoFilter);
    if (autoFilterXml) {
      xml += autoFilterXml;
    }

    const conditionalFormattingXml = buildConditionalFormattingsXML(
      worksheet.conditionalFormattings,
      styleRegistry,
    );
    if (conditionalFormattingXml) {
      xml += conditionalFormattingXml;
    }

    // Hyperlinks
    const dataValidationsXml = buildDataValidationsXML(
      worksheet.dataValidations,
    );
    if (dataValidationsXml) {
      xml += dataValidationsXml;
    }

    xml += buildSheetProtectionXML(worksheet.protection);
    xml += buildWorksheetHyperlinksXML(hyperlinkEntries);
    xml += buildPageMarginsXML(worksheet.pageMargins);
    xml += buildPageSetupXML(worksheet.pageSetup);
    xml += buildHeaderFooterXML(worksheet.headerFooter);
    xml += featureArtifacts.xmlPartsBeforeClose.join('');

    xml += '</worksheet>';
    return {
      xml,
      relationships,
      extraFiles: featureArtifacts.extraFiles,
      mediaExtensions: featureArtifacts.mediaExtensions,
      commentCount: featureArtifacts.commentCount,
      drawingCount: featureArtifacts.drawingCount,
      tableCount: featureArtifacts.tableCount,
    };
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

  function buildWorkbookDefinedNames(workbookInput: Workbook): DefinedName[] {
    const definedNames = [...(workbookInput.definedNames ?? [])];
    for (let i = 0; i < workbookInput.worksheets.length; i++) {
      const worksheet = workbookInput.worksheets[i];
      if (!worksheet.printArea) continue;
      definedNames.push({
        name: '_xlnm.Print_Area',
        localSheetId: i,
        refersTo: `${quoteSheetName(worksheet.name)}!${absoluteCellRef(
          worksheet.printArea.startRow,
          worksheet.printArea.startCol,
        )}:${absoluteCellRef(
          worksheet.printArea.endRow,
          worksheet.printArea.endCol,
        )}`,
      });
    }
    return definedNames;
  }

  /**
   * Check if a hyperlink target is external (URL/email) vs internal (sheet ref)
   */
  function isExternalHyperlink(target: string): boolean {
    return (
      target.startsWith('http://') ||
      target.startsWith('https://') ||
      target.startsWith('mailto:') ||
      target.startsWith('ftp://')
    );
  }

  // Emit worksheets sequentially so their XML trees of strings do not accumulate.
  const sheetNames = workbook.worksheets.map((ws) => ws.name);
  const workbookSheets = workbook.worksheets.map((worksheet) => ({
    name: worksheet.name,
    state: worksheet.state,
  }));
  const definedNames = buildWorkbookDefinedNames(workbook);
  // Keep workbook metadata first for selective readers to resolve sheets early.
  yield [
    'xl/_rels/workbook.xml.rels',
    encoder.encode(buildWorkbookRels(sheetNames.length)),
  ];
  yield [
    'xl/workbook.xml',
    encoder.encode(
      buildWorkbookXML(workbookSheets, { definedNames, view: workbook.views }),
    ),
  ];
  const sheetCounters = {
    nextCommentsIndex: 1,
    nextDrawingIndex: 1,
    nextTableIndex: 1,
  };
  let commentsCount = 0;
  let drawingsCount = 0;
  let tablesCount = 0;
  const mediaExtensions = new Set<string>();
  for (let i = 0; i < workbook.worksheets.length; i++) {
    const result = buildWorksheetXML(workbook.worksheets[i]);
    commentsCount += result.commentCount;
    drawingsCount += result.drawingCount;
    tablesCount += result.tableCount;
    for (const extension of result.mediaExtensions)
      mediaExtensions.add(extension);
    const bytes = encoder.encode(result.xml);
    result.xml = '';
    yield [`xl/worksheets/sheet${i + 1}.xml`, bytes];
    for (const file of result.extraFiles) yield [file.path, file.content];
    if (result.relationships.length) {
      yield [
        `xl/worksheets/_rels/sheet${i + 1}.xml.rels`,
        encoder.encode(buildSheetRelsXML(result.relationships)),
      ];
    }
  }
  sharedStringMap.clear();

  const workbookCreator = options?.creator ?? workbook.creator;
  const workbookCreated = options?.created ?? workbook.created;
  const workbookModified = options?.modified ?? workbook.modified;

  // Build ZIP structure
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': encoder.encode(
      buildContentTypes(sheetNames.length, {
        commentsCount,
        drawingsCount,
        tablesCount,
        includeVml: commentsCount > 0,
        mediaExtensions: [...mediaExtensions],
      }),
    ),
    '_rels/.rels': encoder.encode(buildRootRels()),
    'docProps/app.xml': encoder.encode(buildAppPropsXML(sheetNames)),
    'docProps/core.xml': encoder.encode(
      buildCorePropsXML({
        creator: workbookCreator,
        created: workbookCreated,
        modified: workbookModified,
      }),
    ),
    'xl/styles.xml': encoder.encode(styleRegistry.buildStylesXML()),
    'xl/sharedStrings.xml': encoder.encode(buildSharedStrings(sharedStrings)),
  };

  yield* Object.entries(files);
}

/**
 * Convert Date to Excel serial number
 */
function dateToExcelSerial(date: Date): number {
  const epoch = new Date(Date.UTC(1899, 11, 30));
  const diff = date.getTime() - epoch.getTime();
  return diff / (24 * 60 * 60 * 1000);
}

/**
 * Convert Excel serial number to Date
 */
export function excelSerialToDate(serial: number): Date {
  const epoch = new Date(Date.UTC(1899, 11, 30));
  return new Date(epoch.getTime() + serial * 24 * 60 * 60 * 1000);
}
