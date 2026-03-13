// ============================================
// XLSX Writer — Bun-optimized Excel writing
// ============================================

import { resolve } from 'node:path';
import { type Zippable, zipSync } from 'fflate';
import type { ExcelWriteOptions, Workbook, Worksheet } from '../types';
import { StyleRegistry } from './style-builder';
import {
  buildCellRef,
  buildContentTypes,
  buildRootRels,
  buildSharedStrings,
  buildWorkbookRels,
  buildWorkbookXML,
  escapeXML,
} from './xml-builder';

const encoder = new TextEncoder();

/**
 * Write a Workbook to an Excel (.xlsx) file
 * Uses Bun.write() for optimized file output
 */
export async function writeExcel(
  path: string,
  workbook: Workbook,
  options?: ExcelWriteOptions,
): Promise<void> {
  const buffer = buildExcelBuffer(workbook, options);

  // Validate and resolve path
  const resolvedPath = resolve(path);
  if (resolvedPath.includes('\0')) {
    throw new Error('Invalid file path: contains null bytes');
  }

  // Use Bun.write() for optimized writing
  await Bun.write(resolvedPath, buffer);
}

/**
 * Build Excel buffer in memory (returns Uint8Array)
 * Useful for sending as HTTP response or further processing
 */
export function buildExcelBuffer(
  workbook: Workbook,
  options?: ExcelWriteOptions,
): Uint8Array {
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

  /**
   * Build worksheet XML + collect hyperlink relationships
   */
  function buildWorksheetXML(
    worksheet: Worksheet,
    _sheetIndex: number,
  ): { xml: string; hyperlinkRels: { rId: string; target: string }[] } {
    const hyperlinkRels: { rId: string; target: string }[] = [];
    const hyperlinkEntries: {
      ref: string;
      rId?: string;
      location?: string;
      tooltip?: string;
    }[] = [];
    let hyperlinkRelCounter = 1;

    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml +=
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';

    // Sheet views (freeze pane)
    if (worksheet.freezePane) {
      xml += '<sheetViews>';
      xml += '<sheetView tabSelected="1" workbookViewId="0">';
      const { row, col } = worksheet.freezePane;
      const topLeftCell = buildCellRef(row, col);
      xml += `<pane xSplit="${col}" ySplit="${row}" topLeftCell="${topLeftCell}" activePane="bottomRight" state="frozen"/>`;
      xml += '</sheetView>';
      xml += '</sheetViews>';
    }

    // Sheet format properties
    xml += `<sheetFormatPr defaultRowHeight="${worksheet.defaultRowHeight || 15}"`;
    if (worksheet.defaultColWidth) {
      xml += ` defaultColWidth="${worksheet.defaultColWidth}"`;
    }
    xml += '/>';

    // Columns
    if (worksheet.columns && worksheet.columns.length > 0) {
      xml += '<cols>';
      for (let c = 0; c < worksheet.columns.length; c++) {
        const col = worksheet.columns[c];
        if (col.width) {
          xml += `<col min="${c + 1}" max="${c + 1}" width="${col.width}" customWidth="1"/>`;
        }
      }
      xml += '</cols>';
    }

    // Sheet data
    xml += '<sheetData>';

    for (let r = 0; r < worksheet.rows.length; r++) {
      const row = worksheet.rows[r];
      if (!row) continue;

      let rowAttrs = ` r="${r + 1}"`;
      if (row.height) {
        rowAttrs += ` ht="${row.height}" customHeight="1"`;
      }

      // Register row-level style
      const rowStyleIdx = row.style
        ? styleRegistry.registerStyle(row.style)
        : 0;
      if (rowStyleIdx > 0) {
        rowAttrs += ` s="${rowStyleIdx}" customFormat="1"`;
      }

      xml += `<row${rowAttrs}>`;

      for (let c = 0; c < row.cells.length; c++) {
        const cell = row.cells[c];
        if (!cell) continue;

        const ref = buildCellRef(r, c);

        // Determine style — cell style takes priority, then row style
        const cellStyle = cell.style || row.style;
        const styleIdx = styleRegistry.registerStyle(cellStyle);

        const { value } = cell;

        if (value === null || value === undefined) {
          // Cell might still have a formula or hyperlink even with null display value
          if (cell.formula) {
            xml += `<c r="${ref}"${styleIdx > 0 ? ` s="${styleIdx}"` : ''}>`;
            xml += `<f>${escapeXML(cell.formula)}</f>`;
            if (cell.formulaResult !== undefined) {
              xml += `<v>${escapeXML(String(cell.formulaResult))}</v>`;
            }
            xml += '</c>';
          } else if (styleIdx > 0) {
            xml += `<c r="${ref}" s="${styleIdx}"/>`;
          }
          // Collect hyperlink even for null-value cells
          if (cell.hyperlink) {
            const hl = cell.hyperlink;
            if (isExternalHyperlink(hl.target)) {
              const rId = `rId${hyperlinkRelCounter++}`;
              hyperlinkRels.push({ rId, target: hl.target });
              hyperlinkEntries.push({ ref, rId, tooltip: hl.tooltip });
            } else {
              hyperlinkEntries.push({
                ref,
                location: hl.target,
                tooltip: hl.tooltip,
              });
            }
          }
          continue;
        }

        // Formula cells
        if (cell.formula) {
          xml += `<c r="${ref}"${styleIdx > 0 ? ` s="${styleIdx}"` : ''}>`;
          xml += `<f>${escapeXML(cell.formula)}</f>`;
          if (cell.formulaResult !== undefined) {
            xml += `<v>${escapeXML(String(cell.formulaResult))}</v>`;
          } else if (value !== null && value !== undefined) {
            // Use the value as cached result
            if (typeof value === 'string') {
              const ssIdx = getSharedStringIndex(value);
              xml += `<v>${ssIdx}</v>`;
            } else if (
              typeof value === 'number' ||
              typeof value === 'boolean'
            ) {
              xml += `<v>${value}</v>`;
            }
          }
          xml += '</c>';
        } else if (typeof value === 'string') {
          const ssIdx = getSharedStringIndex(value);
          xml += `<c r="${ref}" t="s"${styleIdx > 0 ? ` s="${styleIdx}"` : ''}>`;
          xml += `<v>${ssIdx}</v>`;
          xml += '</c>';
        } else if (typeof value === 'number') {
          xml += `<c r="${ref}"${styleIdx > 0 ? ` s="${styleIdx}"` : ''}>`;
          xml += `<v>${value}</v>`;
          xml += '</c>';
        } else if (typeof value === 'boolean') {
          xml += `<c r="${ref}" t="b"${styleIdx > 0 ? ` s="${styleIdx}"` : ''}>`;
          xml += `<v>${value ? 1 : 0}</v>`;
          xml += '</c>';
        } else if (value instanceof Date) {
          const excelDate = dateToExcelSerial(value);
          xml += `<c r="${ref}"${styleIdx > 0 ? ` s="${styleIdx}"` : ''}>`;
          xml += `<v>${excelDate}</v>`;
          xml += '</c>';
        }

        // Collect hyperlinks
        if (cell.hyperlink) {
          const hl = cell.hyperlink;
          if (isExternalHyperlink(hl.target)) {
            const rId = `rId${hyperlinkRelCounter++}`;
            hyperlinkRels.push({ rId, target: hl.target });
            hyperlinkEntries.push({ ref, rId, tooltip: hl.tooltip });
          } else {
            hyperlinkEntries.push({
              ref,
              location: hl.target,
              tooltip: hl.tooltip,
            });
          }
        }
      }

      xml += '</row>';
    }

    xml += '</sheetData>';

    // Merge cells
    if (worksheet.mergeCells && worksheet.mergeCells.length > 0) {
      xml += `<mergeCells count="${worksheet.mergeCells.length}">`;
      for (const mc of worksheet.mergeCells) {
        const startRef = buildCellRef(mc.startRow, mc.startCol);
        const endRef = buildCellRef(mc.endRow, mc.endCol);
        xml += `<mergeCell ref="${startRef}:${endRef}"/>`;
      }
      xml += '</mergeCells>';
    }

    // Hyperlinks
    if (hyperlinkEntries.length > 0) {
      xml += '<hyperlinks>';
      for (const hl of hyperlinkEntries) {
        xml += `<hyperlink ref="${hl.ref}"`;
        if (hl.rId) xml += ` r:id="${hl.rId}"`;
        if (hl.location) xml += ` location="${escapeXML(hl.location)}"`;
        if (hl.tooltip) xml += ` tooltip="${escapeXML(hl.tooltip)}"`;
        xml += '/>';
      }
      xml += '</hyperlinks>';
    }

    xml += '</worksheet>';
    return { xml, hyperlinkRels };
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

  // Build all worksheet XMLs
  const sheetNames = workbook.worksheets.map((ws) => ws.name);
  const sheetResults: {
    xml: string;
    hyperlinkRels: { rId: string; target: string }[];
  }[] = [];

  for (let si = 0; si < workbook.worksheets.length; si++) {
    sheetResults.push(buildWorksheetXML(workbook.worksheets[si], si));
  }

  // Build ZIP structure
  const files: Zippable = {
    '[Content_Types].xml': encoder.encode(buildContentTypes(sheetNames.length)),
    '_rels/.rels': encoder.encode(buildRootRels()),
    'xl/_rels/workbook.xml.rels': encoder.encode(
      buildWorkbookRels(sheetNames.length),
    ),
    'xl/workbook.xml': encoder.encode(buildWorkbookXML(sheetNames)),
    'xl/styles.xml': encoder.encode(styleRegistry.buildStylesXML()),
    'xl/sharedStrings.xml': encoder.encode(buildSharedStrings(sharedStrings)),
  };

  for (let i = 0; i < sheetResults.length; i++) {
    files[`xl/worksheets/sheet${i + 1}.xml`] = encoder.encode(
      sheetResults[i].xml,
    );

    // Build per-sheet .rels file for hyperlinks
    const hypRels = sheetResults[i].hyperlinkRels;
    if (hypRels.length > 0) {
      let relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
      relsXml +=
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
      for (const rel of hypRels) {
        relsXml += `<Relationship Id="${rel.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${escapeXML(rel.target)}" TargetMode="External"/>`;
      }
      relsXml += '</Relationships>';
      files[`xl/worksheets/_rels/sheet${i + 1}.xml.rels`] =
        encoder.encode(relsXml);
    }
  }

  // Create ZIP
  return zipSync(files, { level: options?.compress !== false ? 6 : 0 });
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
