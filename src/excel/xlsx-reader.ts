// ============================================
// XLSX Reader — Bun-optimized Excel file reader
// ============================================

import { resolve } from 'node:path';
import { unzipSync } from 'fflate';
import type {
  AlignmentStyle,
  BorderEdgeStyle,
  BorderStyle,
  Cell,
  CellStyle,
  CellValue,
  ColumnConfig,
  ExcelReadOptions,
  FillStyle,
  FontStyle,
  Row,
  Workbook,
  Worksheet,
} from '../types';
import { parseConditionalFormattings } from './conditional-formatting';
import { parseDataValidations } from './data-validation';
import { letterToColIndex } from './xml-builder';
import {
  findChild,
  findChildren,
  getTextContent,
  parseXML,
  type XMLNode,
} from './xml-parser';

// Top-level regex for performance (biome: useTopLevelRegex)
const CELL_REF_REGEX = /^([A-Z]+)(\d+)$/;
const RGB_PREFIX_REGEX = /^FF/;
const COL_LETTER_REGEX = /^([A-Z]+)/;

/** Security limits */
const MAX_FILE_SIZE = 200 * 1024 * 1024; // 200MB max file size
const MAX_DECOMPRESSED_SIZE = 1024 * 1024 * 1024; // 1GB max decompressed size
const MAX_ZIP_ENTRIES = 10_000; // 10K max entries in ZIP
const MAX_ROWS = 1_048_576; // Excel max rows
const MAX_COLS = 16_384; // Excel max columns (XFD)
const MAX_SHARED_STRINGS = 10_000_000; // 10M max shared strings

/**
 * Read an Excel file and return a Workbook
 * Uses Bun.file().arrayBuffer() for optimized binary reading
 */
export async function readExcel(
  path: string,
  options?: ExcelReadOptions,
): Promise<Workbook> {
  const opts = options || {};

  // Validate and resolve path
  const resolvedPath = resolve(path);
  if (resolvedPath.includes('\0')) {
    throw new Error('Invalid file path: contains null bytes');
  }

  // Use Bun.file() for optimized file reading
  const file = Bun.file(resolvedPath);
  const exists = await file.exists();
  if (!exists) {
    throw new Error(`File not found: ${resolvedPath}`);
  }

  // Read as ArrayBuffer and unzip with pre-decompression size check
  const buffer = await file.arrayBuffer();
  if (buffer.byteLength > MAX_FILE_SIZE) {
    throw new Error(
      `File too large: ${buffer.byteLength} bytes (max: ${MAX_FILE_SIZE})`,
    );
  }

  // Zip bomb prevention — check sizes BEFORE decompression via filter callback.
  // fflate's filter receives originalSize (uncompressed) for each entry
  // before it is decompressed, so we can reject without allocating memory.
  let totalDeclaredSize = 0;
  let entryCount = 0;
  const zip = unzipSync(new Uint8Array(buffer), {
    filter(file) {
      entryCount++;
      if (entryCount > MAX_ZIP_ENTRIES) {
        throw new Error(
          `ZIP has too many entries: ${entryCount} (max: ${MAX_ZIP_ENTRIES})`,
        );
      }
      totalDeclaredSize += file.originalSize;
      if (totalDeclaredSize > MAX_DECOMPRESSED_SIZE) {
        throw new Error(
          `Declared decompressed size exceeds limit (max: ${MAX_DECOMPRESSED_SIZE} bytes) — potential zip bomb`,
        );
      }
      return true; // extract this entry
    },
  });

  // Zip Slip prevention — validate all paths inside zip
  for (const path of Object.keys(zip)) {
    if (path.startsWith('/') || path.startsWith('\\') || path.includes('..')) {
      throw new Error(
        `Malicious zip entry detected: "${path}" — potential Zip Slip attack`,
      );
    }
  }

  const decoder = new TextDecoder('utf-8');

  // Parse shared strings
  const sharedStrings = parseSharedStrings(zip, decoder);

  // Parse styles
  const styles =
    opts.includeStyles !== false
      ? parseStyles(zip, decoder)
      : { cellStyles: [], differentialStyles: [] };

  // Parse workbook to get sheet info
  const workbookXML = decoder.decode(zip['xl/workbook.xml']);
  const workbookDoc = parseXML(workbookXML);
  const sheetsNode = findChild(workbookDoc.children[0], 'sheets');
  const sheetNodes = sheetsNode ? findChildren(sheetsNode, 'sheet') : [];

  // Parse workbook rels to get sheet paths
  const relsXML = decoder.decode(zip['xl/_rels/workbook.xml.rels']);
  const relsDoc = parseXML(relsXML);
  const relMap = new Map<string, string>();
  for (const rel of relsDoc.children[0]?.children || []) {
    relMap.set(rel.attributes.Id, rel.attributes.Target);
  }

  // Parse worksheets
  const worksheets: Worksheet[] = [];

  for (let i = 0; i < sheetNodes.length; i++) {
    const sheetNode = sheetNodes[i];
    const sheetName = sheetNode.attributes.name;
    const rId = sheetNode.attributes['r:id'];

    // Filter sheets if specified
    if (opts.sheets) {
      const matchesName = (
        opts.sheets as readonly (string | number)[]
      ).includes(sheetName);
      const matchesIndex = (
        opts.sheets as readonly (string | number)[]
      ).includes(i);
      if (!matchesName && !matchesIndex) continue;
    }

    const target = relMap.get(rId);
    if (!target) continue;

    const sheetPath = target.startsWith('/') ? target.slice(1) : `xl/${target}`;

    // Zip Slip: validate resolved sheet path
    if (sheetPath.includes('..') || sheetPath.startsWith('/')) {
      continue; // skip suspicious paths
    }

    const sheetData = zip[sheetPath];
    if (!sheetData) continue;

    const sheetXML = decoder.decode(sheetData);

    // Parse per-sheet rels for hyperlinks
    const sheetRelsPath = `xl/worksheets/_rels/${target.includes('/') ? target.split('/').pop() : target}.rels`;
    const sheetRelsData = zip[sheetRelsPath];
    const hyperlinkRelMap = new Map<string, string>();
    if (sheetRelsData) {
      const relsXMLStr = decoder.decode(sheetRelsData);
      const relsDoc2 = parseXML(relsXMLStr);
      for (const rel of relsDoc2.children[0]?.children || []) {
        if (rel.attributes.Type?.includes('hyperlink')) {
          hyperlinkRelMap.set(rel.attributes.Id, rel.attributes.Target);
        }
      }
    }

    const worksheet = parseWorksheet(
      sheetXML,
      sheetName,
      sharedStrings,
      styles.cellStyles,
      styles.differentialStyles,
      hyperlinkRelMap,
    );
    worksheets.push(worksheet);
  }

  return { worksheets };
}

/**
 * Parse shared strings from XLSX
 */
function parseSharedStrings(
  zip: Record<string, Uint8Array>,
  decoder: TextDecoder,
): string[] {
  const data = zip['xl/sharedStrings.xml'];
  if (!data) return [];

  const xml = decoder.decode(data);
  const doc = parseXML(xml);
  const strings: string[] = [];

  const siNodes = findChildren(doc.children[0], 'si');
  if (siNodes.length > MAX_SHARED_STRINGS) {
    throw new Error(
      `Too many shared strings: ${siNodes.length} (max: ${MAX_SHARED_STRINGS})`,
    );
  }
  for (const si of siNodes) {
    const tNode = findChild(si, 't');
    if (tNode) {
      strings.push(getTextContent(tNode));
    } else {
      // Handle rich text <r><t>...</t></r>
      let text = '';
      const rNodes = findChildren(si, 'r');
      for (const r of rNodes) {
        const t = findChild(r, 't');
        if (t) text += getTextContent(t);
      }
      strings.push(text);
    }
  }

  return strings;
}

/**
 * Parse styles from XLSX
 */
function parseStyles(
  zip: Record<string, Uint8Array>,
  decoder: TextDecoder,
): { cellStyles: CellStyle[]; differentialStyles: CellStyle[] } {
  const data = zip['xl/styles.xml'];
  if (!data) return { cellStyles: [], differentialStyles: [] };

  const xml = decoder.decode(data);
  const doc = parseXML(xml);
  const root = doc.children[0];
  if (!root) return { cellStyles: [], differentialStyles: [] };

  // Parse fonts
  const fonts: FontStyle[] = [];
  const fontsNode = findChild(root, 'fonts');
  if (fontsNode) {
    for (const fontNode of findChildren(fontsNode, 'font')) {
      fonts.push(parseFontNode(fontNode));
    }
  }

  // Parse fills
  const fills: FillStyle[] = [];
  const fillsNode = findChild(root, 'fills');
  if (fillsNode) {
    for (const fillNode of findChildren(fillsNode, 'fill')) {
      fills.push(parseFillNode(fillNode));
    }
  }

  // Parse borders
  const bordersList: BorderStyle[] = [];
  const bordersNode = findChild(root, 'borders');
  if (bordersNode) {
    for (const borderNode of findChildren(bordersNode, 'border')) {
      bordersList.push(parseBorderNode(borderNode));
    }
  }

  // Parse number formats
  const numFmtMap = new Map<number, string>();
  const numFmtsNode = findChild(root, 'numFmts');
  if (numFmtsNode) {
    for (const nf of findChildren(numFmtsNode, 'numFmt')) {
      numFmtMap.set(
        Number.parseInt(nf.attributes.numFmtId, 10),
        nf.attributes.formatCode,
      );
    }
  }

  // Parse cell xfs
  const cellStyles: CellStyle[] = [];
  const cellXfsNode = findChild(root, 'cellXfs');
  if (cellXfsNode) {
    for (const xf of findChildren(cellXfsNode, 'xf')) {
      cellStyles.push(
        parseCellXfStyle(xf, fonts, fills, bordersList, numFmtMap),
      );
    }
  }

  const differentialStyles: CellStyle[] = [];
  const dxfsNode = findChild(root, 'dxfs');
  if (dxfsNode) {
    for (const dxf of findChildren(dxfsNode, 'dxf')) {
      const style: CellStyle = {};
      const font = findChild(dxf, 'font');
      const fill = findChild(dxf, 'fill');
      const border = findChild(dxf, 'border');
      const numFmt = findChild(dxf, 'numFmt');
      const alignment = findChild(dxf, 'alignment');

      if (font) style.font = parseFontNode(font);
      if (fill) style.fill = parseFillNode(fill);
      if (border) style.border = parseBorderNode(border);
      if (numFmt?.attributes.formatCode) {
        style.numberFormat = numFmt.attributes.formatCode;
      }
      if (alignment) {
        style.alignment = parseAlignmentNode(alignment);
      }

      differentialStyles.push(style);
    }
  }

  return { cellStyles, differentialStyles };
}

function parseFontNode(fontNode: XMLNode): FontStyle {
  const font: FontStyle = {};
  if (findChild(fontNode, 'b')) font.bold = true;
  if (findChild(fontNode, 'i')) font.italic = true;
  if (findChild(fontNode, 'u')) font.underline = true;
  if (findChild(fontNode, 'strike')) font.strike = true;

  const sz = findChild(fontNode, 'sz');
  if (sz?.attributes.val) {
    font.size = Number.parseFloat(sz.attributes.val);
  }

  const name = findChild(fontNode, 'name');
  if (name?.attributes.val) {
    font.name = name.attributes.val;
  }

  const color = findChild(fontNode, 'color');
  if (color?.attributes.rgb) {
    font.color = color.attributes.rgb.replace(RGB_PREFIX_REGEX, '');
  }

  return font;
}

function parseFillNode(fillNode: XMLNode): FillStyle {
  const patternFill =
    fillNode.tag === 'patternFill'
      ? fillNode
      : findChild(fillNode, 'patternFill');

  if (patternFill) {
    const fill: FillStyle = {
      type: 'pattern',
      pattern:
        (patternFill.attributes.patternType as FillStyle['pattern']) || 'none',
    };
    const fgColor = findChild(patternFill, 'fgColor');
    if (fgColor?.attributes.rgb) {
      fill.fgColor = fgColor.attributes.rgb.replace(RGB_PREFIX_REGEX, '');
    }
    const bgColor = findChild(patternFill, 'bgColor');
    if (bgColor?.attributes.rgb) {
      fill.bgColor = bgColor.attributes.rgb.replace(RGB_PREFIX_REGEX, '');
    }
    return fill;
  }

  return { type: 'pattern', pattern: 'none' };
}

function parseBorderNode(borderNode: XMLNode): BorderStyle {
  const border: BorderStyle = {};
  for (const side of ['left', 'right', 'top', 'bottom'] as const) {
    const sideNode = findChild(borderNode, side);
    if (sideNode?.attributes.style) {
      const edge: BorderEdgeStyle = {
        style: sideNode.attributes.style as BorderEdgeStyle['style'],
      };
      const color = findChild(sideNode, 'color');
      if (color?.attributes.rgb) {
        edge.color = color.attributes.rgb.replace(RGB_PREFIX_REGEX, '');
      }
      border[side] = edge;
    }
  }
  return border;
}

function parseAlignmentNode(alignment: XMLNode): AlignmentStyle {
  const style: AlignmentStyle = {};
  if (alignment.attributes.horizontal) {
    style.horizontal = alignment.attributes
      .horizontal as AlignmentStyle['horizontal'];
  }
  if (alignment.attributes.vertical) {
    style.vertical = alignment.attributes
      .vertical as AlignmentStyle['vertical'];
  }
  if (alignment.attributes.wrapText === '1') {
    style.wrapText = true;
  }
  if (alignment.attributes.textRotation) {
    style.textRotation = Number.parseInt(alignment.attributes.textRotation, 10);
  }
  if (alignment.attributes.indent) {
    style.indent = Number.parseInt(alignment.attributes.indent, 10);
  }
  return style;
}

function parseCellXfStyle(
  xf: XMLNode,
  fonts: FontStyle[],
  fills: FillStyle[],
  bordersList: BorderStyle[],
  numFmtMap: Map<number, string>,
): CellStyle {
  const style: CellStyle = {};
  const fontId = Number.parseInt(xf.attributes.fontId || '0', 10);
  const fillId = Number.parseInt(xf.attributes.fillId || '0', 10);
  const borderId = Number.parseInt(xf.attributes.borderId || '0', 10);
  const numFmtId = Number.parseInt(xf.attributes.numFmtId || '0', 10);

  if (xf.attributes.applyFont === '1' && fonts[fontId]) {
    style.font = fonts[fontId];
  }
  if (xf.attributes.applyFill === '1' && fills[fillId]) {
    style.fill = fills[fillId];
  }
  if (xf.attributes.applyBorder === '1' && bordersList[borderId]) {
    style.border = bordersList[borderId];
  }
  if (xf.attributes.applyNumberFormat === '1' && numFmtId > 0) {
    style.numberFormat = numFmtMap.get(numFmtId) || '';
  }

  const alignment = findChild(xf, 'alignment');
  if (alignment) {
    style.alignment = parseAlignmentNode(alignment);
  }

  return style;
}

/**
 * Parse a single worksheet
 */
function parseWorksheet(
  xml: string,
  name: string,
  sharedStrings: string[],
  styles: CellStyle[],
  differentialStyles: CellStyle[],
  hyperlinkRelMap: Map<string, string>,
): Worksheet {
  const doc = parseXML(xml);
  const root = doc.children[0];

  const worksheet: Worksheet = { name, rows: [] };

  // Parse columns
  const colsNode = findChild(root, 'cols');
  if (colsNode) {
    const colConfigs: ColumnConfig[] = [];
    for (const col of findChildren(colsNode, 'col')) {
      const min = Number.parseInt(col.attributes.min, 10) - 1;
      const max = Number.parseInt(col.attributes.max, 10) - 1;
      const width = col.attributes.width
        ? Number.parseFloat(col.attributes.width)
        : undefined;
      for (let c = min; c <= max; c++) {
        if (c >= MAX_COLS) break; // cap column expansion
        while (colConfigs.length <= c) colConfigs.push({});
        colConfigs[c] = { width };
      }
    }
    worksheet.columns = colConfigs;
  }

  // Parse merge cells
  const mergeCellsNode = findChild(root, 'mergeCells');
  if (mergeCellsNode) {
    worksheet.mergeCells = [];
    for (const mc of findChildren(mergeCellsNode, 'mergeCell')) {
      const ref = mc.attributes.ref;
      const [start, end] = ref.split(':');
      const startMatch = start.match(CELL_REF_REGEX);
      const endMatch = end.match(CELL_REF_REGEX);
      if (startMatch && endMatch) {
        worksheet.mergeCells.push({
          startRow: Number.parseInt(startMatch[2], 10) - 1,
          startCol: letterToColIndex(startMatch[1]),
          endRow: Number.parseInt(endMatch[2], 10) - 1,
          endCol: letterToColIndex(endMatch[1]),
        });
      }
    }
  }

  const conditionalFormattings = parseConditionalFormattings(
    root,
    differentialStyles,
  );
  if (conditionalFormattings.length > 0) {
    worksheet.conditionalFormattings = conditionalFormattings;
  }

  // Parse sheet data
  const sheetDataNode = findChild(root, 'sheetData');
  if (!sheetDataNode) return worksheet;

  const rows: Row[] = [];
  let maxRow = 0;

  for (const rowNode of findChildren(sheetDataNode, 'row')) {
    const rowIndex = Number.parseInt(rowNode.attributes.r, 10) - 1;
    if (rowIndex < 0 || rowIndex >= MAX_ROWS) continue; // bounds check

    const height = rowNode.attributes.ht
      ? Number.parseFloat(rowNode.attributes.ht)
      : undefined;

    // Ensure rows array is filled up to this index
    while (rows.length <= rowIndex) {
      rows.push({ cells: [] });
    }

    const row: Row = { cells: [], height };
    const cells: Cell[] = [];

    for (const cellNode of findChildren(rowNode, 'c')) {
      const ref = cellNode.attributes.r;
      const match = ref.match(COL_LETTER_REGEX);
      if (!match) continue;

      const colIndex = letterToColIndex(match[1]);
      if (colIndex < 0 || colIndex >= MAX_COLS) continue; // bounds check

      const cellType = cellNode.attributes.t;
      const styleIndex = Number.parseInt(cellNode.attributes.s || '0', 10);

      // Ensure cells array is filled
      while (cells.length <= colIndex) {
        cells.push({ value: null });
      }

      const vNode = findChild(cellNode, 'v');
      let value: CellValue = null;
      let type: Cell['type'] = 'string';

      if (vNode) {
        const rawValue = getTextContent(vNode);

        if (cellType === 's') {
          // Shared string
          const idx = Number.parseInt(rawValue, 10);
          // Bounds check shared string index
          if (idx >= 0 && idx < sharedStrings.length) {
            value = sharedStrings[idx];
          } else {
            value = rawValue;
          }
          type = 'string';
        } else if (cellType === 'b') {
          value = rawValue === '1';
          type = 'boolean';
        } else if (cellType === 'str' || cellType === 'inlineStr') {
          value = rawValue;
          type = 'string';
        } else {
          // Number or date
          const num = Number.parseFloat(rawValue);
          value = Number.isNaN(num) ? rawValue : num;
          type = 'number';
        }
      }

      // Check inline string
      const isNode = findChild(cellNode, 'is');
      if (isNode) {
        const tNode = findChild(isNode, 't');
        if (tNode) {
          value = getTextContent(tNode);
          type = 'string';
        }
      }

      const cell: Cell = { value, type };

      // Parse formula
      const fNode = findChild(cellNode, 'f');
      if (fNode) {
        cell.formula = getTextContent(fNode);
        cell.type = 'formula';
      }

      // Apply style
      if (styleIndex > 0 && styles[styleIndex]) {
        cell.style = styles[styleIndex];
      }

      cells[colIndex] = cell;
    }

    row.cells = cells;
    rows[rowIndex] = row;
    maxRow = Math.max(maxRow, rowIndex);
  }

  worksheet.rows = rows.slice(0, maxRow + 1);

  // Parse hyperlinks
  const dataValidations = parseDataValidations(root);
  if (dataValidations.length > 0) {
    worksheet.dataValidations = dataValidations;
  }

  // Parse hyperlinks
  const hyperlinksNode = findChild(root, 'hyperlinks');
  if (hyperlinksNode) {
    for (const hlNode of findChildren(hyperlinksNode, 'hyperlink')) {
      const ref = hlNode.attributes.ref;
      if (!ref) continue;

      const refMatch = ref.match(CELL_REF_REGEX);
      if (!refMatch) continue;

      const hlCol = letterToColIndex(refMatch[1]);
      const hlRow = Number.parseInt(refMatch[2], 10) - 1;

      if (hlRow >= 0 && hlRow < worksheet.rows.length) {
        const row = worksheet.rows[hlRow];
        if (row && hlCol >= 0 && hlCol < row.cells.length) {
          const cell = row.cells[hlCol];
          if (cell) {
            const rId = hlNode.attributes['r:id'];
            const location = hlNode.attributes.location;
            const tooltip = hlNode.attributes.tooltip;

            let target = '';
            if (rId && hyperlinkRelMap.has(rId)) {
              target = hyperlinkRelMap.get(rId) ?? '';
            } else if (location) {
              target = location;
            }

            if (target) {
              cell.hyperlink = { target, tooltip };
            }
          }
        }
      }
    }
  }

  // Parse freeze pane
  const sheetViews = findChild(root, 'sheetViews');
  if (sheetViews) {
    const sheetView = findChild(sheetViews, 'sheetView');
    if (sheetView) {
      const pane = findChild(sheetView, 'pane');
      if (pane) {
        const xSplit = Number.parseInt(pane.attributes.xSplit || '0', 10);
        const ySplit = Number.parseInt(pane.attributes.ySplit || '0', 10);
        if (xSplit > 0 || ySplit > 0) {
          worksheet.freezePane = { row: ySplit, col: xSplit };
        }
      }
    }
  }

  return worksheet;
}
