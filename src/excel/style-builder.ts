// ============================================
// Style Builder for XLSX styles.xml
// ============================================

import type {
  AlignmentStyle,
  BorderEdgeStyle,
  BorderStyle,
  CellStyle,
  FillStyle,
  FontStyle,
} from '../types';
import { escapeXML, getFiniteNumberOr } from './xml-builder';

/** Internal style registry to deduplicate and index styles */
export class StyleRegistry {
  private fonts: string[] = [];
  private fills: string[] = [];
  private borders: string[] = [];
  private numberFormats: Map<string, number> = new Map();
  private cellXfs: string[] = [];
  private styleMap: Map<string, number> = new Map();
  private dxfs: string[] = [];
  private dxfMap: Map<string, number> = new Map();
  private nextNumFmtId = 164; // Custom number formats start at 164

  constructor() {
    // Default font (index 0)
    this.fonts.push(
      '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>',
    );
    // Default fills (indices 0 and 1 are required)
    this.fills.push('<fill><patternFill patternType="none"/></fill>');
    this.fills.push('<fill><patternFill patternType="gray125"/></fill>');
    // Default border (index 0)
    this.borders.push(
      '<border><left/><right/><top/><bottom/><diagonal/></border>',
    );
    // Default cell xf (index 0)
    this.cellXfs.push('<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>');
  }

  /**
   * Register a style and return its xf index
   */
  registerStyle(style: CellStyle | undefined): number {
    if (!style) return 0;

    const key = JSON.stringify(style);
    const existing = this.styleMap.get(key);
    if (existing !== undefined) return existing;

    const fontId = style.font ? this.registerFont(style.font) : 0;
    const fillId = style.fill ? this.registerFill(style.fill) : 0;
    const borderId = style.border ? this.registerBorder(style.border) : 0;
    let numFmtId = 0;
    if (style.numberFormat) {
      numFmtId = this.registerNumberFormat(style.numberFormat);
    }

    let xf = `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}"`;

    if (fontId > 0) xf += ' applyFont="1"';
    if (fillId > 0) xf += ' applyFill="1"';
    if (borderId > 0) xf += ' applyBorder="1"';
    if (numFmtId > 0) xf += ' applyNumberFormat="1"';

    if (style.alignment) {
      xf += ' applyAlignment="1">';
      xf += this.buildAlignment(style.alignment);
      xf += '</xf>';
    } else {
      xf += '/>';
    }

    const index = this.cellXfs.length;
    this.cellXfs.push(xf);
    this.styleMap.set(key, index);
    return index;
  }

  registerDifferentialStyle(style: CellStyle | undefined): number | undefined {
    if (!style) return undefined;

    const xml = this.buildDifferentialStyle(style);
    if (!xml) return undefined;

    const existing = this.dxfMap.get(xml);
    if (existing !== undefined) return existing;

    const index = this.dxfs.length;
    this.dxfs.push(xml);
    this.dxfMap.set(xml, index);
    return index;
  }

  private registerFont(font: FontStyle): number {
    let xml = '<font>';
    if (font.bold) xml += '<b/>';
    if (font.italic) xml += '<i/>';
    if (font.underline) xml += '<u/>';
    if (font.strike) xml += '<strike/>';
    xml += `<sz val="${getFiniteNumberOr(font.size, 11)}"/>`;
    if (font.color) {
      xml += `<color rgb="FF${escapeXML(font.color)}"/>`;
    } else {
      xml += '<color theme="1"/>';
    }
    xml += `<name val="${escapeXML(font.name || 'Calibri')}"/>`;
    xml += '<family val="2"/>';
    xml += '</font>';

    const existing = this.fonts.indexOf(xml);
    if (existing !== -1) return existing;
    this.fonts.push(xml);
    return this.fonts.length - 1;
  }

  private registerFill(fill: FillStyle): number {
    let xml = '<fill>';
    if (fill.type === 'pattern') {
      xml += `<patternFill patternType="${escapeXML(fill.pattern || 'solid')}">`;
      if (fill.fgColor) {
        xml += `<fgColor rgb="FF${escapeXML(fill.fgColor)}"/>`;
      }
      if (fill.bgColor) {
        xml += `<bgColor rgb="FF${escapeXML(fill.bgColor)}"/>`;
      } else if (fill.fgColor && fill.pattern === 'solid') {
        xml += `<bgColor indexed="64"/>`;
      }
      xml += '</patternFill>';
    } else {
      xml += '<patternFill patternType="none"/>';
    }
    xml += '</fill>';

    const existing = this.fills.indexOf(xml);
    if (existing !== -1) return existing;
    this.fills.push(xml);
    return this.fills.length - 1;
  }

  private registerBorder(border: BorderStyle): number {
    let xml = '<border>';
    xml += this.buildBorderEdge('left', border.left);
    xml += this.buildBorderEdge('right', border.right);
    xml += this.buildBorderEdge('top', border.top);
    xml += this.buildBorderEdge('bottom', border.bottom);
    xml += '<diagonal/>';
    xml += '</border>';

    const existing = this.borders.indexOf(xml);
    if (existing !== -1) return existing;
    this.borders.push(xml);
    return this.borders.length - 1;
  }

  private buildBorderEdge(side: string, edge?: BorderEdgeStyle): string {
    if (!edge || !edge.style) return `<${escapeXML(side)}/>`;
    let xml = `<${escapeXML(side)} style="${escapeXML(edge.style)}">`;
    if (edge.color) {
      xml += `<color rgb="FF${escapeXML(edge.color)}"/>`;
    }
    xml += `</${escapeXML(side)}>`;
    return xml;
  }

  private registerNumberFormat(format: string): number {
    // Built-in formats
    const builtIn: Record<string, number> = {
      General: 0,
      '0': 1,
      '0.00': 2,
      '#,##0': 3,
      '#,##0.00': 4,
      '0%': 9,
      '0.00%': 10,
      '0.00E+00': 11,
      'mm-dd-yy': 14,
      'd-mmm-yy': 15,
      'd-mmm': 16,
      'mmm-yy': 17,
      'h:mm AM/PM': 18,
      'h:mm:ss AM/PM': 19,
      'h:mm': 20,
      'h:mm:ss': 21,
      'm/d/yy h:mm': 22,
      'yyyy-mm-dd': 14,
    };

    if (builtIn[format] !== undefined) return builtIn[format];

    const existing = this.numberFormats.get(format);
    if (existing !== undefined) return existing;

    const id = this.nextNumFmtId++;
    this.numberFormats.set(format, id);
    return id;
  }

  private buildAlignment(align: AlignmentStyle): string {
    let xml = '<alignment';
    if (align.horizontal)
      xml += ` horizontal="${escapeXML(String(align.horizontal))}"`;
    if (align.vertical)
      xml += ` vertical="${escapeXML(String(align.vertical))}"`;
    if (align.wrapText) xml += ' wrapText="1"';
    if (align.textRotation !== undefined)
      xml += ` textRotation="${getFiniteNumberOr(align.textRotation, 0)}"`;
    if (align.indent !== undefined)
      xml += ` indent="${getFiniteNumberOr(align.indent, 0)}"`;
    xml += '/>';
    return xml;
  }

  private buildDifferentialStyle(style: CellStyle): string {
    let xml = '<dxf>';

    if (style.font) {
      xml += this.buildDifferentialFont(style.font);
    }
    if (style.fill) {
      xml += this.buildDifferentialFill(style.fill);
    }
    if (style.border) {
      xml += this.buildDifferentialBorder(style.border);
    }
    if (style.numberFormat) {
      const numFmtId = this.registerNumberFormat(style.numberFormat);
      xml += `<numFmt numFmtId="${numFmtId}" formatCode="${escapeXML(style.numberFormat)}"/>`;
    }
    if (style.alignment) {
      xml += this.buildAlignment(style.alignment);
    }

    xml += '</dxf>';
    return xml === '<dxf></dxf>' ? '' : xml;
  }

  private buildDifferentialFont(font: FontStyle): string {
    let xml = '<font>';
    let hasContent = false;

    if (font.bold) {
      xml += '<b/>';
      hasContent = true;
    }
    if (font.italic) {
      xml += '<i/>';
      hasContent = true;
    }
    if (font.underline) {
      xml += '<u/>';
      hasContent = true;
    }
    if (font.strike) {
      xml += '<strike/>';
      hasContent = true;
    }
    if (font.size !== undefined) {
      xml += `<sz val="${getFiniteNumberOr(font.size, 11)}"/>`;
      hasContent = true;
    }
    if (font.color) {
      xml += `<color rgb="FF${escapeXML(font.color)}"/>`;
      hasContent = true;
    }
    if (font.name) {
      xml += `<name val="${escapeXML(font.name)}"/>`;
      hasContent = true;
    }

    xml += '</font>';
    return hasContent ? xml : '';
  }

  private buildDifferentialFill(fill: FillStyle): string {
    if (fill.type !== 'pattern') return '';

    let xml = '<fill>';
    let hasContent = false;
    xml += `<patternFill patternType="${escapeXML(fill.pattern || 'solid')}">`;
    if (fill.fgColor) {
      xml += `<fgColor rgb="FF${escapeXML(fill.fgColor)}"/>`;
      hasContent = true;
    }
    if (fill.bgColor) {
      xml += `<bgColor rgb="FF${escapeXML(fill.bgColor)}"/>`;
      hasContent = true;
    } else if (fill.fgColor && fill.pattern === 'solid') {
      xml += '<bgColor indexed="64"/>';
      hasContent = true;
    }
    xml += '</patternFill>';
    xml += '</fill>';
    return hasContent || fill.pattern !== undefined ? xml : '';
  }

  private buildDifferentialBorder(border: BorderStyle): string {
    const hasEdges = border.left || border.right || border.top || border.bottom;
    if (!hasEdges) return '';

    let xml = '<border>';
    xml += this.buildBorderEdge('left', border.left);
    xml += this.buildBorderEdge('right', border.right);
    xml += this.buildBorderEdge('top', border.top);
    xml += this.buildBorderEdge('bottom', border.bottom);
    xml += '<diagonal/>';
    xml += '</border>';
    return xml;
  }

  /**
   * Build the complete styles.xml content
   */
  buildStylesXML(): string {
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml +=
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';

    // Number formats
    if (this.numberFormats.size > 0) {
      xml += `<numFmts count="${this.numberFormats.size}">`;
      for (const [format, id] of this.numberFormats) {
        xml += `<numFmt numFmtId="${id}" formatCode="${escapeXML(format)}"/>`;
      }
      xml += '</numFmts>';
    }

    // Fonts
    xml += `<fonts count="${this.fonts.length}">`;
    xml += this.fonts.join('');
    xml += '</fonts>';

    // Fills
    xml += `<fills count="${this.fills.length}">`;
    xml += this.fills.join('');
    xml += '</fills>';

    // Borders
    xml += `<borders count="${this.borders.length}">`;
    xml += this.borders.join('');
    xml += '</borders>';

    // Cell style xfs (required even if empty)
    xml += '<cellStyleXfs count="1">';
    xml += '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>';
    xml += '</cellStyleXfs>';

    // Cell xfs
    xml += `<cellXfs count="${this.cellXfs.length}">`;
    xml += this.cellXfs.join('');
    xml += '</cellXfs>';

    // Cell styles (required)
    xml += '<cellStyles count="1">';
    xml += '<cellStyle name="Normal" xfId="0" builtinId="0"/>';
    xml += '</cellStyles>';

    // Differential styles (for conditional formatting)
    xml += `<dxfs count="${this.dxfs.length}">`;
    xml += this.dxfs.join('');
    xml += '</dxfs>';

    xml += '</styleSheet>';
    return xml;
  }
}
