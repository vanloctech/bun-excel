import type {
  CellRange,
  CellStyle,
  ConditionalFormatThreshold,
  ConditionalFormatting,
  ConditionalFormattingRule,
} from '../types';
import type { StyleRegistry } from './style-builder';
import {
  buildRangeRef,
  escapeXML,
  getFiniteNumber,
  getFiniteNumberOr,
  getNonNegativeIntegerOr,
  parseCellRef,
} from './xml-builder';
import {
  findChild,
  findChildren,
  getTextContent,
  type XMLNode,
} from './xml-parser';

const WHITESPACE_SPLIT_REGEX = /\s+/;
const RGB_PREFIX_REGEX = /^FF/;
const HASH_PREFIX_REGEX = /^#/;

function dateToExcelSerial(date: Date): number {
  const epoch = new Date(Date.UTC(1899, 11, 30));
  return (date.getTime() - epoch.getTime()) / (24 * 60 * 60 * 1000);
}

function normalizeColor(color: string): string {
  const normalized = color.replace(HASH_PREFIX_REGEX, '').toUpperCase();
  if (normalized.length === 8) return normalized;
  return `FF${normalized}`;
}

function denormalizeColor(color: string): string {
  return color.replace(RGB_PREFIX_REGEX, '');
}

function normalizeRange(range: CellRange): CellRange {
  const startRow = getNonNegativeIntegerOr(range.startRow, 0);
  const startCol = getNonNegativeIntegerOr(range.startCol, 0);
  const endRow = getNonNegativeIntegerOr(range.endRow, startRow);
  const endCol = getNonNegativeIntegerOr(range.endCol, startCol);
  return {
    startRow: Math.min(startRow, endRow),
    startCol: Math.min(startCol, endCol),
    endRow: Math.max(startRow, endRow),
    endCol: Math.max(startCol, endCol),
  };
}

function normalizeRanges(range: CellRange | CellRange[]): CellRange[] {
  return (Array.isArray(range) ? range : [range]).map(normalizeRange);
}

function buildSqref(range: CellRange | CellRange[]): string {
  return normalizeRanges(range)
    .map((entry) =>
      buildRangeRef(entry.startRow, entry.startCol, entry.endRow, entry.endCol),
    )
    .join(' ');
}

function parseSqref(sqref: string): CellRange[] {
  const ranges: CellRange[] = [];

  for (const part of sqref.split(WHITESPACE_SPLIT_REGEX)) {
    if (!part) continue;
    const [startRef, endRef] = part.split(':');
    try {
      const start = parseCellRef(startRef);
      const end = parseCellRef(endRef || startRef);
      ranges.push({
        startRow: Math.min(start.row, end.row),
        startCol: Math.min(start.col, end.col),
        endRow: Math.max(start.row, end.row),
        endCol: Math.max(start.col, end.col),
      });
    } catch {
      // Skip malformed ranges
    }
  }

  return ranges;
}

function stripLeadingEquals(value: string): string {
  return value.startsWith('=') ? value.slice(1) : value;
}

function serializeFormulaValue(
  value: string | number | Date | undefined,
): string | undefined {
  if (value === undefined) return undefined;
  if (value instanceof Date) return String(dateToExcelSerial(value));
  if (typeof value === 'number') {
    const numeric = getFiniteNumber(value);
    return numeric !== undefined ? String(numeric) : undefined;
  }
  return stripLeadingEquals(String(value));
}

function parseFormulaValue(raw: string): string | number {
  const numeric = getFiniteNumber(raw);
  return numeric !== undefined ? numeric : raw;
}

function serializeThreshold(threshold: ConditionalFormatThreshold): string {
  let xml = `<cfvo type="${escapeXML(threshold.type)}"`;
  const value = serializeFormulaValue(
    threshold.value as string | number | Date | undefined,
  );
  if (value !== undefined) {
    xml += ` val="${escapeXML(value)}"`;
  }
  if (threshold.gte !== undefined) {
    xml += ` gte="${threshold.gte ? 1 : 0}"`;
  }
  xml += '/>';
  return xml;
}

function parseThreshold(node: XMLNode): ConditionalFormatThreshold {
  return {
    type: node.attributes.type as ConditionalFormatThreshold['type'],
    value:
      node.attributes.val !== undefined
        ? parseFormulaValue(node.attributes.val)
        : undefined,
    gte: node.attributes.gte === '1',
  };
}

function buildRuleAttributes(
  rule: ConditionalFormattingRule,
  priority: number,
  styleRegistry: StyleRegistry,
): string {
  let attrs = ` type="${escapeXML(rule.type)}" priority="${priority}"`;
  if ('stopIfTrue' in rule && rule.stopIfTrue) {
    attrs += ' stopIfTrue="1"';
  }

  if (
    (rule.type === 'cellIs' || rule.type === 'expression') &&
    'style' in rule &&
    rule.style
  ) {
    const dxfId = styleRegistry.registerDifferentialStyle(rule.style);
    if (dxfId !== undefined) {
      attrs += ` dxfId="${dxfId}"`;
    }
  }

  if (rule.type === 'cellIs') {
    attrs += ` operator="${escapeXML(rule.operator)}"`;
  }

  return attrs;
}

function buildRuleXML(
  rule: ConditionalFormattingRule,
  priority: number,
  styleRegistry: StyleRegistry,
): string {
  let xml = `<cfRule${buildRuleAttributes(rule, priority, styleRegistry)}>`;

  if (rule.type === 'cellIs') {
    const formula1 = serializeFormulaValue(rule.formula1);
    const formula2 = serializeFormulaValue(rule.formula2);
    if (formula1 !== undefined) {
      xml += `<formula>${escapeXML(formula1)}</formula>`;
    }
    if (formula2 !== undefined) {
      xml += `<formula>${escapeXML(formula2)}</formula>`;
    }
  } else if (rule.type === 'expression') {
    xml += `<formula>${escapeXML(stripLeadingEquals(rule.formula))}</formula>`;
  } else if (rule.type === 'colorScale') {
    xml += '<colorScale>';
    for (const threshold of rule.thresholds) {
      xml += serializeThreshold(threshold);
    }
    for (const color of rule.colors) {
      xml += `<color rgb="${escapeXML(normalizeColor(color))}"/>`;
    }
    xml += '</colorScale>';
  } else if (rule.type === 'dataBar') {
    const min = rule.min || { type: 'min' };
    const max = rule.max || { type: 'max' };
    xml += '<dataBar';
    if (rule.showValue !== undefined) {
      xml += ` showValue="${rule.showValue ? 1 : 0}"`;
    }
    if (rule.minLength !== undefined) {
      xml += ` minLength="${getFiniteNumberOr(rule.minLength, 10)}"`;
    }
    if (rule.maxLength !== undefined) {
      xml += ` maxLength="${getFiniteNumberOr(rule.maxLength, 90)}"`;
    }
    xml += '>';
    xml += serializeThreshold(min);
    xml += serializeThreshold(max);
    xml += `<color rgb="${escapeXML(normalizeColor(rule.color))}"/>`;
    xml += '</dataBar>';
  } else if (rule.type === 'iconSet') {
    xml += `<iconSet iconSet="${escapeXML(rule.iconSet)}"`;
    if (rule.showValue !== undefined) {
      xml += ` showValue="${rule.showValue ? 1 : 0}"`;
    }
    if (rule.reverse) {
      xml += ' reverse="1"';
    }
    xml += '>';
    for (const threshold of rule.thresholds) {
      xml += serializeThreshold(threshold);
    }
    xml += '</iconSet>';
  }

  xml += '</cfRule>';
  return xml;
}

export function buildConditionalFormattingsXML(
  formattings: ConditionalFormatting[] | undefined,
  styleRegistry: StyleRegistry,
): string {
  if (!formattings || formattings.length === 0) return '';

  let nextPriority = 1;
  let xml = '';

  for (const formatting of formattings) {
    const sqref = buildSqref(formatting.range);
    if (!sqref || formatting.rules.length === 0) continue;

    xml += `<conditionalFormatting sqref="${escapeXML(sqref)}">`;
    for (const rule of formatting.rules) {
      const priority = rule.priority ?? nextPriority++;
      nextPriority = Math.max(nextPriority, priority + 1);
      xml += buildRuleXML(rule, priority, styleRegistry);
    }
    xml += '</conditionalFormatting>';
  }

  return xml;
}

function parseColorScaleRule(node: XMLNode, priority: number) {
  const colorScale = findChild(node, 'colorScale');
  if (!colorScale) return undefined;

  return {
    type: 'colorScale' as const,
    priority,
    thresholds: findChildren(colorScale, 'cfvo').map(parseThreshold),
    colors: findChildren(colorScale, 'color')
      .map((color) => color.attributes.rgb)
      .filter((color): color is string => Boolean(color))
      .map(denormalizeColor),
  };
}

function parseDataBarRule(node: XMLNode, priority: number) {
  const dataBar = findChild(node, 'dataBar');
  if (!dataBar) return undefined;

  const thresholds = findChildren(dataBar, 'cfvo').map(parseThreshold);
  const color = findChild(dataBar, 'color')?.attributes.rgb;

  if (!color) return undefined;

  return {
    type: 'dataBar' as const,
    priority,
    min: thresholds[0],
    max: thresholds[1],
    color: denormalizeColor(color),
    showValue:
      dataBar.attributes.showValue !== undefined
        ? dataBar.attributes.showValue === '1'
        : undefined,
    minLength:
      dataBar.attributes.minLength !== undefined
        ? getFiniteNumberOr(dataBar.attributes.minLength, 10)
        : undefined,
    maxLength:
      dataBar.attributes.maxLength !== undefined
        ? getFiniteNumberOr(dataBar.attributes.maxLength, 90)
        : undefined,
  };
}

function parseIconSetRule(node: XMLNode, priority: number) {
  const iconSet = findChild(node, 'iconSet');
  if (!iconSet || !iconSet.attributes.iconSet) return undefined;

  return {
    type: 'iconSet' as const,
    priority,
    iconSet: iconSet.attributes.iconSet,
    thresholds: findChildren(iconSet, 'cfvo').map(parseThreshold),
    showValue:
      iconSet.attributes.showValue !== undefined
        ? iconSet.attributes.showValue === '1'
        : undefined,
    reverse: iconSet.attributes.reverse === '1',
  };
}

function parseCellStyleRule(
  node: XMLNode,
  priority: number,
  differentialStyles: CellStyle[],
) {
  const styleId = Number.parseInt(node.attributes.dxfId || '-1', 10);
  const style =
    styleId >= 0 && styleId < differentialStyles.length
      ? differentialStyles[styleId]
      : undefined;

  if (node.attributes.type === 'cellIs' && node.attributes.operator) {
    const formulas = findChildren(node, 'formula').map(getTextContent);
    return {
      type: 'cellIs' as const,
      operator: node.attributes.operator as Extract<
        ConditionalFormattingRule,
        { type: 'cellIs' }
      >['operator'],
      formula1: formulas[0] ? parseFormulaValue(formulas[0]) : '',
      formula2: formulas[1] ? parseFormulaValue(formulas[1]) : undefined,
      style,
      priority,
      stopIfTrue: node.attributes.stopIfTrue === '1',
    };
  }

  if (node.attributes.type === 'expression') {
    const formula = findChild(node, 'formula');
    return {
      type: 'expression' as const,
      formula: formula ? getTextContent(formula) : '',
      style,
      priority,
      stopIfTrue: node.attributes.stopIfTrue === '1',
    };
  }

  return undefined;
}

export function parseConditionalFormattings(
  root: XMLNode,
  differentialStyles: CellStyle[],
): ConditionalFormatting[] {
  const blocks: ConditionalFormatting[] = [];

  for (const formattingNode of findChildren(root, 'conditionalFormatting')) {
    const sqref = formattingNode.attributes.sqref;
    if (!sqref) continue;

    const ranges = parseSqref(sqref);
    if (ranges.length === 0) continue;

    const rules: ConditionalFormattingRule[] = [];
    for (const ruleNode of findChildren(formattingNode, 'cfRule')) {
      const priority = Number.parseInt(ruleNode.attributes.priority || '0', 10);
      const type = ruleNode.attributes.type;
      if (!type) continue;

      if (type === 'cellIs' || type === 'expression') {
        const rule = parseCellStyleRule(ruleNode, priority, differentialStyles);
        if (rule) rules.push(rule);
      } else if (type === 'colorScale') {
        const rule = parseColorScaleRule(ruleNode, priority);
        if (rule) rules.push(rule);
      } else if (type === 'dataBar') {
        const rule = parseDataBarRule(ruleNode, priority);
        if (rule) rules.push(rule);
      } else if (type === 'iconSet') {
        const rule = parseIconSetRule(ruleNode, priority);
        if (rule) rules.push(rule);
      }
    }

    if (rules.length > 0) {
      blocks.push({
        range: ranges.length === 1 ? ranges[0] : ranges,
        rules,
      });
    }
  }

  return blocks;
}
