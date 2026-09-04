import { describe, expect, test } from 'bun:test';
import {
  findChild,
  findChildren,
  getTextContent,
  MAX_XML_SIZE,
  parseXML,
} from '../src/excel/native-xml';

describe('Native worksheet row parser', () => {
  test('returns the native tree directly with document-order text', () => {
    const root = parseXML(
      '<row r="3"><c r="A3"><is><t>Tiếng Việt &amp; &#x27;</t></is></c></row>',
    );
    expect(root.name).toBe('row');
    expect(root.attributes.r).toBe('3');
    expect(getTextContent(root)).toBe("Tiếng Việt & '");
    expect('tag' in root).toBe(false);
    expect(getTextContent(parseXML('<t>before<r>middle</r>after</t>'))).toBe(
      'beforemiddleafter',
    );
  });

  test('supports single quotes, CDATA, comments and whitespace-only text', () => {
    const row = parseXML(
      "<row r='1'><!-- note --><?test data?><c r='A1'><is><t><![CDATA[<>& Tiếng Việt]]></t></is></c><c r='B1'><is><t xml:space='preserve'>  </t></is></c></row>",
    );
    expect(row?.attributes.r).toBe('1');
    expect(findChildren(row, 'c')).toHaveLength(2);
    expect(getTextContent(findChild(row, 'c'))).toBe('<>& Tiếng Việt');
    expect(getTextContent(findChildren(row, 'c')[1])).toBe('  ');
  });

  test('rejects malformed XML instead of silently using the old parser', () => {
    expect(() => parseXML('<row><c></row>')).toThrow();
    expect(() => parseXML('<row><c>&undeclared;</c></row>')).toThrow();
  });

  test('blocks declarations and strips dangerous attribute keys', () => {
    expect(() =>
      parseXML('<!DOCTYPE row [<!ENTITY a "data">]><row>&a;</row>'),
    ).toThrow('DTD');
    const node = parseXML(
      '<row __proto__="x" constructor="y" prototype="z" r="1"/>',
    );
    expect(Object.hasOwn(node.attributes, '__proto__')).toBe(false);
    expect(Object.keys(node?.attributes ?? {})).toEqual(['r']);
  });

  test('enforces size, depth and node-count limits', () => {
    expect(() => parseXML(' '.repeat(MAX_XML_SIZE + 1))).toThrow('too large');
    expect(() =>
      parseXML(`<row>${'<r>'.repeat(100)}${'</r>'.repeat(100)}</row>`),
    ).toThrow();
    expect(() => parseXML(`<row>${'<c/>'.repeat(500_000)}</row>`)).toThrow(
      'node count',
    );
  });
});
