/** Bun's native, non-compact XML tree. No intermediate application tree. */
export interface XMLNode {
  name: string;
  attributes: Record<string, string>;
  children: Array<
    XMLNode | string | { comment: string } | { target: string; data: string }
  >;
}

// Structural declaration for consumers using older @types/bun releases.
const xml = (
  Bun as typeof Bun & {
    XML: { parse(input: string, options: { compact: false }): XMLNode };
  }
).XML;
export const MAX_XML_SIZE = 50 * 1024 * 1024;

export function isElement(node: XMLNode['children'][number]): node is XMLNode {
  return typeof node === 'object' && 'name' in node;
}

export function* elementChildren(node: XMLNode): Generator<XMLNode> {
  for (const child of node.children) if (isElement(child)) yield child;
}

export function parseXML(input: string): XMLNode {
  if (input.length > MAX_XML_SIZE) throw new Error('XML input too large');
  if (!xml) throw new Error('bun-excel requires Bun 1.4.0 or newer');
  // OOXML does not need DTDs. Skip literal declarations inside comments,
  // CDATA and processing instructions, and reject real declarations up front.
  if (input.includes('<!')) {
    let offset = 0;
    while ((offset = input.indexOf('<', offset)) !== -1) {
      let terminator: string | undefined;
      if (input.startsWith('<!--', offset)) terminator = '-->';
      else if (input.startsWith('<![CDATA[', offset)) terminator = ']]>';
      else if (input.startsWith('<?', offset)) terminator = '?>';
      if (terminator) {
        const end = input.indexOf(terminator, offset + 2);
        if (end < 0) break;
        offset = end + terminator.length;
      } else {
        if (input.startsWith('<!', offset))
          throw new Error('DTD and declarations are not allowed in XLSX XML');
        offset++;
      }
    }
  }
  const root = xml.parse(input, { compact: false });
  let count = 0;
  function validate(node: XMLNode, depth: number): void {
    if (depth >= 100)
      throw new Error('XML parsing aborted: exceeded maximum depth (100)');
    if (++count > 500_000)
      throw new Error(
        'XML parsing aborted: exceeded maximum node count (500000)',
      );
    for (const key of ['__proto__', 'constructor', 'prototype']) {
      if (Object.hasOwn(node.attributes, key)) delete node.attributes[key];
    }
    for (const child of node.children)
      if (isElement(child)) validate(child, depth + 1);
  }
  validate(root, 1);
  return root;
}

export function findChild(node: XMLNode, tag: string): XMLNode | undefined {
  for (const child of node.children) {
    if (
      isElement(child) &&
      (child.name === tag || child.name.endsWith(`:${tag}`))
    )
      return child;
  }
  return undefined;
}

export function findChildren(node: XMLNode, tag: string): XMLNode[] {
  return node.children.filter(
    (child): child is XMLNode =>
      isElement(child) &&
      (child.name === tag || child.name.endsWith(`:${tag}`)),
  );
}

export function getTextContent(node: XMLNode | undefined): string {
  if (!node) return '';
  let text = '';
  for (const child of node.children) {
    if (typeof child === 'string') text += child;
    else if (isElement(child)) text += getTextContent(child);
  }
  return text;
}
