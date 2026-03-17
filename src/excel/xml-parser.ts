// ============================================
// XML Parser — Lightweight SAX-style XML parser
// Security-hardened against XML bomb, injection, prototype pollution
// ============================================

// Top-level regex for performance (biome: useTopLevelRegex)
const TAG_SEPARATOR_REGEX = /[\s/]/;
const ATTR_REGEX = /(\w[\w:.]*)\s*=\s*"([^"]*)"/g;

/** Maximum nesting depth to prevent XML bomb / billion laughs attacks */
const MAX_DEPTH = 100;
/** Maximum total nodes to prevent excessive memory consumption */
const MAX_NODES = 500_000;
/** Maximum input size (50MB) */
const MAX_XML_SIZE = 50 * 1024 * 1024;

export interface XMLNode {
  tag: string;
  attributes: Record<string, string>;
  children: XMLNode[];
  text: string;
}

/**
 * Parse XML string into a tree of XMLNodes
 * Lightweight parser — only handles what XLSX needs
 */
export function parseXML(xml: string): XMLNode {
  if (xml.length > MAX_XML_SIZE) {
    throw new Error(
      `XML input too large: ${xml.length} bytes (max: ${MAX_XML_SIZE})`,
    );
  }

  const root: XMLNode = {
    tag: 'root',
    attributes: Object.create(null),
    children: [],
    text: '',
  };
  const stack: XMLNode[] = [root];
  let i = 0;
  let nodeCount = 0;

  while (i < xml.length) {
    if (xml[i] === '<') {
      // Check for processing instruction <?...?>
      if (xml[i + 1] === '?') {
        const end = xml.indexOf('?>', i);
        if (end === -1) break; // malformed PI
        i = end + 2;
        continue;
      }

      // Check for comment <!--...-->
      if (xml.substring(i, i + 4) === '<!--') {
        const end = xml.indexOf('-->', i);
        if (end === -1) break; // malformed comment
        i = end + 3;
        continue;
      }

      // Closing tag
      if (xml[i + 1] === '/') {
        const end = xml.indexOf('>', i);
        if (end === -1) break; // malformed closing tag
        if (stack.length > 1) stack.pop(); // guard against underflow
        i = end + 1;
        continue;
      }

      // Opening tag
      const end = xml.indexOf('>', i);
      if (end === -1) break; // malformed opening tag

      const tagContent = xml.substring(i + 1, end);
      const selfClosing = tagContent.endsWith('/');
      const cleanContent = selfClosing
        ? tagContent.slice(0, -1).trim()
        : tagContent.trim();

      // Parse tag name and attributes
      const spaceIdx = cleanContent.search(TAG_SEPARATOR_REGEX);
      const tag =
        spaceIdx === -1 ? cleanContent : cleanContent.substring(0, spaceIdx);
      const attrStr = spaceIdx === -1 ? '' : cleanContent.substring(spaceIdx);

      // Use Object.create(null) to prevent prototype pollution
      const attributes: Record<string, string> = Object.create(null);
      ATTR_REGEX.lastIndex = 0;
      let match: RegExpExecArray | null;
      while ((match = ATTR_REGEX.exec(attrStr)) !== null) {
        const attrName = match[1];
        // Prevent prototype pollution — block dangerous keys
        if (
          attrName === '__proto__' ||
          attrName === 'constructor' ||
          attrName === 'prototype'
        ) {
          continue;
        }
        attributes[attrName] = decodeXMLEntities(match[2]);
      }

      // Check node count limit
      nodeCount++;
      if (nodeCount > MAX_NODES) {
        throw new Error(
          `XML parsing aborted: exceeded maximum node count (${MAX_NODES})`,
        );
      }

      const node: XMLNode = { tag, attributes, children: [], text: '' };
      const parent = stack[stack.length - 1];
      parent.children.push(node);

      if (!selfClosing) {
        // Check depth limit
        if (stack.length >= MAX_DEPTH) {
          throw new Error(
            `XML parsing aborted: exceeded maximum depth (${MAX_DEPTH})`,
          );
        }
        stack.push(node);
      }

      i = end + 1;
    } else {
      // Text content
      const nextTag = xml.indexOf('<', i);
      const text =
        nextTag === -1 ? xml.substring(i) : xml.substring(i, nextTag);
      if (text.trim().length > 0 && stack.length > 0) {
        const current = stack[stack.length - 1];
        current.text += decodeXMLEntities(text);
      }
      i = nextTag === -1 ? xml.length : nextTag;
    }
  }

  return root;
}

/**
 * Find a child node by tag name (supports namespace-prefixed tags)
 */
export function findChild(node: XMLNode, tag: string): XMLNode | undefined {
  return node.children.find((c) => c.tag === tag || c.tag.endsWith(`:${tag}`));
}

/**
 * Find all children by tag name
 */
export function findChildren(node: XMLNode, tag: string): XMLNode[] {
  return node.children.filter(
    (c) => c.tag === tag || c.tag.endsWith(`:${tag}`),
  );
}

/**
 * Get text content of a node (recursive)
 */
export function getTextContent(node: XMLNode): string {
  let text = node.text;
  for (const child of node.children) {
    text += getTextContent(child);
  }
  return text;
}

/**
 * Decode XML entities
 */
function decodeXMLEntities(str: string): string {
  return str
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, num) => {
      const code = Number.parseInt(num, 10);
      // Validate code point range to prevent invalid characters
      if (code < 0 || code > 0x10ffff || (code >= 0xd800 && code <= 0xdfff)) {
        return ''; // invalid code point
      }
      return String.fromCodePoint(code);
    })
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => {
      const code = Number.parseInt(hex, 16);
      if (code < 0 || code > 0x10ffff || (code >= 0xd800 && code <= 0xdfff)) {
        return ''; // invalid code point
      }
      return String.fromCodePoint(code);
    })
    .replace(/&amp;/g, '&');
}
