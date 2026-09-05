import { elementChildren, MAX_XML_SIZE, parseXML } from './native-xml';
import { awaitRead, cancelReadOnAbort } from './read-control';

const TAG_END = /"[^"]*"|'[^']*'|[^"'>]+|>/y;
const XML_DECLARATION = /^<\?xml(?:\s|\?)/i;
const NAME_END = /[\s/>]/;
const XML_BATCH_SIZE = 96 * 1024;
const SPREADSHEET_NAMESPACES = new Set([
  '', // Keep accepting namespace-free workbooks.
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  'http://purl.oclc.org/ooxml/spreadsheetml/main',
]);

interface EnvelopeElement {
  name: string;
  opening: string;
  namespaces: Record<string, string>;
}

function elementScope(name: string, namespaces: Record<string, string>) {
  const colon = name.indexOf(':');
  return {
    local: colon < 0 ? name : name.slice(colon + 1),
    namespace:
      namespaces[colon < 0 ? '' : name.slice(0, colon)] ??
      (colon < 0 ? '' : undefined),
  };
}

function matchesPath(
  name: string,
  localName: string,
  namespaces: Record<string, string>,
  ancestors: EnvelopeElement[],
): boolean {
  const path = localName === 'row' ? ['worksheet', 'sheetData'] : ['sst'];
  if (ancestors.length !== path.length) return false;
  const scope = elementScope(name, namespaces);
  if (
    scope.local !== localName ||
    !SPREADSHEET_NAMESPACES.has(scope.namespace ?? '?')
  )
    return false;
  return ancestors.every((ancestor, index) => {
    const parent = elementScope(ancestor.name, ancestor.namespaces);
    return parent.local === path[index] && parent.namespace === scope.namespace;
  });
}

/** Retain and validate only the markup outside selected data elements. */
class XmlEnvelope {
  readonly ancestors: EnvelopeElement[] = [];
  rootSeen = false;
  prologStart = true;
  private buffer = '';

  private readonly localName: string;
  private readonly scoped: boolean;

  constructor(localName: string, scoped: boolean) {
    this.localName = localName;
    this.scoped = scoped;
  }

  validate(): void {
    if (!this.buffer) return;
    const closing = this.ancestors
      .map(({ name }) => `</${name}>`)
      .reverse()
      .join('');
    parseXML(`<validation>${this.buffer}${closing}</validation>`);
    this.buffer = this.ancestors.map(({ opening }) => opening).join('');
  }

  append(value: string): void {
    this.buffer += value;
    if (this.buffer.length >= XML_BATCH_SIZE) this.validate();
  }

  declaration(token: string): boolean {
    if (!XML_DECLARATION.test(token)) return false;
    if (!this.prologStart) throw new Error('Misplaced XML declaration');
    parseXML(`${token}<validation/>`);
    return true;
  }

  tag(opening: string) {
    const closing = opening[1] === '/';
    const selfClosing = opening[opening.length - 2] === '/';
    const nameStart = closing ? 2 : 1;
    let nameEnd = nameStart;
    while (nameEnd < opening.length && !NAME_END.test(opening[nameEnd]))
      nameEnd++;
    const name = opening.slice(nameStart, nameEnd);
    let namespaces = this.ancestors.at(-1)?.namespaces ?? Object.create(null);
    if (!closing && opening.includes('xmlns')) {
      const header = parseXML(selfClosing ? opening : `${opening}</${name}>`);
      namespaces = { ...namespaces };
      for (const [key, value] of Object.entries(header.attributes)) {
        if (key === 'xmlns') namespaces[''] = value;
        else if (key.startsWith('xmlns:')) namespaces[key.slice(6)] = value;
      }
    }
    const matches = this.scoped
      ? matchesPath(name, this.localName, namespaces, this.ancestors)
      : name === this.localName || name.endsWith(`:${this.localName}`);
    if (closing) {
      if (this.ancestors.pop()?.name !== name)
        throw new Error('Mismatched XML closing tag');
    } else {
      if (!this.ancestors.length) {
        if (this.rootSeen) throw new Error('Multiple XML root elements');
        this.rootSeen = true;
      }
      if (!selfClosing) {
        this.ancestors.push({ name, opening, namespaces });
        if (this.ancestors.length >= 100)
          throw new Error('XML maximum depth exceeded');
      }
    }
    return { name, closing, selfClosing, matches };
  }
}

function tagEnd(buffer: string, start: number): number {
  TAG_END.lastIndex = start;
  let token: RegExpExecArray | null;
  while ((token = TAG_END.exec(buffer))) {
    if (token[0] === '>') return token.index + 1;
  }
  return -1;
}

function scanFragment(
  buffer: string,
  initialCursor: number,
  markers: RegExp,
  initialDepth: number,
  nameLength: number,
) {
  let cursor = initialCursor;
  let depth = initialDepth;
  while (true) {
    markers.lastIndex = cursor;
    const marker = markers.exec(buffer);
    if (!marker)
      return {
        cursor: Math.max(cursor, buffer.length - nameLength - 12),
        depth,
        end: -1,
      };
    const index = marker.index;
    if (marker[0].startsWith('<!') || marker[0] === '<?') {
      let terminator = '?>';
      if (marker[0] === '<!--') terminator = '-->';
      else if (marker[0] === '<![CDATA[') terminator = ']]>';
      const close = buffer.indexOf(terminator, index + marker[0].length);
      if (close < 0) return { cursor: index, depth, end: -1 };
      cursor = close + terminator.length;
      continue;
    }
    const end = tagEnd(buffer, index + marker[0].length);
    if (end < 0) return { cursor: index, depth, end: -1 };
    if (marker[0].startsWith('</')) depth--;
    else if (buffer[end - 2] !== '/') depth++;
    cursor = end;
    if (depth === 0) return { cursor, depth, end };
  }
}

/**
 * Frame complete elements; XML syntax and values are parsed only by Bun.XML.
 * At most one incomplete element plus one input chunk is retained. Comments,
 * CDATA, quoted attributes and processing instructions cannot delimit a row.
 */
export async function* streamXmlElements(
  stream: ReadableStream<Uint8Array>,
  localName: string,
  maxSize = MAX_XML_SIZE,
  scoped = false,
  signal?: AbortSignal,
): AsyncGenerator<string> {
  const reader = stream.getReader();
  const removeAbort = cancelReadOnAbort(reader, signal);
  const decoder = new TextDecoder('utf-8', { fatal: true });
  let buffer = '';
  let cursor = 0;
  let start = -1;
  let depth = 0;
  let fragmentName = '';
  let fragmentMarkers: RegExp | undefined;
  const envelope = new XmlEnvelope(localName, scoped);
  const ancestors = envelope.ancestors;
  let finished = false;
  try {
    while (!finished) {
      signal?.throwIfAborted();
      const chunk = signal
        ? await awaitRead(reader.read(), signal)
        : await reader.read();
      signal?.throwIfAborted();
      finished = chunk.done;
      buffer += finished
        ? decoder.decode()
        : decoder.decode(chunk.value, { stream: true });
      while (true) {
        signal?.throwIfAborted();
        // Skip ordinary child markup: Bun validates it when the complete
        // fragment is parsed. Only matching boundaries and literal sections
        // can change where the fragment ends.
        if (start >= 0 && fragmentMarkers) {
          const result = scanFragment(
            buffer,
            cursor,
            fragmentMarkers,
            depth,
            fragmentName.length,
          );
          cursor = result.cursor;
          depth = result.depth;
          if (result.end < 0) break;
          const end = result.end;
          if (end - start > maxSize)
            throw new Error('XML element exceeds buffer limit');
          ancestors.pop();
          const element = buffer.slice(start, end);
          buffer = buffer.slice(end);
          cursor = 0;
          start = -1;
          yield element;
          continue;
        }
        const open = buffer.indexOf('<', cursor);
        if (open < 0) {
          if (!ancestors.length && buffer.slice(cursor).trim())
            throw new Error('Text outside XML root');
          // Retain incomplete text so entities and ]]> cannot be split across
          // separate native validation calls.
          if (finished) {
            envelope.append(buffer.slice(cursor));
            cursor = buffer.length;
          }
          break;
        }
        if (!ancestors.length && buffer.slice(cursor, open).trim())
          throw new Error('Text outside XML root');
        const text = buffer.slice(cursor, open);
        if (text) {
          envelope.prologStart = false;
          envelope.append(text);
        }
        cursor = open;
        let end = -1;
        let declaration = false;
        let fragment = false;
        if (buffer.startsWith('<!--', open)) {
          const close = buffer.indexOf('-->', open + 4);
          if (close >= 0) end = close + 3;
        } else if (buffer.startsWith('<![CDATA[', open)) {
          if (!ancestors.length) throw new Error('CDATA outside XML root');
          const close = buffer.indexOf(']]>', open + 9);
          if (close >= 0) end = close + 3;
        } else if (buffer.startsWith('<?', open)) {
          const close = buffer.indexOf('?>', open + 2);
          if (close >= 0) {
            end = close + 2;
            const token = buffer.slice(open, end);
            declaration = envelope.declaration(token);
          }
        } else {
          // Wait for enough bytes to distinguish declarations split across chunks.
          if (
            buffer.length - open < 9 &&
            !finished &&
            buffer[open + 1] === '!'
          ) {
            cursor = open;
            break;
          }
          if (buffer.startsWith('<!', open))
            throw new Error('DTD and declarations are not allowed in XLSX XML');
          end = tagEnd(buffer, open + 1);
          if (end >= 0) {
            const { name, closing, selfClosing, matches } = envelope.tag(
              buffer.slice(open, end),
            );
            fragment = matches && !closing;
            if (start < 0 && fragment) {
              start = open;
              depth = 0;
              if (fragmentName !== name) {
                fragmentName = name;
                const escapedName = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                fragmentMarkers = new RegExp(
                  `<!--|<!\\[CDATA\\[|<\\?|</?${escapedName}(?=[\\s/>])`,
                  'g',
                );
              }
            }
            if (start >= 0) {
              if (closing) depth--;
              else if (!selfClosing) depth++;
              if (depth === 0) {
                if (end - start > maxSize)
                  throw new Error('XML element exceeds buffer limit');
                const element = buffer.slice(start, end);
                buffer = buffer.slice(end);
                cursor = 0;
                start = -1;
                yield element;
                continue;
              }
            }
          }
        }
        if (end < 0) {
          cursor = open;
          break;
        }
        envelope.prologStart = false;
        if (!declaration && !fragment) envelope.append(buffer.slice(open, end));
        cursor = end;
      }
      // Discard consumed document markup; keep only an unfinished token/element.
      const retain = start >= 0 ? start : cursor;
      buffer = buffer.slice(retain);
      cursor -= retain;
      if (start >= 0) start = 0;
      if (buffer.length > maxSize)
        throw new Error('XML element exceeds buffer limit');
    }
    if (!envelope.rootSeen || ancestors.length || start >= 0 || buffer.trim())
      throw new Error('Truncated XML element');
    envelope.validate();
  } finally {
    removeAbort();
    try {
      const cancellation = reader.cancel();
      if (signal?.aborted) void cancellation.catch(() => {});
      else await cancellation;
    } finally {
      reader.releaseLock();
    }
  }
}

/** Bound native parser calls and retained trees to small XML batches. */
export async function* streamNativeElements(
  stream: ReadableStream<Uint8Array>,
  name: string,
  signal?: AbortSignal,
): AsyncGenerator<import('./native-xml').XMLNode> {
  const parts: string[] = [];
  let size = 0;
  for await (const element of streamXmlElements(
    stream,
    name,
    MAX_XML_SIZE,
    true,
    signal,
  )) {
    if (parts.length && size + element.length > XML_BATCH_SIZE) {
      yield* elementChildren(parseXML(`<batch>${parts.join('')}</batch>`));
      parts.length = 0;
      size = 0;
    }
    if (element.length > XML_BATCH_SIZE) {
      yield parseXML(element);
    } else {
      parts.push(element);
      size += element.length;
    }
  }
  if (parts.length)
    yield* elementChildren(parseXML(`<batch>${parts.join('')}</batch>`));
}
