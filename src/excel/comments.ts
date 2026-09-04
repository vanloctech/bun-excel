import type { CellComment } from '../types';
import {
  findChild,
  findChildren,
  getTextContent,
  parseXML,
} from './native-xml';
import { buildCellRef, escapeXML, parseCellRef } from './xml-builder';

export interface CommentEntry {
  ref: string;
  comment: CellComment;
}

export interface ParsedCommentEntry {
  row: number;
  col: number;
  comment: CellComment;
}

export function buildCommentsXML(entries: CommentEntry[]): string {
  const authors: string[] = [];
  const authorIndex = new Map<string, number>();

  for (const entry of entries) {
    const author = entry.comment.author || 'Author';
    if (!authorIndex.has(author)) {
      authorIndex.set(author, authors.length);
      authors.push(author);
    }
  }

  let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  xml +=
    '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
  xml += '<authors>';
  for (const author of authors) {
    xml += `<author>${escapeXML(author)}</author>`;
  }
  xml += '</authors><commentList>';

  for (const entry of entries) {
    const author = entry.comment.author || 'Author';
    const authorId = authorIndex.get(author) ?? 0;
    xml += `<comment ref="${escapeXML(entry.ref)}" authorId="${authorId}">`;
    xml += `<text><t xml:space="preserve">${escapeXML(
      entry.comment.text,
    )}</t></text>`;
    xml += '</comment>';
  }

  xml += '</commentList></comments>';
  return xml;
}

export function buildCommentsVML(entries: CommentEntry[]): string {
  let xml = '<?xml version="1.0" encoding="UTF-8"?>\n';
  xml +=
    '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">';
  xml +=
    '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>';
  xml +=
    '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">';
  xml += '<v:stroke joinstyle="miter"/>';
  xml += '<v:path gradientshapeok="t" o:connecttype="rect"/>';
  xml += '</v:shapetype>';

  for (let i = 0; i < entries.length; i++) {
    const { row, col } = parseCellRef(entries[i].ref);
    xml += `<v:shape id="_x0000_s${1025 + i}" type="#_x0000_t202" style="position:absolute;margin-left:80pt;margin-top:5pt;width:104pt;height:64pt;z-index:${i + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">`;
    xml += '<v:fill color2="#ffffe1"/>';
    xml += '<v:shadow on="t" color="black" obscured="t"/>';
    xml += '<v:path o:connecttype="none"/>';
    xml +=
      '<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox>';
    xml += '<x:ClientData ObjectType="Note">';
    xml += '<x:MoveWithCells/><x:SizeWithCells/><x:AutoFill>False</x:AutoFill>';
    xml += `<x:Row>${row}</x:Row><x:Column>${col}</x:Column>`;
    xml += '</x:ClientData></v:shape>';
  }

  xml += '</xml>';
  return xml;
}

export function parseCommentsXML(xml: string): ParsedCommentEntry[] {
  const root = parseXML(xml);
  if (!root) return [];

  const authorsNode = findChild(root, 'authors');
  const authors = authorsNode
    ? findChildren(authorsNode, 'author').map((node) => getTextContent(node))
    : [];

  const commentList = findChild(root, 'commentList');
  if (!commentList) return [];

  const entries: ParsedCommentEntry[] = [];
  for (const commentNode of findChildren(commentList, 'comment')) {
    const ref = commentNode.attributes.ref;
    if (!ref) continue;
    const { row, col } = parseCellRef(ref);
    const authorId = Number.parseInt(
      commentNode.attributes.authorId || '0',
      10,
    );
    const textNode = findChild(commentNode, 'text');
    if (!textNode) continue;

    let text = '';
    const directText = findChild(textNode, 't');
    if (directText) {
      text = getTextContent(directText);
    } else {
      for (const rNode of findChildren(textNode, 'r')) {
        const tNode = findChild(rNode, 't');
        if (tNode) text += getTextContent(tNode);
      }
    }

    entries.push({
      row,
      col,
      comment: {
        text,
        author: authors[authorId],
      },
    });
  }

  return entries;
}

export function commentRefFromCoords(row: number, col: number): string {
  return buildCellRef(row, col);
}
