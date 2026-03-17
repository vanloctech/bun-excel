import type { BinaryData, WorksheetImage } from '../types';
import { escapeXML, getFiniteNumber } from './xml-builder';
import { findChild, parseXML } from './xml-parser';

export interface ImagePart {
  path: string;
  data: Uint8Array;
}

export interface DrawingArtifacts {
  drawingXml: string;
  drawingRelsXml: string;
  media: ImagePart[];
}

const MAX_IMAGE_ROW = 1_048_575;
const MAX_IMAGE_COL = 16_383;

function toUint8Array(data: BinaryData): Uint8Array {
  return data instanceof Uint8Array ? data : new Uint8Array(data);
}

function normalizeImageFormat(
  format: WorksheetImage['format'],
): 'png' | 'jpeg' | 'gif' {
  return format === 'jpg' ? 'jpeg' : format;
}

export function getImageContentType(format: WorksheetImage['format']): string {
  const normalized = normalizeImageFormat(format);
  return normalized === 'jpeg' ? 'image/jpeg' : `image/${normalized}`;
}

function buildMarkerXML(
  tagName: 'xdr:from' | 'xdr:to',
  row: number,
  col: number,
): string {
  return `<${tagName}><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></${tagName}>`;
}

function normalizeImageCoordinate(
  value: unknown,
  max: number,
  label: string,
): number {
  const coordinate = getFiniteNumber(value);
  if (
    coordinate === undefined ||
    !Number.isInteger(coordinate) ||
    coordinate < 0 ||
    coordinate > max
  ) {
    throw new Error(`Invalid image ${label}: ${String(value)}`);
  }
  return coordinate;
}

export function buildDrawingArtifacts(
  sheetIndex: number,
  images: WorksheetImage[],
): DrawingArtifacts {
  let drawingXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  drawingXml +=
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';

  let drawingRelsXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  drawingRelsXml +=
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';

  const media: ImagePart[] = [];
  for (let i = 0; i < images.length; i++) {
    const image = images[i];
    const startRow = normalizeImageCoordinate(
      image.range.startRow,
      MAX_IMAGE_ROW,
      'startRow',
    );
    const startCol = normalizeImageCoordinate(
      image.range.startCol,
      MAX_IMAGE_COL,
      'startCol',
    );
    const endRow = normalizeImageCoordinate(
      image.range.endRow,
      MAX_IMAGE_ROW,
      'endRow',
    );
    const endCol = normalizeImageCoordinate(
      image.range.endCol,
      MAX_IMAGE_COL,
      'endCol',
    );
    if (endRow < startRow || endCol < startCol) {
      throw new Error('Invalid image range: end must be >= start');
    }

    const imageIndex = i + 1;
    const relId = `rId${imageIndex}`;
    const mediaFormat = normalizeImageFormat(image.format);
    const mediaPath = `xl/media/image${sheetIndex}-${imageIndex}.${mediaFormat}`;
    media.push({
      path: mediaPath,
      data: toUint8Array(image.data),
    });

    drawingXml += '<xdr:twoCellAnchor>';
    drawingXml += buildMarkerXML('xdr:from', startRow, startCol);
    drawingXml += buildMarkerXML('xdr:to', endRow + 1, endCol + 1);
    drawingXml += '<xdr:pic>';
    drawingXml += '<xdr:nvPicPr>';
    drawingXml += `<xdr:cNvPr id="${imageIndex}" name="${escapeXML(
      image.name || `Image ${imageIndex}`,
    )}"`;
    if (image.description) {
      drawingXml += ` descr="${escapeXML(image.description)}"`;
    }
    drawingXml += '/><xdr:cNvPicPr/>';
    drawingXml += '</xdr:nvPicPr>';
    drawingXml += '<xdr:blipFill>';
    drawingXml += `<a:blip r:embed="${relId}"/>`;
    drawingXml += '<a:stretch><a:fillRect/></a:stretch>';
    drawingXml += '</xdr:blipFill>';
    drawingXml +=
      '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>';
    drawingXml += '</xdr:pic><xdr:clientData/></xdr:twoCellAnchor>';

    drawingRelsXml += `<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image${sheetIndex}-${imageIndex}.${mediaFormat}"/>`;
  }

  drawingXml += '</xdr:wsDr>';
  drawingRelsXml += '</Relationships>';

  return { drawingXml, drawingRelsXml, media };
}

export function parseDrawingImages(
  drawingXml: string,
  drawingRelsXml: string,
  zip: Record<string, Uint8Array>,
): WorksheetImage[] {
  const drawingDoc = parseXML(drawingXml);
  const drawingRoot = drawingDoc.children[0];
  if (!drawingRoot) return [];

  const relsDoc = parseXML(drawingRelsXml);
  const relRoot = relsDoc.children[0];
  const relMap = new Map<string, string>();
  for (const rel of relRoot?.children || []) {
    relMap.set(rel.attributes.Id, rel.attributes.Target);
  }

  const images: WorksheetImage[] = [];
  for (const anchor of drawingRoot.children.filter(
    (node) => node.tag === 'xdr:twoCellAnchor' || node.tag === 'twoCellAnchor',
  )) {
    const from = findChild(anchor, 'xdr:from') || findChild(anchor, 'from');
    const to = findChild(anchor, 'xdr:to') || findChild(anchor, 'to');
    const pic = findChild(anchor, 'xdr:pic') || findChild(anchor, 'pic');
    const blipFill = pic
      ? findChild(pic, 'xdr:blipFill') || findChild(pic, 'blipFill')
      : undefined;
    const blip = blipFill
      ? findChild(blipFill, 'a:blip') || findChild(blipFill, 'blip')
      : undefined;
    const nvPicPr = pic
      ? findChild(pic, 'xdr:nvPicPr') || findChild(pic, 'nvPicPr')
      : undefined;
    const cNvPr = nvPicPr
      ? findChild(nvPicPr, 'xdr:cNvPr') || findChild(nvPicPr, 'cNvPr')
      : undefined;
    if (!from || !to || !blip) continue;

    const relId = blip.attributes['r:embed'];
    const target = relMap.get(relId);
    if (!target) continue;
    const imagePath = target.startsWith('../')
      ? `xl/${target.slice(3)}`
      : `xl/drawings/${target}`;
    const imageData = zip[imagePath];
    if (!imageData) continue;

    const fromRow = Number.parseInt(
      (findChild(from, 'xdr:row') || findChild(from, 'row'))?.text || '0',
      10,
    );
    const fromCol = Number.parseInt(
      (findChild(from, 'xdr:col') || findChild(from, 'col'))?.text || '0',
      10,
    );
    const toRowRaw = Number.parseInt(
      (findChild(to, 'xdr:row') || findChild(to, 'row'))?.text || '0',
      10,
    );
    const toColRaw = Number.parseInt(
      (findChild(to, 'xdr:col') || findChild(to, 'col'))?.text || '0',
      10,
    );
    const format = (imagePath.split('.').pop() ||
      'png') as WorksheetImage['format'];

    images.push({
      data: imageData,
      format,
      range: {
        startRow: fromRow,
        startCol: fromCol,
        endRow: Math.max(fromRow, toRowRaw - 1),
        endCol: Math.max(fromCol, toColRaw - 1),
      },
      name: cNvPr?.attributes.name,
      description: cNvPr?.attributes.descr,
    });
  }

  return images;
}
