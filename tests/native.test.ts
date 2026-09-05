import { describe, expect, test } from 'bun:test';
import { unzipSync } from 'fflate';
import { escapeXML } from '../src/excel/xml-builder';
import { zipBuffer } from '../src/excel/zip-buffer';

describe('Native XML escaping', () => {
  test('preserves entity spelling, Unicode and literal entity text', () => {
    expect(escapeXML(`<>&"' Tiếng Việt 日本語 😀 &#x27; &apos;`)).toBe(
      '&lt;&gt;&amp;&quot;&apos; Tiếng Việt 日本語 😀 &amp;#x27; &amp;apos;',
    );
    expect(escapeXML('')).toBe('');
  });
});

describe('Native buffered ZIP', () => {
  for (const compress of [true, false]) {
    test(`round-trips empty, binary and sliced entries (compress=${compress})`, () => {
      const files = {
        'empty.xml': new Uint8Array(0),
        'xl/文字.xml': new TextEncoder().encode('<t>Tiếng Việt 😀</t>'),
        'media.bin': new Uint8Array([255, 0, 1, 128, 255]).subarray(1, 4),
        'shared.bin': new Uint8Array(new SharedArrayBuffer(3)),
      };
      expect<Record<string, Uint8Array>>(
        unzipSync(zipBuffer(files, compress)),
      ).toEqual(files);
      expect(unzipSync(zipBuffer({}, compress))).toEqual({});
    });

    test(`writes standard CRC32 and compression metadata (compress=${compress})`, () => {
      const data = new TextEncoder().encode('123456789');
      const bytes = zipBuffer({ 'test.txt': data }, compress);
      const view = new DataView(
        bytes.buffer,
        bytes.byteOffset,
        bytes.byteLength,
      );
      // The end-of-central-directory record points to the central header.
      const centralOffset = view.getUint32(bytes.length - 6, true);
      expect(view.getUint32(centralOffset, true)).toBe(0x02014b50);
      expect(view.getUint16(centralOffset + 10, true)).toBe(compress ? 8 : 0);
      expect(view.getUint32(centralOffset + 16, true)).toBe(0xcbf43926);
      expect(view.getUint32(centralOffset + 24, true)).toBe(9);
    });
  }
});
