import { expect, test } from 'bun:test';
import { getTextContent, parseXML } from '../src/excel/native-xml';
import {
  streamNativeElements,
  streamXmlElements,
} from '../src/excel/xml-elements';

function input(xml: string, size: number, onCancel?: () => void) {
  const bytes = new TextEncoder().encode(xml);
  let offset = 0;
  return new ReadableStream<Uint8Array>({
    pull(controller) {
      if (offset >= bytes.length) {
        controller.close();
        return;
      }
      controller.enqueue(bytes.subarray(offset, offset + size));
      offset += size;
    },
    cancel: onCancel,
  });
}

for (const size of [1, 2, 7, 64, 65536]) {
  test(`frames native XML elements across ${size}-byte chunks`, async () => {
    const xml = `<?xml version="1.0"?><worksheet><!-- <row/> --><sheetData><row r="1" custom="a > b"><c><is><t><![CDATA[Tiếng Việt 😀 </row> <!DOCTYPE fake>]]></t></is></c></row><?skip <row/>?><x:row xmlns:x="urn:test" r='2'/></sheetData></worksheet>`;
    const values = [];
    for await (const fragment of streamXmlElements(input(xml, size), 'row'))
      values.push(parseXML(fragment));
    expect(values).toHaveLength(2);
    expect(getTextContent(values[0])).toBe(
      'Tiếng Việt 😀 </row> <!DOCTYPE fake>',
    );
    expect(values[0].attributes.custom).toBe('a > b');
    expect(values[1].name).toBe('x:row');
  });
}

test('rejects oversized complete and incomplete elements and truncated tokens', async () => {
  for (const xml of [
    `<row>${'x'.repeat(100)}</row>`,
    `<row>${'x'.repeat(100)}`,
    '<row><c>',
  ]) {
    const read = async () => {
      for await (const _ of streamXmlElements(input(xml, 8), 'row', 32)) {
        /* consume */
      }
    };
    await expect(read()).rejects.toThrow();
  }
});

test('rejects DTD declarations even when split across chunks', async () => {
  const read = async () => {
    for await (const _ of streamXmlElements(
      input('<!DOCTYPE worksheet><worksheet><row/></worksheet>', 1),
      'row',
    )) {
      /* consume */
    }
  };
  await expect(read()).rejects.toThrow('DTD');
});

test('consumer cancellation stops reading and cancels the source', async () => {
  let cancelled = false;
  let pulls = 0;
  const stream = new ReadableStream<Uint8Array>({
    pull(controller) {
      pulls++;
      controller.enqueue(new TextEncoder().encode('<row/>'));
    },
    cancel() {
      cancelled = true;
    },
  });
  for await (const fragment of streamXmlElements(stream, 'row')) {
    expect(fragment).toBe('<row/>');
    break;
  }
  expect(cancelled).toBe(true);
  expect(pulls).toBeLessThanOrEqual(2);
});

test('rejects malformed document envelopes and incomplete attributes', async () => {
  for (const xml of [
    '<worksheet><row/></wrong>',
    '<worksheet><row/>',
    '<worksheet/><worksheet/>',
    '<worksheet><row a="unfinished >',
  ]) {
    const read = async () => {
      for await (const _ of streamXmlElements(input(xml, 3), 'row')) {
        /* consume */
      }
    };
    await expect(read()).rejects.toThrow();
  }
});

test('native batches keep read-ahead bounded for a slow consumer', async () => {
  let pulls = 0;
  let cancelled = false;
  const chunk = new TextEncoder().encode(
    `<row><c><t>${'x'.repeat(100)}</t></c></row>`.repeat(32),
  );
  const stream = new ReadableStream<Uint8Array>({
    pull(controller) {
      pulls++;
      if (pulls === 1)
        controller.enqueue(new TextEncoder().encode('<worksheet><sheetData>'));
      else controller.enqueue(chunk);
    },
    cancel() {
      cancelled = true;
    },
  });
  const iterator = streamNativeElements(stream, 'row');
  expect((await iterator.next()).value?.name).toBe('row');
  expect(pulls).toBeLessThan(40);
  await iterator.return(undefined);
  expect(cancelled).toBe(true);
});

for (const size of [1, 7, 65536]) {
  test(`native batches reject invalid envelope syntax across ${size}-byte chunks`, async () => {
    for (const xml of [
      '<worksheet broken=unquoted><sheetData><row/></sheetData></worksheet>',
      '<worksheet><sheetData><row/></sheetData><!-- invalid -- comment --></worksheet>',
      '<worksheet><sheetData><row/></sheetData><extra a="1" a="2"/></worksheet>',
      '<worksheet><sheetData><row/></sheetData>&unknown;</worksheet>',
      '<worksheet><sheetData><row/></sheetData>]]></worksheet>',
      '<worksheet><sheetData><row/></sheetData></worksheet extra="1">',
      '<worksheet><?xml version="1.0"?><sheetData><row/></sheetData></worksheet>',
      '<![CDATA[text]]><worksheet/>',
    ]) {
      const read = async () => {
        for await (const _ of streamNativeElements(input(xml, size), 'row')) {
          /* consume */
        }
      };
      await expect(read()).rejects.toThrow();
    }
  });

  test(`selects only worksheet rows and direct shared strings across ${size}-byte chunks`, async () => {
    const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    const worksheet = `<s:worksheet xmlns:s="${ns}" xmlns:x="urn:extension"><s:sheetData><s:row r="1"/><x:row r="2"/><s:row xmlns:s="urn:extension" r="3"/><s:row r="4"/></s:sheetData><extLst><ext><s:row r="5"/></ext></extLst></s:worksheet>`;
    const rows = [];
    for await (const row of streamNativeElements(input(worksheet, size), 'row'))
      rows.push(row.attributes.r);
    expect(rows).toEqual(['1', '4']);
    const shared = `<sst xmlns="${ns}"><extLst><ext><si><t>extension</t></si></ext></extLst><si><t>real</t></si><si xmlns="urn:extension"><t>foreign</t></si></sst>`;
    const strings = [];
    for await (const si of streamNativeElements(input(shared, size), 'si'))
      strings.push(getTextContent(si));
    expect(strings).toEqual(['real']);
  });
}

test('validates large envelopes in bounded batches without losing ancestor context', async () => {
  const metadata =
    '<extra a="1">&amp;<![CDATA[<literal>]]><!-- valid --></extra>'.repeat(
      5000,
    );
  const xml = `<?xml version="1.0"?><worksheet>${metadata}<sheetData><row r="1"/></sheetData>${metadata}</worksheet><!-- tail -->`;
  const rows = [];
  for await (const row of streamNativeElements(input(xml, 65536), 'row'))
    rows.push(row.attributes.r);
  expect(rows).toEqual(['1']);
  const invalid = xml.replace('<!-- tail -->', '<!-- invalid -- tail -->');
  const read = async () => {
    for await (const _ of streamNativeElements(input(invalid, 65536), 'row')) {
      /* consume */
    }
  };
  await expect(read()).rejects.toThrow();
});
