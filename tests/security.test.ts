import { describe, expect, test } from 'bun:test';
import { readdirSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { unzipSync, zipSync } from 'fflate';
import { ExcelTemplate, readExcel, readExcelStream, writeExcel } from '../src';

const CMD_START_REGEX = /^=CMD/m;
const XML_ATTR_INJECTION_MARKER = 'attacker=';
const MALICIOUS_ZIP_ENTRY_REGEX = /Malicious zip entry detected/;

function readZipTextEntry(
  zip: Record<string, Uint8Array>,
  path: string,
  decoder: TextDecoder,
): string {
  const entry = zip[path];
  if (!entry) {
    throw new Error(`Missing ZIP entry: ${path}`);
  }
  return decoder.decode(entry);
}

describe('Security - Path Validation', () => {
  test('rejects path with null bytes', async () => {
    await expect(
      writeExcel('test\x00.xlsx', {
        worksheets: [{ name: 'S', rows: [] }],
      }),
    ).rejects.toThrow();
  });

  test('rejects path with null bytes on read', async () => {
    await expect(readExcel('test\x00.xlsx')).rejects.toThrow();
  });

  test('rejects empty path', async () => {
    await expect(
      writeExcel('', {
        worksheets: [{ name: 'S', rows: [] }],
      }),
    ).rejects.toThrow();
  });
});

describe('Security - Input Validation', () => {
  test('handles very long string values', async () => {
    const longStr = 'A'.repeat(50000);
    const path = './tests/.tmp/long-string.xlsx';

    const { mkdirSync } = await import('node:fs');
    mkdirSync('./tests/.tmp', { recursive: true });

    await writeExcel(path, {
      worksheets: [
        {
          name: 'Long',
          rows: [{ cells: [{ value: longStr }] }],
        },
      ],
    });

    const wb = await readExcel(path);
    expect(wb.worksheets[0].rows[0].cells[0].value).toBe(longStr);
  });

  test('handles null and undefined cell values', async () => {
    const path = './tests/.tmp/nulls.xlsx';

    const { mkdirSync } = await import('node:fs');
    mkdirSync('./tests/.tmp', { recursive: true });

    await writeExcel(path, {
      worksheets: [
        {
          name: 'Nulls',
          rows: [
            {
              cells: [
                { value: null },
                { value: undefined },
                { value: '' },
                { value: 0 },
              ],
            },
          ],
        },
      ],
    });

    const file = Bun.file(path);
    expect(file.size).toBeGreaterThan(0);
  });

  test('handles worksheet with no columns config', async () => {
    const path = './tests/.tmp/no-cols.xlsx';

    const { mkdirSync } = await import('node:fs');
    mkdirSync('./tests/.tmp', { recursive: true });

    await writeExcel(path, {
      worksheets: [
        {
          name: 'NoCols',
          rows: [{ cells: [{ value: 'test' }] }],
        },
      ],
    });

    const wb = await readExcel(path);
    expect(wb.worksheets[0].rows[0].cells[0].value).toBe('test');
  });
});

describe('Edge Cases', () => {
  test('handles empty cells in sparse rows', async () => {
    const path = './tests/.tmp/sparse.xlsx';

    const { mkdirSync } = await import('node:fs');
    mkdirSync('./tests/.tmp', { recursive: true });

    await writeExcel(path, {
      worksheets: [
        {
          name: 'Sparse',
          rows: [
            {
              cells: [
                { value: 'A' },
                { value: null },
                { value: null },
                { value: 'D' },
              ],
            },
          ],
        },
      ],
    });

    const wb = await readExcel(path);
    const cells = wb.worksheets[0].rows[0].cells;
    expect(cells[0].value).toBe('A');
    expect(cells[3].value).toBe('D');
  });

  test('handles many sheets', async () => {
    const path = './tests/.tmp/many-sheets.xlsx';

    const { mkdirSync } = await import('node:fs');
    mkdirSync('./tests/.tmp', { recursive: true });

    const worksheets = Array.from({ length: 10 }, (_, i) => ({
      name: `Sheet${i + 1}`,
      rows: [{ cells: [{ value: `Data from sheet ${i + 1}` }] }],
    }));

    await writeExcel(path, { worksheets });

    const wb = await readExcel(path);
    expect(wb.worksheets).toHaveLength(10);
    expect(wb.worksheets[9].name).toBe('Sheet10');
  });

  test('handles mixed value types in a row', async () => {
    const path = './tests/.tmp/mixed-types.xlsx';

    const { mkdirSync } = await import('node:fs');
    mkdirSync('./tests/.tmp', { recursive: true });

    await writeExcel(path, {
      worksheets: [
        {
          name: 'Mixed',
          rows: [
            {
              cells: [
                { value: 'text' },
                { value: 42 },
                { value: 3.14 },
                { value: true },
                { value: false },
                { value: null },
              ],
            },
          ],
        },
      ],
    });

    const wb = await readExcel(path);
    const cells = wb.worksheets[0].rows[0].cells;
    expect(cells[0].value).toBe('text');
    expect(cells[1].value).toBe(42);
    expect(cells[2].value).toBe(3.14);
    expect(cells[3].value).toBe(true);
    expect(cells[4].value).toBe(false);
  });

  test('handles all border styles', async () => {
    const path = './tests/.tmp/borders.xlsx';

    const { mkdirSync } = await import('node:fs');
    mkdirSync('./tests/.tmp', { recursive: true });

    await writeExcel(path, {
      worksheets: [
        {
          name: 'Borders',
          rows: [
            {
              cells: [
                {
                  value: 'Bordered',
                  style: {
                    border: {
                      top: { style: 'thin', color: '000000' },
                      bottom: { style: 'medium', color: 'FF0000' },
                      left: { style: 'thick', color: '00FF00' },
                      right: { style: 'dashed', color: '0000FF' },
                    },
                  },
                },
              ],
            },
          ],
        },
      ],
    });

    const wb = await readExcel(path);
    const cell = wb.worksheets[0].rows[0].cells[0];
    expect(cell.style?.border?.top?.style).toBe('thin');
    expect(cell.style?.border?.bottom?.style).toBe('medium');
  });
});

describe('Security - CSV Formula Injection', () => {
  test('neutralizes values starting with = by prefixing with quote', async () => {
    const path = './tests/.tmp/formula-inject.csv';

    const { mkdirSync } = await import('node:fs');
    mkdirSync('./tests/.tmp', { recursive: true });

    const { writeCSV } = await import('../src');
    await writeCSV(path, [['=CMD("calc")', 'safe']]);

    const content = await Bun.file(path).text();
    expect(content).toContain("'=CMD");
    expect(content).not.toMatch(CMD_START_REGEX);
  });

  test('neutralizes values starting with +, -, @', async () => {
    const path = './tests/.tmp/formula-inject2.csv';

    const { mkdirSync } = await import('node:fs');
    mkdirSync('./tests/.tmp', { recursive: true });

    const { writeCSV } = await import('../src');
    await writeCSV(path, [['+cmd|echo', '-1+1', '@SUM(A1)']]);

    const content = await Bun.file(path).text();
    expect(content).toContain("'+cmd");
    expect(content).toContain("'-1+1");
    expect(content).toContain("'@SUM");
  });

  test('CSV stream writer also neutralizes formula triggers', async () => {
    const path = './tests/.tmp/formula-inject-stream.csv';

    const { mkdirSync } = await import('node:fs');
    mkdirSync('./tests/.tmp', { recursive: true });

    const { createCSVStream } = await import('../src');
    const stream = createCSVStream(path);
    stream.writeRow(['=HYPERLINK("http://evil.com")']);
    await stream.end();

    const content = await Bun.file(path).text();
    expect(content).toContain("'=HYPERLINK");
  });
});

describe('Security - Stream Field Length', () => {
  test('readCSVStream rejects oversized fields', async () => {
    const path = './tests/.tmp/huge-field-stream.csv';

    const { mkdirSync } = await import('node:fs');
    mkdirSync('./tests/.tmp', { recursive: true });

    // Create a CSV with a field > 1MB
    const hugeField = 'X'.repeat(1_100_000);
    await Bun.write(path, `a,${hugeField}\n`);

    const { readCSVStream } = await import('../src');
    const consume = async () => {
      for await (const _row of readCSVStream(path)) {
        // consume
      }
    };

    await expect(consume()).rejects.toThrow('maximum length');
  });
});

describe('Security - Numeric Attribute Sanitization', () => {
  test('writeExcel sanitizes numeric runtime values before writing XML attributes', async () => {
    const path = './tests/.tmp/runtime-attr-sanitize.xlsx';
    const attacker = '1" attacker="1';

    const { mkdirSync } = await import('node:fs');
    mkdirSync('./tests/.tmp', { recursive: true });

    await writeExcel(path, {
      worksheets: [
        {
          name: 'Sanitized',
          freezePane: {
            row: attacker as unknown as number,
            col: attacker as unknown as number,
          },
          defaultRowHeight: attacker as unknown as number,
          defaultColWidth: attacker as unknown as number,
          columns: [{ width: attacker as unknown as number }],
          rows: [
            {
              height: attacker as unknown as number,
              cells: [
                {
                  value: 'safe',
                  style: {
                    font: {
                      size: attacker as unknown as number,
                    },
                  },
                },
              ],
            },
          ],
        },
      ],
    });

    const zip = unzipSync(new Uint8Array(await Bun.file(path).arrayBuffer()));
    const decoder = new TextDecoder('utf-8');
    const sheetXml = readZipTextEntry(zip, 'xl/worksheets/sheet1.xml', decoder);
    const stylesXml = readZipTextEntry(zip, 'xl/styles.xml', decoder);

    expect(sheetXml).not.toContain(XML_ATTR_INJECTION_MARKER);
    expect(stylesXml).not.toContain(XML_ATTR_INJECTION_MARKER);

    const wb = await readExcel(path);
    expect(wb.worksheets[0].rows[0].cells[0].value).toBe('safe');
  });

  test('streaming writers sanitize numeric runtime values before writing XML attributes', async () => {
    const attacker = '1" attacker="1';

    const { mkdirSync } = await import('node:fs');
    mkdirSync('./tests/.tmp', { recursive: true });

    const { createChunkedExcelStream, createExcelStream } = await import(
      '../src'
    );

    const cases = [
      {
        path: './tests/.tmp/runtime-attr-stream.xlsx',
        writer: createExcelStream('./tests/.tmp/runtime-attr-stream.xlsx', {
          freezePane: {
            row: attacker as unknown as number,
            col: attacker as unknown as number,
          },
          defaultRowHeight: attacker as unknown as number,
          columns: [{ width: attacker as unknown as number }],
        }),
      },
      {
        path: './tests/.tmp/runtime-attr-chunked.xlsx',
        writer: createChunkedExcelStream(
          './tests/.tmp/runtime-attr-chunked.xlsx',
          {
            freezePane: {
              row: attacker as unknown as number,
              col: attacker as unknown as number,
            },
            defaultRowHeight: attacker as unknown as number,
            columns: [{ width: attacker as unknown as number }],
          },
        ),
      },
    ];

    for (const entry of cases) {
      entry.writer.writeRow({
        height: attacker as unknown as number,
        cells: [
          {
            value: 'safe',
            style: {
              font: {
                size: attacker as unknown as number,
              },
            },
          },
        ],
      });
      await entry.writer.end();

      const zip = unzipSync(
        new Uint8Array(await Bun.file(entry.path).arrayBuffer()),
      );
      const decoder = new TextDecoder('utf-8');
      const sheetXml = readZipTextEntry(
        zip,
        'xl/worksheets/sheet1.xml',
        decoder,
      );
      const stylesXml = readZipTextEntry(zip, 'xl/styles.xml', decoder);

      expect(sheetXml).not.toContain(XML_ATTR_INJECTION_MARKER);
      expect(stylesXml).not.toContain(XML_ATTR_INJECTION_MARKER);

      const wb = await readExcel(entry.path);
      expect(wb.worksheets[0].rows[0].cells[0].value).toBe('safe');
    }
  });

  test('writeExcel rejects invalid runtime image coordinates', async () => {
    const path = './tests/.tmp/runtime-image-coords.xlsx';

    const { mkdirSync } = await import('node:fs');
    mkdirSync('./tests/.tmp', { recursive: true });

    await expect(
      writeExcel(path, {
        worksheets: [
          {
            name: 'Images',
            rows: [],
            images: [
              {
                format: 'png',
                data: new Uint8Array([1, 2, 3]),
                range: {
                  startRow: '0</xdr:row><attack/>' as unknown as number,
                  startCol: 0,
                  endRow: 0,
                  endCol: 0,
                },
              },
            ],
          },
        ],
      }),
    ).rejects.toThrow('Invalid image startRow');
  });
});

describe('Security - Streaming XLSX Cleanup', () => {
  test('readExcelStream cleans up temp worksheet files when unzip fails', async () => {
    const path = './tests/.tmp/read-stream-cleanup.xlsx';

    const { mkdirSync } = await import('node:fs');
    mkdirSync('./tests/.tmp', { recursive: true });

    const encoder = new TextEncoder();
    const sheetXml = encoder.encode(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"/></sheetData></worksheet>',
    );
    const zipBytes = zipSync({
      'xl/worksheets/sheet1.xml': sheetXml,
      '../evil.xml': encoder.encode('boom'),
    });
    await Bun.write(path, zipBytes);

    const before = readdirSync(tmpdir()).filter((name) =>
      name.startsWith('bun-excel-stream-'),
    ).length;

    const iterator = readExcelStream(path);
    await expect(iterator.next()).rejects.toThrow(MALICIOUS_ZIP_ENTRY_REGEX);

    const after = readdirSync(tmpdir()).filter((name) =>
      name.startsWith('bun-excel-stream-'),
    ).length;

    expect(after).toBe(before);
  });
});

describe('Security - Template Mode Bounds', () => {
  test('template mode rejects defined names outside Excel worksheet bounds', () => {
    const template = new ExcelTemplate({
      worksheets: [{ name: 'Sheet1', rows: [] }],
      definedNames: [
        {
          name: 'HugeRange',
          refersTo: '=Sheet1!A1048577',
        },
      ],
    });

    expect(() => template.setDefinedName('HugeRange', 'value')).toThrow(
      'outside Excel worksheet bounds',
    );
  });

  test('template mode ignores prototype-polluting patch keys', () => {
    const template = new ExcelTemplate({
      worksheets: [{ name: 'Sheet1', rows: [] }],
    });
    const pollutedInput = JSON.parse(
      '{"value":"safe","__proto__":{"polluted":"yes"}}',
    ) as Record<string, unknown>;

    template.setCell('Sheet1', 'A1', pollutedInput as never);

    expect(template.workbook.worksheets[0].rows[0].cells[0].value).toBe('safe');
    expect(({} as Record<string, unknown>).polluted).toBeUndefined();
  });
});

describe('Security - XML Entity Decoding', () => {
  test('readExcel does not double-unescape ampersand-escaped entities', async () => {
    const path = './tests/.tmp/double-unescape.xlsx';

    const { mkdirSync } = await import('node:fs');
    mkdirSync('./tests/.tmp', { recursive: true });

    await writeExcel(path, {
      worksheets: [
        {
          name: 'Sheet1',
          rows: [{ cells: [{ value: '&quot;' }] }],
        },
      ],
    });

    const workbook = await readExcel(path);
    expect(workbook.worksheets[0].rows[0].cells[0].value).toBe('&quot;');
  });
});
