import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { mkdirSync, rmSync } from 'node:fs';
import { readCSV, writeCSV } from '../src';

const TMP = './tests/.tmp';

beforeAll(() => {
  mkdirSync(TMP, { recursive: true });
});

afterAll(() => {
  rmSync(TMP, { recursive: true, force: true });
});

describe('CSV Writer', () => {
  test('writes basic CSV file', async () => {
    const path = `${TMP}/basic.csv`;
    await writeCSV(path, [
      ['Name', 'Age'],
      ['Alice', 28],
      ['Bob', 32],
    ]);

    const content = await Bun.file(path).text();
    expect(content).toContain('Name');
    expect(content).toContain('Alice');
    expect(content).toContain('28');
  });

  test('writes CSV with custom delimiter', async () => {
    const path = `${TMP}/semicolon.csv`;
    await writeCSV(
      path,
      [
        ['A', 'B'],
        ['1', '2'],
      ],
      { delimiter: ';' },
    );

    const content = await Bun.file(path).text();
    expect(content).toContain('A;B');
  });

  test('writes CSV with BOM', async () => {
    const path = `${TMP}/bom.csv`;
    await writeCSV(path, [['Test']], { bom: true });

    const buf = await Bun.file(path).arrayBuffer();
    const bytes = new Uint8Array(buf);
    expect(bytes[0]).toBe(0xef);
    expect(bytes[1]).toBe(0xbb);
    expect(bytes[2]).toBe(0xbf);
  });

  test('writes CSV with headers option', async () => {
    const path = `${TMP}/headers.csv`;
    await writeCSV(path, [['val1', 'val2']], {
      includeHeader: true,
      headers: ['Col1', 'Col2'],
    });

    const content = await Bun.file(path).text();
    const lines = content.trim().split('\n');
    expect(lines[0]).toContain('Col1');
    expect(lines[1]).toContain('val1');
  });

  test('escapes quotes in CSV values', async () => {
    const path = `${TMP}/quotes.csv`;
    await writeCSV(path, [['He said "hello"', 'normal']]);

    const content = await Bun.file(path).text();
    expect(content).toContain('"He said ""hello"""');
  });
});

describe('CSV Reader', () => {
  test('reads basic CSV file', async () => {
    const path = `${TMP}/read-basic.csv`;
    await Bun.write(path, 'Name,Age\nAlice,28\nBob,32');

    const wb = await readCSV(path);
    expect(wb.worksheets).toHaveLength(1);

    const sheet = wb.worksheets[0];
    expect(sheet.rows.length).toBeGreaterThanOrEqual(2);
  });

  test('auto-detects number type', async () => {
    const path = `${TMP}/read-types.csv`;
    await Bun.write(path, 'val\n42\n3.14\ntrue\nhello');

    const wb = await readCSV(path);
    const rows = wb.worksheets[0].rows;

    expect(rows[1].cells[0].value).toBe(42);
    expect(rows[2].cells[0].value).toBe(3.14);
    expect(rows[3].cells[0].value).toBe(true);
    expect(rows[4].cells[0].value).toBe('hello');
  });

  test('handles empty lines', async () => {
    const path = `${TMP}/read-empty.csv`;
    await Bun.write(path, 'A\n\nB\n');

    const wb = await readCSV(path, { skipEmptyLines: true });
    const rows = wb.worksheets[0].rows;
    expect(rows.length).toBe(2);
  });

  test('handles quoted fields with commas', async () => {
    const path = `${TMP}/read-quoted.csv`;
    await Bun.write(path, 'name,address\nAlice,"123 Main St, Apt 4"');

    const wb = await readCSV(path);
    const rows = wb.worksheets[0].rows;
    expect(rows[1].cells[1].value).toBe('123 Main St, Apt 4');
  });

  test('round-trip: write then read preserves data', async () => {
    const path = `${TMP}/round-trip.csv`;
    await writeCSV(path, [
      ['Name', 'Score'],
      ['Alice', 95],
      ['Bob', 87],
    ]);

    const wb = await readCSV(path);
    const rows = wb.worksheets[0].rows;

    expect(rows[0].cells[0].value).toBe('Name');
    expect(rows[1].cells[1].value).toBe(95);
    expect(rows[2].cells[1].value).toBe(87);
  });
});
