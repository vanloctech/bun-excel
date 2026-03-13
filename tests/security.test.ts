import { describe, expect, test } from 'bun:test';
import { readExcel, writeExcel } from '../src';

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
