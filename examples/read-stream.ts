import { mkdirSync } from 'node:fs';
import { readExcelStream, writeExcel } from '../src';

const OUTPUT = './output';
const inputPath = `${OUTPUT}/read-stream-source.xlsx`;

mkdirSync(OUTPUT, { recursive: true });

await writeExcel(inputPath, {
  creator: 'bun-excel example',
  worksheets: [
    {
      name: 'Orders',
      rows: [
        {
          cells: [
            { value: 'Order ID', style: { font: { bold: true } } },
            { value: 'Customer', style: { font: { bold: true } } },
            { value: 'Created At', style: { font: { bold: true } } },
            { value: 'Total', style: { font: { bold: true } } },
          ],
        },
        {
          cells: [
            { value: 'ORD-001' },
            { value: 'Alice' },
            {
              value: new Date('2026-03-16T08:00:00.000Z'),
              style: { numberFormat: 'yyyy-mm-dd hh:mm' },
            },
            {
              value: 125,
              type: 'formula',
              formula: 'SUM(100,25)',
              formulaResult: 125,
            },
          ],
        },
        {
          cells: [
            { value: 'ORD-002' },
            { value: 'Bob' },
            {
              value: new Date('2026-03-16T10:30:00.000Z'),
              style: { numberFormat: 'yyyy-mm-dd hh:mm' },
            },
            {
              value: 80,
              type: 'formula',
              formula: 'SUM(50,30)',
              formulaResult: 80,
            },
          ],
        },
      ],
    },
    {
      name: 'Summary',
      rows: [
        { cells: [{ value: 'Metric' }, { value: 'Value' }] },
        { cells: [{ value: 'Orders' }, { value: 2 }] },
      ],
    },
  ],
});

console.log(`Streaming rows from ${inputPath}\n`);

for await (const entry of readExcelStream(inputPath, { sheets: ['Orders'] })) {
  const values = entry.row.cells.map((cell) => {
    if (cell.value instanceof Date) {
      return cell.value.toISOString();
    }
    return String(cell.value ?? '');
  });

  console.log(
    `[${entry.sheetName}] row ${entry.rowIndex + 1}: ${values.join(' | ')}`,
  );
}

// Preview up to 100 data rows, selecting columns A, C and D.
// Coordinates remain zero-based worksheet indices, even after filtering.
const columns = [0, 2, 3] as const;
for await (const { rowIndex, row } of readExcelStream(inputPath, {
  sheets: ['Orders'],
  startRow: 1,
  maxRows: 100,
  columns,
})) {
  console.log(
    `Selected row ${rowIndex + 1}:`,
    columns.map((col) => row.cells[col]?.value ?? null),
  );
}

// Rectangular range B2:D2: both row bounds are inclusive.
for await (const { row } of readExcelStream(inputPath, {
  sheets: ['Orders'],
  startRow: 1,
  endRow: 1,
  columns: [1, 2, 3],
})) {
  console.log(
    'B2:D2:',
    [1, 2, 3].map((col) => row.cells[col]?.value ?? null),
  );
}
