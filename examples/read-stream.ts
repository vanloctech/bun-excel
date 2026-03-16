import { mkdirSync } from 'node:fs';
import { readExcelStream, writeExcel } from '../src';

const OUTPUT = './output';
const inputPath = `${OUTPUT}/read-stream-source.xlsx`;

mkdirSync(OUTPUT, { recursive: true });

await writeExcel(inputPath, {
  creator: 'bun-spreadsheet example',
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
