import { mkdirSync } from 'node:fs';
import { exportExcelRowsToResponse, readExcel } from '../src';

const OUTPUT = './output';
mkdirSync(OUTPUT, { recursive: true });

const { response, diagnostics } = await exportExcelRowsToResponse({
  filename: 'orders.xlsx',
  sheetName: 'Orders',
  mode: 'chunked',
  rows: [
    ['Order ID', 'Customer', 'Total'],
    ['ORD-001', 'Alice', 125],
    ['ORD-002', 'Bob', 80],
    ['ORD-003', 'Carol', 210],
  ],
});

console.log('Diagnostics');
console.log(diagnostics);

const outputPath = `${OUTPUT}/response-streaming.xlsx`;
await Bun.write(outputPath, new Uint8Array(await response.arrayBuffer()));

const workbook = await readExcel(outputPath);
console.log(
  workbook.worksheets[0]?.rows.map((row) =>
    row.cells.map((cell) => cell.value),
  ),
);
