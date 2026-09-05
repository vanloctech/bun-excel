import { type Workbook, writeExcel } from '../src';

// EXCEL_PASSWORD=<your-password> bun run examples/encrypted-export.ts output.xlsx
const password = process.env.EXCEL_PASSWORD;
if (!password)
  throw new Error('Set EXCEL_PASSWORD before running this example');
const target = process.argv[2] ?? './output/private.xlsx';
const workbook: Workbook = {
  worksheets: [
    {
      name: 'Private',
      rows: [
        { cells: [{ value: 'Name' }, { value: 'Amount' }] },
        { cells: [{ value: 'Nguyễn An' }, { value: 1500000 }] },
      ],
    },
  ],
};
await writeExcel(target, workbook, { password });
console.log(`Encrypted workbook written to ${target}`);
