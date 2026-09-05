import { ExcelHeaderError, readExcelObjectsStream } from '../src';

// bun run examples/read-objects.ts products.xlsx
// Products headers: Name | Price | Active
const source = process.argv[2];
if (!source)
  throw new Error('Usage: bun run examples/read-objects.ts products.xlsx');
try {
  for await (const result of readExcelObjectsStream(source, {
    sheets: ['Products'],
    schema: {
      name: { header: 'Name', type: 'string', required: true },
      price: { header: 'Price', type: 'number', required: true },
      active: { header: 'Active', type: 'boolean' },
    },
  })) {
    if (result.ok) console.log(result.data);
    else console.error(result.sheetName, result.rowIndex, result.errors);
  }
} catch (error) {
  if (error instanceof ExcelHeaderError) {
    console.error(error.code, error.sheetName, error.rowIndex, error.header);
    process.exitCode = 1;
  } else {
    throw error;
  }
}
