import { readExcelValuesStream } from '../src';

// Run: bun run examples/read-values.ts input.xlsx
const source = process.argv[2];
if (!source)
  throw new Error('Usage: bun run examples/read-values.ts input.xlsx');
for await (const batch of readExcelValuesStream(source, { batchSize: 256 })) {
  for (let i = 0; i < batch.rows.length; i++) {
    console.log(batch.sheetName, batch.rowIndices[i], batch.rows[i]);
  }
  // Await your database insert here to keep pending batches bounded.
}
