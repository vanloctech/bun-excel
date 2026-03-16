// ============================================
// Template Mode — load, fill, and write
// ============================================

import { mkdirSync } from 'node:fs';
import { loadExcelTemplate, type Workbook, writeExcel } from '../src/index';

const OUTPUT = './output';
const TEMPLATE_PATH = `${OUTPUT}/invoice-template.xlsx`;
const FILLED_PATH = `${OUTPUT}/invoice-filled.xlsx`;

mkdirSync(OUTPUT, { recursive: true });

console.log('Template Mode Example');
console.log('='.repeat(60));
console.log('\n1. Creating template workbook...');

const templateWorkbook: Workbook = {
  creator: 'bun-spreadsheet',
  definedNames: [
    { name: 'InvoiceNumber', refersTo: "'Invoice'!$B$2" },
    { name: 'InvoiceDate', refersTo: "'Invoice'!$B$3" },
    { name: 'LineItems', refersTo: "'Invoice'!$A$6:$C$8" },
    { name: 'InvoiceTotal', refersTo: "'Invoice'!$C$10" },
  ],
  worksheets: [
    {
      name: 'Invoice',
      freezePane: { row: 1, col: 0 },
      rows: [
        {
          cells: [
            {
              value: 'Invoice Template',
              style: {
                font: { bold: true, size: 18, color: '1F4E78' },
              },
            },
          ],
        },
        {
          cells: [
            { value: 'Invoice #' },
            {
              value: 'TBD',
              style: { font: { bold: true, color: '1F4E78' } },
              comment: { text: 'Filled by template mode', author: 'Loc' },
            },
          ],
        },
        { cells: [{ value: 'Invoice Date' }, { value: null }] },
        { cells: [] },
        {
          cells: [
            { value: 'Item', style: { font: { bold: true } } },
            { value: 'Qty', style: { font: { bold: true } } },
            { value: 'Price', style: { font: { bold: true } } },
          ],
        },
        { cells: [{ value: '' }, { value: 0 }, { value: 0 }] },
        { cells: [{ value: '' }, { value: 0 }, { value: 0 }] },
        { cells: [{ value: '' }, { value: 0 }, { value: 0 }] },
        { cells: [] },
        {
          cells: [
            { value: 'Total', style: { font: { bold: true } } },
            { value: null },
            {
              value: 0,
              style: { font: { bold: true }, numberFormat: '$#,##0.00' },
            },
          ],
        },
      ],
      tables: [
        {
          name: 'ItemsTable',
          range: { startRow: 4, startCol: 0, endRow: 7, endCol: 2 },
          columns: [{ name: 'Item' }, { name: 'Qty' }, { name: 'Price' }],
          style: { name: 'TableStyleMedium2', showRowStripes: true },
        },
      ],
    },
  ],
};

await writeExcel(TEMPLATE_PATH, templateWorkbook);
console.log(`      -> ${TEMPLATE_PATH}`);

console.log('\n2. Loading template...');
const template = await loadExcelTemplate(TEMPLATE_PATH);

console.log('3. Filling named ranges and cells...');
template.setDefinedName('InvoiceNumber', 'INV-2026-001');
template.setDefinedName('InvoiceDate', new Date('2026-03-16T00:00:00Z'));
template.setDefinedName('LineItems', [
  ['Apple', 2, 12.5],
  ['Orange', 5, 8.25],
  ['Banana', 3, 5.75],
]);
template.setDefinedName('InvoiceTotal', 84.0);
template.setCell('Invoice', 'A1', 'Invoice 2026');
template.setCell('Invoice', 'B2', {
  value: 'INV-2026-001',
  style: { font: { bold: true, color: 'C00000' } },
});

console.log('\n4. Writing filled workbook...');
await template.write(FILLED_PATH);
console.log(`      -> ${FILLED_PATH}`);

console.log('\n5. Done');
console.log('\nYou can compare the template and filled files in ./output/');
