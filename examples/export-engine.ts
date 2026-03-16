import { mkdirSync } from 'node:fs';
import {
  buildExcelResponse,
  exportExcelRows,
  exportMultiSheetExcel,
} from '../src';

const OUTPUT = './output';
mkdirSync(OUTPUT, { recursive: true });

console.log('Single-sheet production export\n');

const singleSheetResult = await exportExcelRows({
  target: `${OUTPUT}/production-export-single.xlsx`,
  sheetName: 'Orders',
  mode: 'chunked',
  progressIntervalRows: 2,
  rows: [
    ['Order ID', 'Customer', 'Total'],
    ['ORD-001', 'Alice', 125],
    ['ORD-002', 'Bob', 80],
    ['ORD-003', 'Carol', 210],
  ],
  onProgress(progress) {
    console.log(
      `[single] ${progress.stage} rows=${progress.rowsWritten} rss=${Math.round(
        progress.memory.rssBytes / 1024 / 1024,
      )}MB`,
    );
  },
});

console.log(singleSheetResult);

console.log('\nMulti-sheet production export\n');

const multiSheetResult = await exportMultiSheetExcel({
  target: `${OUTPUT}/production-export-multi.xlsx`,
  creator: 'bun-spreadsheet example',
  progressIntervalRows: 2,
  sheets: [
    {
      name: 'Orders',
      rows: [
        ['Order ID', 'Customer', 'Total'],
        ['ORD-001', 'Alice', 125],
        ['ORD-002', 'Bob', 80],
      ],
    },
    {
      name: 'Summary',
      rows: [
        ['Metric', 'Value'],
        ['Orders', 2],
        ['Revenue', 205],
      ],
    },
  ],
  onProgress(progress) {
    console.log(
      `[multi] ${progress.stage} sheet=${progress.sheetName ?? '-'} rows=${progress.rowsWritten}`,
    );
  },
});

console.log(multiSheetResult);

const response = await buildExcelResponse(
  {
    worksheets: [
      {
        name: 'Download',
        rows: [{ cells: [{ value: 'Ready' }, { value: true }] }],
      },
    ],
  },
  { filename: 'download.xlsx' },
);

console.log('\nResponse headers');
console.log(response.headers.get('content-type'));
console.log(response.headers.get('content-disposition'));
