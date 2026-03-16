// ============================================
// Large Report — 30 columns x 30,000 rows
// realistic business-report workload
// ============================================

import { mkdirSync } from 'node:fs';
import { writeExcel } from '../src';
import {
  buildLargeReportWorkbook,
  createLargeReportMergeCells,
  LARGE_REPORT_COL_COUNT,
  LARGE_REPORT_DATA_ROWS,
  LARGE_REPORT_OUTPUT,
  largeReportNumericCols,
} from './large-report-workload';

const OUTPUT = './output';

mkdirSync(OUTPUT, { recursive: true });

console.log('Large Report Generator');
console.log('='.repeat(60));
console.log(
  `\nGenerating ${LARGE_REPORT_DATA_ROWS.toLocaleString()} rows x ${LARGE_REPORT_COL_COUNT} columns...`,
);

const startTime = Bun.nanoseconds();
const workbook = buildLargeReportWorkbook();
const genTime = (Bun.nanoseconds() - startTime) / 1_000_000;

console.log(`      Data generated in ${genTime.toFixed(0)}ms`);
console.log('\nWriting XLSX file...');

const writeStart = Bun.nanoseconds();
await writeExcel(LARGE_REPORT_OUTPUT, workbook);
const writeTime = (Bun.nanoseconds() - writeStart) / 1_000_000;
const totalTime = (Bun.nanoseconds() - startTime) / 1_000_000;

const fileInfo = Bun.file(LARGE_REPORT_OUTPUT);
const fileSizeMB = (fileInfo.size / (1024 * 1024)).toFixed(2);
const rows = workbook.worksheets[0]?.rows ?? [];
const mergeCells = createLargeReportMergeCells();

console.log(`      -> ${LARGE_REPORT_OUTPUT}`);
console.log(`\n${'='.repeat(60)}`);
console.log('Summary:\n');
console.log(
  `  Dimensions:  ${LARGE_REPORT_COL_COUNT} columns x ${LARGE_REPORT_DATA_ROWS.toLocaleString()} data rows`,
);
console.log(
  `  Total rows:  ${rows.length.toLocaleString()} (incl. title, header, footer)`,
);
console.log(
  `  Merge cells: ${mergeCells.length} (title, subtitle, 4 footer labels)`,
);
console.log(
  `  Formulas:    ${largeReportNumericCols.length * 4} (SUM + AVG + MAX + MIN x ${largeReportNumericCols.length} columns)`,
);
console.log(`  File size:   ${fileSizeMB} MB`);
console.log(`  Gen time:    ${genTime.toFixed(0)}ms`);
console.log(`  Write time:  ${writeTime.toFixed(0)}ms`);
console.log(`  Total time:  ${totalTime.toFixed(0)}ms`);
