// ============================================
// bun-spreadsheet — Main Entry Point
// ============================================

// CSV
export { readCSV, readCSVStream } from './csv/csv-reader';
export { CSVStreamWriter, createCSVStream, writeCSV } from './csv/csv-writer';
export {
  ExcelTemplate,
  loadExcelTemplate,
  type TemplateCellInput,
  type TemplateCellMatrix,
  type TemplateCellRef,
  type TemplateSheetRef,
} from './excel/template-mode';
export {
  type ChunkedExcelStreamOptions,
  createChunkedExcelStream,
  ExcelChunkedStreamWriter,
} from './excel/xlsx-chunked-stream-writer';
// Excel
export { readExcel, readExcelStream } from './excel/xlsx-reader';
export {
  createExcelStream,
  createMultiSheetExcelStream,
  type ExcelStreamOptions,
  ExcelStreamWriter,
  MultiSheetExcelStreamWriter,
} from './excel/xlsx-stream-writer';
export {
  buildExcelBuffer,
  excelSerialToDate,
  writeExcel,
} from './excel/xlsx-writer';
// Types
export type {
  AlignmentStyle,
  BinaryData,
  BorderEdgeStyle,
  BorderStyle,
  Cell,
  CellComment,
  CellRange,
  CellStyle,
  CellValue,
  ColumnConfig,
  ConditionalFormatCellRule,
  ConditionalFormatColorScaleRule,
  ConditionalFormatDataBarRule,
  ConditionalFormatExpressionRule,
  ConditionalFormatIconSetRule,
  ConditionalFormatThreshold,
  ConditionalFormatting,
  ConditionalFormattingRule,
  CSVReadOptions,
  CSVWriteOptions,
  DataValidation,
  DefinedName,
  ExcelReadOptions,
  ExcelReadStreamRow,
  ExcelWriteOptions,
  FileSource,
  FileTarget,
  FillStyle,
  FontStyle,
  HeaderFooter,
  HeaderFooterSection,
  Hyperlink,
  MergeCell,
  PageMargins,
  PageSetup,
  RichTextRun,
  Row,
  SplitPane,
  StreamWriter,
  Workbook,
  WorkbookView,
  Worksheet,
  WorksheetImage,
  WorksheetProtection,
  WorksheetState,
  WorksheetTable,
  WorksheetTableColumn,
  WorksheetTableStyle,
} from './types';
export {
  duplicateRow,
  insertColumns,
  insertRows,
  spliceRows,
} from './worksheet-ops';
