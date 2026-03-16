// ============================================
// bun-spreadsheet — Core Type Definitions
// ============================================

/** Input source for reading spreadsheet data */
export type FileSource = string | Bun.BunFile | Bun.S3File;

/** Output target for writing spreadsheet data */
export type FileTarget = string | Bun.BunFile | Bun.S3File;

/** Cell value types */
export type CellValue = string | number | boolean | Date | null | undefined;

/** Font style */
export interface FontStyle {
  name?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string; // hex color e.g. "FF0000"
}

/** Rich text run within a single cell */
export interface RichTextRun {
  text: string;
  font?: FontStyle;
}

/** Fill style */
export interface FillStyle {
  type: 'pattern' | 'gradient';
  pattern?: 'solid' | 'darkGray' | 'mediumGray' | 'lightGray' | 'none';
  fgColor?: string; // hex color
  bgColor?: string; // hex color
}

/** Split pane configuration */
export interface SplitPane {
  x: number;
  y: number;
  topLeftCell?: { row: number; col: number };
}

/** Border edge style */
export interface BorderEdgeStyle {
  style?:
    | 'thin'
    | 'medium'
    | 'thick'
    | 'dotted'
    | 'dashed'
    | 'double'
    | 'hair'
    | 'dashDot'
    | 'dashDotDot'
    | 'mediumDashed'
    | 'mediumDashDot'
    | 'mediumDashDotDot'
    | 'slantDashDot';
  color?: string; // hex color
}

/** Border style */
export interface BorderStyle {
  top?: BorderEdgeStyle;
  bottom?: BorderEdgeStyle;
  left?: BorderEdgeStyle;
  right?: BorderEdgeStyle;
}

/** Alignment style */
export interface AlignmentStyle {
  horizontal?: 'left' | 'center' | 'right' | 'fill' | 'justify';
  vertical?: 'top' | 'center' | 'bottom';
  wrapText?: boolean;
  textRotation?: number;
  indent?: number;
}

/** Complete cell style */
export interface CellStyle {
  font?: FontStyle;
  fill?: FillStyle;
  border?: BorderStyle;
  alignment?: AlignmentStyle;
  numberFormat?: string; // e.g. "#,##0.00", "yyyy-mm-dd"
  protection?: {
    locked?: boolean;
    hidden?: boolean;
  };
}

/** Hyperlink */
export interface Hyperlink {
  /** URL target (http, https, mailto, or internal sheet reference like "Sheet2!A1") */
  target: string;
  /** Optional tooltip text shown on hover */
  tooltip?: string;
}

/** Cell comment/note */
export interface CellComment {
  text: string;
  author?: string;
}

/** Binary data accepted for embedded workbook assets */
export type BinaryData = Uint8Array | ArrayBuffer;

/** Cell/range coordinates */
export interface CellRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

/** Data validation rule */
export interface DataValidation {
  /** Target range(s) for this validation rule */
  range: CellRange | CellRange[];
  /** Validation type */
  type:
    | 'list'
    | 'whole'
    | 'decimal'
    | 'date'
    | 'time'
    | 'textLength'
    | 'custom';
  /** Comparison operator for non-list/custom validations */
  operator?:
    | 'between'
    | 'notBetween'
    | 'equal'
    | 'notEqual'
    | 'greaterThan'
    | 'lessThan'
    | 'greaterThanOrEqual'
    | 'lessThanOrEqual';
  /** Whether blank cells are allowed */
  allowBlank?: boolean;
  /** Show the input prompt when the cell is selected */
  showInputMessage?: boolean;
  /** Show the error alert when invalid data is entered */
  showErrorMessage?: boolean;
  /** Error alert style */
  errorStyle?: 'stop' | 'warning' | 'information';
  /** Input prompt title */
  promptTitle?: string;
  /** Input prompt body */
  prompt?: string;
  /** Error alert title */
  errorTitle?: string;
  /** Error alert body */
  error?: string;
  /** First formula or literal list values */
  formula1?: string | number | Date | string[];
  /** Optional second formula */
  formula2?: string | number | Date;
}

/** Conditional formatting threshold */
export interface ConditionalFormatThreshold {
  type: 'min' | 'max' | 'num' | 'percent' | 'percentile' | 'formula';
  value?: string | number | Date;
  gte?: boolean;
}

/** Conditional formatting rule: highlight cells by comparison */
export interface ConditionalFormatCellRule {
  type: 'cellIs';
  operator:
    | 'between'
    | 'notBetween'
    | 'equal'
    | 'notEqual'
    | 'greaterThan'
    | 'lessThan'
    | 'greaterThanOrEqual'
    | 'lessThanOrEqual';
  formula1: string | number | Date;
  formula2?: string | number | Date;
  style?: CellStyle;
  priority?: number;
  stopIfTrue?: boolean;
}

/** Conditional formatting rule: custom formula */
export interface ConditionalFormatExpressionRule {
  type: 'expression';
  formula: string;
  style?: CellStyle;
  priority?: number;
  stopIfTrue?: boolean;
}

/** Conditional formatting rule: color scale */
export interface ConditionalFormatColorScaleRule {
  type: 'colorScale';
  thresholds: ConditionalFormatThreshold[];
  colors: string[];
  priority?: number;
}

/** Conditional formatting rule: data bar */
export interface ConditionalFormatDataBarRule {
  type: 'dataBar';
  min?: ConditionalFormatThreshold;
  max?: ConditionalFormatThreshold;
  color: string;
  showValue?: boolean;
  minLength?: number;
  maxLength?: number;
  priority?: number;
}

/** Conditional formatting rule: icon set */
export interface ConditionalFormatIconSetRule {
  type: 'iconSet';
  iconSet: string;
  thresholds: ConditionalFormatThreshold[];
  showValue?: boolean;
  reverse?: boolean;
  priority?: number;
}

/** Conditional formatting rule union */
export type ConditionalFormattingRule =
  | ConditionalFormatCellRule
  | ConditionalFormatExpressionRule
  | ConditionalFormatColorScaleRule
  | ConditionalFormatDataBarRule
  | ConditionalFormatIconSetRule;

/** Worksheet conditional formatting block */
export interface ConditionalFormatting {
  range: CellRange | CellRange[];
  rules: ConditionalFormattingRule[];
}

/** A single cell */
export interface Cell {
  value: CellValue;
  style?: CellStyle;
  type?: 'string' | 'number' | 'boolean' | 'date' | 'formula';
  richText?: RichTextRun[];
  /** Formula expression (without leading '='), e.g. "SUM(A1:A10)" */
  formula?: string;
  /** Cached result of the formula (shown before recalculation) */
  formulaResult?: string | number | boolean;
  /** Hyperlink on this cell */
  hyperlink?: Hyperlink;
  /** Optional cell comment/note */
  comment?: CellComment;
}

/** A row of cells */
export interface Row {
  cells: Cell[];
  height?: number;
  style?: CellStyle;
  hidden?: boolean;
  collapsed?: boolean;
  outlineLevel?: number;
}

/** Column configuration */
export interface ColumnConfig {
  width?: number;
  style?: CellStyle;
  header?: string;
  hidden?: boolean;
  collapsed?: boolean;
  outlineLevel?: number;
}

/** Merge cell range */
export interface MergeCell extends CellRange {}

export type WorksheetState = 'visible' | 'hidden' | 'veryHidden';

export interface WorksheetProtection {
  password?: string;
  sheet?: boolean;
  objects?: boolean;
  scenarios?: boolean;
  formatCells?: boolean;
  formatColumns?: boolean;
  formatRows?: boolean;
  insertColumns?: boolean;
  insertRows?: boolean;
  insertHyperlinks?: boolean;
  deleteColumns?: boolean;
  deleteRows?: boolean;
  selectLockedCells?: boolean;
  sort?: boolean;
  autoFilter?: boolean;
  pivotTables?: boolean;
  selectUnlockedCells?: boolean;
}

export interface PageMargins {
  left?: number;
  right?: number;
  top?: number;
  bottom?: number;
  header?: number;
  footer?: number;
}

export interface PageSetup {
  orientation?: 'portrait' | 'landscape';
  paperSize?: number;
  scale?: number;
  fitToWidth?: number;
  fitToHeight?: number;
  firstPageNumber?: number;
  useFirstPageNumber?: boolean;
}

export interface HeaderFooterSection {
  left?: string;
  center?: string;
  right?: string;
}

export interface HeaderFooter {
  differentFirst?: boolean;
  differentOddEven?: boolean;
  oddHeader?: HeaderFooterSection;
  oddFooter?: HeaderFooterSection;
  evenHeader?: HeaderFooterSection;
  evenFooter?: HeaderFooterSection;
  firstHeader?: HeaderFooterSection;
  firstFooter?: HeaderFooterSection;
}

export interface WorksheetImage {
  data: BinaryData;
  format: 'png' | 'jpeg' | 'jpg' | 'gif';
  range: CellRange;
  name?: string;
  description?: string;
}

export interface WorksheetTableColumn {
  name: string;
  totalsRowLabel?: string;
  totalsRowFunction?:
    | 'sum'
    | 'average'
    | 'count'
    | 'countNums'
    | 'max'
    | 'min'
    | 'stdDev'
    | 'var'
    | 'custom';
}

export interface WorksheetTableStyle {
  name?: string;
  showFirstColumn?: boolean;
  showLastColumn?: boolean;
  showRowStripes?: boolean;
  showColumnStripes?: boolean;
}

export interface WorksheetTable {
  name: string;
  displayName?: string;
  range: CellRange;
  headerRow?: boolean;
  totalsRow?: boolean;
  columns?: WorksheetTableColumn[];
  style?: WorksheetTableStyle;
}

/** Worksheet */
export interface Worksheet {
  name: string;
  rows: Row[];
  columns?: ColumnConfig[];
  mergeCells?: MergeCell[];
  autoFilter?: CellRange;
  dataValidations?: DataValidation[];
  conditionalFormattings?: ConditionalFormatting[];
  freezePane?: { row: number; col: number };
  splitPane?: SplitPane;
  defaultRowHeight?: number;
  defaultColWidth?: number;
  state?: WorksheetState;
  protection?: WorksheetProtection;
  pageMargins?: PageMargins;
  pageSetup?: PageSetup;
  headerFooter?: HeaderFooter;
  printArea?: CellRange;
  images?: WorksheetImage[];
  tables?: WorksheetTable[];
}

export interface DefinedName {
  name: string;
  refersTo: string;
  comment?: string;
  hidden?: boolean;
  localSheetId?: number;
}

export interface WorkbookView {
  activeTab?: number;
  firstSheet?: number;
  visibility?: 'visible' | 'hidden' | 'veryHidden';
  xWindow?: number;
  yWindow?: number;
  windowWidth?: number;
  windowHeight?: number;
}

/** Workbook */
export interface Workbook {
  worksheets: Worksheet[];
  creator?: string;
  created?: Date;
  modified?: Date;
  definedNames?: DefinedName[];
  views?: WorkbookView;
}

/** Row emitted by the streaming XLSX reader */
export interface ExcelReadStreamRow {
  sheetIndex: number;
  sheetName: string;
  rowIndex: number;
  row: Row;
}

/** CSV read options */
export interface CSVReadOptions {
  delimiter?: string;
  quoteChar?: string;
  escapeChar?: string;
  hasHeader?: boolean;
  encoding?: string;
  skipEmptyLines?: boolean;
}

/** CSV write options */
export interface CSVWriteOptions {
  delimiter?: string;
  quoteChar?: string;
  lineEnding?: string;
  includeHeader?: boolean;
  headers?: string[];
  bom?: boolean;
}

/** Excel read options */
export interface ExcelReadOptions {
  sheets?: string[] | number[];
  includeStyles?: boolean;
}

/** Excel write options */
export interface ExcelWriteOptions {
  creator?: string;
  created?: Date;
  modified?: Date;
  compress?: boolean;
  definedNames?: DefinedName[];
  views?: WorkbookView;
}

/** Stream writer interface */
export interface StreamWriter<T = void> {
  writeRow(row: Row | CellValue[]): void;
  flush(): void | Promise<void>;
  end(): T | Promise<T>;
}
