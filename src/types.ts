// ============================================
// bun-spreadsheet — Core Type Definitions
// ============================================

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

/** Fill style */
export interface FillStyle {
  type: 'pattern' | 'gradient';
  pattern?: 'solid' | 'darkGray' | 'mediumGray' | 'lightGray' | 'none';
  fgColor?: string; // hex color
  bgColor?: string; // hex color
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
}

/** Hyperlink */
export interface Hyperlink {
  /** URL target (http, https, mailto, or internal sheet reference like "Sheet2!A1") */
  target: string;
  /** Optional tooltip text shown on hover */
  tooltip?: string;
}

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

/** A single cell */
export interface Cell {
  value: CellValue;
  style?: CellStyle;
  type?: 'string' | 'number' | 'boolean' | 'date' | 'formula';
  /** Formula expression (without leading '='), e.g. "SUM(A1:A10)" */
  formula?: string;
  /** Cached result of the formula (shown before recalculation) */
  formulaResult?: string | number | boolean;
  /** Hyperlink on this cell */
  hyperlink?: Hyperlink;
}

/** A row of cells */
export interface Row {
  cells: Cell[];
  height?: number;
  style?: CellStyle;
}

/** Column configuration */
export interface ColumnConfig {
  width?: number;
  style?: CellStyle;
  header?: string;
}

/** Merge cell range */
export interface MergeCell extends CellRange {}

/** Worksheet */
export interface Worksheet {
  name: string;
  rows: Row[];
  columns?: ColumnConfig[];
  mergeCells?: MergeCell[];
  dataValidations?: DataValidation[];
  freezePane?: { row: number; col: number };
  defaultRowHeight?: number;
  defaultColWidth?: number;
}

/** Workbook */
export interface Workbook {
  worksheets: Worksheet[];
  creator?: string;
  created?: Date;
  modified?: Date;
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
  compress?: boolean;
}

/** Stream writer interface */
export interface StreamWriter<T = void> {
  writeRow(row: Row | CellValue[]): void;
  flush(): void | Promise<void>;
  end(): T | Promise<T>;
}
