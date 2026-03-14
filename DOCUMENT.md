# API Documentation

Complete API reference for bun-spreadsheet.

---

## Table of Contents

- [Excel](#excel)
  - [writeExcel](#writeexcelpath-workbook-options)
  - [readExcel](#readexcelpath-options)
  - [buildExcelBuffer](#buildexcelbufferworkbook-options)
- [CSV](#csv)
  - [writeCSV](#writecsvpath-data-options)
  - [readCSV](#readcsvpath-options)
  - [readCSVStream](#readcsvstreampath-options)
  - [createCSVStream](#createcsvstreampath-options)
- [Excel Streaming](#excel-streaming)
  - [createExcelStream](#createexcelstreampath-options)
  - [createMultiSheetExcelStream](#createmultisheetexcelstreampath-options)
  - [createChunkedExcelStream](#createchunkedexcelstreampath-options)
- [Types](#types)
  - [Workbook](#workbook)
  - [Worksheet](#worksheet)
  - [Row](#row)
  - [Cell](#cell)
  - [CellValue](#cellvalue)
  - [ColumnConfig](#columnconfig)
  - [MergeCell](#mergecell)
  - [Hyperlink](#hyperlink)
  - [DataValidation](#datavalidation)
  - [ConditionalFormatting](#conditionalformatting)
- [Styles](#styles)
  - [CellStyle](#cellstyle)
  - [FontStyle](#fontstyle)
  - [FillStyle](#fillstyle)
  - [BorderStyle](#borderstyle)
  - [AlignmentStyle](#alignmentstyle)
  - [Number Formats](#number-formats)
- [Features](#features)
  - [Formulas](#formulas)
  - [Hyperlinks](#hyperlinks-1)
  - [Merge Cells](#merge-cells)
  - [Freeze Panes](#freeze-panes)
  - [Data Validation](#data-validation)
  - [Conditional Formatting](#conditional-formatting)
- [Writing Modes Comparison](#writing-modes-comparison)

---

## Excel

### `writeExcel(path, workbook, options?)`

Write a Workbook to an `.xlsx` file.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `path` | `string` | Yes | Output file path |
| `workbook` | `Workbook` | Yes | Workbook data to write |
| `options` | `ExcelWriteOptions` | No | Write options |

**ExcelWriteOptions:**

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `creator` | `string` | `undefined` | Author name in file metadata |
| `compress` | `boolean` | `true` | Enable ZIP compression |

**Returns:** `Promise<void>`

**Example:**

```typescript
import { writeExcel, type Workbook } from "bun-spreadsheet";

const workbook: Workbook = {
  worksheets: [{
    name: "Sheet1",
    columns: [{ width: 20 }, { width: 15 }],
    rows: [
      {
        cells: [
          { value: "Name", style: { font: { bold: true } } },
          { value: "Score", style: { font: { bold: true } } },
        ],
      },
      { cells: [{ value: "Alice" }, { value: 95 }] },
      { cells: [{ value: "Bob" }, { value: 87 }] },
    ],
  }],
};

await writeExcel("report.xlsx", workbook, { creator: "My App" });
```

---

### `readExcel(path, options?)`

Read an `.xlsx` file into a Workbook object.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `path` | `string` | Yes | Path to .xlsx file |
| `options` | `ExcelReadOptions` | No | Read options |

**ExcelReadOptions:**

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `sheets` | `string[] \| number[]` | all sheets | Specific sheets to read (by name or index) |
| `includeStyles` | `boolean` | `true` | Whether to parse and include cell styles |

**Returns:** `Promise<Workbook>`

**Example:**

```typescript
import { readExcel } from "bun-spreadsheet";

// Read all sheets
const workbook = await readExcel("report.xlsx");

// Read specific sheets only
const partial = await readExcel("report.xlsx", {
  sheets: ["Sheet1"],
  includeStyles: false,  // faster if you don't need styles
});

// Iterate over data
for (const sheet of workbook.worksheets) {
  console.log(`Sheet: ${sheet.name}`);
  for (const row of sheet.rows) {
    console.log(row.cells.map(c => c.value).join(" | "));
  }
}
```

---

### `buildExcelBuffer(workbook, options?)`

Build an Excel file as a `Uint8Array` buffer without writing to disk. Useful for sending as HTTP response or further processing.

**Parameters:** Same as `writeExcel` except no `path`.

**Returns:** `Uint8Array`

**Example:**

```typescript
import { buildExcelBuffer } from "bun-spreadsheet";

const buffer = buildExcelBuffer(workbook);
// Use buffer for HTTP response, upload, etc.
```

---

## CSV

### `writeCSV(path, data, options?)`

Write data to a CSV file.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `path` | `string` | Yes | Output file path |
| `data` | `Workbook \| CellValue[][]` | Yes | Data to write |
| `options` | `CSVWriteOptions` | No | Write options |

**CSVWriteOptions:**

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `delimiter` | `string` | `","` | Field delimiter |
| `quoteChar` | `string` | `"\""` | Quote character |
| `lineEnding` | `string` | `"\n"` | Line ending |
| `includeHeader` | `boolean` | `false` | Whether to include header row |
| `headers` | `string[]` | `undefined` | Custom header names |
| `bom` | `boolean` | `false` | Add UTF-8 BOM (for Excel compatibility) |

**Returns:** `Promise<void>`

**Example:**

```typescript
import { writeCSV } from "bun-spreadsheet";

// Simple array data
await writeCSV("data.csv", [
  ["Name", "Age", "City"],
  ["Alice", 28, "Hanoi"],
  ["Bob", 32, "Ho Chi Minh"],
]);

// With options
await writeCSV("export.csv", data, {
  delimiter: ";",
  bom: true,           // Excel-compatible UTF-8
  includeHeader: true,
  headers: ["ID", "Name", "Value"],
});
```

---

### `readCSV(path, options?)`

Read a CSV file into a Workbook object (single worksheet).

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `path` | `string` | Yes | Path to CSV file |
| `options` | `CSVReadOptions` | No | Read options |

**CSVReadOptions:**

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `delimiter` | `string` | `","` | Field delimiter |
| `quoteChar` | `string` | `"\""` | Quote character |
| `escapeChar` | `string` | `"\""` | Escape character |
| `hasHeader` | `boolean` | `false` | Whether first row is header |
| `encoding` | `string` | `"utf-8"` | File encoding |
| `skipEmptyLines` | `boolean` | `false` | Skip empty lines |

**Returns:** `Promise<Workbook>`

**Example:**

```typescript
import { readCSV } from "bun-spreadsheet";

const workbook = await readCSV("data.csv", {
  hasHeader: true,
  skipEmptyLines: true,
});

const rows = workbook.worksheets[0].rows;
// Values are auto-detected: numbers, booleans, dates, strings
```

Auto-type detection converts:
- `"42"` -> `42` (number)
- `"3.14"` -> `3.14` (number)
- `"true"` / `"false"` -> `true` / `false` (boolean)
- `"2024-01-15"` -> `Date` object
- Everything else -> `string`

---

### `readCSVStream(path, options?)`

Stream-read a large CSV file row by row. Returns an `AsyncGenerator`.

**Parameters:** Same as `readCSV`.

**Returns:** `AsyncGenerator<Row>`

**Example:**

```typescript
import { readCSVStream } from "bun-spreadsheet";

for await (const row of readCSVStream("large.csv")) {
  const values = row.cells.map(c => c.value);
  // Process each row without loading entire file into memory
}
```

---

### `createCSVStream(path, options?)`

Create a streaming CSV writer. Writes rows directly to disk.

**Parameters:** Same as `writeCSV`.

**Returns:** `CSVStreamWriter`

**CSVStreamWriter methods:**

| Method | Description |
|--------|-------------|
| `writeRow(values: CellValue[])` | Write a single row |
| `flush()` | Flush buffer to disk |
| `end(): Promise<void>` | Finalize and close the file |

**Example:**

```typescript
import { createCSVStream } from "bun-spreadsheet";

const stream = createCSVStream("output.csv", {
  headers: ["ID", "Name", "Value"],
  includeHeader: true,
});

for (let i = 0; i < 100000; i++) {
  stream.writeRow([i + 1, `Item_${i}`, Math.random() * 1000]);
}

await stream.end();
```

---

## Excel Streaming

Three streaming modes for different scenarios:

| Mode | Memory | Best For | Shared Strings |
|------|--------|----------|----------------|
| `createExcelStream` | Moderate | Most use cases (< 100K rows) | Yes (in-memory) |
| `createMultiSheetExcelStream` | Moderate | Multiple sheets | Yes (in-memory) |
| `createChunkedExcelStream` | Constant (~low) | Very large files (100K+ rows) | No (inline strings) |

---

### `createExcelStream(path, options?)`

Create a streaming Excel writer. Serializes each row to XML immediately but keeps shared strings in memory.

**ExcelStreamOptions:**

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `sheetName` | `string` | `"Sheet1"` | Name of the worksheet |
| `columns` | `ColumnConfig[]` | `undefined` | Column width configurations |
| `defaultRowHeight` | `number` | `15` | Default row height |
| `freezePane` | `{ row, col }` | `undefined` | Freeze pane position |
| `mergeCells` | `MergeCell[]` | `undefined` | Merge cell ranges |
| `creator` | `string` | `undefined` | Author name |
| `compress` | `boolean` | `true` | Enable ZIP compression |

**Returns:** `ExcelStreamWriter`

**ExcelStreamWriter methods:**

| Method | Description |
|--------|-------------|
| `writeRow(row: Row \| CellValue[])` | Write a single row (as Row object or plain array) |
| `flush()` | Flush buffer |
| `end(): Promise<void>` | Finalize ZIP and write to disk |

**Example:**

```typescript
import { createExcelStream } from "bun-spreadsheet";

const stream = createExcelStream("report.xlsx", {
  sheetName: "Data",
  columns: [{ width: 10 }, { width: 25 }, { width: 15 }],
  freezePane: { row: 1, col: 0 },
});

// Header row with styles (using Row object)
stream.writeRow({
  cells: [
    { value: "ID", style: { font: { bold: true } } },
    { value: "Name", style: { font: { bold: true } } },
    { value: "Price", style: { font: { bold: true } } },
  ],
});

// Data rows (using plain arrays for convenience)
for (let i = 0; i < 50000; i++) {
  stream.writeRow([i + 1, `Product_${i}`, Math.random() * 1000]);
}

await stream.end();
```

---

### `createMultiSheetExcelStream(path, options?)`

Create a streaming Excel writer with support for multiple sheets.

**Returns:** `MultiSheetExcelStreamWriter`

**MultiSheetExcelStreamWriter methods:**

| Method | Description |
|--------|-------------|
| `addSheet(name, options?)` | Add a new sheet and switch to it |
| `writeRow(row: Row \| CellValue[])` | Write a row to the current sheet |
| `flush()` | Flush buffer |
| `end(): Promise<void>` | Finalize |

**addSheet options:**

| Option | Type | Description |
|--------|------|-------------|
| `columns` | `ColumnConfig[]` | Column configurations for this sheet |
| `freezePane` | `{ row, col }` | Freeze pane |
| `mergeCells` | `MergeCell[]` | Merge cell ranges |

**Example:**

```typescript
import { createMultiSheetExcelStream } from "bun-spreadsheet";

const stream = createMultiSheetExcelStream("multi.xlsx");

stream.addSheet("Revenue", {
  columns: [{ width: 15 }, { width: 12 }],
  freezePane: { row: 1, col: 0 },
});
stream.writeRow(["Month", "Amount"]);
stream.writeRow(["January", 50000]);
stream.writeRow(["February", 62000]);

stream.addSheet("Expenses", {
  columns: [{ width: 15 }, { width: 12 }],
});
stream.writeRow(["Category", "Amount"]);
stream.writeRow(["Salaries", 30000]);

await stream.end();
```

---

### `createChunkedExcelStream(path, options?)`

Create a chunked streaming Excel writer with constant memory usage. Row XML is written to a temporary file on disk, then assembled into ZIP at the end.

**ChunkedExcelStreamOptions:** Same as `ExcelStreamOptions`.

**Returns:** `ExcelChunkedStreamWriter`

**ExcelChunkedStreamWriter methods:**

| Method | Description |
|--------|-------------|
| `writeRow(row: Row \| CellValue[])` | Write a row (serialized to disk immediately) |
| `writeStyledRow(values, styles)` | Write a row with per-cell styles |
| `writeRows(rows)` | Write multiple rows at once |
| `flush()` | Flush temp file buffer |
| `end(): Promise<void>` | Assemble ZIP from temp file and write output |
| `currentRowCount` | Get current row count |

**How it works:**

1. `writeRow()` - serializes row XML and writes to a temp file on disk (no memory retention)
2. `end()` - reads temp file, wraps with worksheet XML, creates ZIP, writes output, deletes temp file

Uses **inline strings** (`<is><t>...</t></is>`) instead of shared string table, so no string tracking is needed in memory.

**Example:**

```typescript
import { createChunkedExcelStream } from "bun-spreadsheet";

const stream = createChunkedExcelStream("huge_report.xlsx", {
  sheetName: "Report",
  columns: [{ width: 14 }, { width: 20 }, { width: 12 }],
  freezePane: { row: 1, col: 0 },
  mergeCells: [
    { startRow: 0, startCol: 0, endRow: 0, endCol: 2 },
  ],
});

// Header
stream.writeRow({
  cells: [
    { value: "ID", style: { font: { bold: true } } },
    { value: "Name", style: { font: { bold: true } } },
    { value: "Value", style: { font: { bold: true } } },
  ],
});

// 1 million rows -- memory stays constant
for (let i = 0; i < 1_000_000; i++) {
  stream.writeRow([i, `Row ${i}`, Math.random() * 10000]);
}

await stream.end();
```

---

## Types

### Workbook

```typescript
interface Workbook {
  worksheets: Worksheet[];
  creator?: string;      // Author name
  created?: Date;        // Created timestamp
  modified?: Date;       // Modified timestamp
}
```

### Worksheet

```typescript
interface Worksheet {
  name: string;                              // Sheet name
  rows: Row[];                               // Array of rows
  columns?: ColumnConfig[];                  // Column configurations
  mergeCells?: MergeCell[];                  // Merged cell ranges
  dataValidations?: DataValidation[];        // Data validation rules
  conditionalFormattings?: ConditionalFormatting[]; // Conditional formatting rules
  freezePane?: { row: number; col: number }; // Freeze pane position
  defaultRowHeight?: number;                 // Default row height
  defaultColWidth?: number;                  // Default column width
}
```

### Row

```typescript
interface Row {
  cells: Cell[];       // Array of cells in this row
  height?: number;     // Custom row height
  style?: CellStyle;   // Default style for all cells in this row
}
```

### Cell

```typescript
interface Cell {
  value: CellValue;                                // Cell value
  style?: CellStyle;                               // Cell style
  type?: "string" | "number" | "boolean" | "date" | "formula";
  formula?: string;                                // Formula (without "=")
  formulaResult?: string | number | boolean;       // Cached formula result
  hyperlink?: Hyperlink;                           // Hyperlink on this cell
}
```

### CellValue

```typescript
type CellValue = string | number | boolean | Date | null | undefined;
```

### ColumnConfig

```typescript
interface ColumnConfig {
  width?: number;      // Column width in characters
  style?: CellStyle;   // Default column style
  header?: string;     // Header text
}
```

### MergeCell

```typescript
interface MergeCell {
  startRow: number;    // Start row (0-indexed)
  startCol: number;    // Start column (0-indexed)
  endRow: number;      // End row (0-indexed)
  endCol: number;      // End column (0-indexed)
}
```

### Hyperlink

```typescript
interface Hyperlink {
  target: string;      // URL, mailto:, or internal ref like "Sheet2!A1"
  tooltip?: string;    // Tooltip text on hover
}
```

### DataValidation

```typescript
interface DataValidation {
  range: CellRange | CellRange[];    // Target range(s), 0-indexed
  type: "list" | "whole" | "decimal" | "date" | "time" | "textLength" | "custom";
  operator?: "between" | "notBetween" | "equal" | "notEqual"
           | "greaterThan" | "lessThan" | "greaterThanOrEqual" | "lessThanOrEqual";
  allowBlank?: boolean;
  showInputMessage?: boolean;
  showErrorMessage?: boolean;
  errorStyle?: "stop" | "warning" | "information";
  promptTitle?: string;
  prompt?: string;
  errorTitle?: string;
  error?: string;
  formula1?: string | number | Date | string[];
  formula2?: string | number | Date;
}

interface CellRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}
```

**Notes:**

- `list` supports either a formula/range string (for example `"Sheet2!A1:A10"` or `"=$A$1:$A$10"`) or inline string arrays like `["Low", "Medium", "High"]`.
- `whole`, `decimal`, `date`, `time`, and `textLength` typically use `formula1` and optional `formula2` with an `operator`.
- `custom` uses `formula1` as the validation formula. You may include a leading `=`, but it is not required.
- All ranges are 0-based in the API and are written as Excel A1 references internally.

### ConditionalFormatting

```typescript
interface ConditionalFormatting {
  range: CellRange | CellRange[];    // Target range(s), 0-indexed
  rules: ConditionalFormattingRule[];
}

type ConditionalFormattingRule =
  | ConditionalFormatCellRule
  | ConditionalFormatExpressionRule
  | ConditionalFormatColorScaleRule
  | ConditionalFormatDataBarRule
  | ConditionalFormatIconSetRule;

interface ConditionalFormatCellRule {
  type: "cellIs";
  operator: "between" | "notBetween" | "equal" | "notEqual"
          | "greaterThan" | "lessThan" | "greaterThanOrEqual" | "lessThanOrEqual";
  formula1: string | number | Date;
  formula2?: string | number | Date;
  style: CellStyle;
  priority?: number;
  stopIfTrue?: boolean;
}

interface ConditionalFormatExpressionRule {
  type: "expression";
  formula: string;
  style: CellStyle;
  priority?: number;
  stopIfTrue?: boolean;
}

interface ConditionalFormatColorScaleRule {
  type: "colorScale";
  thresholds: [
    ConditionalFormatThreshold,
    ConditionalFormatThreshold,
    ConditionalFormatThreshold?
  ];
  colors: [string, string, string?]; // Hex colors without #
  priority?: number;
}

interface ConditionalFormatDataBarRule {
  type: "dataBar";
  color: string;                     // Hex color without #
  min?: ConditionalFormatThreshold;
  max?: ConditionalFormatThreshold;
  showValue?: boolean;
  priority?: number;
}

interface ConditionalFormatIconSetRule {
  type: "iconSet";
  iconSet: "3Arrows" | "3ArrowsGray" | "3Flags" | "3TrafficLights1" | "3TrafficLights2"
         | "3Signs" | "3Symbols" | "3Symbols2" | "4Arrows" | "4ArrowsGray"
         | "4RedToBlack" | "4Rating" | "4TrafficLights" | "5Arrows" | "5ArrowsGray"
         | "5Rating" | "5Quarters";
  thresholds?: ConditionalFormatThreshold[];
  showValue?: boolean;
  reverse?: boolean;
  priority?: number;
}

interface ConditionalFormatThreshold {
  type: "min" | "max" | "num" | "percent" | "percentile" | "formula";
  value?: string | number;
}
```

**Notes:**

- `cellIs` applies style-driven highlight rules such as `>`, `<`, or `between`.
- `expression` evaluates a custom Excel formula against the top-left cell of the target range.
- `colorScale`, `dataBar`, and `iconSet` use Excel's built-in visual rules and do not require a `style`.
- `priority` follows Excel's rule order. Lower numbers are evaluated first.
- All ranges are 0-based in the API and are written as Excel A1 references internally.

---

## Styles

### CellStyle

```typescript
interface CellStyle {
  font?: FontStyle;
  fill?: FillStyle;
  border?: BorderStyle;
  alignment?: AlignmentStyle;
  numberFormat?: string;       // e.g. "#,##0.00", "yyyy-mm-dd"
}
```

### FontStyle

```typescript
interface FontStyle {
  name?: string;       // Font name, e.g. "Arial", "Calibri"
  size?: number;       // Font size in points
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;    // Strikethrough
  color?: string;      // Hex color without #, e.g. "FF0000" for red
}
```

**Example:**

```typescript
// Bold red header
{ font: { bold: true, color: "FF0000", size: 14 } }

// Italic with specific font
{ font: { italic: true, name: "Times New Roman" } }
```

### FillStyle

```typescript
interface FillStyle {
  type: "pattern" | "gradient";
  pattern?: "solid" | "darkGray" | "mediumGray" | "lightGray" | "none";
  fgColor?: string;    // Foreground hex color
  bgColor?: string;    // Background hex color
}
```

**Example:**

```typescript
// Yellow background
{ fill: { type: "pattern", pattern: "solid", fgColor: "FFFF00" } }

// Light gray
{ fill: { type: "pattern", pattern: "solid", fgColor: "D9D9D9" } }
```

### BorderStyle

```typescript
interface BorderStyle {
  top?: BorderEdgeStyle;
  bottom?: BorderEdgeStyle;
  left?: BorderEdgeStyle;
  right?: BorderEdgeStyle;
}

interface BorderEdgeStyle {
  style?: "thin" | "medium" | "thick" | "dotted" | "dashed" | "double"
        | "hair" | "dashDot" | "dashDotDot" | "mediumDashed"
        | "mediumDashDot" | "mediumDashDotDot" | "slantDashDot";
  color?: string;      // Hex color
}
```

**Example:**

```typescript
// Full thin border
{
  border: {
    top: { style: "thin", color: "000000" },
    bottom: { style: "thin", color: "000000" },
    left: { style: "thin", color: "000000" },
    right: { style: "thin", color: "000000" },
  }
}

// Bottom-only thick red border
{
  border: {
    bottom: { style: "thick", color: "FF0000" },
  }
}
```

### AlignmentStyle

```typescript
interface AlignmentStyle {
  horizontal?: "left" | "center" | "right" | "fill" | "justify";
  vertical?: "top" | "center" | "bottom";
  wrapText?: boolean;       // Wrap text in cell
  textRotation?: number;    // Text rotation in degrees (0-180)
  indent?: number;          // Indent level
}
```

**Example:**

```typescript
// Centered with wrap
{ alignment: { horizontal: "center", vertical: "center", wrapText: true } }

// Right-aligned with indent
{ alignment: { horizontal: "right", indent: 2 } }
```

### Number Formats

Set `numberFormat` on CellStyle to control how numbers and dates display.

**Common formats:**

| Format | Example Output | Description |
|--------|---------------|-------------|
| `"0"` | `1234` | Integer |
| `"0.00"` | `1234.56` | 2 decimal places |
| `"#,##0"` | `1,234` | Thousands separator |
| `"#,##0.00"` | `1,234.56` | Thousands + decimals |
| `"0%"` | `12%` | Percentage |
| `"0.00%"` | `12.34%` | Percentage with decimals |
| `"$#,##0.00"` | `$1,234.56` | Currency |
| `"yyyy-mm-dd"` | `2024-01-15` | Date (ISO) |
| `"dd/mm/yyyy"` | `15/01/2024` | Date (EU) |
| `"mm/dd/yyyy"` | `01/15/2024` | Date (US) |
| `"hh:mm:ss"` | `14:30:00` | Time |
| `"yyyy-mm-dd hh:mm"` | `2024-01-15 14:30` | Date + time |

**Example:**

```typescript
// Currency cell
{ value: 1234.56, style: { numberFormat: "$#,##0.00" } }

// Percentage
{ value: 0.1234, style: { numberFormat: "0.00%" } }

// Date
{ value: new Date("2024-01-15"), style: { numberFormat: "yyyy-mm-dd" } }
```

---

## Features

### Formulas

Write formulas with optional cached results. The cached result is shown before Excel recalculates.

```typescript
{
  cells: [
    { value: 10 },
    { value: 20 },
    { value: 30 },
    {
      value: null,
      formula: "SUM(A1:C1)",        // Formula without leading "="
      formulaResult: 60,             // Cached result
      style: { numberFormat: "#,##0" },
    },
  ],
}
```

All standard Excel functions are supported: SUM, AVERAGE, IF, MAX, MIN, COUNT, VLOOKUP, HLOOKUP, INDEX, MATCH, CONCATENATE, LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER, DATE, TODAY, NOW, ROUND, ABS, and more.

---

### Hyperlinks

Three types of hyperlinks:

```typescript
// 1. External URL
{
  value: "Visit Website",
  hyperlink: { target: "https://example.com", tooltip: "Click to open" },
  style: { font: { color: "0563C1", underline: true } },
}

// 2. Email
{
  value: "Contact Us",
  hyperlink: { target: "mailto:hello@example.com" },
}

// 3. Internal sheet reference
{
  value: "Go to Summary",
  hyperlink: { target: "Sheet2!A1" },
}
```

---

### Merge Cells

Merge a range of cells. Only the top-left cell's value and style are visible.

```typescript
const worksheet: Worksheet = {
  name: "Report",
  rows: [
    // Row 0: merged title spanning A1:F1
    {
      cells: [
        { value: "Annual Report 2024", style: { font: { bold: true, size: 16 } } },
        { value: null }, { value: null }, { value: null }, { value: null }, { value: null },
      ],
    },
    // Row 1: headers
    { cells: [{ value: "Month" }, { value: "Revenue" }, /* ... */] },
  ],
  mergeCells: [
    { startRow: 0, startCol: 0, endRow: 0, endCol: 5 },  // A1:F1
  ],
};
```

> Note: All row/column indices are 0-based.

---

### Freeze Panes

Freeze rows and/or columns so they stay visible when scrolling.

```typescript
const worksheet: Worksheet = {
  name: "Data",
  rows: [/* ... */],
  freezePane: { row: 1, col: 0 },  // Freeze first row (header)
};
```

| Value | Effect |
|-------|--------|
| `{ row: 1, col: 0 }` | Freeze first row |
| `{ row: 0, col: 1 }` | Freeze first column |
| `{ row: 1, col: 1 }` | Freeze first row and first column |
| `{ row: 2, col: 0 }` | Freeze top 2 rows |

---

### Data Validation

Use worksheet-level `dataValidations` to enforce dropdown lists, numeric ranges, date windows, or custom formulas in Excel.

```typescript
const worksheet: Worksheet = {
  name: "Validated",
  rows: [
    { cells: [{ value: "Priority" }, { value: "Score" }, { value: "Due Date" }] },
    { cells: [{ value: null }, { value: null }, { value: null }] },
  ],
  dataValidations: [
    {
      range: { startRow: 1, startCol: 0, endRow: 100, endCol: 0 },
      type: "list",
      formula1: ["Low", "Medium", "High"],
      allowBlank: true,
    },
    {
      range: { startRow: 1, startCol: 1, endRow: 100, endCol: 1 },
      type: "whole",
      operator: "between",
      formula1: 1,
      formula2: 10,
      errorTitle: "Invalid score",
      error: "Score must be between 1 and 10",
    },
    {
      range: { startRow: 1, startCol: 2, endRow: 100, endCol: 2 },
      type: "date",
      operator: "between",
      formula1: new Date("2026-01-01"),
      formula2: new Date("2026-12-31"),
    },
  ],
};
```

**Custom formula example:**

```typescript
{
  range: { startRow: 1, startCol: 0, endRow: 100, endCol: 0 },
  type: "custom",
  formula1: "COUNTIF($A:$A,A2)=1",
  promptTitle: "Unique value",
  prompt: "Each value in column A must be unique",
}
```

**Common patterns:**

```typescript
// Dropdown from inline values
{ type: "list", range, formula1: ["New", "In Progress", "Done"] }

// Dropdown from another sheet/range
{ type: "list", range, formula1: "Lookup!$A$1:$A$20" }

// Decimal greater than zero
{ type: "decimal", range, operator: "greaterThan", formula1: 0 }

// Text length limit
{ type: "textLength", range, operator: "lessThanOrEqual", formula1: 20 }
```

---

### Conditional Formatting

Use worksheet-level `conditionalFormattings` to highlight cells, apply color scales, render data bars, or show icon sets.

```typescript
const worksheet: Worksheet = {
  name: "Dashboard",
  rows: [
    { cells: [{ value: "Score" }, { value: "Trend" }, { value: "Variance" }] },
    { cells: [{ value: 92 }, { value: 0.92 }, { value: 12 }] },
    { cells: [{ value: 68 }, { value: 0.54 }, { value: -8 }] },
  ],
  conditionalFormattings: [
    {
      range: { startRow: 1, startCol: 0, endRow: 100, endCol: 0 },
      rules: [{
        type: "cellIs",
        operator: "greaterThanOrEqual",
        formula1: 90,
        style: {
          fill: { type: "pattern", pattern: "solid", fgColor: "C6EFCE" },
          font: { color: "006100", bold: true },
        },
      }],
    },
    {
      range: { startRow: 1, startCol: 1, endRow: 100, endCol: 1 },
      rules: [{
        type: "dataBar",
        color: "5B9BD5",
      }],
    },
    {
      range: { startRow: 1, startCol: 2, endRow: 100, endCol: 2 },
      rules: [{
        type: "iconSet",
        iconSet: "3Arrows",
      }],
    },
  ],
};
```

**Common patterns:**

```typescript
// Highlight negative values
{
  range,
  rules: [{
    type: "cellIs",
    operator: "lessThan",
    formula1: 0,
    style: {
      font: { color: "9C0006" },
      fill: { type: "pattern", pattern: "solid", fgColor: "FFC7CE" },
    },
  }],
}

// Use a custom formula
{
  range,
  rules: [{
    type: "expression",
    formula: "MOD(ROW(),2)=0",
    style: {
      fill: { type: "pattern", pattern: "solid", fgColor: "F2F2F2" },
    },
  }],
}

// Apply a 3-color scale
{
  range,
  rules: [{
    type: "colorScale",
    thresholds: [
      { type: "min" },
      { type: "percentile", value: 50 },
      { type: "max" },
    ],
    colors: ["F8696B", "FFEB84", "63BE7B"],
  }],
}
```

---

## Writing Modes Comparison

| Feature | `writeExcel` | `createExcelStream` | `createChunkedExcelStream` |
|---------|-------------|--------------------|-----------------------------|
| Memory | Entire workbook in RAM | Row XML buffers in RAM | Constant (~low) |
| Shared Strings | Yes | Yes | No (inline) |
| Multiple Sheets | Yes | Single sheet | Single sheet |
| Multi-sheet | Via Workbook | `createMultiSheetExcelStream` | Not supported |
| Styles | Full support | Full support | Full support |
| Formulas | Full support | Full support | Full support |
| Hyperlinks | Full support | Full support | Full support |
| Merge Cells | Full support | Full support | Full support |
| Freeze Panes | Full support | Full support | Full support |
| Data Validation | Full support | Full support | Full support |
| Conditional Formatting | Full support | Full support | Full support |
| Best For | Small-medium files | Medium-large files | Very large files (100K+) |

**When to use which:**

- **`writeExcel`** -- You have all data ready in memory. Simplest API.
- **`createExcelStream`** -- Data is generated row-by-row (e.g., from database query). Good balance of features and memory.
- **`createChunkedExcelStream`** -- Extreme large files where memory is a concern. Trades some disk I/O for constant memory usage.
