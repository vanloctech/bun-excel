# API Documentation

Complete API reference for bun-excel.

---

## Table of Contents

- [Excel](#excel)
  - [writeExcel](#writeexceltarget-workbook-options)
  - [readExcel](#readexcelsource-options)
  - [readExcelStream](#readexcelstreamsource-options)
  - [exportExcelRows](#exportexcelrowsoptions)
  - [exportMultiSheetExcel](#exportmultisheetexceloptions)
  - [buildExcelResponse](#buildexcelresponseworkbook-options)
  - [exportExcelRowsToResponse](#exportexcelrowstoresponseoptions)
  - [exportMultiSheetExcelToResponse](#exportmultisheetexceltoresponseoptions)
  - [buildExcelBuffer](#buildexcelbufferworkbook-options)
  - [loadExcelTemplate](#loadexceltemplatesource-options)
- [CSV](#csv)
  - [writeCSV](#writecsvtarget-data-options)
  - [readCSV](#readcsvsource-options)
  - [readCSVStream](#readcsvstreamsource-options)
  - [createCSVStream](#createcsvstreamtarget-options)
- [Excel Streaming](#excel-streaming)
  - [createExcelStream](#createexcelstreamtarget-options)
  - [createMultiSheetExcelStream](#createmultisheetexcelstreamtarget-options)
  - [createChunkedExcelStream](#createchunkedexcelstreamtarget-options)
- [Types](#types)
  - [FileSource](#filesource)
  - [FileTarget](#filetarget)
  - [Workbook](#workbook)
  - [Worksheet](#worksheet)
  - [Row](#row)
  - [ExcelReadStreamRow](#excelreadstreamrow)
  - [ExcelExportProgress](#excelexportprogress)
  - [ExcelExportDiagnostics](#excelexportdiagnostics)
  - [Cell](#cell)
  - [CellValue](#cellvalue)
  - [CellComment](#cellcomment)
  - [BinaryData](#binarydata)
  - [ColumnConfig](#columnconfig)
  - [MergeCell](#mergecell)
  - [SplitPane](#splitpane)
  - [Hyperlink](#hyperlink)
  - [DataValidation](#datavalidation)
  - [ConditionalFormatting](#conditionalformatting)
  - [WorksheetImage](#worksheetimage)
  - [WorksheetTable](#worksheettable)
- [Styles](#styles)
  - [CellStyle](#cellstyle)
  - [FontStyle](#fontstyle)
  - [FillStyle](#fillstyle)
  - [BorderStyle](#borderstyle)
  - [AlignmentStyle](#alignmentstyle)
  - [Number Formats](#number-formats)
- [Features](#features)
  - [Formulas](#formulas)
  - [Hyperlinks](#hyperlinks)
  - [Merge Cells](#merge-cells)
  - [Auto Filters](#auto-filters)
  - [Freeze Panes](#freeze-panes)
  - [Split Views](#split-views)
  - [Cell Comments / Notes](#cell-comments--notes)
  - [Images](#images)
  - [Tables](#tables)
  - [Data Validation](#data-validation)
  - [Conditional Formatting](#conditional-formatting)
- [Writing Modes Comparison](#writing-modes-comparison)

---

## Bun Runtime Targets

For Bun-native workflows, most read/write APIs accept both local files and Bun runtime file objects:

- `FileSource` = `string | Bun.BunFile | Bun.S3File`
- `FileTarget` = `string | Bun.BunFile | Bun.S3File`

This means you can:

- read from local paths like `"./report.xlsx"`
- read from `Bun.file("./report.xlsx")`
- read from S3 via `new Bun.S3Client().file("reports/report.xlsx")`
- write or stream directly to an S3 object target using Bun's `S3File.writer()` path

**Example:**

```typescript
import {
  createChunkedExcelStream,
  readExcel,
  writeExcel,
} from "bun-excel";

const s3 = new Bun.S3Client();
const remoteFile = s3.file("reports/monthly.xlsx");

await writeExcel(remoteFile, workbook);

const workbookFromS3 = await readExcel(remoteFile);

const stream = createChunkedExcelStream(remoteFile, {
  sheetName: "Report",
});
stream.writeRow(["ID", "Value"]);
await stream.end();
```

---

## Excel

### `writeExcel(target, workbook, options?)`

Write a Workbook to an `.xlsx` file.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `target` | `FileTarget` | Yes | Output target: local path, `Bun.file(...)`, or `S3File` |
| `workbook` | `Workbook` | Yes | Workbook data to write |
| `options` | `ExcelWriteOptions` | No | Write options |

**ExcelWriteOptions:**

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `creator` | `string` | `undefined` | Author name in file metadata |
| `created` | `Date` | `undefined` | Created timestamp in workbook metadata |
| `modified` | `Date` | `undefined` | Modified timestamp in workbook metadata |
| `compress` | `boolean` | `true` | Enable ZIP compression |
| `s3WriterOptions` | `S3WriterOptions` | `undefined` | Passed through to Bun `S3File.writer()` for multipart upload tuning |

**Returns:** `Promise<void>`

**Example:**

```typescript
import { writeExcel, type Workbook } from "bun-excel";

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

const s3 = new Bun.S3Client();
await writeExcel(s3.file("exports/report.xlsx"), workbook, {
  creator: "My App",
});
```

---

### `readExcel(source, options?)`

Read an `.xlsx` file into a Workbook object.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `source` | `FileSource` | Yes | Input source: local path, `Bun.file(...)`, or `S3File` |
| `options` | `ExcelReadOptions` | No | Read options |

**ExcelReadOptions:**

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `sheets` | `string[] \| number[]` | all sheets | Specific sheets to read (by name or index) |
| `includeStyles` | `boolean` | `true` | Whether to parse and include cell styles |

**Returns:** `Promise<Workbook>`

**Example:**

```typescript
import { readExcel } from "bun-excel";

// Read all sheets
const workbook = await readExcel("report.xlsx");

// Read from Bun.file(...)
const fromLocalBlob = await readExcel(Bun.file("./report.xlsx"));

// Read from S3
const s3 = new Bun.S3Client();
const fromS3 = await readExcel(s3.file("reports/report.xlsx"));

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

### `readExcelStream(source, options?)`

Read an `.xlsx` file as an async row stream instead of materializing the full workbook in memory.

This API is designed for large worksheet reads on Bun runtime paths:

- local files
- `Bun.file(...)`
- `Bun.S3File`

It streams the ZIP container with Bun-native streams, spools worksheet XML to temp files, and yields rows one by one.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `source` | `FileSource` | Yes | Input source: local path, `Bun.file(...)`, or `S3File` |
| `options` | `ExcelReadOptions` | No | Same sheet/style options as `readExcel()` |

**Returns:** `AsyncGenerator<ExcelReadStreamRow>`

**ExcelReadStreamRow:**

| Field | Type | Description |
|------|------|-------------|
| `sheetIndex` | `number` | Zero-based sheet index in workbook order |
| `sheetName` | `string` | Worksheet name |
| `rowIndex` | `number` | Zero-based row index |
| `row` | `Row` | Parsed row payload |

**Current scope:**

- supports `sharedStrings`
- supports style-based date detection when `includeStyles !== false`
- returns formulas with cached values when present
- supports `sheets` filtering by name or index
- optimized for row-by-row processing of large worksheets

**Current limitations:**

- it does not emit workbook/sheet metadata events
- it does not currently stream worksheet comments, images, tables, or hyperlinks
- it is intended for data extraction workloads, not full workbook round-trip fidelity

**Example:**

```typescript
import { readExcelStream } from "bun-excel";

for await (const entry of readExcelStream("report.xlsx", {
  sheets: ["Orders"],
})) {
  const values = entry.row.cells.map((cell) => cell.value);
  console.log(entry.sheetName, entry.rowIndex, values);
}

const s3 = new Bun.S3Client();
for await (const entry of readExcelStream(s3.file("reports/big.xlsx"))) {
  // Process each row immediately without loading the full workbook
  console.log(entry.row.cells[0]?.value);
}
```

Use `readExcel()` when you need the full workbook object with worksheet features like images, comments, tables, validations, and formatting metadata. Use `readExcelStream()` when you need lower memory row-by-row reads.

---

### `exportExcelRows(options)`

Production-oriented Excel export helper for a single worksheet.

This helper wraps the streaming writers with:

- `AbortSignal` support
- progress callbacks
- output diagnostics
- direct local / `Bun.file(...)` / `S3File` targets

**Key options:**

| Option | Type | Description |
|--------|------|-------------|
| `target` | `FileTarget` | Output target |
| `rows` | `Iterable<Row \| CellValue[]> \| AsyncIterable<Row \| CellValue[]>` | Row source |
| `mode` | `"stream" \| "chunked"` | Writer mode |
| `sheetName` | `string` | Worksheet name |
| `signal` | `AbortSignal` | Cancel the export |
| `progressIntervalRows` | `number` | Emit progress every N rows |
| `onProgress` | `(progress) => void \| Promise<void>` | Progress callback |

**Returns:** `Promise<ExcelExportDiagnostics>`

**Example:**

```typescript
import { exportExcelRows } from "bun-excel";

const result = await exportExcelRows({
  target: "orders.xlsx",
  sheetName: "Orders",
  mode: "chunked",
  rows,
  signal,
  onProgress(progress) {
    console.log(progress.stage, progress.rowsWritten);
  },
});
```

Use this when you are exporting generated data from a job worker, HTTP handler, or queue consumer.

---

### `exportMultiSheetExcel(options)`

Production-oriented Excel export helper for multi-sheet workbooks.

It uses the disk-backed multi-sheet stream writer and adds the same production hooks as `exportExcelRows()`.

**Key options:**

| Option | Type | Description |
|--------|------|-------------|
| `target` | `FileTarget` | Output target |
| `sheets` | `ExportExcelSheetRows[]` | Sheet definitions and row sources |
| `signal` | `AbortSignal` | Cancel the export |
| `progressIntervalRows` | `number` | Emit progress every N rows across all sheets |
| `onProgress` | `(progress) => void \| Promise<void>` | Progress callback |
| `creator` / `created` / `modified` | workbook metadata | Workbook metadata |

**Returns:** `Promise<ExcelExportDiagnostics>`

**Example:**

```typescript
import { exportMultiSheetExcel } from "bun-excel";

await exportMultiSheetExcel({
  target: "report.xlsx",
  sheets: [
    {
      name: "Orders",
      rows: orderRows,
    },
    {
      name: "Summary",
      rows: summaryRows,
    },
  ],
  onProgress(progress) {
    console.log(progress.sheetName, progress.rowsWritten);
  },
});
```

---

### `buildExcelResponse(workbook, options?)`

Build a Bun/Fetch `Response` object for Excel downloads.

This is useful for HTTP handlers that want to return an `.xlsx` file directly.

**Options:**

| Option | Type | Description |
|--------|------|-------------|
| `filename` | `string` | Sets `Content-Disposition` |
| `headers` | `HeadersInit` | Additional response headers |

**Returns:** `Promise<Response>`

**Example:**

```typescript
import { buildExcelResponse } from "bun-excel";

return await buildExcelResponse(workbook, {
  filename: "report.xlsx",
});
```

---

### `exportExcelRowsToResponse(options)`

Export a single-sheet workbook to a streaming `Response`.

Unlike `buildExcelResponse()`, this path writes to a temp `.xlsx` file first and then streams that file back as the response body. This is a better fit for large exports because your HTTP handler does not need to keep the full workbook buffer in memory.

**Returns:** `Promise<{ response: Response; diagnostics: ExcelExportDiagnostics }>`

**Example:**

```typescript
import { exportExcelRowsToResponse } from "bun-excel";

const { response, diagnostics } = await exportExcelRowsToResponse({
  filename: "orders.xlsx",
  sheetName: "Orders",
  mode: "chunked",
  rows,
});

console.log(diagnostics.rowsWritten);
return response;
```

---

### `exportMultiSheetExcelToResponse(options)`

Export a multi-sheet workbook to a streaming `Response`.

This uses the same temp-file-to-response flow as `exportExcelRowsToResponse()`, but for multi-sheet workloads.

**Returns:** `Promise<{ response: Response; diagnostics: ExcelExportDiagnostics }>`

**Example:**

```typescript
import { exportMultiSheetExcelToResponse } from "bun-excel";

const { response } = await exportMultiSheetExcelToResponse({
  filename: "report.xlsx",
  sheets: [
    { name: "Orders", rows: orderRows },
    { name: "Summary", rows: summaryRows },
  ],
});

return response;
```

---

### `buildExcelBuffer(workbook, options?)`

Build an Excel file as a `Uint8Array` buffer without writing to disk. Useful for sending as HTTP response or further processing.

**Parameters:** Same as `writeExcel` except no `path`.

**Returns:** `Uint8Array`

**Example:**

```typescript
import { buildExcelBuffer } from "bun-excel";

const buffer = buildExcelBuffer(workbook);
// Use buffer for HTTP response, upload, etc.
```

---

### `loadExcelTemplate(source, options?)`

Load an existing `.xlsx` file as a mutable template, update cells or named ranges, then write it back out.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `source` | `FileSource` | Yes | Template source: local path, `Bun.file(...)`, or `S3File` |
| `options` | `ExcelReadOptions` | No | Read options passed to `readExcel()` |

**Returns:** `Promise<ExcelTemplate>`

**ExcelTemplate methods:**

| Method | Description |
|--------|-------------|
| `getWorksheet(sheetRef)` | Get a worksheet by name or index |
| `getDefinedName(name, scope?)` | Get a workbook/sheet scoped defined name |
| `setCell(sheetRef, ref, valueOrPatch)` | Update one cell by A1 ref or `{ row, col }` |
| `fillRange(sheetRef, startRef, matrix)` | Fill a rectangular range from a 2D array |
| `setDefinedName(name, valueOrMatrix, scope?)` | Write into a named cell or named range |
| `build(options?)` | Build the updated workbook to a `Uint8Array` |
| `write(target, options?)` | Write the updated workbook to a file target |

**Example:**

```typescript
import {
  loadExcelTemplate,
  writeExcel,
  type Workbook,
} from "bun-excel";

const templateWorkbook: Workbook = {
  definedNames: [
    { name: "InvoiceNumber", refersTo: "'Invoice'!$B$2" },
    { name: "LineItems", refersTo: "'Invoice'!$A$6:$B$7" },
  ],
  worksheets: [
    {
      name: "Invoice",
      rows: [
        { cells: [{ value: "Invoice" }] },
        { cells: [{ value: "Invoice #" }, { value: "TBD" }] },
        { cells: [] },
        { cells: [] },
        { cells: [{ value: "Item" }, { value: "Qty" }] },
        { cells: [{ value: "" }, { value: 0 }] },
        { cells: [{ value: "" }, { value: 0 }] },
      ],
    },
  ],
};

await writeExcel("invoice-template.xlsx", templateWorkbook);

const template = await loadExcelTemplate("invoice-template.xlsx");
template.setDefinedName("InvoiceNumber", "INV-001");
template.setDefinedName("LineItems", [
  ["Apple", 2],
  ["Orange", 5],
]);
template.setCell("Invoice", "A1", "Invoice 2026");
await template.write("invoice-filled.xlsx");
```

**Notes:**

- Template mode works best when the source workbook uses features already supported by `readExcel()` and `writeExcel()`.
- Existing styles, comments, images, tables, freeze panes, and other supported worksheet metadata are preserved because the workbook is loaded, mutated, then written back out.
- `setDefinedName()` accepts a scalar for single-cell names and a 2D array for multi-cell named ranges.
- For a multi-cell named range, passing a scalar throws an error to avoid accidental partial writes.

---

## CSV

### `writeCSV(target, data, options?)`

Write data to a CSV file.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `target` | `FileTarget` | Yes | Output target: local path, `Bun.file(...)`, or `S3File` |
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
| `s3WriterOptions` | `S3WriterOptions` | `undefined` | Passed through to Bun `S3File.writer()` when the target is an `S3File` |

**Returns:** `Promise<void>`

**Example:**

```typescript
import { writeCSV } from "bun-excel";

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

const s3 = new Bun.S3Client();
await writeCSV(s3.file("exports/data.csv"), data);
```

---

### `readCSV(source, options?)`

Read a CSV file into a Workbook object (single worksheet).

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `source` | `FileSource` | Yes | Input source: local path, `Bun.file(...)`, or `S3File` |
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
import { readCSV } from "bun-excel";

const workbook = await readCSV("data.csv", {
  hasHeader: true,
  skipEmptyLines: true,
});

const workbookFromFile = await readCSV(Bun.file("./data.csv"));

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

### `readCSVStream(source, options?)`

Stream-read a large CSV file row by row. Returns an `AsyncGenerator`.

**Parameters:** Same as `readCSV`.

**Returns:** `AsyncGenerator<Row>`

**Example:**

```typescript
import { readCSVStream } from "bun-excel";

for await (const row of readCSVStream("large.csv")) {
  const values = row.cells.map(c => c.value);
  // Process each row without loading entire file into memory
}

const s3 = new Bun.S3Client();
for await (const row of readCSVStream(s3.file("imports/large.csv"))) {
  // Stream rows directly from S3
}
```

---

### `createCSVStream(target, options?)`

Create a streaming CSV writer. Writes rows directly to the target.

**Parameters:** Same as `writeCSV`.

**Returns:** `CSVStreamWriter`

**CSVStreamWriter methods:**

| Method | Description |
|--------|-------------|
| `writeRow(values: CellValue[])` | Write a single row |
| `flush()` | Flush buffered output |
| `end(): Promise<void>` | Finalize and close the file |

**Example:**

```typescript
import { createCSVStream } from "bun-excel";

const stream = createCSVStream("output.csv", {
  headers: ["ID", "Name", "Value"],
  includeHeader: true,
});

for (let i = 0; i < 100000; i++) {
  stream.writeRow([i + 1, `Item_${i}`, Math.random() * 1000]);
}

await stream.end();

const s3 = new Bun.S3Client();
const remoteStream = createCSVStream(s3.file("exports/output.csv"), {
  headers: ["ID", "Name"],
  includeHeader: true,
});
remoteStream.writeRow([1, "Alice"]);
await remoteStream.end();
```

---

## Excel Streaming

Three streaming modes for different scenarios:

| Mode | Memory | Best For | Shared Strings |
|------|--------|----------|----------------|
| `createExcelStream` | Low (disk-backed) | Most single-sheet streaming exports | No (inline strings) |
| `createMultiSheetExcelStream` | Low-moderate (disk-backed per sheet) | Multiple sheets | No (inline strings) |
| `createChunkedExcelStream` | Constant (~low) | Very large files (100K+ rows) | No (inline strings) |

---

### `createExcelStream(target, options?)`

Create a streaming Excel writer. Uses disk-backed temp files and inline strings, then finalizes the workbook into the target.

**ExcelStreamOptions:**

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `sheetName` | `string` | `"Sheet1"` | Name of the worksheet |
| `columns` | `ColumnConfig[]` | `undefined` | Column width configurations |
| `defaultRowHeight` | `number` | `15` | Default row height |
| `images` | `WorksheetImage[]` | `undefined` | Worksheet images to embed |
| `tables` | `WorksheetTable[]` | `undefined` | Structured Excel tables |
| `freezePane` | `{ row, col }` | `undefined` | Freeze pane position |
| `splitPane` | `SplitPane` | `undefined` | Split view configuration |
| `mergeCells` | `MergeCell[]` | `undefined` | Merge cell ranges |
| `autoFilter` | `CellRange` | `undefined` | Auto filter range |
| `creator` | `string` | `undefined` | Author name |
| `created` | `Date` | `undefined` | Created timestamp in workbook metadata |
| `modified` | `Date` | `undefined` | Modified timestamp in workbook metadata |
| `compress` | `boolean` | `true` | Enable ZIP compression |

**Returns:** `ExcelStreamWriter`

**ExcelStreamWriter methods:**

| Method | Description |
|--------|-------------|
| `writeRow(row: Row \| CellValue[])` | Write a single row (as Row object or plain array) |
| `flush()` | Flush buffer |
| `end(): Promise<void>` | Finalize ZIP and write/upload the output target |

**Example:**

```typescript
import { createExcelStream } from "bun-excel";

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

const s3 = new Bun.S3Client();
const remoteStream = createExcelStream(s3.file("exports/report.xlsx"), {
  sheetName: "Data",
});
remoteStream.writeRow(["ID", "Name"]);
await remoteStream.end();
```

---

### `createMultiSheetExcelStream(target, options?)`

Create a streaming Excel writer with support for multiple sheets. Each sheet is staged on local temp files and the final workbook is written to the target.

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
| `images` | `WorksheetImage[]` | Images embedded in this sheet |
| `tables` | `WorksheetTable[]` | Structured tables written for this sheet |
| `freezePane` | `{ row, col }` | Freeze pane |
| `splitPane` | `SplitPane` | Split view |
| `mergeCells` | `MergeCell[]` | Merge cell ranges |
| `autoFilter` | `CellRange` | Auto filter range |

**Example:**

```typescript
import { createMultiSheetExcelStream } from "bun-excel";

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

const s3 = new Bun.S3Client();
const remoteMulti = createMultiSheetExcelStream(
  s3.file("exports/multi.xlsx"),
);
remoteMulti.addSheet("Sheet1");
remoteMulti.writeRow(["Hello"]);
await remoteMulti.end();
```

---

### `createChunkedExcelStream(target, options?)`

Create a chunked streaming Excel writer with constant memory usage. Row XML is written to temporary files on disk, then assembled into ZIP and streamed to the target at the end.

**ChunkedExcelStreamOptions:** Same as `ExcelStreamOptions`.

**Returns:** `ExcelChunkedStreamWriter`

**ExcelChunkedStreamWriter methods:**

| Method | Description |
|--------|-------------|
| `writeRow(row: Row \| CellValue[])` | Write a row (serialized to disk immediately) |
| `writeStyledRow(values, styles)` | Write a row with per-cell styles |
| `writeRows(rows)` | Write multiple rows at once |
| `flush()` | Flush temp file buffer |
| `end(): Promise<void>` | Assemble ZIP from temp files and write/upload output |
| `currentRowCount` | Get current row count |

**How it works:**

1. `writeRow()` - serializes row XML and writes to a temp file on disk (no memory retention)
2. `end()` - reads temp file, wraps with worksheet XML, creates ZIP, writes output, deletes temp file

Uses **inline strings** (`<is><t>...</t></is>`) instead of shared string table, so no string tracking is needed in memory.

**Example:**

```typescript
import { createChunkedExcelStream } from "bun-excel";

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

const s3 = new Bun.S3Client();
const remoteChunked = createChunkedExcelStream(
  s3.file("exports/huge_report.xlsx"),
  { sheetName: "Report" },
);
remoteChunked.writeRow(["ID", "Value"]);
await remoteChunked.end();
```

---

## Types

### FileSource

```typescript
type FileSource = string | Bun.BunFile | Bun.S3File
```

Used by read APIs such as `readExcel()`, `readCSV()`, and `readCSVStream()`.

### FileTarget

```typescript
type FileTarget = string | Bun.BunFile | Bun.S3File
```

Used by write APIs such as `writeExcel()`, `writeCSV()`, `createCSVStream()`, `createExcelStream()`, `createMultiSheetExcelStream()`, and `createChunkedExcelStream()`.

### S3WriterOptions

```typescript
type S3WriterOptions = Parameters<Bun.S3File["writer"]>[0];
```

Pass-through options for Bun's `S3File.writer()`. Useful for tuning multipart upload behavior on large exports.

Common fields supported by Bun include:

- `partSize`
- `queueSize`
- `retry`
- `type`
- `contentDisposition`

### Workbook

```typescript
interface Workbook {
  worksheets: Worksheet[];
  creator?: string;      // Author name
  created?: Date;        // Created timestamp
  modified?: Date;       // Modified timestamp
  definedNames?: DefinedName[];
  views?: WorkbookView;
}
```

`creator`, `created`, and `modified` are written into workbook metadata and returned by `readExcel()`.

### Worksheet

```typescript
interface Worksheet {
  name: string;                              // Sheet name
  rows: Row[];                               // Array of rows
  columns?: ColumnConfig[];                  // Column configurations
  mergeCells?: MergeCell[];                  // Merged cell ranges
  autoFilter?: CellRange;                    // Auto filter range
  dataValidations?: DataValidation[];        // Data validation rules
  conditionalFormattings?: ConditionalFormatting[]; // Conditional formatting rules
  freezePane?: { row: number; col: number }; // Freeze pane position
  splitPane?: SplitPane;                     // Split view configuration
  defaultRowHeight?: number;                 // Default row height
  defaultColWidth?: number;                  // Default column width
  images?: WorksheetImage[];                 // Embedded worksheet images
  tables?: WorksheetTable[];                 // Structured tables
}
```

`images` and `tables` are worksheet-level features. Comments/notes are attached per cell via `Cell.comment`.

### Row

```typescript
interface Row {
  cells: Cell[];       // Array of cells in this row
  height?: number;     // Custom row height
  style?: CellStyle;   // Default style for all cells in this row
}
```

### ExcelReadStreamRow

```typescript
interface ExcelReadStreamRow {
  sheetIndex: number;
  sheetName: string;
  rowIndex: number;
  row: Row;
}
```

Returned by `readExcelStream()` for each parsed worksheet row.

### ExcelExportProgress

```typescript
interface ExcelExportProgress {
  stage: "writing" | "finalizing" | "completed" | "aborted";
  mode: "stream" | "chunked" | "multi-sheet";
  rowsWritten: number;
  elapsedMs: number;
  target: string;
  memory: ExcelExportMemorySnapshot;
  sheetName?: string;
  sheetIndex?: number;
}
```

Returned to `onProgress` callbacks in `exportExcelRows()` and `exportMultiSheetExcel()`.

### ExcelExportDiagnostics

```typescript
interface ExcelExportDiagnostics {
  mode: "stream" | "chunked" | "multi-sheet";
  target: string;
  rowsWritten: number;
  startedAt: Date;
  finishedAt: Date;
  durationMs: number;
  outputSizeBytes: number;
  sheetCount?: number;
  memory: {
    baseline: ExcelExportMemorySnapshot;
    peak: ExcelExportMemorySnapshot;
    end: ExcelExportMemorySnapshot;
  };
}
```

Returned after a production export finishes successfully.

### Cell

```typescript
interface Cell {
  value: CellValue;                                // Cell value
  style?: CellStyle;                               // Cell style
  type?: "string" | "number" | "boolean" | "date" | "formula";
  richText?: RichTextRun[];                        // Partial formatting within one cell
  formula?: string;                                // Formula (without "=")
  formulaResult?: string | number | boolean;       // Cached formula result
  hyperlink?: Hyperlink;                           // Hyperlink on this cell
  comment?: CellComment;                           // Cell comment / note
}
```

### CellValue

```typescript
type CellValue = string | number | boolean | Date | null | undefined;
```

### CellComment

```typescript
interface CellComment {
  text: string;
  author?: string;
}
```

### BinaryData

```typescript
type BinaryData = Uint8Array | ArrayBuffer;
```

Used for embedded images. Pass raw bytes that are already available in memory.

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

### SplitPane

```typescript
interface SplitPane {
  x: number;                               // Horizontal split position
  y: number;                               // Vertical split position
  topLeftCell?: { row: number; col: number }; // Optional top-left visible cell
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

### WorksheetImage

```typescript
interface WorksheetImage {
  data: BinaryData;
  format: "png" | "jpeg" | "jpg" | "gif";
  range: CellRange;
  name?: string;
  description?: string;
}
```

**Notes:**

- `range` controls the image anchor rectangle in worksheet coordinates.
- `data` must be image bytes already loaded in memory.
- `readExcel()` returns the embedded image bytes back in `worksheet.images`.

### WorksheetTable

```typescript
interface WorksheetTable {
  name: string;
  displayName?: string;
  range: CellRange;
  headerRow?: boolean;
  totalsRow?: boolean;
  columns?: WorksheetTableColumn[];
  style?: WorksheetTableStyle;
}

interface WorksheetTableColumn {
  name: string;
  totalsRowLabel?: string;
  totalsRowFunction?: "sum" | "average" | "count" | "countNums" | "max" | "min"
    | "stdDev" | "var" | "custom";
}

interface WorksheetTableStyle {
  name?: string;
  showFirstColumn?: boolean;
  showLastColumn?: boolean;
  showRowStripes?: boolean;
  showColumnStripes?: boolean;
}
```

**Notes:**

- `range` includes the header row and totals row if they are enabled.
- `name` must be unique within the workbook.
- In stream and chunked writers, pass `columns` explicitly for predictable table headers.

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

// Simple gradient using first/last stop colors
{ fill: { type: "gradient", fgColor: "FFF2CC", bgColor: "F4B183" } }
```

For `gradient`, `fgColor` is used as the first stop and `bgColor` as the last stop.

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

When reading XLSX files, numeric cells with date/time number formats are automatically returned as `Date`.

---

## Features

### Formulas

Write and read formulas with optional cached results. The cached result is shown before Excel recalculates.

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

### Auto Filters

Use `autoFilter` to enable Excel's dropdown filters on a header/data range.

```typescript
const worksheet: Worksheet = {
  name: "Filtered",
  rows: [
    { cells: [{ value: "Name" }, { value: "Score" }] },
    { cells: [{ value: "Alice" }, { value: 95 }] },
    { cells: [{ value: "Bob" }, { value: 87 }] },
  ],
  autoFilter: { startRow: 0, startCol: 0, endRow: 100, endCol: 1 },
};
```

This writes Excel's `<autoFilter>` range and `readExcel()` returns the same range back.

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

### Split Views

Use `splitPane` when you want scrollable split views instead of frozen headers.

```typescript
const worksheet: Worksheet = {
  name: "Split",
  rows: [/* ... */],
  splitPane: {
    x: 1200,
    y: 1800,
    topLeftCell: { row: 1, col: 1 },
  },
};
```

`freezePane` and `splitPane` are different Excel view modes. If both are provided, `freezePane` takes precedence.

---

### Cell Comments / Notes

Use `cell.comment` to write an Excel comment/note on a specific cell.

```typescript
const worksheet: Worksheet = {
  name: "Comments",
  rows: [
    {
      cells: [
        {
          value: "Status",
          comment: {
            text: "This column is filled by the reviewer",
            author: "Loc",
          },
        },
        { value: "Owner" },
      ],
    },
    {
      cells: [
        {
          value: "Pending",
          comment: { text: "Waiting for approval" },
        },
        { value: "Alice" },
      ],
    },
  ],
};
```

**Notes:**

- `author` is optional. If omitted, Excel still shows the note text.
- Comments are attached to cells, not worksheet-level options.
- In streaming writers, use `Row` objects when you need comments because plain arrays cannot carry metadata.
- `readExcel()` returns comments in `cell.comment`.

---

### Images

Use `worksheet.images` to embed PNG, JPEG, or GIF assets anchored to a cell range.

```typescript
const logoBytes = await Bun.file("./logo.png").bytes();

const worksheet: Worksheet = {
  name: "Dashboard",
  rows: [
    { cells: [{ value: "Sales Dashboard" }] },
    { cells: [{ value: "Q1 2026" }] },
  ],
  images: [
    {
      data: logoBytes,
      format: "png",
      range: { startRow: 0, startCol: 3, endRow: 3, endCol: 5 },
      name: "Company Logo",
      description: "Top-right dashboard logo",
    },
  ],
};
```

**Notes:**

- Supported formats are `png`, `jpeg` / `jpg`, and `gif`.
- `data` must be raw bytes (`Uint8Array` or `ArrayBuffer`), not a file path.
- The image is anchored using `range`; larger ranges produce larger rendered images in Excel.
- `createExcelStream`, `createMultiSheetExcelStream`, and `createChunkedExcelStream` support worksheet images via their sheet options.
- `readExcel()` returns embedded images as `worksheet.images` with their original bytes and anchor range.

---

### Tables

Use `worksheet.tables` to create structured Excel tables with headers, optional totals rows, and built-in table styles.

```typescript
const worksheet: Worksheet = {
  name: "Orders",
  rows: [
    { cells: [{ value: "Order ID" }, { value: "Region" }, { value: "Amount" }] },
    { cells: [{ value: "A-1001" }, { value: "North" }, { value: 1250 }] },
    { cells: [{ value: "A-1002" }, { value: "South" }, { value: 980 }] },
    { cells: [{ value: "A-1003" }, { value: "West" }, { value: 1640 }] },
  ],
  tables: [
    {
      name: "OrdersTable",
      range: { startRow: 0, startCol: 0, endRow: 3, endCol: 2 },
      headerRow: true,
      totalsRow: false,
      columns: [
        { name: "Order ID" },
        { name: "Region" },
        { name: "Amount" },
      ],
      style: {
        name: "TableStyleMedium2",
        showRowStripes: true,
      },
    },
  ],
};
```

**Totals row example:**

```typescript
{
  name: "SalesTable",
  range: { startRow: 0, startCol: 0, endRow: 10, endCol: 2 },
  headerRow: true,
  totalsRow: true,
  columns: [
    { name: "Month" },
    { name: "Revenue", totalsRowFunction: "sum" },
    { name: "Owner", totalsRowLabel: "Total" },
  ],
}
```

**Notes:**

- `range` must cover the entire table area.
- `columns` should match the number of columns in `range`.
- In stream and chunked writers, pass `columns` explicitly. Those writers do not keep all worksheet rows in memory, so explicit headers are the reliable option.
- `readExcel()` returns table definitions in `worksheet.tables`.

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
| Memory | Entire workbook in RAM | Low (disk-backed temp files) | Constant (~low) |
| Shared Strings | Yes | No (inline) | No (inline) |
| Multiple Sheets | Yes | Single sheet | Single sheet |
| Multi-sheet | Via Workbook | `createMultiSheetExcelStream` | Not supported |
| Styles | Full support | Full support | Full support |
| Formulas | Full support | Full support | Full support |
| Workbook properties | Full support | Full support | Full support |
| Hyperlinks | Full support | Full support | Full support |
| Comments / Notes | Full support | Full support | Full support |
| Images | Full support | Full support | Full support |
| Tables | Full support | Full support | Full support |
| Merge Cells | Full support | Full support | Full support |
| Auto Filters | Full support | Full support | Full support |
| Freeze Panes | Full support | Full support | Full support |
| Split Views | Full support | Full support | Full support |
| Data Validation | Full support | Full support | Full support |
| Conditional Formatting | Full support | Full support | Full support |
| Best For | Small-medium files | Medium-large files | Very large files (100K+) |

**When to use which:**

- **`writeExcel`** -- You have all data ready in memory. Simplest API.
- **`createExcelStream`** -- Data is generated row-by-row (e.g., from database query). Disk-backed and suitable for local files or `S3File` targets.
- **`createChunkedExcelStream`** -- Extreme large files where memory is a concern. Trades some disk I/O for constant memory usage.
- For comments, use `Row` objects rather than plain arrays so you can attach `cell.comment`.
- For streaming tables, pass `table.columns` explicitly for stable header metadata.
