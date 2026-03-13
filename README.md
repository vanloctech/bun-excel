# bun-spreadsheet

High-performance, Bun-optimized library for reading and writing Excel (.xlsx) and CSV files with streaming, styling, formulas, and hyperlinks.

## Features

- **Bun-native** — Built on `Bun.file()`, `Bun.write()`, and `FileSink` for maximum performance
- **Excel (.xlsx)** — Full read/write with styles, formulas, hyperlinks, merge cells, freeze panes
- **CSV** — Read/write with auto-type detection and streaming support
- **3 write modes** — Normal, streaming (memory-efficient), chunked streaming (constant memory)
- **Rich styling** — Fonts, fills, borders, alignment, number formats
- **Hyperlinks** — External URLs, email, internal sheet references
- **Formulas** — Read/write with cached results (SUM, AVERAGE, IF, etc.)
- **Security hardened** — XML bomb prevention, path traversal protection, input validation
- **Minimal deps** — Only [fflate](https://github.com/101arrowz/fflate) for ZIP compression

## Install

```bash
bun add bun-spreadsheet
```

## Quick Start

### Write Excel

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

await writeExcel("report.xlsx", workbook);
```

### Read Excel

```typescript
import { readExcel } from "bun-spreadsheet";

const workbook = await readExcel("report.xlsx");
for (const sheet of workbook.worksheets) {
  console.log(`Sheet: ${sheet.name}`);
  for (const row of sheet.rows) {
    console.log(row.cells.map(c => c.value).join(" | "));
  }
}
```

### CSV

```typescript
import { readCSV, writeCSV } from "bun-spreadsheet";

// Write
await writeCSV("data.csv", [
  [{ value: "Name" }, { value: "Age" }],
  [{ value: "Alice" }, { value: 28 }],
]);

// Read
const csv = await readCSV("data.csv");
```

## API Reference

### Excel — Read/Write

#### `readExcel(path, options?): Promise<Workbook>`

Read an .xlsx file into a Workbook object.

```typescript
const workbook = await readExcel("file.xlsx", {
  sheets: ["Sheet1"],     // Optional: specific sheets to read
  includeStyles: true,    // Optional: include cell styles (default: true)
});
```

#### `writeExcel(path, workbook, options?): Promise<void>`

Write a Workbook to an .xlsx file.

```typescript
await writeExcel("output.xlsx", workbook, {
  creator: "My App",
  compress: true,  // ZIP compression (default: true)
});
```

### Styles

```typescript
const style: CellStyle = {
  font: {
    name: "Arial",
    size: 12,
    bold: true,
    italic: false,
    underline: false,
    strike: false,
    color: "FF0000",  // hex color
  },
  fill: {
    type: "pattern",
    pattern: "solid",
    fgColor: "FFFF00",
  },
  border: {
    top: { style: "thin", color: "000000" },
    bottom: { style: "medium", color: "000000" },
    left: { style: "thin", color: "000000" },
    right: { style: "thin", color: "000000" },
  },
  alignment: {
    horizontal: "center",  // left | center | right | fill | justify
    vertical: "center",    // top | center | bottom
    wrapText: true,
    textRotation: 45,
    indent: 1,
  },
  numberFormat: "#,##0.00",
};
```

### Formulas

```typescript
{
  cells: [
    {
      value: null,
      formula: "SUM(A1:A10)",
      formulaResult: 550,  // cached result (shown before recalculation)
      style: { numberFormat: "#,##0" },
    },
  ],
}
```

Supported Excel functions: SUM, AVERAGE, IF, MAX, MIN, COUNT, VLOOKUP, and all other standard Excel formulas.

### Hyperlinks

```typescript
{
  cells: [
    // External URL
    {
      value: "Visit Website",
      hyperlink: { target: "https://example.com", tooltip: "Click to open" },
      style: { font: { color: "0563C1", underline: true } },
    },
    // Email
    {
      value: "Contact Us",
      hyperlink: { target: "mailto:hello@example.com" },
    },
    // Internal sheet reference
    {
      value: "Go to Summary",
      hyperlink: { target: "Sheet2!A1" },
    },
  ],
}
```

### Merge Cells & Freeze Panes

```typescript
const worksheet: Worksheet = {
  name: "Report",
  rows: [...],
  mergeCells: [
    { startRow: 0, startCol: 0, endRow: 0, endCol: 5 },  // Merge A1:F1
  ],
  freezePane: { row: 1, col: 0 },  // Freeze first row
};
```

### CSV — Read/Write

#### `readCSV(path, options?): Promise<Workbook>`

```typescript
const workbook = await readCSV("data.csv", {
  delimiter: ",",
  hasHeader: true,
  skipEmptyLines: true,
});
```

#### `writeCSV(path, rows, options?): Promise<void>`

```typescript
await writeCSV("output.csv", rows, {
  delimiter: ",",
  includeHeader: true,
  headers: ["Name", "Age", "City"],
  bom: true,  // UTF-8 BOM for Excel compatibility
});
```

#### `readCSVStream(path, options?): AsyncGenerator`

Stream large CSV files row by row:

```typescript
import { readCSVStream } from "bun-spreadsheet";

for await (const row of readCSVStream("large.csv")) {
  // Process each row without loading entire file
}
```

#### `createCSVStream(path, options?): CSVStreamWriter`

Stream write CSV files:

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

## Streaming Excel

### `createExcelStream(path, options?)` — Memory-Efficient

Serializes each row to XML immediately. Good for most use cases.

```typescript
import { createExcelStream } from "bun-spreadsheet";

const stream = createExcelStream("report.xlsx", {
  sheetName: "Data",
  columns: [{ width: 10 }, { width: 25 }],
  freezePane: { row: 1, col: 0 },
});

stream.writeRow({
  cells: [
    { value: "ID", style: { font: { bold: true } } },
    { value: "Name", style: { font: { bold: true } } },
  ],
});

for (let i = 0; i < 50000; i++) {
  stream.writeRow([i + 1, `Product_${i}`]);
}

await stream.end();
```

### `createChunkedExcelStream(path, options?)` — Constant Memory

Writes row XML to a temp file on disk. Memory stays constant regardless of row count. Best for very large files (100K+ rows).

```typescript
import { createChunkedExcelStream } from "bun-spreadsheet";

const stream = createChunkedExcelStream("huge_report.xlsx", {
  sheetName: "Report",
  columns: [{ width: 14 }, { width: 20 }],
  freezePane: { row: 1, col: 0 },
  mergeCells: [
    { startRow: 0, startCol: 0, endRow: 0, endCol: 5 },
  ],
});

// Each row is serialized and written to disk immediately
for (let i = 0; i < 1_000_000; i++) {
  stream.writeRow([i, `Row ${i}`, Math.random() * 10000]);
}

await stream.end(); // Assembles ZIP from temp file
```

### `createMultiSheetExcelStream(path, options?)`

```typescript
import { createMultiSheetExcelStream } from "bun-spreadsheet";

const stream = createMultiSheetExcelStream("multi.xlsx");

stream.addSheet("Revenue", { columns: [{ width: 15 }] });
stream.writeRow(["January"]);
stream.writeRow(["February"]);

stream.addSheet("Expenses", { columns: [{ width: 15 }] });
stream.writeRow(["Salaries"]);

await stream.end();
```

## Benchmarks

30 columns × 30,000 rows with styles and formulas (MacOS, Bun 1.3):

| Method | Total Time | Heap | File Size |
|--------|-----------|------|-----------|
| `writeExcel` | ~1,500ms | ~129MB | 6.85MB |
| `createExcelStream` | ~1,460ms | ~163MB | 6.85MB |
| `createChunkedExcelStream` | ~1,325ms | ~75MB | 6.63MB |

Run benchmarks yourself:

```bash
bun run benchmark
```

## Examples

```bash
# Run all examples
bun run demo

# Large report (30 col × 30K rows)
bun run large-report

# Benchmarks
bun run benchmark
```

## Type Reference

```typescript
interface Workbook {
  worksheets: Worksheet[];
  creator?: string;
  created?: Date;
  modified?: Date;
}

interface Worksheet {
  name: string;
  rows: Row[];
  columns?: ColumnConfig[];
  mergeCells?: MergeCell[];
  freezePane?: { row: number; col: number };
}

interface Cell {
  value: string | number | boolean | Date | null | undefined;
  style?: CellStyle;
  type?: "string" | "number" | "boolean" | "date" | "formula";
  formula?: string;
  formulaResult?: string | number | boolean;
  hyperlink?: { target: string; tooltip?: string };
}
```

## Security

This library is security-hardened:

- **XML bomb prevention** — Depth limits, node count caps, input size validation
- **Path traversal protection** — `path.resolve()` + null byte checks on all file paths
- **Zip slip prevention** — Validates all paths within ZIP archives
- **Input validation** — Max rows (1M), max columns (16K), max file size (200MB)
- **XML injection prevention** — All user values properly escaped
- **Prototype pollution prevention** — `Object.create(null)` for dynamic maps

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for development setup and guidelines.

## License

[MIT](LICENSE)
