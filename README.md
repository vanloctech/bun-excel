# bun-spreadsheet

[![CI](https://github.com/vanloctech/bun-spreadsheet/actions/workflows/ci.yml/badge.svg)](https://github.com/vanloctech/bun-spreadsheet/actions/workflows/ci.yml)
[![npm version](https://img.shields.io/npm/v/bun-spreadsheet.svg)](https://www.npmjs.com/package/bun-spreadsheet)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Bun](https://img.shields.io/badge/Bun-%E2%89%A51.0-black?logo=bun)](https://bun.sh)
[![TypeScript](https://img.shields.io/badge/TypeScript-%E2%89%A55.0-blue?logo=typescript)](https://www.typescriptlang.org/)

[![English](https://img.shields.io/badge/lang-English-blue)](README.md) [![中文](https://img.shields.io/badge/lang-%E4%B8%AD%E6%96%87-red)](README.zh-CN.md)

High-performance, Bun-optimized Excel (.xlsx) and CSV library for TypeScript.

> Runtime note: `bun-spreadsheet` uses Bun-specific APIs such as `Bun.file()`, `Bun.write()`, and `FileSink`. It is intended for Bun and is not compatible with Node.js or Deno.

## Features

- **Bun-native** — Built on `Bun.file()`, `Bun.write()`, and `FileSink` for maximum performance
- **Excel (.xlsx)** — Full read/write with styles, formulas, hyperlinks, merge cells, freeze panes
- **CSV** — Read/write with auto-type detection and streaming support
- **3 write modes** — Normal, streaming (memory-efficient), chunked streaming (constant memory)
- **Rich styling** — Fonts, fills, borders, alignment, number formats
- **Hyperlinks** — External URLs, email, internal sheet references
- **Formulas** — Read/write with cached results (SUM, AVERAGE, IF, etc.)
- **Data validation** — Dropdown lists, number ranges, date limits, custom formulas
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

## Documentation

See [DOCUMENT.md](DOCUMENT.md) for the complete API reference, including:

- All functions (`writeExcel`, `readExcel`, `writeCSV`, `readCSV`, streaming APIs)
- Type definitions (`Workbook`, `Worksheet`, `Cell`, `Row`, etc.)
- Styles guide (font, fill, border, alignment, number formats)
- Features (formulas, hyperlinks, merge cells, freeze panes, data validation)
- Writing modes comparison (normal vs streaming vs chunked)

## Benchmarks

30 columns x 30,000 rows with styles and formulas (MacOS, Bun 1.3):

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

# Large report (30 col x 30K rows)
bun run large-report

# Benchmarks
bun run benchmark
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
