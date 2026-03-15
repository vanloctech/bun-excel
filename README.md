# bun-spreadsheet

[![CI](https://github.com/vanloctech/bun-spreadsheet/actions/workflows/ci.yml/badge.svg)](https://github.com/vanloctech/bun-spreadsheet/actions/workflows/ci.yml)
[![npm version](https://img.shields.io/npm/v/bun-spreadsheet.svg)](https://www.npmjs.com/package/bun-spreadsheet)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Bun](https://img.shields.io/badge/Bun-%E2%89%A51.0-black?logo=bun)](https://bun.sh)
[![TypeScript](https://img.shields.io/badge/TypeScript-%E2%89%A55.0-blue?logo=typescript)](https://www.typescriptlang.org/)

[![English](https://img.shields.io/badge/lang-English-blue)](README.md) [![中文](https://img.shields.io/badge/lang-%E4%B8%AD%E6%96%87-red)](README.zh-CN.md)

High-performance, Bun-optimized Excel (.xlsx) and CSV library for TypeScript.

> ⚠️ **Note**: Runtime note: `bun-spreadsheet` uses Bun-specific APIs. It is intended for Bun and is not compatible with Node.js or Deno.

## Why This Package

- **Built for Bun, not adapted from Node-first abstractions** — The core file paths use `Bun.file()`, `Bun.write()`, `FileSink`, and Bun-native streaming APIs directly.
- **TypeScript-first spreadsheet model** — `Workbook`, `Worksheet`, `Row`, `Cell`, and style objects are explicit and practical to work with in Bun apps.
- **Focused on real report workflows** — Styles, formulas, hyperlinks, data validation, conditional formatting, auto filters, freeze/split panes, and workbook metadata are supported where they matter for business exports.
- **Multiple write strategies for different workloads** — Use normal writes for simplicity, stream writes for lower memory pressure, and chunked disk-backed writes for large exports.

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

Measured on Bun `1.3.10` / `darwin arm64` with a single worksheet, compressed `.xlsx`, and `1,000,000` rows x `10` columns:

| Mode | Total time | Finalize time | Rows/sec | Peak RSS | Peak JS heap | File size |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `createExcelStream()` | `12.5s` | `8.4s` | `79,857` | `132.0MB` | `5.1MB` | `54.33MB` |
| `createChunkedExcelStream()` | `12.1s` | `8.4s` | `82,882` | `121.5MB` | `5.1MB` | `54.33MB` |

`createExcelStream()` now uses the same disk-backed low-memory path as the chunked writer for single-sheet exports, so the numbers are expected to be close. Re-run the large benchmark on your machine with:

```bash
bun run benchmark:1m
```

For the smaller comparison benchmark across normal, streaming, and chunked modes, run:

```bash
bun run benchmark
```

## Examples

```bash
# Run all examples
bun run demo

# Large report (30 col x 30K rows)
bun run large-report

# Benchmarks (normal vs stream vs chunked)
bun run benchmark

# 1M-row Excel benchmark (stream vs chunked)
bun run benchmark:1m
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
