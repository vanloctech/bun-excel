# bun-excel

[![CI](https://github.com/vanloctech/bun-excel/actions/workflows/ci.yml/badge.svg)](https://github.com/vanloctech/bun-excel/actions/workflows/ci.yml)
[![npm version](https://img.shields.io/npm/v/bun-excel.svg)](https://www.npmjs.com/package/bun-excel)
[![GitHub stars](https://img.shields.io/github/stars/vanloctech/bun-excel?style=social)](https://github.com/vanloctech/bun-excel/stargazers)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Bun](https://img.shields.io/badge/Bun-%E2%89%A51.0-black?logo=bun)](https://bun.sh)
[![TypeScript](https://img.shields.io/badge/TypeScript-%E2%89%A55.0-blue?logo=typescript)](https://www.typescriptlang.org/)

[![English](https://img.shields.io/badge/lang-English-blue)](README.md) [![中文](https://img.shields.io/badge/lang-%E4%B8%AD%E6%96%87-red)](README.zh-CN.md)

High-performance, Bun-optimized Excel and CSV library for TypeScript.

> ⚠️ **Note**: Runtime note: `bun-excel` uses Bun-specific APIs. It is intended for Bun and is not compatible with Node.js or Deno.
> 
> Please read our [Code of Conduct](CODE_OF_CONDUCT.md) before participating in the project.
> 
> Some security scanners may flag `schemas.openxmlformats.org` or similar URLs in this package. These are OOXML namespace and relationship identifiers required by the Excel file format, not runtime network requests.

## Why This Package

- **Built for Bun, not adapted from Node-first abstractions** — The core file paths use `Bun.file()`, `Bun.write()`, `FileSink`, and Bun-native streaming APIs directly.
- **Works naturally with Bun-native file targets, including S3** — Read from and write to local paths, `Bun.file(...)`, and Bun `S3File` objects, including direct streaming exports to S3 destinations.
- **Production export helpers for Bun backends** — Supports progress callbacks, `AbortSignal`, export diagnostics, streaming `Response` helpers, and S3 multipart tuning through Bun-native writer options.
- **TypeScript-first spreadsheet model** — `Workbook`, `Worksheet`, `Row`, `Cell`, and style objects are explicit and practical to work with in Bun apps.
- **Focused on real report workflows** — Styles, formulas, hyperlinks, data validation, conditional formatting, auto filters, freeze/split panes, and workbook metadata are supported where they matter for business exports.
- **Multiple write strategies for different workloads** — Use normal writes for simplicity, stream writes for lower memory pressure, and chunked disk-backed writes for large exports.

## Install

```bash
bun add bun-excel
```

## Quick Start

### Write Excel

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

await writeExcel("report.xlsx", workbook);
```

### Read Excel

```typescript
import { readExcel } from "bun-excel";

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
import { readCSV, writeCSV } from "bun-excel";

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

Measured on Bun `1.3.10` / `MacOS ARM` with a single worksheet, compressed `.xlsx`, and `1,000,000` rows x `10` columns:

| Mode | Total time | Finalize time | Rows/sec | Peak RSS | Peak heapUsed | File size |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `createExcelStream()` | `13.1s` | `8.9s` | `76,363` | `110.6MB` | `5.1MB` | `54.33MB` |
| `createChunkedExcelStream()` | `11.9s` | `8.5s` | `84,029` | `120.9MB` | `5.1MB` | `54.33MB` |

`createExcelStream()` now uses the same disk-backed low-memory path as the chunked writer for single-sheet exports, so the numbers are expected to be close. Re-run the large benchmark on your machine with:

```bash
bun run benchmark:1m
```

For the realistic `large-report` workload benchmark across normal, streaming, and chunked modes, measured on Bun `1.3.10` / `MacOS ARM` with a single worksheet, compressed `.xlsx`, and `30` columns x `30,000` rows using the same styles, merge cells, and footer formulas as `examples/large-report.ts`:

| Method | Total time | Peak RSS delta | Peak heapUsed delta | File size |
| --- | ---: | ---: | ---: | ---: |
| `writeExcel()` | `1.92s` | `518.6MB` | `154.0MB` | `6.20MB` |
| `createExcelStream()` | `1.98s` | `48.0MB` | `39.2MB` | `6.35MB` |
| `createChunkedExcelStream()` | `2.01s` | `2.5MB` | `3.1MB` | `6.35MB` |

These memory columns are peak deltas over the baseline process memory during the benchmark, not just the memory difference before and after the write.

Re-run this benchmark on your machine with:

```bash
bun run benchmark
```

## Examples

```bash
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
