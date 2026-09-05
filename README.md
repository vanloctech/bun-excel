# bun-excel

[![CI](https://github.com/vanloctech/bun-excel/actions/workflows/ci.yml/badge.svg)](https://github.com/vanloctech/bun-excel/actions/workflows/ci.yml)
[![npm version](https://img.shields.io/npm/v/bun-excel.svg)](https://www.npmjs.com/package/bun-excel)
[![GitHub stars](https://img.shields.io/github/stars/vanloctech/bun-excel?style=social)](https://github.com/vanloctech/bun-excel/stargazers)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Bun](https://img.shields.io/badge/Bun-%E2%89%A51.4-black?logo=bun)](https://bun.sh)
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

Requires Bun 1.4.0 or newer.

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

For previews and selective imports, [`readExcelStream()`](DOCUMENT.md#readexcelstreamsource-options) supports row ranges, per-sheet row limits, column selection, cancellation and progress callbacks.

Use [`readExcelInfo()`](DOCUMENT.md#readexcelinfosource) to inspect sheet names, visibility and workbook metadata before importing rows.

Use [`readExcelObjectsStream()`](DOCUMENT.md#readexcelobjectsstreamsource-options) to map headers to typed objects and receive cell-level validation errors while streaming.

[`readExcel()`](DOCUMENT.md#readexcelsource-options) can skip images, comments and tables with `includeImages: false`, `includeComments: false` and `includeTables: false`.

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

Measured on Bun `1.4.0` / macOS ARM64 on 2026-09-04. Values are medians of 3 runs, with each mode in a fresh process. The write workloads export a single compressed `.xlsx` worksheet.

**1,000,000 rows × 10 columns**

| Mode | Total time | Finalize time | Rows/sec | Peak RSS | File size |
| --- | ---: | ---: | ---: | ---: | ---: |
| `createExcelStream()` | `7.43s` | `4.01s` | `134,505` | `88.2 MiB` | `54.31 MiB` |
| `createChunkedExcelStream()` | `6.80s` | `3.28s` | `147,004` | `94.4 MiB` | `54.31 MiB` |

```bash
bun run benchmark:1m
```

**Large report: 30,000 data rows × 30 columns**, including styles, merged cells and footer formulas from `examples/large-report.ts`.

| Method | Total time | Peak RSS | Sampled peak heapUsed | File size |
| --- | ---: | ---: | ---: | ---: |
| `writeExcel()` | `0.86s` | `337.0 MiB` | `121.8 MiB` | `5.90 MiB` |
| `createExcelStream()` | `0.94s` | `81.7 MiB` | `6.3 MiB` | `6.03 MiB` |
| `createChunkedExcelStream()` | `0.96s` | `82.5 MiB` | `8.5 MiB` | `6.03 MiB` |

```bash
bun run benchmark
```

**Streaming reads: `readExcelStream()`**

Previous reader (custom XML parser) versus the current Bun XML native reader, on the same Bun `1.4.0` runtime. Each fixture contains 30,009 rows, including report headers and footers. Results are medians of 3 alternating runs per implementation in fresh processes; time includes reading all rows and hashing their JSON output. Output checksums matched.

| XLSX fixture | Previous time | Native time | Previous peak RSS | Native peak RSS |
| --- | ---: | ---: | ---: | ---: |
| Shared strings (`bench-normal.xlsx`) | `2.107s` | `1.709s` | `184.7 MiB` | `174.6 MiB` |
| Inline strings (`bench-stream.xlsx`) | `2.214s` | `1.709s` | `153.5 MiB` | `149.4 MiB` |

This compares the complete readers, including XML batching and ZIP decompression, rather than XML parsing alone.

Peak RSS is the OS-recorded process maximum, including runtime memory; heapUsed is sampled and may miss brief peaks. Memory and file sizes use MiB. Results vary by machine and system load; RSS values are not directly comparable to the previous shared-process memory deltas.

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
