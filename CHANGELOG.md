# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.0.0] - 2026-03-17

### Added
- Initial public release of `bun-excel`, renamed from `bun-spreadsheet`
- Core Excel (.xlsx) read/write with full styling support
- CSV read/write with auto-type detection
- Streaming CSV (read + write) with `AsyncGenerator` and `FileSink`
- Streaming Excel write with immediate row serialization (`createExcelStream`)
- Chunked streaming Excel write with disk-backed low-memory output (`createChunkedExcelStream`)
- Multi-sheet Excel streaming (`createMultiSheetExcelStream`)
- Bun-native `readExcelStream()` for row-by-row XLSX reads from local files, `Bun.file(...)`, and `S3File` sources
- Template mode via `loadExcelTemplate()` and `ExcelTemplate` for loading an existing workbook, updating cells or named ranges, and writing the result back out
- Production-oriented Excel export helpers via `exportExcelRows()`, `exportMultiSheetExcel()`, `buildExcelResponse()`, and streaming `Response` helpers with progress callbacks, abort support, and export diagnostics
- Cell styles: font, fill, border, alignment, number formats
- Rich text support for partial cell formatting
- Cell comments / notes support across normal, streaming, chunked, and multi-sheet Excel writers
- Hyperlinks (external URL, mailto, internal sheet references)
- Worksheet image embedding support for PNG, JPEG, and GIF assets via `worksheet.images`
- Structured table support via `worksheet.tables`, including read/write of table definitions and streaming writer support
- Merged cells (read + write)
- Freeze panes and split views (read + write)
- Formula support (read + write with cached results)
- Data validation support for dropdown lists, number ranges, date limits, and custom formulas
- Conditional formatting support for highlight rules, color scales, data bars, and icon sets
- Auto filter support via worksheet-level `autoFilter`
- Workbook metadata read/write support for `creator`, `created`, and `modified`
- Workbook and worksheet view settings, page setup, header/footer, outline levels, sheet state, protection, and defined names
- Worksheet row/column utility operations
- Bun-native file target support for local paths, `Bun.file(...)`, and `S3File`
- Direct Bun S3 object-target support for CSV and Excel streaming writers via `S3File.writer()`
- Comprehensive benchmarks, examples, and API documentation

### Changed
- Publish built runtime files from `dist/` instead of exposing TypeScript source files directly
- Use standard `dist/index.js` + `dist/index.d.ts` entrypoints for normal npm consumption
- Replace the publish-time `prepare` hook setup with a dedicated `setup:hooks` development script
- Switch `createExcelStream()` to the same disk-backed Bun `FileSink`/temp-file path instead of buffering serialized row XML in memory
- Refactor `createChunkedExcelStream()` to stream worksheet XML and ZIP output incrementally instead of materializing the full worksheet/archive in memory at `end()`
- Refactor multi-sheet streaming writes to assemble worksheets from disk-backed temp files rather than buffering all rows in memory
- Use `Bun.file().bytes()` for XLSX reads and `Bun.file(path).delete()` for temp-file cleanup on Bun runtime paths
- Add `FileSource` / `FileTarget` support so CSV/XLSX read and write APIs accept `BunFile` and `S3File` in addition to local path strings
- Rename the project and npm package from `bun-spreadsheet` to `bun-excel`
- Add Bun-only runtime compatibility notes to the README files

### Security
- Harden CSV export against formula injection
- Add ZIP bomb and oversized input guards for CSV/XLSX reading paths
- Sanitize numeric runtime values before writing XML attributes
- Clean up temp worksheet files on `readExcelStream()` failure paths
- Validate image coordinates and template-mode worksheet bounds against Excel limits
