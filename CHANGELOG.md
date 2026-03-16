# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- Add template mode via `loadExcelTemplate()` and `ExcelTemplate` for loading an existing workbook, updating cells or named ranges, and writing the result back out
- Add Bun-native `readExcelStream()` for row-by-row XLSX reads from local files, `Bun.file(...)`, and `S3File` sources

## [1.1.0] - 2026-03-16

### Added
- Add cell comments / notes support across normal, streaming, chunked, and multi-sheet Excel writers, with `readExcel()` round-trip parsing
- Add worksheet image embedding support for PNG, JPEG, and GIF assets via `worksheet.images`
- Add structured table support via `worksheet.tables`, including read/write of table definitions and streaming writer support

### Changed
- Switch `createExcelStream()` to the same disk-backed Bun `FileSink`/temp-file path instead of buffering serialized row XML in memory
- Refactor multi-sheet streaming writes to assemble worksheets from disk-backed temp files rather than buffering all rows in memory
- Use `Bun.file().bytes()` for XLSX reads and `Bun.file(path).delete()` for temp-file cleanup on Bun runtime paths
- Add `FileSource` / `FileTarget` support so CSV/XLSX read and write APIs accept `BunFile` and `S3File` in addition to local path strings
- Add direct Bun S3 object-target support for CSV and Excel streaming writers via `S3File.writer()`

## [1.0.2] - 2026-03-15

### Changed
- Refactor `createChunkedExcelStream()` to stream worksheet XML and ZIP output incrementally instead of materializing the full worksheet/archive in memory at `end()`
- Move chunked-writer hyperlink metadata to temp files and add internal temp/output flush management to reduce RAM growth during large exports

## [1.0.1] - 2026-03-14

### Changed
- Publish built runtime files from `dist/` instead of exposing TypeScript source files directly
- Update package exports to standard `dist/index.js` + `dist/index.d.ts` entrypoints for normal npm consumption
- Replace the publish-time `prepare` hook setup with a dedicated `setup:hooks` development script
- Add Bun-only runtime compatibility notes to the README files

## [1.0.0] - 2024-01-20

### Added
- Core Excel (.xlsx) read/write with full styling support
- CSV read/write with auto-type detection
- Streaming CSV (read + write) with `AsyncGenerator` and `FileSink`
- Streaming Excel write with immediate row serialization (`createExcelStream`)
- Chunked streaming Excel write with constant memory (`createChunkedExcelStream`)
- Multi-sheet Excel streaming (`createMultiSheetExcelStream`)
- Cell styles: font, fill, border, alignment, number formats
- Merged cells (read + write)
- Freeze panes (read + write)
- Formula support (read + write with cached results)
- Hyperlinks (external URL, mailto, internal sheet references)
- Security hardening: XML bomb prevention, path traversal protection, input validation
- Comprehensive benchmarks and examples
- Excel data validation support for dropdown lists, number ranges, date limits, and custom formulas
- Read/write support for worksheet-level `dataValidations` across normal, streaming, and chunked Excel writers
- Conditional formatting support for highlight rules, color scales, data bars, and icon sets
- Auto filter support via worksheet-level `autoFilter`
- Workbook metadata read/write support for `creator`, `created`, and `modified`
- Split view support via worksheet-level `splitPane`
- Gradient fill XML support for cell styles
- Automatic date conversion on read for numeric cells with date/time number formats
