# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

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
