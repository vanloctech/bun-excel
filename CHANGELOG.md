# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
