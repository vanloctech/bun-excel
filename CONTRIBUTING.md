# Contributing to bun-spreadsheet

Thank you for your interest in contributing! This guide will help you get started.

## Development Setup

### Prerequisites

- [Bun](https://bun.sh) >= 1.0.0
- Git

### Getting Started

```bash
# Clone the repository
git clone https://github.com/<your-username>/bun-spreadsheet.git
cd bun-spreadsheet

# Install dependencies
bun install

# Run type check
bun run typecheck

# Run examples
bun run demo
bun run benchmark
```

## Project Structure

```
src/
├── index.ts                            # Public API exports
├── types.ts                            # All TypeScript interfaces
├── csv/
│   ├── csv-reader.ts                   # CSV read + streaming read
│   └── csv-writer.ts                   # CSV write + streaming write
└── excel/
    ├── xml-parser.ts                   # Lightweight XML parser
    ├── xml-builder.ts                  # XML generation utilities
    ├── style-builder.ts                # Style registry + styles.xml
    ├── xlsx-reader.ts                  # XLSX file reader
    ├── xlsx-writer.ts                  # XLSX file writer
    ├── xlsx-stream-writer.ts           # Streaming XLSX writer
    └── xlsx-chunked-stream-writer.ts   # Constant-memory chunked writer
```

## Development Workflow

### Making Changes

1. Create a new branch from `main`
2. Make your changes
3. Run type checking: `bun run typecheck`
4. Run the demo to verify: `bun run demo`
5. Run the benchmark if performance-related: `bun run benchmark`
6. Submit a pull request

### Code Style

- **TypeScript strict mode** — all files must pass `tsc --noEmit`
- **No `any` types** — use proper typing or `unknown`
- **Security first** — validate all user inputs, escape XML values, check file paths
- **Bun-native** — prefer Bun APIs (`Bun.file()`, `Bun.write()`, `FileSink`) over Node.js equivalents

### Adding New Features

1. Add types to `src/types.ts`
2. Implement the feature in the appropriate module
3. Export from `src/index.ts`
4. Add an example in `examples/`
5. Update `README.md` with API documentation

### Security Guidelines

This library processes user-supplied data and files. Keep these in mind:

- **XML injection**: Always use `escapeXML()` for user-controlled values
- **Path traversal**: Validate file paths with `path.resolve()` and null byte checks
- **Resource limits**: Cap maximum rows, columns, file sizes to prevent DoS
- **Prototype pollution**: Use `Object.create(null)` for dynamic key maps

## Reporting Issues

- Use [GitHub Issues](https://github.com/<your-username>/bun-spreadsheet/issues)
- Include Bun version (`bun --version`)
- Include a minimal reproduction if possible

## License

By contributing, you agree that your contributions will be licensed under the [MIT License](LICENSE).
