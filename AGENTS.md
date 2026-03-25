# AGENTS.md

Guidance for AI agents and new contributors working in this codebase.

## Project Overview

NPOI is a .NET port of [Apache POI](https://poi.apache.org/) for reading and
writing Microsoft Office file formats without requiring COM or an Office
installation. Targets net472, netstandard2.0, netstandard2.1, and net8.0.

## Code Style

- Use 4 spaces for indentation (default for .NET projects)
- No comments unless explaining non-obvious business logic
- Prefer `var` for local variables when type is clear from right-hand side
- Follow existing patterns for nullable reference types (`string?` vs `string`)
- Use XML documentation (`/// <summary>`) for public APIs only

## Directory → Project Map

| Directory         | Project                      | Purpose                                                        |
|-------------------|------------------------------|----------------------------------------------------------------|
| `main/`           | NPOI.Core                    | HSSF (.xls), HPSF (doc properties), SS model, DDF, utilities  |
| `ooxml/`          | NPOI.OOXML                   | XSSF (.xlsx), SXSSF (streaming), XWPF (.docx), XSLF (.pptx)  |
| `OpenXmlFormats/` | NPOI.OpenXmlFormats          | C# classes mapping to OOXML XML schemas (data-only, no logic)  |
| `openxml4Net/`    | NPOI.OpenXml4Net             | OPC packaging: zip I/O, content types, part relationships      |
| `scratchpad/`     | (various)                    | HWPF (.doc binary), HSLF (.ppt binary), other legacy formats   |
| `testcases/`      | Test project                 | Mirrors source structure; test data in `testcases/test-data/`  |
| `build/`          | _build (Nuke)                | Build automation                                               |
| `benchmarks/`     | NPOI.Benchmarks              | BenchmarkDotNet performance tests                              |
| `solution/`       | Solution/packaging files     | `.sln` files and NuGet pack project                            |

## Format Acronyms

These names come from Apache POI and appear throughout the codebase:

- **HSSF** — Horrible SpreadSheet Format (binary `.xls`)
- **XSSF** — XML SpreadSheet Format (`.xlsx`)
- **SXSSF** — Streaming XSSF (write-only sliding window over XSSF)
- **HWPF** — Horrible Word Processor Format (binary `.doc`)
- **XWPF** — XML Word Processor Format (`.docx`)
- **HSLF** — Horrible SLide Format (binary `.ppt`)
- **XSLF** — XML SLide Format (`.pptx`)
- **HPSF** — Horrible Property Set Format (document metadata)
- **SS** — Common spreadsheet interfaces (`ISheet`, `IRow`, `ICell`) shared by HSSF and XSSF
- **DDF** — Dreadful Drawing Format (shapes/drawing records in binary formats)

## Format Selection Guide

| Format | Use When | Memory | Limitations |
|--------|----------|--------|-------------|
| **HSSF** | Reading/writing `.xls` (Excel 97-2003) | Medium | 65,536 row limit |
| **XSSF** | Reading/writing `.xlsx`, full features needed | High | Memory scales with file size |
| **SXSSF** | Writing large `.xlsx` files (>100K rows) | Low (constant) | Write-only, limited random access |

## Key Architecture Patterns

- **SS.UserModel interfaces** (`IWorkbook`, `ISheet`, `IRow`, `ICell`) abstract
  over both HSSF and XSSF. Prefer these when writing format-agnostic code.

- **SXSSF streaming layer** — classes in `ooxml/XSSF/Streaming/` wrap the
  corresponding XSSF classes with a configurable row window for low-memory
  streaming writes.

- **OpenXmlFormats are POCOs** — the classes in `OpenXmlFormats/` map directly
  to the ECMA-376 (OOXML) XML schemas. Business logic that operates on these
  types lives in `ooxml/` or `main/`, not alongside the schema classes.

- **OPC packaging** (`openxml4Net/`) handles the zip container structure shared
  by `.xlsx`, `.docx`, and `.pptx` files. Format-specific layers read and write
  individual XML parts through this packaging layer.

- **Design patterns** — NPOI uses:
  - **Strategy**: format-specific implementations behind interfaces
  - **Factory**: `CreateSheet()`, `CreateRow()`, `CreateCell()` maintain parent-child consistency
  - **Decorator**: SXSSF wraps XSSF, intercepting writes for streaming
  - **Flyweight**: `StylesTable` and `SharedStringsTable` deduplicate formatting and strings
  - **Template Method**: interface methods with HSSF/XSSF format-specific implementations

- **HSSF uses records** — `InternalWorkbook`/`InternalSheet` aggregate BIFF8 records
  (`SSTRecord`, `ExtendedFormatRecord`, `RowRecord`, etc.)

- **XSSF uses XML beans** — OpenXmlFormats classes (`CT_Workbook`, `CT_Worksheet`,
  `CT_Row`, `CT_Cell`) map to OOXML XML elements

## Building and Testing

```bash
# Build
dotnet build solution/NPOI.Core.sln

# Run tests
dotnet test solution/NPOI.Core.Test.sln

# Run benchmarks
dotnet run -c Release --project benchmarks/NPOI.Benchmarks/
```

## Code Verification

After making changes:
1. Build the solution to catch compilation errors
2. Run relevant tests to verify functionality
3. Check for style violations if present (look for editorconfig or style guidelines)

## Common Tasks

### Adding a new format/feature
1. Check existing implementations in `main/` or `ooxml/` for patterns
2. Add tests in `testcases/` with sample files in `testcases/test-data/`
3. Follow the SS.UserModel interface pattern for format-agnostic features

### Working with OOXML formats
1. Schema classes live in `OpenXmlFormats/` (do not add logic here)
2. Business logic goes in `ooxml/` using the schema classes
3. Use the OPC packaging layer (`openxml4Net/`) for zip operations

## Test Data

Test fixture files (`.xls`, `.xlsx`, `.docx`, etc.) live under
`testcases/test-data/` organized by format. Tests reference these files via
relative paths from the test project root.
