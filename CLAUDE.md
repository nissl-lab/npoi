# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

NPOI is the .NET version of Apache POI, enabling reading/writing Office 2003/2007 files (xls, xlsx, docx). The project is organized into several core modules with distinct responsibilities.

## Core Architecture

The codebase is organized into the following main components:

- **main/** - Core NPOI functionality (HSSF for Excel .xls, HPSF for document properties, DDF for drawing)
- **ooxml/** - OOXML format support (XSSF for Excel .xlsx, encryption, digital signatures)
- **openxml4Net/** - Low-level OpenXML package handling
- **OpenXmlFormats/** - OpenXML format definitions and schemas
- **scratchpad/** - Additional format support (HSLF for PowerPoint, HWPF for Word)
- **testcases/** - Comprehensive unit tests for all modules
- **solution/** - Main solution files and packaging project

## Build Commands

The project uses Nuke build system. Use these commands:

```bash
# Build the project
./build.sh

# Build with specific target
./build.sh Compile

# Run tests (requires fonts to be installed on Linux CI)
./build.sh Test

# Create NuGet packages
./build.sh Pack

# Clean build artifacts
./build.sh Clean
```

On Windows, use `build.cmd` or `build.ps1` instead of `build.sh`.

## Solution Structure

- **NPOI.Core.sln** - Main solution with core libraries
- **NPOI.Core.Test.sln** - Test solution
- **NPOI.Pack.csproj** - NuGet packaging project (in solution/)

The build system automatically manages dependencies between projects and handles multi-targeting (.NET Framework 4.7.2, .NET Standard 2.0/2.1, .NET 8.0).

## Key Dependencies

- **BouncyCastle.Cryptography** - Cryptographic operations
- **SharpZipLib** - ZIP compression
- **SixLabors.ImageSharp** & **SixLabors.Fonts** - Image and font handling
- **MathNet.Numerics** - Mathematical computations
- **Microsoft.IO.RecyclableMemoryStream** - Memory-efficient stream handling

## Testing

Tests are comprehensive and include both unit tests and integration tests. The test suite requires specific fonts on Linux systems (automatically handled in CI).

Use `./build.sh Test` to run all tests across all target frameworks.