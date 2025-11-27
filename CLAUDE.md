# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

md2excelxml is a Python CLI tool that converts Markdown tables into Office Open XML format (Excel-compatible). It processes Markdown files and generates XML files that can be opened in Microsoft Excel with proper formatting, borders, and styles.

## Installation and Setup

```bash
# Install in development mode with uv
uv pip install -e .

# For Mermaid diagram support (optional)
# Install Playwright browser for rendering
uv run python -m playwright install chromium

# Install system dependencies (Linux/WSL)
sudo apt-get install -y libnspr4 libnss3 libasound2t64

# Install Japanese fonts (for Japanese text in Mermaid diagrams)
sudo apt-get install -y fonts-noto-cjk fonts-noto-cjk-extra

# Or use playwright's helper (installs all dependencies)
sudo uv run python -m playwright install-deps
```

All dependencies including Spire.XLS are managed by uv and defined in `pyproject.toml`.

**Important:**
- Python 3.15+ is not currently supported due to Spire.XLS compatibility issues. Use Python 3.8-3.14.
- Mermaid diagram support requires `mermaid-cli` and Playwright Chromium browser with system dependencies.

## Development Commands

```bash
# Run the tool on test file
uv run md2excel test_sample.md

# Run with custom output path
uv run md2excel test_sample.md -o output.xlsx

# The generated file can be opened directly in Excel
```

## Docker Usage

A Dockerfile is provided for containerized execution with all dependencies pre-installed:

```bash
# Build the Docker image
./docker-build.sh
# or
docker build -t md2excel:latest .

# Run the tool
docker run -v $(pwd):/data md2excel:latest input.md -o output.xlsx

# The Docker image includes:
# - Python 3.12
# - All Python dependencies (spire-xls, mermaid-cli)
# - Playwright with Chromium
# - Japanese fonts (Noto CJK, IPA)
# - All system dependencies
```

## Architecture

The tool follows a three-stage pipeline architecture:

1. **Parsing** (`parser.py`): `MarkdownParser` class parses Markdown content and extracts structured data
   - Recognizes `# Heading1` as sheet boundaries
   - Treats `## Heading2` as section headings within sheets (doesn't create new sheets) and marks them with `type='heading'`
   - Ignores `### Heading3` and deeper levels
   - Extracts both tables and text content (paragraphs)
   - Extracts Mermaid code blocks (` ```mermaid ... ``` `) and marks them with `type='mermaid'`
   - Processes list syntax within table cells
   - Maintains the order of headings, text, tables, and Mermaid diagrams as they appear in the source

2. **Generation** (`excel_generator.py`, `xml_generator.py`, and `mermaid_renderer.py`): Excel file generation
   - `ExcelGenerator`: Uses Spire.XLS library to generate Excel files in various formats
     - For `.xml` output: Uses `FileFormat.XML` to generate Excel XML Spreadsheet (SpreadsheetML) format
     - For `.xlsx` output: Uses `ExcelVersion.Version2013` to generate Office Open XML format
     - Supports rendering Mermaid diagrams as PNG images embedded in the Excel file
   - `MermaidRenderer`: Uses mermaid-cli (Playwright + Chromium) to render Mermaid diagrams to PNG
     - Converts Mermaid code to PNG images locally (no external services)
     - Gracefully handles missing dependencies with warning messages
   - `XMLGenerator`: Legacy generator using xml.etree for SpreadsheetML format (kept for backward compatibility)
   - All generators support text blocks, tables, and Mermaid diagrams

3. **Entry Point** (`main.py`): CLI interface using argparse
   - Validates input files
   - Orchestrates parser and generator
   - Handles error reporting

### Data Flow

```
Markdown File → MarkdownParser.parse_file() → List[Dict]
                                                ↓
                                                sheets = [
                                                  {
                                                    'name': str,
                                                    'items': [
                                                      {
                                                        'type': 'text',
                                                        'content': str
                                                      },
                                                      {
                                                        'type': 'table',
                                                        'headers': List[str],
                                                        'rows': List[List[str]]
                                                      }
                                                    ]
                                                  }
                                                ]
                                                ↓
                    ExcelGenerator.generate_excel() → Excel XML File
```

**Data Structure:**
- Each sheet contains a list of `items` (not `tables`)
- Each item has a `type` field: `'text'`, `'heading'`, `'table'`, or `'mermaid'`
- Heading items have a `level` field (2 for H2) and `content` field
- Text items have a `content` field
- Table items have `headers` and `rows` fields
- Mermaid items have a `content` field (diagram code) and `language` field ('mermaid')
- Items are rendered in the order they appear

### List Processing Logic

The parser transforms Markdown list syntax into formatted cell content:

- Bullet lists (`- Item`) → `・ Item` with newline separation
- Numbered lists (`1. Step`) → `1) Step` with newline separation
- Multiple list items in one cell are split and formatted separately
- `<br>` tags are supported for manual line breaks

This processing happens in `parser.py:_process_cell_content()` which handles complex patterns like:
- Inline bullet lists: `- Item1 - Item2`
- Inline numbered lists: `1. Step1 2. Step2`
- Mixed content with `<br>` tags

### Style Application

`ExcelGenerator` applies consistent styling:
- Heading cells (H2): Bold font, size 13, no borders, no text wrapping, top-left alignment
- Text cells: Normal font, no borders, no text wrapping, top-left alignment
- Table header rows: Bold font with borders, text wrapping enabled
- Table data cells: Normal font with borders, text wrapping enabled, top-left alignment
- Table cells get thin borders on all sides
- Mermaid diagrams: Rendered as PNG images with 600px width, positioned at the current row
- Columns auto-fit to content width (with fallback to default 15 units)

### Mermaid Diagram Support

Mermaid diagrams in code blocks are automatically rendered as images:
- Parser detects ` ```mermaid ... ``` ` code blocks
- `MermaidRenderer` uses mermaid-cli with Playwright to render diagrams locally
- Images are embedded directly in the Excel file as PNG
- If rendering fails (missing dependencies), a warning message is displayed in the cell
- No external services are used - all rendering is done locally

## Key Implementation Details

- **Sheet naming**: Excel has restrictions (31 char max, no `:\/?*[]` chars). The generator sanitizes names via `_sanitize_sheet_name()`
- **1-indexed rows/columns**: Spire.XLS uses 1-based indexing, not 0-based
- **Empty rows**: Added between tables within a sheet for visual separation
- **Encoding**: All file operations use UTF-8 encoding
- **Borders API**: Use `BordersLineType` enum (e.g., `BordersLineType.EdgeLeft`) instead of numeric indices to set cell borders
- **File formats**:
  - `.xml` files use `FileFormat.XML` (Excel XML Spreadsheet format)
  - `.xlsx` files use `ExcelVersion.Version2013` (Office Open XML format)
  - Default output is Excel 2013 format if extension is not recognized

## Testing

The repository includes a test file demonstrating all features:
- `test_sample.md`: Example Markdown with multiple sheets, sections, and list formats
- Generated outputs: `test_sample.xml`, `test_sample_v2.xml`, `test_sample_final.xml`

To verify changes, run the tool on the test file and open the output in Excel to check formatting.

## Dependencies

- **Spire.XLS** (>=2.0.0): Commercial Python library for Excel manipulation (primary dependency)
- **mermaid-cli** (>=0.1.0): Optional dependency for Mermaid diagram rendering
  - Requires Playwright and Chromium browser
  - Requires system dependencies: libnspr4, libnss3, libasound2t64 (Linux/WSL)
- **Python** >=3.8 required
- Standard library: `xml.etree.ElementTree`, `xml.dom.minidom`, `argparse`, `pathlib`, `re`, `os`, `tempfile`, `logging`, `asyncio`

All dependencies are managed through uv and defined in `pyproject.toml`.

### Optional: Mermaid Support

Mermaid diagram support is optional. If mermaid-cli is not installed, the tool will:
- Still process Markdown files normally
- Display a warning message in place of Mermaid diagrams
- Continue processing other content without errors

To enable Mermaid support, follow the installation steps in the "Installation and Setup" section above.

## Common Issues

- If tables aren't being detected, ensure there's proper spacing between the separator line (`|---|---|`) and data rows
- Sheet names longer than 31 characters will be truncated automatically
- The tool requires at least one H1 heading and one table to generate output successfully
- **Mermaid diagrams not rendering**: Ensure system dependencies are installed:
  ```bash
  # Install Chromium browser for Playwright
  uv run python -m playwright install chromium

  # Install system dependencies (Linux/WSL)
  sudo apt-get install -y libnspr4 libnss3 libasound2t64
  ```
- **Japanese text garbled in Mermaid diagrams**: Install Japanese fonts:
  ```bash
  # Install Noto CJK fonts (recommended)
  sudo apt-get install -y fonts-noto-cjk fonts-noto-cjk-extra

  # Or install IPA fonts
  sudo apt-get install -y fonts-ipafont fonts-ipaexfont
  ```
  The renderer uses the following font fallback chain:
  - Noto Sans JP (preferred)
  - Hiragino Sans
  - Yu Gothic
  - Meiryo
  - sans-serif
- **WSL/Linux headless environment**: Playwright may require additional libraries for headless browser operation
