# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

md2excelxml is a Python CLI tool that converts Markdown tables into Office Open XML format (Excel-compatible). It processes Markdown files and generates XML files that can be opened in Microsoft Excel with proper formatting, borders, and styles.

## Installation and Setup

```bash
# Install in development mode with uv
uv pip install -e .
```

All dependencies including Spire.XLS are managed by uv and defined in `pyproject.toml`.

**Important:** Python 3.15+ is not currently supported due to Spire.XLS compatibility issues. Use Python 3.8-3.14.

## Development Commands

```bash
# Run the tool on test file
uv run md2excelxml test_sample.md

# Run with custom output path
uv run md2excelxml test_sample.md -o output.xml

# The generated XML can be opened directly in Excel
```

## Architecture

The tool follows a three-stage pipeline architecture:

1. **Parsing** (`parser.py`): `MarkdownParser` class parses Markdown content and extracts structured data
   - Recognizes `# Heading1` as sheet boundaries
   - Treats `## Heading2` as section headings within sheets (doesn't create new sheets) and marks them with `type='heading'`
   - Ignores `### Heading3` and deeper levels
   - Extracts both tables and text content (paragraphs)
   - Processes list syntax within table cells
   - Maintains the order of headings, text, and tables as they appear in the source

2. **Generation** (`excel_generator.py` and `xml_generator.py`): Excel file generation
   - `ExcelGenerator`: Uses Spire.XLS library to generate Excel files in various formats
     - For `.xml` output: Uses `FileFormat.XML` to generate Excel XML Spreadsheet (SpreadsheetML) format
     - For `.xlsx` output: Uses `ExcelVersion.Version2013` to generate Office Open XML format
   - `XMLGenerator`: Legacy generator using xml.etree for SpreadsheetML format (kept for backward compatibility)
   - Both generators support text blocks and tables

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
- Each item has a `type` field: `'text'`, `'heading'`, or `'table'`
- Heading items have a `level` field (2 for H2) and `content` field
- Text items have a `content` field
- Table items have `headers` and `rows` fields
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
- Columns auto-fit to content width (with fallback to default 15 units)

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
- **Python** >=3.8 required
- Standard library: `xml.etree.ElementTree`, `xml.dom.minidom`, `argparse`, `pathlib`, `re`

All dependencies are managed through uv and defined in `pyproject.toml`.

## Common Issues

- If tables aren't being detected, ensure there's proper spacing between the separator line (`|---|---|`) and data rows
- Sheet names longer than 31 characters will be truncated automatically
- The tool requires at least one H1 heading and one table to generate output successfully
