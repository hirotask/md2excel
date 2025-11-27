"""Main entry point for md2excel converter."""

import sys
import argparse
from pathlib import Path
from .parser import MarkdownParser
from .excel_generator import ExcelGenerator


def main():
    """Main function to convert Markdown to Excel format."""
    parser = argparse.ArgumentParser(
        description='Convert Markdown tables to Excel (XLSX) format'
    )
    parser.add_argument(
        'input',
        type=str,
        help='Input Markdown file path'
    )
    parser.add_argument(
        '-o', '--output',
        type=str,
        help='Output Excel file path (default: input filename with .xlsx extension)'
    )

    args = parser.parse_args()

    # Validate input file
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Error: Input file '{args.input}' does not exist.", file=sys.stderr)
        sys.exit(1)

    if not input_path.is_file():
        print(f"Error: '{args.input}' is not a file.", file=sys.stderr)
        sys.exit(1)

    # Determine output path
    if args.output:
        output_path = Path(args.output)
    else:
        output_path = input_path.with_suffix('.xlsx')

    try:
        # Parse Markdown file
        print(f"Parsing Markdown file: {input_path}")
        md_parser = MarkdownParser()
        sheets = md_parser.parse_file(str(input_path))

        if not sheets:
            print("Warning: No sheets found in the Markdown file.", file=sys.stderr)
            print("Make sure your Markdown file contains at least one heading 1 (# ) and a table.")
            sys.exit(1)

        # Generate Excel file
        print(f"Generating Excel file: {output_path}")
        excel_generator = ExcelGenerator()
        excel_generator.generate_excel(sheets, str(output_path))

        print(f"Successfully converted {input_path} to {output_path}")
        print(f"Generated {len(sheets)} sheet(s)")

    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
