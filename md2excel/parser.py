"""Markdown parser for converting Markdown tables to structured data."""

import re
from typing import List, Dict, Tuple


class MarkdownParser:
    """Parse Markdown files and extract table data."""

    def __init__(self):
        self.sheets = []

    def parse_file(self, filepath: str) -> List[Dict]:
        """
        Parse a Markdown file and extract tables grouped by sheets.

        Args:
            filepath: Path to the Markdown file

        Returns:
            List of dictionaries containing sheet information
        """
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()

        return self.parse_content(content)

    def parse_content(self, content: str) -> List[Dict]:
        """
        Parse Markdown content and extract tables, text, and code blocks grouped by sheets.

        Args:
            content: Markdown content as string

        Returns:
            List of dictionaries containing sheet information
        """
        lines = content.split('\n')
        sheets = []
        current_sheet = None
        current_section = None
        in_table = False
        table_lines = []
        text_buffer = []
        in_code_block = False
        code_block_lines = []
        code_block_language = None

        def save_text_buffer():
            """Save accumulated text lines as a text item."""
            if text_buffer and current_sheet is not None:
                # Join text lines and strip whitespace
                text_content = '\n'.join(text_buffer).strip()
                if text_content:
                    current_sheet['items'].append({
                        'type': 'text',
                        'content': text_content
                    })
                text_buffer.clear()

        def save_code_block():
            """Save accumulated code block as a mermaid item."""
            if code_block_lines and current_sheet is not None and code_block_language == 'mermaid':
                code_content = '\n'.join(code_block_lines).strip()
                if code_content:
                    current_sheet['items'].append({
                        'type': 'mermaid',
                        'content': code_content,
                        'language': 'mermaid'
                    })
                code_block_lines.clear()

        for line in lines:
            # Check for code block delimiter (```)
            if line.strip().startswith('```'):
                if not in_code_block:
                    # Starting a code block
                    # Save text buffer before code block
                    save_text_buffer()

                    # Extract language identifier
                    language = line.strip()[3:].strip()
                    code_block_language = language if language else None
                    in_code_block = True
                    code_block_lines = []
                else:
                    # Ending a code block
                    in_code_block = False
                    save_code_block()
                    code_block_language = None
                continue

            # If inside code block, accumulate lines
            if in_code_block:
                code_block_lines.append(line)
                continue

            # Check for H1 (# ) - new sheet
            if line.startswith('# ') and not line.startswith('## '):
                # Save previous text buffer
                save_text_buffer()

                # Save previous table if exists
                if in_table and table_lines:
                    table_data = self._parse_table(table_lines)
                    if table_data and current_sheet is not None:
                        table_data['type'] = 'table'
                        current_sheet['items'].append(table_data)
                    table_lines = []
                    in_table = False

                # Create new sheet
                sheet_name = line[2:].strip()
                current_sheet = {
                    'name': sheet_name,
                    'items': []
                }
                sheets.append(current_sheet)
                current_section = None

            # Check for H2 (## ) - section within sheet
            elif line.startswith('## ') and not line.startswith('### '):
                # Save previous text buffer
                save_text_buffer()

                # Save previous table if exists
                if in_table and table_lines:
                    table_data = self._parse_table(table_lines)
                    if table_data and current_sheet is not None:
                        table_data['type'] = 'table'
                        current_sheet['items'].append(table_data)
                    table_lines = []
                    in_table = False

                current_section = line[3:].strip()
                # Add section as heading
                if current_sheet is not None:
                    current_sheet['items'].append({
                        'type': 'heading',
                        'level': 2,
                        'content': current_section
                    })

            # Check for H3 or higher - ignore
            elif line.startswith('### '):
                continue

            # Check for table line
            elif '|' in line and line.strip():
                # Save text buffer before table
                save_text_buffer()

                in_table = True
                table_lines.append(line)

            # If we were in a table and hit a non-table line, save the table
            elif in_table and not line.strip():
                table_data = self._parse_table(table_lines)
                if table_data and current_sheet is not None:
                    table_data['type'] = 'table'
                    current_sheet['items'].append(table_data)
                table_lines = []
                in_table = False

            # Regular text line
            elif line.strip() and not in_table and current_sheet is not None:
                text_buffer.append(line.strip())

        # Save final text buffer
        save_text_buffer()

        # Save final table if exists
        if in_table and table_lines:
            table_data = self._parse_table(table_lines)
            if table_data and current_sheet is not None:
                table_data['type'] = 'table'
                current_sheet['items'].append(table_data)

        # Save final code block if exists
        if in_code_block and code_block_lines:
            save_code_block()

        return sheets

    def _parse_table(self, lines: List[str]) -> Dict:
        """
        Parse table lines and extract header and data rows.

        Args:
            lines: List of table lines

        Returns:
            Dictionary containing table data
        """
        if len(lines) < 2:
            return {}

        # Parse header
        header_line = lines[0]
        headers = self._parse_row(header_line)

        # Skip separator line (line with hyphens)
        # Parse data rows
        data_rows = []
        for line in lines[2:]:
            if not line.strip():
                continue
            row_data = self._parse_row(line)
            if row_data:
                data_rows.append(row_data)

        return {
            'headers': headers,
            'rows': data_rows
        }

    def _parse_row(self, line: str) -> List[str]:
        """
        Parse a table row and extract cell contents.

        Args:
            line: Table row line

        Returns:
            List of cell contents
        """
        # Remove leading and trailing pipes
        line = line.strip()
        if line.startswith('|'):
            line = line[1:]
        if line.endswith('|'):
            line = line[:-1]

        # Split by pipe and clean up cells
        cells = [cell.strip() for cell in line.split('|')]

        # Process cell content (handle lists)
        processed_cells = []
        for cell in cells:
            processed_cell = self._process_cell_content(cell)
            processed_cells.append(processed_cell)

        return processed_cells

    def _process_cell_content(self, content: str) -> str:
        """
        Process cell content to handle lists and formatting.

        Args:
            content: Cell content

        Returns:
            Processed cell content
        """
        # Handle bullet lists (- ) -> ・
        # Handle numbered lists (1. ) -> 1)

        # First, split by <br> tags if they exist
        lines = content.split('<br>')
        processed_lines = []

        for line in lines:
            line = line.strip()

            # Check if line contains multiple list items separated by " - " or similar patterns
            # Handle multiple bullet points in a single line
            if ' - ' in line or line.startswith('- '):
                # Remove leading '- ' if present
                if line.startswith('- '):
                    line = line[2:]

                # Split by ' - ' pattern and filter out empty strings
                bullet_items = [item.strip() for item in line.split(' - ') if item.strip()]

                for item in bullet_items:
                    # Check if item starts with a number (numbered list)
                    if re.match(r'^\d+\.', item):
                        numbered_match = re.match(r'^(\d+)\.\s*(.+)', item)
                        if numbered_match:
                            num = numbered_match.group(1)
                            text = numbered_match.group(2)
                            processed_lines.append(f'{num}) {text}')
                    else:
                        # Regular bullet list item
                        processed_lines.append(f'・ {item}')
                continue

            # Handle numbered lists: "1. Item 2. Item"
            if re.search(r'\d+\.\s+', line):
                # Split by numbered list pattern
                numbered_items = re.split(r'(\d+\.\s+)', line)
                current_item = ''
                for part in numbered_items:
                    if re.match(r'^\d+\.\s+$', part):
                        # This is a number marker
                        if current_item:
                            processed_lines.append(current_item.strip())
                        num = part.strip().rstrip('.')
                        current_item = f'{num})'
                    else:
                        # This is content
                        if current_item and part.strip():
                            current_item += ' ' + part.strip()
                        elif part.strip() and not current_item:
                            current_item = part.strip()

                if current_item:
                    processed_lines.append(current_item.strip())
                continue

            # No special pattern found, add line as-is if not empty
            if line:
                processed_lines.append(line)

        return '\n'.join(processed_lines) if processed_lines else content
