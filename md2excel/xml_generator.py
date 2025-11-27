"""Office Open XML generator for creating Excel-compatible XML files."""

import xml.etree.ElementTree as ET
from xml.dom import minidom
from typing import List, Dict


class XMLGenerator:
    """Generate Office Open XML format files from parsed table data."""

    # XML namespaces
    NS = {
        'ss': 'urn:schemas-microsoft-com:office:spreadsheet',
        'o': 'urn:schemas-microsoft-com:office:office',
        'x': 'urn:schemas-microsoft-com:office:excel',
        'html': 'http://www.w3.org/TR/REC-html40'
    }

    def __init__(self):
        """Initialize XML generator."""
        self._register_namespaces()

    def _register_namespaces(self):
        """Register XML namespaces."""
        for prefix, uri in self.NS.items():
            ET.register_namespace(prefix, uri)

    def generate_xml(self, sheets: List[Dict], output_path: str):
        """
        Generate Office Open XML file from sheet data.

        Args:
            sheets: List of sheet dictionaries containing table data
            output_path: Path to output XML file
        """
        # Create root Workbook element
        workbook = ET.Element(f"{{{self.NS['ss']}}}Workbook")
        workbook.set(f"{{{self.NS['ss']}}}xmlns", self.NS['ss'])

        # Add document properties
        self._add_document_properties(workbook)

        # Add styles
        self._add_styles(workbook)

        # Add worksheets
        for sheet in sheets:
            self._add_worksheet(workbook, sheet)

        # Write to file with pretty formatting
        tree = ET.ElementTree(workbook)
        xml_string = ET.tostring(workbook, encoding='utf-8')
        pretty_xml = minidom.parseString(xml_string).toprettyxml(indent='  ', encoding='utf-8')

        with open(output_path, 'wb') as f:
            f.write(pretty_xml)

    def _add_document_properties(self, workbook):
        """Add document properties to workbook."""
        doc_props = ET.SubElement(workbook, f"{{{self.NS['o']}}}DocumentProperties")
        ET.SubElement(doc_props, f"{{{self.NS['o']}}}Author").text = "md2excelxml"
        ET.SubElement(doc_props, f"{{{self.NS['o']}}}Created").text = "2025-01-01T00:00:00Z"

        excel_workbook = ET.SubElement(workbook, f"{{{self.NS['x']}}}ExcelWorkbook")
        ET.SubElement(excel_workbook, f"{{{self.NS['x']}}}WindowHeight").text = "12000"
        ET.SubElement(excel_workbook, f"{{{self.NS['x']}}}WindowWidth").text = "24000"

    def _add_styles(self, workbook):
        """Add cell styles to workbook."""
        styles = ET.SubElement(workbook, f"{{{self.NS['ss']}}}Styles")

        # Default style
        default_style = ET.SubElement(styles, f"{{{self.NS['ss']}}}Style")
        default_style.set(f"{{{self.NS['ss']}}}ID", "Default")
        default_style.set(f"{{{self.NS['ss']}}}Name", "Normal")

        default_alignment = ET.SubElement(default_style, f"{{{self.NS['ss']}}}Alignment")
        default_alignment.set(f"{{{self.NS['ss']}}}Vertical", "Top")
        default_alignment.set(f"{{{self.NS['ss']}}}Horizontal", "Left")
        default_alignment.set(f"{{{self.NS['ss']}}}WrapText", "1")

        default_borders = ET.SubElement(default_style, f"{{{self.NS['ss']}}}Borders")
        for position in ['Left', 'Top', 'Right', 'Bottom']:
            border = ET.SubElement(default_borders, f"{{{self.NS['ss']}}}Border")
            border.set(f"{{{self.NS['ss']}}}Position", position)
            border.set(f"{{{self.NS['ss']}}}LineStyle", "Continuous")
            border.set(f"{{{self.NS['ss']}}}Weight", "1")

        # Header style (bold)
        header_style = ET.SubElement(styles, f"{{{self.NS['ss']}}}Style")
        header_style.set(f"{{{self.NS['ss']}}}ID", "Header")

        header_font = ET.SubElement(header_style, f"{{{self.NS['ss']}}}Font")
        header_font.set(f"{{{self.NS['ss']}}}Bold", "1")

        header_alignment = ET.SubElement(header_style, f"{{{self.NS['ss']}}}Alignment")
        header_alignment.set(f"{{{self.NS['ss']}}}Vertical", "Top")
        header_alignment.set(f"{{{self.NS['ss']}}}Horizontal", "Left")
        header_alignment.set(f"{{{self.NS['ss']}}}WrapText", "1")

        header_borders = ET.SubElement(header_style, f"{{{self.NS['ss']}}}Borders")
        for position in ['Left', 'Top', 'Right', 'Bottom']:
            border = ET.SubElement(header_borders, f"{{{self.NS['ss']}}}Border")
            border.set(f"{{{self.NS['ss']}}}Position", position)
            border.set(f"{{{self.NS['ss']}}}LineStyle", "Continuous")
            border.set(f"{{{self.NS['ss']}}}Weight", "1")

        # Heading2 style (bold, larger font, no borders, no wrap)
        heading2_style = ET.SubElement(styles, f"{{{self.NS['ss']}}}Style")
        heading2_style.set(f"{{{self.NS['ss']}}}ID", "Heading2")

        heading2_font = ET.SubElement(heading2_style, f"{{{self.NS['ss']}}}Font")
        heading2_font.set(f"{{{self.NS['ss']}}}Bold", "1")
        heading2_font.set(f"{{{self.NS['ss']}}}Size", "13")

        heading2_alignment = ET.SubElement(heading2_style, f"{{{self.NS['ss']}}}Alignment")
        heading2_alignment.set(f"{{{self.NS['ss']}}}Vertical", "Top")
        heading2_alignment.set(f"{{{self.NS['ss']}}}Horizontal", "Left")
        heading2_alignment.set(f"{{{self.NS['ss']}}}WrapText", "0")

        # Text style (no borders, no wrap)
        text_style = ET.SubElement(styles, f"{{{self.NS['ss']}}}Style")
        text_style.set(f"{{{self.NS['ss']}}}ID", "Text")

        text_alignment = ET.SubElement(text_style, f"{{{self.NS['ss']}}}Alignment")
        text_alignment.set(f"{{{self.NS['ss']}}}Vertical", "Top")
        text_alignment.set(f"{{{self.NS['ss']}}}Horizontal", "Left")
        text_alignment.set(f"{{{self.NS['ss']}}}WrapText", "0")

    def _add_worksheet(self, workbook, sheet: Dict):
        """
        Add a worksheet to the workbook.

        Args:
            workbook: Workbook element
            sheet: Sheet dictionary containing table data
        """
        worksheet = ET.SubElement(workbook, f"{{{self.NS['ss']}}}Worksheet")
        worksheet.set(f"{{{self.NS['ss']}}}Name", self._sanitize_sheet_name(sheet['name']))

        table = ET.SubElement(worksheet, f"{{{self.NS['ss']}}}Table")

        # Calculate maximum number of columns across all items
        max_cols = 1  # At least 1 column for text
        for item in sheet.get('items', []):
            if item.get('type') == 'table':
                if item.get('headers'):
                    max_cols = max(max_cols, len(item['headers']))
                for row in item.get('rows', []):
                    max_cols = max(max_cols, len(row))

        # Add column definitions with auto-width
        for _ in range(max_cols):
            col = ET.SubElement(table, f"{{{self.NS['ss']}}}Column")
            col.set(f"{{{self.NS['ss']}}}AutoFitWidth", "1")
            col.set(f"{{{self.NS['ss']}}}Width", "100")

        # Add items (tables, text, and headings)
        for item in sheet.get('items', []):
            item_type = item.get('type', 'table')

            if item_type == 'heading':
                # Add heading as a single row with Heading2 style
                heading_row = ET.SubElement(table, f"{{{self.NS['ss']}}}Row")
                self._add_cell(heading_row, item.get('content', ''), "Heading2")
            elif item_type == 'text':
                # Add text as a single row with Text style
                text_row = ET.SubElement(table, f"{{{self.NS['ss']}}}Row")
                self._add_cell(text_row, item.get('content', ''), "Text")
            elif item_type == 'table':
                self._add_table_rows(table, item)
                # Add empty row after table
                empty_row = ET.SubElement(table, f"{{{self.NS['ss']}}}Row")

    def _add_table_rows(self, table, table_data: Dict):
        """
        Add table rows to the worksheet table.

        Args:
            table: Table element
            table_data: Table data dictionary
        """
        # Add header row
        if table_data.get('headers'):
            header_row = ET.SubElement(table, f"{{{self.NS['ss']}}}Row")
            for header in table_data['headers']:
                self._add_cell(header_row, header, "Header")

        # Add data rows
        for row_data in table_data.get('rows', []):
            data_row = ET.SubElement(table, f"{{{self.NS['ss']}}}Row")
            for cell_value in row_data:
                self._add_cell(data_row, cell_value, "Default")

    def _add_cell(self, row, value: str, style_id: str):
        """
        Add a cell to a row.

        Args:
            row: Row element
            value: Cell value
            style_id: Style ID to apply
        """
        cell = ET.SubElement(row, f"{{{self.NS['ss']}}}Cell")
        cell.set(f"{{{self.NS['ss']}}}StyleID", style_id)

        data = ET.SubElement(cell, f"{{{self.NS['ss']}}}Data")
        data.set(f"{{{self.NS['ss']}}}Type", "String")
        data.text = value if value else ""

    def _sanitize_sheet_name(self, name: str) -> str:
        """
        Sanitize sheet name to comply with Excel requirements.

        Args:
            name: Original sheet name

        Returns:
            Sanitized sheet name
        """
        # Excel sheet names cannot contain: : \ / ? * [ ]
        # Maximum length is 31 characters
        invalid_chars = [':', '\\', '/', '?', '*', '[', ']']
        for char in invalid_chars:
            name = name.replace(char, '_')

        # Truncate to 31 characters
        if len(name) > 31:
            name = name[:31]

        return name
