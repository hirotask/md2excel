"""Excel generator using Spire.XLS for creating Excel files."""

from typing import List, Dict
from spire.xls import Workbook, Worksheet, CellRange, HorizontalAlignType, VerticalAlignType
from spire.xls import ExcelVersion, LineStyleType, BordersLineType, Color


class ExcelGenerator:
    """Generate Excel files from parsed table data using Spire.XLS."""

    def __init__(self):
        """Initialize Excel generator."""
        pass

    def generate_excel(self, sheets: List[Dict], output_path: str):
        """
        Generate Excel file from sheet data.

        Args:
            sheets: List of sheet dictionaries containing table data
            output_path: Path to output Excel file
        """
        # Create a new workbook
        workbook = Workbook()

        # Remove default sheet
        if workbook.Worksheets.Count > 0:
            workbook.Worksheets.Clear()

        # Add worksheets for each sheet in the data
        for sheet_data in sheets:
            self._add_worksheet(workbook, sheet_data)

        # Save the workbook as XLSX format
        workbook.SaveToFile(output_path, ExcelVersion.Version2013)

        workbook.Dispose()

    def _add_worksheet(self, workbook: Workbook, sheet_data: Dict):
        """
        Add a worksheet to the workbook.

        Args:
            workbook: Workbook object
            sheet_data: Sheet data dictionary
        """
        # Create new worksheet
        worksheet = workbook.Worksheets.Add(self._sanitize_sheet_name(sheet_data['name']))

        current_row = 1  # Start from row 1 (1-indexed in Spire.XLS)

        # Process each item in the sheet (tables and text)
        for item in sheet_data.get('items', []):
            item_type = item.get('type', 'table')

            if item_type == 'heading':
                # Add heading with Heading style
                cell = worksheet.Range[current_row, 1]
                cell.Text = item.get('content', '')

                # Apply heading style based on level
                level = item.get('level', 2)
                if level == 2:
                    # Apply Heading 2 style: larger font, bold, no borders
                    cell.Style.Font.IsBold = True
                    cell.Style.Font.Size = 13
                    cell.Style.HorizontalAlignment = HorizontalAlignType.Left
                    cell.Style.VerticalAlignment = VerticalAlignType.Top
                    cell.Style.WrapText = False

                current_row += 1

            elif item_type == 'text':
                # Add text as a single cell spanning the first column
                cell = worksheet.Range[current_row, 1]
                cell.Text = item.get('content', '')

                # Set text style (no borders, no bold, no wrapping)
                cell.Style.HorizontalAlignment = HorizontalAlignType.Left
                cell.Style.VerticalAlignment = VerticalAlignType.Top
                cell.Style.WrapText = False

                current_row += 1

            elif item_type == 'table':
                # Add header row
                if item.get('headers'):
                    for col_idx, header in enumerate(item['headers'], start=1):
                        cell = worksheet.Range[current_row, col_idx]
                        cell.Text = header

                        # Set header style: bold
                        cell.Style.Font.IsBold = True

                        # Set alignment
                        cell.Style.HorizontalAlignment = HorizontalAlignType.Left
                        cell.Style.VerticalAlignment = VerticalAlignType.Top
                        cell.Style.WrapText = True

                        # Set borders
                        self._set_cell_borders(cell)

                    current_row += 1

                # Add data rows
                for row_data in item.get('rows', []):
                    for col_idx, cell_value in enumerate(row_data, start=1):
                        cell = worksheet.Range[current_row, col_idx]
                        cell.Text = cell_value if cell_value else ""

                        # Set alignment
                        cell.Style.HorizontalAlignment = HorizontalAlignType.Left
                        cell.Style.VerticalAlignment = VerticalAlignType.Top
                        cell.Style.WrapText = True

                        # Set borders
                        self._set_cell_borders(cell)

                    current_row += 1

                # Add empty row after table
                current_row += 1

        # Auto-fit columns
        for col_idx in range(1, worksheet.Columns.Length + 1):
            try:
                worksheet.AutoFitColumn(col_idx)
            except:
                # If auto-fit fails, set a default width
                worksheet.Columns[col_idx - 1].ColumnWidth = 15

    def _set_cell_borders(self, cell: CellRange):
        """
        Set borders for a cell.

        Args:
            cell: Cell range object
        """
        # Set all borders using BordersLineType enum
        cell.Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin
        cell.Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin
        cell.Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin
        cell.Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin

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
