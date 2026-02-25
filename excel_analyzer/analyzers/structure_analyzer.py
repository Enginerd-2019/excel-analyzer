"""Analyzer for worksheet structure (merged cells, dimensions, print settings)"""

import logging
from typing import List, Dict
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter
from ..models.worksheet import (
    ColumnDimensionModel,
    RowDimensionModel,
    PrintSettingsModel,
    HeaderFooterModel,
)
from ..utils.logging_utils import get_logger

logger = get_logger(__name__)


class StructureAnalyzer:
    """Extracts worksheet structure information"""

    def extract_merged_cells(self, worksheet: Worksheet) -> List[str]:
        """
        Extract merged cell ranges.

        Args:
            worksheet: openpyxl Worksheet object

        Returns:
            List of merged cell range strings (e.g., ["A1:B2", "C3:D4"])
        """
        merged_cells = []
        for merged_range in worksheet.merged_cells.ranges:
            merged_cells.append(str(merged_range))
        return merged_cells

    def extract_column_dimensions(self, worksheet: Worksheet) -> Dict[str, ColumnDimensionModel]:
        """
        Extract column widths and properties.

        Args:
            worksheet: openpyxl Worksheet object

        Returns:
            Dictionary mapping column letters to ColumnDimensionModel
        """
        column_dimensions = {}

        for col_letter, col_dim in worksheet.column_dimensions.items():
            column_dimensions[col_letter] = ColumnDimensionModel(
                column=col_letter,
                width=col_dim.width or 8.43,  # Default Excel column width
                hidden=col_dim.hidden or False,
                custom_width=col_dim.customWidth or False,
            )

        return column_dimensions

    def extract_row_dimensions(self, worksheet: Worksheet) -> Dict[int, RowDimensionModel]:
        """
        Extract row heights and properties.

        Args:
            worksheet: openpyxl Worksheet object

        Returns:
            Dictionary mapping row numbers to RowDimensionModel
        """
        row_dimensions = {}

        for row_num, row_dim in worksheet.row_dimensions.items():
            row_dimensions[row_num] = RowDimensionModel(
                row=row_num,
                height=row_dim.height or 15.0,  # Default Excel row height
                hidden=row_dim.hidden or False,
                custom_height=row_dim.customHeight or False,
            )

        return row_dimensions

    def extract_print_settings(self, worksheet: Worksheet) -> PrintSettingsModel:
        """
        Extract print settings and page setup.

        Args:
            worksheet: openpyxl Worksheet object

        Returns:
            PrintSettingsModel with print settings
        """
        page_setup = worksheet.page_setup
        page_margins = worksheet.page_margins
        print_options = worksheet.print_options

        # Extract print titles (repeat rows/columns)
        print_titles_rows = None
        print_titles_cols = None
        if worksheet.print_titles:
            print_titles = worksheet.print_titles
            if print_titles:
                # Parse print titles (can be row range, column range, or both)
                parts = print_titles.split(',')
                for part in parts:
                    part = part.strip()
                    if '$' in part:
                        # Remove sheet reference
                        if '!' in part:
                            part = part.split('!')[1]
                        # Check if it's a row or column range
                        if ':$' in part or part.startswith('$'):
                            if part[1].isdigit():
                                print_titles_rows = part
                            else:
                                print_titles_cols = part

        return PrintSettingsModel(
            orientation=page_setup.orientation or "portrait",
            paper_size=page_setup.paperSize,
            scale=page_setup.scale or 100,
            fit_to_width=page_setup.fitToWidth,
            fit_to_height=page_setup.fitToHeight,
            margin_left=page_margins.left,
            margin_right=page_margins.right,
            margin_top=page_margins.top,
            margin_bottom=page_margins.bottom,
            margin_header=page_margins.header,
            margin_footer=page_margins.footer,
            print_area=worksheet.print_area.ref if (worksheet.print_area and hasattr(worksheet.print_area, 'ref')) else (worksheet.print_area if isinstance(worksheet.print_area, str) else None),
            print_titles_rows=print_titles_rows,
            print_titles_cols=print_titles_cols,
            print_gridlines=print_options.gridLines or False,
            print_headings=print_options.headings or False,
        )

    def extract_header_footer(self, worksheet: Worksheet) -> HeaderFooterModel:
        """
        Extract header and footer settings.

        Args:
            worksheet: openpyxl Worksheet object

        Returns:
            HeaderFooterModel with header/footer text
        """
        hf = worksheet.HeaderFooter

        return HeaderFooterModel(
            odd_header=hf.oddHeader,
            odd_footer=hf.oddFooter,
            even_header=hf.evenHeader,
            even_footer=hf.evenFooter,
            first_header=hf.firstHeader,
            first_footer=hf.firstFooter,
            different_odd_even=hf.differentOddEven or False,
            different_first=hf.differentFirst or False,
            scale_with_doc=hf.scaleWithDoc if hf.scaleWithDoc is not None else True,
            align_with_margins=hf.alignWithMargins if hf.alignWithMargins is not None else True,
        )

    def extract_freeze_panes(self, worksheet: Worksheet) -> str:
        """
        Extract freeze panes setting.

        Args:
            worksheet: openpyxl Worksheet object

        Returns:
            Cell coordinate of freeze panes or None
        """
        if worksheet.freeze_panes:
            return worksheet.freeze_panes
        return None

    def extract_auto_filter(self, worksheet: Worksheet) -> str:
        """
        Extract auto filter range.

        Args:
            worksheet: openpyxl Worksheet object

        Returns:
            Cell range string or None
        """
        if worksheet.auto_filter and worksheet.auto_filter.ref:
            return worksheet.auto_filter.ref
        return None

    def extract_tab_color(self, worksheet: Worksheet) -> str:
        """
        Extract worksheet tab color.

        Args:
            worksheet: openpyxl Worksheet object

        Returns:
            RGB color string or None
        """
        if worksheet.sheet_properties and worksheet.sheet_properties.tabColor:
            from ..utils.color_utils import convert_color
            color_model = convert_color(worksheet.sheet_properties.tabColor)
            if color_model:
                return color_model.value
        return None
