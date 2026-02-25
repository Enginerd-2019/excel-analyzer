"""Text output formatter"""

import logging
from pathlib import Path
from tabulate import tabulate
from .base_formatter import BaseFormatter
from ..models.workbook import WorkbookModel
from ..utils.logging_utils import get_logger
from .. import __version__

logger = get_logger(__name__)


class TextFormatter(BaseFormatter):
    """Formats workbook analysis as human-readable text"""

    def format(self, workbook_model: WorkbookModel, output_path: str, verbose: bool = False):
        """
        Generate text output.

        Args:
            workbook_model: WorkbookModel to format
            output_path: Output file path
            verbose: If True, log progress

        Returns:
            Output file path
        """
        if verbose:
            logger.info(f"Generating text output: {output_path}")

        output_lines = []

        # Header
        output_lines.append("=" * 80)
        output_lines.append("EXCEL WORKBOOK ANALYSIS")
        output_lines.append("=" * 80)
        output_lines.append("")

        # File info
        output_lines.append(f"File: {workbook_model.file_path}")
        output_lines.append(f"Format: {workbook_model.file_format.upper()}")
        output_lines.append(f"Analyzer Version: {__version__}")
        output_lines.append("")

        # Properties
        output_lines.append("-" * 80)
        output_lines.append("WORKBOOK PROPERTIES")
        output_lines.append("-" * 80)
        props = workbook_model.properties
        if props.title:
            output_lines.append(f"Title: {props.title}")
        if props.creator:
            output_lines.append(f"Creator: {props.creator}")
        if props.created:
            output_lines.append(f"Created: {props.created}")
        if props.modified:
            output_lines.append(f"Modified: {props.modified}")
        output_lines.append(f"Total Worksheets: {len(workbook_model.worksheets)}")
        output_lines.append("")

        # Defined Names
        if workbook_model.defined_names:
            output_lines.append("-" * 80)
            output_lines.append("DEFINED NAMES")
            output_lines.append("-" * 80)
            for dn in workbook_model.defined_names:
                output_lines.append(f"  {dn.name} = {dn.value}")
            output_lines.append("")

        # Worksheets
        for ws in workbook_model.worksheets:
            output_lines.append("=" * 80)
            output_lines.append(f"WORKSHEET: {ws.name}")
            output_lines.append("=" * 80)
            output_lines.append("")

            # Basic info
            output_lines.append(f"Sheet Index: {ws.index}")
            output_lines.append(f"Total Cells: {len(ws.cells)}")
            output_lines.append(f"Sheet State: {ws.sheet_state}")
            if ws.tab_color:
                output_lines.append(f"Tab Color: {ws.tab_color}")
            output_lines.append("")

            # Merged cells
            if ws.merged_cells:
                output_lines.append(f"Merged Cells ({len(ws.merged_cells)}):")
                for mc in ws.merged_cells:
                    output_lines.append(f"  - {mc}")
                output_lines.append("")

            # Freeze panes
            if ws.freeze_panes:
                output_lines.append(f"Freeze Panes: {ws.freeze_panes}")
                output_lines.append("")

            # Auto filter
            if ws.auto_filter:
                output_lines.append(f"Auto Filter: {ws.auto_filter}")
                output_lines.append("")

            # Cell data (first 100 non-empty cells)
            if ws.cells:
                output_lines.append(f"Cell Data (showing first 100 cells):")
                output_lines.append("")

                # Prepare table data
                table_data = []
                for idx, cell in enumerate(ws.cells[:100]):
                    value_str = str(cell.value)[:50]  # Truncate long values
                    if len(str(cell.value)) > 50:
                        value_str += "..."

                    formula_str = ""
                    if cell.formula:
                        formula_str = str(cell.formula)[:50]

                    table_data.append([
                        cell.coordinate,
                        cell.data_type,
                        value_str,
                        formula_str,
                        cell.number_format,
                    ])

                output_lines.append(tabulate(
                    table_data,
                    headers=["Cell", "Type", "Value", "Formula", "Format"],
                    tablefmt="grid"
                ))
                output_lines.append("")

                if len(ws.cells) > 100:
                    output_lines.append(f"... and {len(ws.cells) - 100} more cells")
                    output_lines.append("")

            # Cell Formatting (colors, fonts)
            cells_with_formatting = []
            for cell in ws.cells:
                if not cell.formatting:
                    continue

                formatting_details = []

                # Check font color
                if cell.formatting.font and cell.formatting.font.color:
                    color = cell.formatting.font.color
                    if color.value and color.value not in ['#000000', None]:
                        formatting_details.append(f"Font Color: {color.value}")

                # Check background color
                if cell.formatting.fill and cell.formatting.fill.fg_color:
                    if cell.formatting.fill.pattern_type != 'none':
                        color = cell.formatting.fill.fg_color
                        if color.value and color.value not in ['#000000', '#FFFFFF', None]:
                            formatting_details.append(f"Background: {color.value}")

                # Check bold/italic
                if cell.formatting.font:
                    if cell.formatting.font.bold:
                        formatting_details.append("Bold")
                    if cell.formatting.font.italic:
                        formatting_details.append("Italic")

                if formatting_details:
                    cells_with_formatting.append((cell.coordinate, ', '.join(formatting_details)))

            if cells_with_formatting:
                output_lines.append(f"Cell Formatting ({len(cells_with_formatting)} cells with custom formatting):")
                for coord, details in cells_with_formatting[:50]:
                    output_lines.append(f"  {coord}: {details}")
                if len(cells_with_formatting) > 50:
                    output_lines.append(f"  ... and {len(cells_with_formatting) - 50} more cells")
                output_lines.append("")

            # Data validations
            if ws.data_validations:
                output_lines.append(f"Data Validations ({len(ws.data_validations)}):")
                for dv in ws.data_validations:
                    output_lines.append(f"  Range: {dv.sqref}")
                    output_lines.append(f"    Type: {dv.validation_type}")
                    if dv.formula1:
                        output_lines.append(f"    Formula: {dv.formula1}")
                output_lines.append("")

            # Conditional formatting
            if ws.conditional_formatting:
                output_lines.append(f"Conditional Formatting ({len(ws.conditional_formatting)}):")
                for cf in ws.conditional_formatting:
                    output_lines.append(f"  Range: {cf.sqref}")
                    output_lines.append(f"    Type: {cf.rule_type}")
                    output_lines.append(f"    Priority: {cf.priority}")
                output_lines.append("")

            # Charts
            if ws.charts:
                output_lines.append(f"Charts ({len(ws.charts)}):")
                for idx, chart in enumerate(ws.charts, 1):
                    output_lines.append(f"  Chart {idx}:")
                    output_lines.append(f"    Type: {chart.chart_type}")
                    if chart.title:
                        output_lines.append(f"    Title: {chart.title}")
                    output_lines.append(f"    Series Count: {len(chart.series)}")
                    for s_idx, series in enumerate(chart.series, 1):
                        output_lines.append(f"      Series {s_idx}:")
                        if series.title:
                            output_lines.append(f"        Title: {series.title}")
                        if series.values:
                            output_lines.append(f"        Values: {series.values}")
                        if series.categories:
                            output_lines.append(f"        Categories: {series.categories}")
                output_lines.append("")

            # Images
            if ws.images:
                output_lines.append(f"Images ({len(ws.images)}):")
                for idx, img in enumerate(ws.images, 1):
                    output_lines.append(f"  Image {idx}:")
                    output_lines.append(f"    Format: {img.format}")
                    output_lines.append(f"    Size: {img.width}x{img.height}")
                    output_lines.append(f"    Anchor: {img.anchor}")
                    output_lines.append(f"    Data Length: {len(img.data)} bytes")
                output_lines.append("")

            # Print settings
            if ws.print_settings:
                ps = ws.print_settings
                output_lines.append("Print Settings:")
                output_lines.append(f"  Orientation: {ps.orientation}")
                output_lines.append(f"  Scale: {ps.scale}%")
                if ps.print_area:
                    output_lines.append(f"  Print Area: {ps.print_area}")
                output_lines.append("")

            output_lines.append("")

        # Write to file
        output_path = Path(output_path)
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(output_lines))

        if verbose:
            logger.info(f"Text output written to: {output_path}")

        return output_path
