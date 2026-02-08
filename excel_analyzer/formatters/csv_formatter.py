"""CSV output formatter"""

import csv
import logging
from pathlib import Path
from typing import List
from .base_formatter import BaseFormatter
from ..models.workbook import WorkbookModel
from ..utils.logging_utils import get_logger

logger = get_logger(__name__)


class CSVFormatter(BaseFormatter):
    """Formats workbook cell data as CSV files (one per worksheet)"""

    def format(self, workbook_model: WorkbookModel, output_path: str, verbose: bool = False) -> List[Path]:
        """
        Generate CSV output (one file per worksheet).

        Args:
            workbook_model: WorkbookModel to format
            output_path: Output directory or base path
            verbose: If True, log progress

        Returns:
            List of output file paths
        """
        if verbose:
            logger.info(f"Generating CSV output")

        output_files = []
        output_dir = Path(output_path)

        # If output_path is a file path, get its directory
        if output_dir.suffix:
            base_name = output_dir.stem
            output_dir = output_dir.parent
        else:
            # Use source filename as base
            base_name = Path(workbook_model.file_path).stem

        # Ensure output directory exists
        output_dir.mkdir(parents=True, exist_ok=True)

        # Generate one CSV per worksheet
        for ws in workbook_model.worksheets:
            # Sanitize worksheet name for filename
            safe_sheet_name = "".join(c if c.isalnum() or c in (' ', '_', '-') else '_' for c in ws.name)

            csv_path = output_dir / f"{base_name}_{safe_sheet_name}.csv"

            if verbose:
                logger.info(f"Writing CSV for worksheet '{ws.name}': {csv_path}")

            self._write_worksheet_csv(ws, csv_path)
            output_files.append(csv_path)

        if verbose:
            logger.info(f"Generated {len(output_files)} CSV files")

        return output_files

    def _write_worksheet_csv(self, worksheet, output_path: Path):
        """Write a single worksheet to CSV"""

        # Find the maximum row and column
        max_row = 0
        max_col = 0
        for cell in worksheet.cells:
            max_row = max(max_row, cell.row)
            max_col = max(max_col, cell.column)

        # Create empty grid
        grid = [['' for _ in range(max_col)] for _ in range(max_row)]

        # Fill grid with cell values
        for cell in worksheet.cells:
            row_idx = cell.row - 1
            col_idx = cell.column - 1

            # Use the value, or calculated_value for formulas
            if cell.formula and cell.calculated_value is not None:
                value = cell.calculated_value
            else:
                value = cell.value

            # Convert to string
            if value is None:
                value = ''
            else:
                value = str(value)

            grid[row_idx][col_idx] = value

        # Write to CSV
        with open(output_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerows(grid)
