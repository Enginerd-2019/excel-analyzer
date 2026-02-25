"""JSON output formatter"""

import json
import logging
from datetime import datetime
from pathlib import Path
from .base_formatter import BaseFormatter
from ..models.workbook import WorkbookModel
from ..utils.logging_utils import get_logger
from .. import __version__

logger = get_logger(__name__)


class JSONFormatter(BaseFormatter):
    """Formats workbook analysis as JSON"""

    def format(self, workbook_model: WorkbookModel, output_path: str, verbose: bool = False):
        """
        Generate JSON output.

        Args:
            workbook_model: WorkbookModel to format
            output_path: Output file path
            verbose: If True, log progress

        Returns:
            Output file path
        """
        if verbose:
            logger.info(f"Generating JSON output: {output_path}")

        # Create output data structure
        output_data = {
            "metadata": {
                "analyzer_version": __version__,
                "analyzed_at": datetime.now().isoformat(),
                "source_file": workbook_model.file_path,
            },
            "workbook": workbook_model.to_dict(),
        }

        # Write JSON file
        output_path = Path(output_path)
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, indent=2, ensure_ascii=False, default=str)

        if verbose:
            logger.info(f"JSON output written to: {output_path}")

        return output_path
