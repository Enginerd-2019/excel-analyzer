"""Base formatter class"""

from abc import ABC, abstractmethod
from pathlib import Path
from ..models.workbook import WorkbookModel


class BaseFormatter(ABC):
    """Abstract base class for output formatters"""

    @abstractmethod
    def format(self, workbook_model: WorkbookModel, output_path: str, verbose: bool = False):
        """
        Generate output file from workbook model.

        Args:
            workbook_model: WorkbookModel to format
            output_path: Path for output file (or directory for multi-file formats)
            verbose: If True, output detailed progress

        Returns:
            Output file path(s)
        """
        pass

    def _get_output_path(self, base_path: str, suffix: str, extension: str) -> Path:
        """
        Generate output path with suffix and extension.

        Args:
            base_path: Base output path
            suffix: Suffix to add (e.g., "_analysis")
            extension: File extension (e.g., ".json")

        Returns:
            Path object for output file
        """
        path = Path(base_path)

        # If base_path is a directory, use the workbook filename
        if path.is_dir():
            return path

        # If base_path includes a filename, use it as-is
        if path.suffix:
            return path

        # Otherwise, add suffix and extension
        return Path(f"{base_path}{suffix}{extension}")
