"""Base analyzer class"""

from abc import ABC, abstractmethod
from ..models.workbook import WorkbookModel


class BaseAnalyzer(ABC):
    """Abstract base class for Excel analyzers"""

    @abstractmethod
    def analyze(self, file_path: str, verbose: bool = False) -> WorkbookModel:
        """
        Analyze an Excel file and return a WorkbookModel.

        Args:
            file_path: Path to Excel file
            verbose: If True, output detailed progress information

        Returns:
            WorkbookModel containing all extracted data
        """
        pass
