"""Data models for cells"""

from dataclasses import dataclass, asdict
from typing import Any, Optional
from .formatting import CellFormattingModel


@dataclass
class CellModel:
    """Represents a cell in Excel"""

    coordinate: str  # e.g., "A1"
    row: int
    column: int
    column_letter: str
    value: Any
    data_type: str  # 's' (string), 'n' (number), 'b' (boolean), 'f' (formula), 'e' (error), 'd' (date)
    number_format: str = "General"
    formula: Optional[str] = None
    calculated_value: Optional[Any] = None  # For formula cells
    is_merged: bool = False
    formatting: Optional[CellFormattingModel] = None
    hyperlink: Optional[str] = None
    comment: Optional[str] = None

    def to_dict(self):
        result = {
            'coordinate': self.coordinate,
            'row': self.row,
            'column': self.column,
            'column_letter': self.column_letter,
            'value': self.value,
            'data_type': self.data_type,
            'number_format': self.number_format,
        }

        if self.formula:
            result['formula'] = self.formula
        if self.calculated_value is not None:
            result['calculated_value'] = self.calculated_value
        if self.is_merged:
            result['is_merged'] = self.is_merged
        if self.formatting:
            result['formatting'] = self.formatting.to_dict()
        if self.hyperlink:
            result['hyperlink'] = self.hyperlink
        if self.comment:
            result['comment'] = self.comment

        return result
