"""Data models for worksheets"""

from dataclasses import dataclass, field
from typing import List, Dict, Optional, Any
from .cell import CellModel
from .chart import ChartModel
from .image import ImageModel


@dataclass
class ColumnDimensionModel:
    """Represents column width"""

    column: str  # Column letter
    width: float
    hidden: bool = False
    custom_width: bool = False

    def to_dict(self):
        return {
            'column': self.column,
            'width': self.width,
            'hidden': self.hidden,
            'custom_width': self.custom_width,
        }


@dataclass
class RowDimensionModel:
    """Represents row height"""

    row: int
    height: float
    hidden: bool = False
    custom_height: bool = False

    def to_dict(self):
        return {
            'row': self.row,
            'height': self.height,
            'hidden': self.hidden,
            'custom_height': self.custom_height,
        }


@dataclass
class DataValidationModel:
    """Represents data validation rule"""

    sqref: str  # Cell range (e.g., "A1:A10")
    validation_type: str  # 'list', 'whole', 'decimal', 'date', 'time', 'textLength', 'custom'
    operator: Optional[str] = None  # 'between', 'notBetween', 'equal', etc.
    formula1: Optional[str] = None
    formula2: Optional[str] = None
    allow_blank: bool = True
    show_input_message: bool = False
    input_title: Optional[str] = None
    input_message: Optional[str] = None
    show_error_message: bool = True
    error_title: Optional[str] = None
    error_message: Optional[str] = None
    error_style: str = "stop"  # 'stop', 'warning', 'information'

    def to_dict(self):
        return {k: v for k, v in self.__dict__.items() if v is not None}


@dataclass
class ConditionalFormattingModel:
    """Represents conditional formatting rule"""

    sqref: str  # Cell range
    rule_type: str  # 'cellIs', 'expression', 'colorScale', 'dataBar', 'iconSet', etc.
    priority: int
    formula: Optional[List[str]] = None
    operator: Optional[str] = None
    stop_if_true: bool = False
    dxf_id: Optional[int] = None
    format_description: Optional[Dict[str, Any]] = None

    def to_dict(self):
        return {k: v for k, v in self.__dict__.items() if v is not None}


@dataclass
class PrintSettingsModel:
    """Represents print settings"""

    orientation: str = "portrait"  # 'portrait' or 'landscape'
    paper_size: Optional[int] = None
    scale: int = 100
    fit_to_width: Optional[int] = None
    fit_to_height: Optional[int] = None
    margin_left: float = 0.7
    margin_right: float = 0.7
    margin_top: float = 0.75
    margin_bottom: float = 0.75
    margin_header: float = 0.3
    margin_footer: float = 0.3
    print_area: Optional[str] = None
    print_titles_rows: Optional[str] = None
    print_titles_cols: Optional[str] = None
    print_gridlines: bool = False
    print_headings: bool = False

    def to_dict(self):
        return {k: v for k, v in self.__dict__.items() if v is not None}


@dataclass
class HeaderFooterModel:
    """Represents headers and footers"""

    odd_header: Optional[str] = None
    odd_footer: Optional[str] = None
    even_header: Optional[str] = None
    even_footer: Optional[str] = None
    first_header: Optional[str] = None
    first_footer: Optional[str] = None
    different_odd_even: bool = False
    different_first: bool = False
    scale_with_doc: bool = True
    align_with_margins: bool = True

    def to_dict(self):
        return {k: v for k, v in self.__dict__.items() if v is not None}


@dataclass
class WorksheetModel:
    """Represents a worksheet in Excel"""

    name: str
    index: int
    cells: List[CellModel] = field(default_factory=list)
    merged_cells: List[str] = field(default_factory=list)  # e.g., ["A1:B2", "C3:D4"]
    column_dimensions: Dict[str, ColumnDimensionModel] = field(default_factory=dict)
    row_dimensions: Dict[int, RowDimensionModel] = field(default_factory=dict)
    data_validations: List[DataValidationModel] = field(default_factory=list)
    conditional_formatting: List[ConditionalFormattingModel] = field(default_factory=list)
    charts: List[ChartModel] = field(default_factory=list)
    images: List[ImageModel] = field(default_factory=list)
    print_settings: Optional[PrintSettingsModel] = None
    header_footer: Optional[HeaderFooterModel] = None
    freeze_panes: Optional[str] = None  # Cell coordinate (e.g., "B2")
    auto_filter: Optional[str] = None  # Cell range (e.g., "A1:D10")
    tab_color: Optional[str] = None  # RGB color
    sheet_state: str = "visible"  # 'visible', 'hidden', 'veryHidden'
    sheet_view: Optional[Dict[str, Any]] = None  # View settings (zoom, etc.)

    def to_dict(self):
        result = {
            'name': self.name,
            'index': self.index,
            'cells': [cell.to_dict() for cell in self.cells],
            'merged_cells': self.merged_cells,
            'column_dimensions': {
                col: dim.to_dict() for col, dim in self.column_dimensions.items()
            },
            'row_dimensions': {
                row: dim.to_dict() for row, dim in self.row_dimensions.items()
            },
            'data_validations': [dv.to_dict() for dv in self.data_validations],
            'conditional_formatting': [cf.to_dict() for cf in self.conditional_formatting],
            'charts': [chart.to_dict() for chart in self.charts],
            'images': [img.to_dict() for img in self.images],
        }

        if self.print_settings:
            result['print_settings'] = self.print_settings.to_dict()
        if self.header_footer:
            result['header_footer'] = self.header_footer.to_dict()
        if self.freeze_panes:
            result['freeze_panes'] = self.freeze_panes
        if self.auto_filter:
            result['auto_filter'] = self.auto_filter
        if self.tab_color:
            result['tab_color'] = self.tab_color
        result['sheet_state'] = self.sheet_state
        if self.sheet_view:
            result['sheet_view'] = self.sheet_view

        return result
