"""Data models for charts"""

from dataclasses import dataclass, asdict
from typing import Optional, List, Dict, Any
from .formatting import ColorModel


@dataclass
class ChartSeriesModel:
    """Represents a chart data series"""

    title: Optional[str] = None
    values: Optional[str] = None  # Cell range reference
    categories: Optional[str] = None  # Cell range reference
    color: Optional[ColorModel] = None
    marker_style: Optional[str] = None
    line_style: Optional[str] = None

    def to_dict(self):
        result = asdict(self)
        if self.color:
            result['color'] = self.color.to_dict()
        return {k: v for k, v in result.items() if v is not None}


@dataclass
class AxisModel:
    """Represents a chart axis"""

    title: Optional[str] = None
    min_value: Optional[float] = None
    max_value: Optional[float] = None
    major_unit: Optional[float] = None
    minor_unit: Optional[float] = None
    number_format: Optional[str] = None
    axis_position: Optional[str] = None  # 'left', 'right', 'top', 'bottom'
    crosses: Optional[str] = None
    delete: bool = False  # If True, axis is hidden

    def to_dict(self):
        return {k: v for k, v in asdict(self).items() if v is not None}


@dataclass
class LegendModel:
    """Represents chart legend"""

    position: str = "right"  # 'top', 'bottom', 'left', 'right', 'top_right'
    overlay: bool = False

    def to_dict(self):
        return asdict(self)


@dataclass
class ChartPositionModel:
    """Represents chart position on worksheet"""

    anchor: str  # Cell coordinate (e.g., "E5")
    x_offset: int = 0
    y_offset: int = 0
    width: Optional[int] = None
    height: Optional[int] = None

    def to_dict(self):
        return {k: v for k, v in asdict(self).items() if v is not None}


@dataclass
class ChartModel:
    """Represents a chart in Excel"""

    chart_type: str  # 'bar', 'line', 'pie', 'scatter', 'area', etc.
    title: Optional[str] = None
    series: List[ChartSeriesModel] = None
    x_axis: Optional[AxisModel] = None
    y_axis: Optional[AxisModel] = None
    legend: Optional[LegendModel] = None
    position: Optional[ChartPositionModel] = None
    style: Optional[int] = None

    def __post_init__(self):
        if self.series is None:
            self.series = []

    def to_dict(self):
        result = {
            'chart_type': self.chart_type,
            'title': self.title,
            'series': [s.to_dict() for s in self.series],
        }
        if self.x_axis:
            result['x_axis'] = self.x_axis.to_dict()
        if self.y_axis:
            result['y_axis'] = self.y_axis.to_dict()
        if self.legend:
            result['legend'] = self.legend.to_dict()
        if self.position:
            result['position'] = self.position.to_dict()
        if self.style is not None:
            result['style'] = self.style
        return result
