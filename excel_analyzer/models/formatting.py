"""Data models for cell formatting"""

from dataclasses import dataclass, asdict
from typing import Optional


@dataclass
class ColorModel:
    """Represents a color in Excel"""

    type: str  # 'rgb', 'theme', 'auto', 'indexed'
    value: Optional[str] = None  # hex color or theme/index number
    tint: Optional[float] = None  # tint/shade modification

    def to_dict(self):
        return {k: v for k, v in asdict(self).items() if v is not None}


@dataclass
class FontModel:
    """Represents font formatting"""

    name: str = "Calibri"
    size: float = 11.0
    bold: bool = False
    italic: bool = False
    underline: str = "none"  # 'none', 'single', 'double', etc.
    strike: bool = False
    color: Optional[ColorModel] = None

    def to_dict(self):
        result = asdict(self)
        if self.color:
            result['color'] = self.color.to_dict()
        return result


@dataclass
class FillModel:
    """Represents fill/background formatting"""

    pattern_type: str = "none"  # 'none', 'solid', 'gray125', etc.
    fg_color: Optional[ColorModel] = None  # foreground color
    bg_color: Optional[ColorModel] = None  # background color

    def to_dict(self):
        result = asdict(self)
        if self.fg_color:
            result['fg_color'] = self.fg_color.to_dict()
        if self.bg_color:
            result['bg_color'] = self.bg_color.to_dict()
        return result


@dataclass
class BorderSideModel:
    """Represents one side of a border"""

    style: str = "none"  # 'none', 'thin', 'medium', 'thick', etc.
    color: Optional[ColorModel] = None

    def to_dict(self):
        result = asdict(self)
        if self.color:
            result['color'] = self.color.to_dict()
        return result


@dataclass
class BorderModel:
    """Represents cell borders"""

    left: Optional[BorderSideModel] = None
    right: Optional[BorderSideModel] = None
    top: Optional[BorderSideModel] = None
    bottom: Optional[BorderSideModel] = None
    diagonal: Optional[BorderSideModel] = None
    diagonal_up: bool = False
    diagonal_down: bool = False

    def to_dict(self):
        result = {}
        if self.left:
            result['left'] = self.left.to_dict()
        if self.right:
            result['right'] = self.right.to_dict()
        if self.top:
            result['top'] = self.top.to_dict()
        if self.bottom:
            result['bottom'] = self.bottom.to_dict()
        if self.diagonal:
            result['diagonal'] = self.diagonal.to_dict()
        result['diagonal_up'] = self.diagonal_up
        result['diagonal_down'] = self.diagonal_down
        return result


@dataclass
class AlignmentModel:
    """Represents cell alignment"""

    horizontal: str = "general"  # 'general', 'left', 'center', 'right', etc.
    vertical: str = "bottom"  # 'top', 'center', 'bottom', etc.
    text_rotation: int = 0
    wrap_text: bool = False
    shrink_to_fit: bool = False
    indent: int = 0

    def to_dict(self):
        return asdict(self)


@dataclass
class ProtectionModel:
    """Represents cell protection"""

    locked: bool = True
    hidden: bool = False

    def to_dict(self):
        return asdict(self)


@dataclass
class CellFormattingModel:
    """Complete cell formatting"""

    font: FontModel
    fill: FillModel
    border: BorderModel
    alignment: AlignmentModel
    protection: ProtectionModel

    def to_dict(self):
        return {
            'font': self.font.to_dict(),
            'fill': self.fill.to_dict(),
            'border': self.border.to_dict(),
            'alignment': self.alignment.to_dict(),
            'protection': self.protection.to_dict(),
        }
