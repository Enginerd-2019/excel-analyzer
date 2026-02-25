"""Analyzer for cell formatting"""

from openpyxl.cell.cell import Cell
from ..models.formatting import (
    CellFormattingModel,
    FontModel,
    FillModel,
    BorderModel,
    BorderSideModel,
    AlignmentModel,
    ProtectionModel,
)
from ..utils.color_utils import convert_color


class FormatAnalyzer:
    """Extracts cell formatting information"""

    def extract_cell_formatting(self, cell: Cell) -> CellFormattingModel:
        """
        Extract complete formatting from a cell.

        Args:
            cell: openpyxl Cell object

        Returns:
            CellFormattingModel with all formatting details
        """
        return CellFormattingModel(
            font=self._extract_font(cell.font),
            fill=self._extract_fill(cell.fill),
            border=self._extract_border(cell.border),
            alignment=self._extract_alignment(cell.alignment),
            protection=self._extract_protection(cell.protection),
        )

    def _extract_font(self, font) -> FontModel:
        """Extract font formatting"""
        if not font:
            return FontModel()

        return FontModel(
            name=font.name or "Calibri",
            size=font.size or 11.0,
            bold=font.bold or False,
            italic=font.italic or False,
            underline=font.underline or "none",
            strike=font.strike or False,
            color=convert_color(font.color),
        )

    def _extract_fill(self, fill) -> FillModel:
        """Extract fill formatting"""
        if not fill:
            return FillModel()

        pattern_type = fill.patternType or "none"

        fg_color = None
        bg_color = None

        if hasattr(fill, 'fgColor'):
            fg_color = convert_color(fill.fgColor)
        if hasattr(fill, 'bgColor'):
            bg_color = convert_color(fill.bgColor)

        return FillModel(
            pattern_type=pattern_type,
            fg_color=fg_color,
            bg_color=bg_color,
        )

    def _extract_border_side(self, side):
        """Extract one side of border"""
        if not side or not side.style:
            return BorderSideModel(style="none")

        return BorderSideModel(
            style=side.style,
            color=convert_color(side.color),
        )

    def _extract_border(self, border) -> BorderModel:
        """Extract border formatting"""
        if not border:
            return BorderModel()

        return BorderModel(
            left=self._extract_border_side(border.left),
            right=self._extract_border_side(border.right),
            top=self._extract_border_side(border.top),
            bottom=self._extract_border_side(border.bottom),
            diagonal=self._extract_border_side(border.diagonal),
            diagonal_up=border.diagonalUp or False,
            diagonal_down=border.diagonalDown or False,
        )

    def _extract_alignment(self, alignment) -> AlignmentModel:
        """Extract alignment formatting"""
        if not alignment:
            return AlignmentModel()

        return AlignmentModel(
            horizontal=alignment.horizontal or "general",
            vertical=alignment.vertical or "bottom",
            text_rotation=alignment.textRotation or 0,
            wrap_text=alignment.wrapText or False,
            shrink_to_fit=alignment.shrinkToFit or False,
            indent=alignment.indent or 0,
        )

    def _extract_protection(self, protection) -> ProtectionModel:
        """Extract protection formatting"""
        if not protection:
            return ProtectionModel()

        return ProtectionModel(
            locked=protection.locked if protection.locked is not None else True,
            hidden=protection.hidden or False,
        )
