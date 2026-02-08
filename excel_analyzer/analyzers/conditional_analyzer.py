"""Analyzer for conditional formatting"""

import logging
from typing import List
from openpyxl.worksheet.worksheet import Worksheet
from ..models.worksheet import ConditionalFormattingModel
from ..utils.logging_utils import get_logger

logger = get_logger(__name__)


class ConditionalAnalyzer:
    """Extracts conditional formatting rules"""

    def extract_rules(self, worksheet: Worksheet) -> List[ConditionalFormattingModel]:
        """
        Extract conditional formatting rules from worksheet.

        Args:
            worksheet: openpyxl Worksheet object

        Returns:
            List of ConditionalFormattingModel objects
        """
        rules = []

        try:
            # Handle different openpyxl versions
            if hasattr(worksheet.conditional_formatting, 'items'):
                cf_items = worksheet.conditional_formatting.items()
            else:
                # For newer openpyxl versions, iterate differently
                cf_items = [(cf.sqref, [cf]) for cf in worksheet.conditional_formatting]

            for cf_range, cf_rules in cf_items:
                sqref = str(cf_range)

                for rule in cf_rules:
                    # Extract formula(s)
                    formulas = []
                    if hasattr(rule, 'formula') and rule.formula:
                        formulas = rule.formula if isinstance(rule.formula, list) else [rule.formula]

                    # Get format description if available
                    format_desc = None
                    if hasattr(rule, 'dxf') and rule.dxf:
                        format_desc = self._extract_dxf_format(rule.dxf)

                    rule_model = ConditionalFormattingModel(
                        sqref=sqref,
                        rule_type=rule.type if hasattr(rule, 'type') else "unknown",
                        priority=rule.priority if hasattr(rule, 'priority') else 0,
                        formula=formulas if formulas else None,
                        operator=rule.operator if hasattr(rule, 'operator') else None,
                        stop_if_true=rule.stopIfTrue if hasattr(rule, 'stopIfTrue') else False,
                        dxf_id=rule.dxfId if hasattr(rule, 'dxfId') else None,
                        format_description=format_desc,
                    )
                    rules.append(rule_model)
        except Exception as e:
            logger.warning(f"Error extracting conditional formatting: {e}")

        return rules

    def _extract_dxf_format(self, dxf):
        """Extract formatting from DXF (Differential Formatting)"""
        format_desc = {}

        try:
            if hasattr(dxf, 'font') and dxf.font:
                font_info = {}
                if dxf.font.bold:
                    font_info['bold'] = True
                if dxf.font.italic:
                    font_info['italic'] = True
                if dxf.font.color:
                    from ..utils.color_utils import convert_color
                    color_model = convert_color(dxf.font.color)
                    if color_model:
                        font_info['color'] = color_model.value
                if font_info:
                    format_desc['font'] = font_info

            if hasattr(dxf, 'fill') and dxf.fill:
                fill_info = {}
                if dxf.fill.patternType:
                    fill_info['pattern_type'] = dxf.fill.patternType
                if hasattr(dxf.fill, 'fgColor') and dxf.fill.fgColor:
                    from ..utils.color_utils import convert_color
                    color_model = convert_color(dxf.fill.fgColor)
                    if color_model:
                        fill_info['fg_color'] = color_model.value
                if fill_info:
                    format_desc['fill'] = fill_info

            if hasattr(dxf, 'border') and dxf.border:
                border_info = {}
                for side in ['left', 'right', 'top', 'bottom']:
                    side_obj = getattr(dxf.border, side, None)
                    if side_obj and side_obj.style:
                        border_info[side] = {'style': side_obj.style}
                if border_info:
                    format_desc['border'] = border_info

        except Exception as e:
            logger.warning(f"Error extracting DXF format: {e}")

        return format_desc if format_desc else None
