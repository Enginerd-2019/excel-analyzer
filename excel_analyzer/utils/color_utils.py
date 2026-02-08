"""Utilities for color conversion"""

from typing import Optional
from ..models.formatting import ColorModel


def rgb_to_hex(rgb: str) -> str:
    """
    Convert ARGB or RGB string to hex color.

    Args:
        rgb: RGB string (e.g., "FFAABBCC" or "AABBCC")

    Returns:
        Hex color string (e.g., "#AABBCC")
    """
    if not rgb:
        return "#000000"

    # Remove alpha channel if present (first 2 chars)
    if len(rgb) == 8:
        rgb = rgb[2:]

    # Ensure it's 6 characters
    if len(rgb) != 6:
        return "#000000"

    return f"#{rgb.upper()}"


def convert_color(color_obj) -> Optional[ColorModel]:
    """
    Convert openpyxl color object to ColorModel.

    Args:
        color_obj: openpyxl color object

    Returns:
        ColorModel or None
    """
    if color_obj is None:
        return None

    try:
        # RGB color
        if hasattr(color_obj, 'rgb') and color_obj.rgb:
            return ColorModel(
                type='rgb',
                value=rgb_to_hex(color_obj.rgb),
                tint=getattr(color_obj, 'tint', None)
            )

        # Theme color
        elif hasattr(color_obj, 'theme') and color_obj.theme is not None:
            return ColorModel(
                type='theme',
                value=str(color_obj.theme),
                tint=getattr(color_obj, 'tint', None)
            )

        # Indexed color
        elif hasattr(color_obj, 'indexed') and color_obj.indexed is not None:
            return ColorModel(
                type='indexed',
                value=str(color_obj.indexed)
            )

        # Auto color
        elif hasattr(color_obj, 'auto') and color_obj.auto:
            return ColorModel(type='auto')

        # Default
        else:
            return ColorModel(type='auto')

    except Exception:
        return ColorModel(type='auto')


def get_color_hex(color_obj) -> str:
    """
    Get hex color string from openpyxl color object.

    Args:
        color_obj: openpyxl color object

    Returns:
        Hex color string
    """
    if color_obj is None:
        return "#000000"

    try:
        if hasattr(color_obj, 'rgb') and color_obj.rgb:
            return rgb_to_hex(color_obj.rgb)
        else:
            return "#000000"
    except Exception:
        return "#000000"
