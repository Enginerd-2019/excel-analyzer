"""Utilities for file handling"""

import os
from pathlib import Path
from typing import Optional


def validate_file(file_path: str) -> bool:
    """
    Check if file exists and is readable.

    Args:
        file_path: Path to file

    Returns:
        True if file is valid, False otherwise
    """
    path = Path(file_path)

    if not path.exists():
        return False

    if not path.is_file():
        return False

    if not os.access(file_path, os.R_OK):
        return False

    return True


def detect_file_format(file_path: str) -> Optional[str]:
    """
    Detect Excel file format (.xlsx or .xls).

    Args:
        file_path: Path to Excel file

    Returns:
        'xlsx', 'xls', or None if not recognized
    """
    path = Path(file_path)
    suffix = path.suffix.lower()

    if suffix == '.xlsx':
        return 'xlsx'
    elif suffix == '.xls':
        return 'xls'
    elif suffix == '.xlsm':
        return 'xlsx'  # Treat macro-enabled as xlsx
    elif suffix == '.xlsb':
        return 'xlsb'  # Binary format (not currently supported)
    else:
        return None


def get_output_filename(input_path: str, suffix: str, extension: str) -> str:
    """
    Generate output filename from input path.

    Args:
        input_path: Path to input file
        suffix: Suffix to add (e.g., "_analysis")
        extension: Output file extension (e.g., ".json")

    Returns:
        Output filename
    """
    path = Path(input_path)
    stem = path.stem
    return f"{stem}{suffix}{extension}"
