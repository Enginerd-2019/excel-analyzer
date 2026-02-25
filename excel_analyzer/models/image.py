"""Data models for images"""

from dataclasses import dataclass, asdict
from typing import Optional


@dataclass
class ImageModel:
    """Represents an image in Excel"""

    format: str  # 'png', 'jpeg', 'bmp', 'gif', etc.
    data: str  # base64 encoded image data
    width: int
    height: int
    anchor: str  # Cell coordinate
    x_offset: int = 0
    y_offset: int = 0
    description: Optional[str] = None

    def to_dict(self):
        return {k: v for k, v in asdict(self).items() if v is not None}
