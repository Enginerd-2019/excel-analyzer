"""Analyzer for images and shapes"""

import logging
import base64
from typing import List
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter
from ..models.image import ImageModel
from ..utils.logging_utils import get_logger

logger = get_logger(__name__)


class ImageAnalyzer:
    """Extracts images and shapes"""

    def extract_images(self, worksheet: Worksheet) -> List[ImageModel]:
        """
        Extract all images from worksheet.

        Args:
            worksheet: openpyxl Worksheet object

        Returns:
            List of ImageModel objects
        """
        images = []

        try:
            for image in worksheet._images:
                image_model = self._extract_image(image)
                if image_model:
                    images.append(image_model)
        except Exception as e:
            logger.warning(f"Error extracting images: {e}")

        return images

    def _extract_image(self, image) -> ImageModel:
        """Extract a single image"""
        try:
            # Get image data
            image_data = image._data()

            # Detect format
            img_format = self._detect_format(image_data)

            # Encode to base64
            encoded_data = base64.b64encode(image_data).decode('utf-8')

            # Get size
            width = image.width if hasattr(image, 'width') else 0
            height = image.height if hasattr(image, 'height') else 0

            # Get anchor position
            anchor_cell = "A1"
            x_offset = 0
            y_offset = 0

            if hasattr(image, 'anchor') and image.anchor:
                anchor = image.anchor
                if hasattr(anchor, '_from'):
                    col = anchor._from.col
                    row = anchor._from.row
                    anchor_cell = f"{get_column_letter(col + 1)}{row + 1}"

                    if hasattr(anchor._from, 'colOff'):
                        x_offset = anchor._from.colOff
                    if hasattr(anchor._from, 'rowOff'):
                        y_offset = anchor._from.rowOff

            # Get description
            description = None
            if hasattr(image, 'description') and image.description:
                description = image.description

            return ImageModel(
                format=img_format,
                data=encoded_data,
                width=width,
                height=height,
                anchor=anchor_cell,
                x_offset=x_offset,
                y_offset=y_offset,
                description=description,
            )

        except Exception as e:
            logger.warning(f"Error extracting image details: {e}")
            return None

    def _detect_format(self, image_data: bytes) -> str:
        """Detect image format from magic bytes"""
        if not image_data or len(image_data) < 12:
            return "unknown"

        # Check magic bytes
        if image_data[:8] == b'\x89PNG\r\n\x1a\n':
            return "png"
        elif image_data[:2] == b'\xff\xd8':
            return "jpeg"
        elif image_data[:2] == b'BM':
            return "bmp"
        elif image_data[:6] in (b'GIF87a', b'GIF89a'):
            return "gif"
        elif image_data[:4] == b'RIFF' and image_data[8:12] == b'WEBP':
            return "webp"
        else:
            return "unknown"
