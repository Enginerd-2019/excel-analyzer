"""Analyzer for data validation rules"""

import logging
from typing import List
from openpyxl.worksheet.worksheet import Worksheet
from ..models.worksheet import DataValidationModel
from ..utils.logging_utils import get_logger

logger = get_logger(__name__)


class ValidationAnalyzer:
    """Extracts data validation rules"""

    def extract_validations(self, worksheet: Worksheet) -> List[DataValidationModel]:
        """
        Extract data validation rules from worksheet.

        Args:
            worksheet: openpyxl Worksheet object

        Returns:
            List of DataValidationModel objects
        """
        validations = []

        try:
            for dv in worksheet.data_validations.dataValidation:
                validation_model = DataValidationModel(
                    sqref=str(dv.sqref) if dv.sqref else "",
                    validation_type=dv.type or "none",
                    operator=dv.operator,
                    formula1=dv.formula1,
                    formula2=dv.formula2,
                    allow_blank=dv.allowBlank if dv.allowBlank is not None else True,
                    show_input_message=dv.showInputMessage or False,
                    input_title=dv.promptTitle,
                    input_message=dv.prompt,
                    show_error_message=dv.showErrorMessage if dv.showErrorMessage is not None else True,
                    error_title=dv.errorTitle,
                    error_message=dv.error,
                    error_style=dv.errorStyle or "stop",
                )
                validations.append(validation_model)
        except Exception as e:
            logger.warning(f"Error extracting data validations: {e}")

        return validations
