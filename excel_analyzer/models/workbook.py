"""Data models for workbooks"""

from dataclasses import dataclass, field
from typing import List, Dict, Optional, Any
from datetime import datetime
from .worksheet import WorksheetModel


@dataclass
class WorkbookPropertiesModel:
    """Represents workbook properties/metadata"""

    title: Optional[str] = None
    subject: Optional[str] = None
    creator: Optional[str] = None
    keywords: Optional[str] = None
    description: Optional[str] = None
    last_modified_by: Optional[str] = None
    created: Optional[datetime] = None
    modified: Optional[datetime] = None
    category: Optional[str] = None
    content_status: Optional[str] = None
    version: Optional[str] = None
    revision: Optional[int] = None
    application: Optional[str] = None

    def to_dict(self):
        result = {}
        for key, value in self.__dict__.items():
            if value is not None:
                if isinstance(value, datetime):
                    result[key] = value.isoformat()
                else:
                    result[key] = value
        return result


@dataclass
class DefinedNameModel:
    """Represents a defined name (named range)"""

    name: str
    value: str  # Formula/reference
    local_sheet_id: Optional[int] = None  # None if workbook-level
    comment: Optional[str] = None
    hidden: bool = False

    def to_dict(self):
        return {k: v for k, v in self.__dict__.items() if v is not None}


@dataclass
class WorkbookModel:
    """Represents an Excel workbook"""

    file_path: str
    file_format: str  # 'xlsx' or 'xls'
    properties: WorkbookPropertiesModel
    worksheets: List[WorksheetModel] = field(default_factory=list)
    defined_names: List[DefinedNameModel] = field(default_factory=list)
    active_sheet_index: int = 0
    calculation_mode: str = "auto"  # 'auto', 'manual', 'autoNoTable'
    workbook_view: Optional[Dict[str, Any]] = None

    def to_dict(self):
        result = {
            'file_path': self.file_path,
            'file_format': self.file_format,
            'properties': self.properties.to_dict(),
            'worksheets': [ws.to_dict() for ws in self.worksheets],
            'defined_names': [dn.to_dict() for dn in self.defined_names],
            'active_sheet_index': self.active_sheet_index,
            'calculation_mode': self.calculation_mode,
        }

        if self.workbook_view:
            result['workbook_view'] = self.workbook_view

        return result
