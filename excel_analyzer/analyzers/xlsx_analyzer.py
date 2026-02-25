"""Main XLSX file analyzer"""

import logging
import openpyxl
from typing import Optional
from tqdm import tqdm
from .base_analyzer import BaseAnalyzer
from .cell_analyzer import CellAnalyzer
from .format_analyzer import FormatAnalyzer
from .structure_analyzer import StructureAnalyzer
from .validation_analyzer import ValidationAnalyzer
from .conditional_analyzer import ConditionalAnalyzer
from .chart_analyzer import ChartAnalyzer
from .image_analyzer import ImageAnalyzer
from ..models.workbook import WorkbookModel, WorkbookPropertiesModel, DefinedNameModel
from ..models.worksheet import WorksheetModel
from ..utils.logging_utils import get_logger

logger = get_logger(__name__)


class XLSXAnalyzer(BaseAnalyzer):
    """Analyzer for .xlsx files using openpyxl"""

    def __init__(self):
        self.cell_analyzer = CellAnalyzer()
        self.format_analyzer = FormatAnalyzer()
        self.structure_analyzer = StructureAnalyzer()
        self.validation_analyzer = ValidationAnalyzer()
        self.conditional_analyzer = ConditionalAnalyzer()
        self.chart_analyzer = ChartAnalyzer()
        self.image_analyzer = ImageAnalyzer()

    def analyze(self, file_path: str, verbose: bool = False) -> WorkbookModel:
        """
        Analyze an XLSX file.

        Args:
            file_path: Path to .xlsx file
            verbose: If True, show detailed progress

        Returns:
            WorkbookModel with complete analysis
        """
        if verbose:
            logger.info(f"Opening workbook: {file_path}")

        # Load workbook (data_only=False to preserve formulas)
        workbook = openpyxl.load_workbook(file_path, data_only=False, keep_vba=False)

        if verbose:
            logger.info(f"Workbook loaded with {len(workbook.worksheets)} worksheets")

        # Extract workbook properties
        properties = self._extract_properties(workbook)

        # Extract defined names
        defined_names = self._extract_defined_names(workbook)

        # Extract workbook-level settings
        active_sheet_index = 0
        if workbook.active:
            try:
                active_sheet_index = workbook.worksheets.index(workbook.active)
            except (ValueError, AttributeError):
                active_sheet_index = 0
        calculation_mode = "auto"  # Default, could extract from workbook.xml if needed

        # Create workbook model
        workbook_model = WorkbookModel(
            file_path=file_path,
            file_format='xlsx',
            properties=properties,
            worksheets=[],
            defined_names=defined_names,
            active_sheet_index=active_sheet_index,
            calculation_mode=calculation_mode,
        )

        # Analyze each worksheet
        worksheets = workbook.worksheets
        if verbose:
            worksheets = tqdm(worksheets, desc="Analyzing worksheets", unit="sheet")

        for idx, worksheet in enumerate(worksheets):
            if verbose and not isinstance(worksheets, tqdm):
                logger.info(f"Analyzing worksheet: {worksheet.title}")

            worksheet_model = self._analyze_worksheet(worksheet, idx, verbose)
            workbook_model.worksheets.append(worksheet_model)

        if verbose:
            logger.info("Analysis complete")

        workbook.close()
        return workbook_model

    def _extract_properties(self, workbook) -> WorkbookPropertiesModel:
        """Extract workbook properties/metadata"""
        props = workbook.properties

        return WorkbookPropertiesModel(
            title=props.title,
            subject=props.subject,
            creator=props.creator,
            keywords=props.keywords,
            description=props.description,
            last_modified_by=props.lastModifiedBy,
            created=props.created,
            modified=props.modified,
            category=props.category,
            content_status=props.contentStatus,
            version=props.version,
            revision=props.revision,
            application=workbook.properties.application if hasattr(workbook.properties, 'application') else None,
        )

    def _extract_defined_names(self, workbook) -> list:
        """Extract defined names (named ranges)"""
        defined_names = []

        try:
            if hasattr(workbook, 'defined_names'):
                for name, defn in workbook.defined_names.items():
                    # Get the destinations (can be multiple for cross-sheet names)
                    try:
                        destinations = list(defn.destinations)
                        if destinations:
                            for sheet_name, cell_ref in destinations:
                                value = f"'{sheet_name}'!{cell_ref}" if sheet_name else cell_ref
                                defined_names.append(DefinedNameModel(
                                    name=name,
                                    value=value,
                                    hidden=defn.hidden if hasattr(defn, 'hidden') else False,
                                ))
                                break  # Only take first destination for simplicity
                    except:
                        # Fallback: use the attr_text if destinations fails
                        if hasattr(defn, 'attr_text'):
                            defined_names.append(DefinedNameModel(
                                name=name,
                                value=defn.attr_text,
                                hidden=defn.hidden if hasattr(defn, 'hidden') else False,
                            ))
        except Exception as e:
            logger.warning(f"Error extracting defined names: {e}")

        return defined_names

    def _analyze_worksheet(self, worksheet, index: int, verbose: bool) -> WorksheetModel:
        """Analyze a single worksheet"""

        # Extract cells
        cells = self.cell_analyzer.extract_cells(worksheet, verbose)

        # Extract merged cells
        merged_cells = self.structure_analyzer.extract_merged_cells(worksheet)

        # Extract column dimensions
        column_dimensions = self.structure_analyzer.extract_column_dimensions(worksheet)

        # Extract row dimensions
        row_dimensions = self.structure_analyzer.extract_row_dimensions(worksheet)

        # Extract data validations
        data_validations = self.validation_analyzer.extract_validations(worksheet)

        # Extract conditional formatting
        conditional_formatting = self.conditional_analyzer.extract_rules(worksheet)

        # Extract charts
        charts = self.chart_analyzer.extract_charts(worksheet)

        # Extract images
        images = self.image_analyzer.extract_images(worksheet)

        # Extract print settings
        print_settings = self.structure_analyzer.extract_print_settings(worksheet)

        # Extract header/footer
        header_footer = self.structure_analyzer.extract_header_footer(worksheet)

        # Extract freeze panes
        freeze_panes = self.structure_analyzer.extract_freeze_panes(worksheet)

        # Extract auto filter
        auto_filter = self.structure_analyzer.extract_auto_filter(worksheet)

        # Extract tab color
        tab_color = self.structure_analyzer.extract_tab_color(worksheet)

        # Extract sheet state
        sheet_state = "visible"
        if hasattr(worksheet, 'sheet_state'):
            sheet_state = worksheet.sheet_state

        # Extract sheet view (zoom, etc.)
        sheet_view = None
        if hasattr(worksheet, 'sheet_view') and worksheet.sheet_view:
            try:
                view = worksheet.sheet_view
                sheet_view = {
                    'zoom_scale': view.zoomScale if hasattr(view, 'zoomScale') else 100,
                    'zoom_scale_normal': view.zoomScaleNormal if hasattr(view, 'zoomScaleNormal') else 100,
                    'show_gridlines': view.showGridLines if hasattr(view, 'showGridLines') else True,
                    'show_row_col_headers': view.showRowColHeaders if hasattr(view, 'showRowColHeaders') else True,
                }
            except:
                pass

        return WorksheetModel(
            name=worksheet.title,
            index=index,
            cells=cells,
            merged_cells=merged_cells,
            column_dimensions=column_dimensions,
            row_dimensions=row_dimensions,
            data_validations=data_validations,
            conditional_formatting=conditional_formatting,
            charts=charts,
            images=images,
            print_settings=print_settings,
            header_footer=header_footer,
            freeze_panes=freeze_panes,
            auto_filter=auto_filter,
            tab_color=tab_color,
            sheet_state=sheet_state,
            sheet_view=sheet_view,
        )
