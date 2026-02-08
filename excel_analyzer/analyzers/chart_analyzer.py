"""Analyzer for charts"""

import logging
from typing import List
from openpyxl.worksheet.worksheet import Worksheet
from ..models.chart import (
    ChartModel,
    ChartSeriesModel,
    AxisModel,
    LegendModel,
    ChartPositionModel,
)
from ..utils.color_utils import convert_color
from ..utils.logging_utils import get_logger

logger = get_logger(__name__)


class ChartAnalyzer:
    """Extracts chart information"""

    def extract_charts(self, worksheet: Worksheet) -> List[ChartModel]:
        """
        Extract all charts from worksheet.

        Args:
            worksheet: openpyxl Worksheet object

        Returns:
            List of ChartModel objects
        """
        charts = []

        try:
            for chart in worksheet._charts:
                chart_model = self._extract_chart(chart)
                if chart_model:
                    charts.append(chart_model)
        except Exception as e:
            logger.warning(f"Error extracting charts: {e}")

        return charts

    def _extract_chart(self, chart) -> ChartModel:
        """Extract a single chart"""
        try:
            # Determine chart type
            chart_type = self._get_chart_type(chart)

            # Extract title
            title = None
            if hasattr(chart, 'title') and chart.title:
                title = str(chart.title.tx.rich.p[0].r[0].t) if hasattr(chart.title, 'tx') else None

            # Extract series
            series = self._extract_series(chart)

            # Extract axes
            x_axis = None
            y_axis = None
            if hasattr(chart, 'x_axis') and chart.x_axis:
                x_axis = self._extract_axis(chart.x_axis)
            if hasattr(chart, 'y_axis') and chart.y_axis:
                y_axis = self._extract_axis(chart.y_axis)

            # Extract legend
            legend = None
            if hasattr(chart, 'legend') and chart.legend:
                legend = self._extract_legend(chart.legend)

            # Extract position
            position = None
            if hasattr(chart, 'anchor') and chart.anchor:
                position = self._extract_position(chart.anchor)

            # Extract style
            style = None
            if hasattr(chart, 'style'):
                style = chart.style

            return ChartModel(
                chart_type=chart_type,
                title=title,
                series=series,
                x_axis=x_axis,
                y_axis=y_axis,
                legend=legend,
                position=position,
                style=style,
            )

        except Exception as e:
            logger.warning(f"Error extracting chart details: {e}")
            return None

    def _get_chart_type(self, chart) -> str:
        """Determine chart type"""
        chart_type_name = type(chart).__name__

        # Map class names to readable types
        type_mapping = {
            'BarChart': 'bar',
            'BarChart3D': 'bar3d',
            'LineChart': 'line',
            'LineChart3D': 'line3d',
            'PieChart': 'pie',
            'PieChart3D': 'pie3d',
            'ScatterChart': 'scatter',
            'AreaChart': 'area',
            'AreaChart3D': 'area3d',
            'DoughnutChart': 'doughnut',
            'RadarChart': 'radar',
            'BubbleChart': 'bubble',
            'StockChart': 'stock',
            'SurfaceChart': 'surface',
            'SurfaceChart3D': 'surface3d',
        }

        return type_mapping.get(chart_type_name, chart_type_name.lower())

    def _extract_series(self, chart) -> List[ChartSeriesModel]:
        """Extract chart series"""
        series_list = []

        try:
            if not hasattr(chart, 'series'):
                return series_list

            for series in chart.series:
                # Extract title
                series_title = None
                if hasattr(series, 'title') and series.title:
                    series_title = str(series.title)

                # Extract values range
                values = None
                if hasattr(series, 'val') and series.val:
                    values = str(series.val)

                # Extract categories range
                categories = None
                if hasattr(series, 'cat') and series.cat:
                    categories = str(series.cat)

                # Extract color (if available)
                color = None
                try:
                    if hasattr(series, 'graphicalProperties') and series.graphicalProperties:
                        gp = series.graphicalProperties
                        if hasattr(gp, 'solidFill') and gp.solidFill:
                            color = convert_color(gp.solidFill)
                except:
                    pass

                series_model = ChartSeriesModel(
                    title=series_title,
                    values=values,
                    categories=categories,
                    color=color,
                )
                series_list.append(series_model)

        except Exception as e:
            logger.warning(f"Error extracting series: {e}")

        return series_list

    def _extract_axis(self, axis) -> AxisModel:
        """Extract axis information"""
        try:
            # Extract title
            title = None
            if hasattr(axis, 'title') and axis.title:
                try:
                    title = str(axis.title.tx.rich.p[0].r[0].t) if hasattr(axis.title, 'tx') else None
                except:
                    pass

            # Extract scaling
            min_value = None
            max_value = None
            if hasattr(axis, 'scaling') and axis.scaling:
                min_value = axis.scaling.min
                max_value = axis.scaling.max

            # Extract units
            major_unit = None
            minor_unit = None
            if hasattr(axis, 'majorUnit'):
                major_unit = axis.majorUnit
            if hasattr(axis, 'minorUnit'):
                minor_unit = axis.minorUnit

            # Extract number format
            number_format = None
            if hasattr(axis, 'number_format') and axis.number_format:
                number_format = axis.number_format

            # Extract position
            axis_position = None
            if hasattr(axis, 'axPos') and axis.axPos:
                axis_position = axis.axPos.val if hasattr(axis.axPos, 'val') else str(axis.axPos)

            # Check if axis is deleted
            delete = False
            if hasattr(axis, 'delete') and axis.delete:
                delete = axis.delete.val if hasattr(axis.delete, 'val') else False

            return AxisModel(
                title=title,
                min_value=min_value,
                max_value=max_value,
                major_unit=major_unit,
                minor_unit=minor_unit,
                number_format=number_format,
                axis_position=axis_position,
                delete=delete,
            )

        except Exception as e:
            logger.warning(f"Error extracting axis: {e}")
            return AxisModel()

    def _extract_legend(self, legend) -> LegendModel:
        """Extract legend information"""
        try:
            # Extract position
            position = "right"  # default
            if hasattr(legend, 'position') and legend.position:
                position = legend.position

            # Extract overlay
            overlay = False
            if hasattr(legend, 'overlay') and legend.overlay:
                overlay = legend.overlay.val if hasattr(legend.overlay, 'val') else False

            return LegendModel(
                position=position,
                overlay=overlay,
            )

        except Exception as e:
            logger.warning(f"Error extracting legend: {e}")
            return LegendModel()

    def _extract_position(self, anchor) -> ChartPositionModel:
        """Extract chart position"""
        try:
            # Get anchor cell
            anchor_cell = "A1"
            if hasattr(anchor, '_from'):
                col = anchor._from.col
                row = anchor._from.row
                from openpyxl.utils import get_column_letter
                anchor_cell = f"{get_column_letter(col + 1)}{row + 1}"

            # Get offsets
            x_offset = 0
            y_offset = 0
            if hasattr(anchor, '_from'):
                if hasattr(anchor._from, 'colOff'):
                    x_offset = anchor._from.colOff
                if hasattr(anchor._from, 'rowOff'):
                    y_offset = anchor._from.rowOff

            # Get size (if available)
            width = None
            height = None
            if hasattr(anchor, 'to'):
                # Calculate approximate size based on anchor points
                # This is approximate as we don't have exact pixel dimensions
                pass

            return ChartPositionModel(
                anchor=anchor_cell,
                x_offset=x_offset,
                y_offset=y_offset,
                width=width,
                height=height,
            )

        except Exception as e:
            logger.warning(f"Error extracting position: {e}")
            return ChartPositionModel(anchor="A1")
