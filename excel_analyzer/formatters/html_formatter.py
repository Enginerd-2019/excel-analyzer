"""HTML output formatter"""

import logging
from pathlib import Path
from jinja2 import Environment, BaseLoader
from .base_formatter import BaseFormatter
from ..models.workbook import WorkbookModel
from ..utils.logging_utils import get_logger
from .. import __version__

logger = get_logger(__name__)


HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel Analysis: {{ workbook.file_path | basename }}</title>
    <style>
        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }

        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.6;
            color: #333;
            background-color: #f5f5f5;
            padding: 20px;
        }

        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            border-radius: 8px;
            overflow: hidden;
        }

        header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
        }

        header h1 {
            margin-bottom: 10px;
        }

        .metadata {
            background: #f8f9fa;
            padding: 20px;
            border-bottom: 1px solid #dee2e6;
        }

        .metadata-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 15px;
        }

        .metadata-item {
            display: flex;
            flex-direction: column;
        }

        .metadata-label {
            font-weight: 600;
            color: #6c757d;
            font-size: 0.875rem;
            margin-bottom: 5px;
        }

        .metadata-value {
            font-size: 1rem;
        }

        .tabs {
            display: flex;
            background: #e9ecef;
            overflow-x: auto;
            border-bottom: 2px solid #dee2e6;
        }

        .tab {
            padding: 15px 25px;
            cursor: pointer;
            border: none;
            background: none;
            font-size: 1rem;
            transition: all 0.3s;
            white-space: nowrap;
        }

        .tab:hover {
            background: #dee2e6;
        }

        .tab.active {
            background: white;
            border-bottom: 3px solid #667eea;
            font-weight: 600;
        }

        .worksheet-content {
            display: none;
            padding: 30px;
        }

        .worksheet-content.active {
            display: block;
        }

        .section {
            margin-bottom: 30px;
        }

        .section-title {
            font-size: 1.5rem;
            color: #495057;
            border-bottom: 2px solid #667eea;
            padding-bottom: 10px;
            margin-bottom: 20px;
        }

        .info-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
            margin-bottom: 20px;
        }

        .info-item {
            background: #f8f9fa;
            padding: 15px;
            border-radius: 5px;
        }

        .info-label {
            font-weight: 600;
            color: #6c757d;
            font-size: 0.875rem;
        }

        .info-value {
            font-size: 1rem;
            margin-top: 5px;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
            font-size: 0.9rem;
        }

        th, td {
            border: 1px solid #dee2e6;
            padding: 10px;
            text-align: left;
        }

        th {
            background: #667eea;
            color: white;
            font-weight: 600;
        }

        tr:nth-child(even) {
            background: #f8f9fa;
        }

        tr:hover {
            background: #e9ecef;
        }

        .list-item {
            background: #f8f9fa;
            padding: 10px 15px;
            margin-bottom: 10px;
            border-left: 3px solid #667eea;
            border-radius: 3px;
        }

        .chart-info, .image-info {
            background: #fff3cd;
            border: 1px solid #ffc107;
            padding: 15px;
            border-radius: 5px;
            margin-bottom: 15px;
        }

        .chart-title, .image-title {
            font-weight: 600;
            color: #856404;
            margin-bottom: 10px;
        }

        code {
            background: #f8f9fa;
            padding: 2px 6px;
            border-radius: 3px;
            font-family: 'Courier New', monospace;
            font-size: 0.9em;
        }

        .badge {
            display: inline-block;
            padding: 3px 8px;
            background: #667eea;
            color: white;
            border-radius: 12px;
            font-size: 0.8rem;
            margin-left: 10px;
        }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>📊 Excel Workbook Analysis</h1>
            <p>{{ workbook.file_path }}</p>
        </header>

        <div class="metadata">
            <div class="metadata-grid">
                <div class="metadata-item">
                    <span class="metadata-label">Format</span>
                    <span class="metadata-value">{{ workbook.file_format | upper }}</span>
                </div>
                <div class="metadata-item">
                    <span class="metadata-label">Worksheets</span>
                    <span class="metadata-value">{{ workbook.worksheets | length }}</span>
                </div>
                {% if workbook.properties.creator %}
                <div class="metadata-item">
                    <span class="metadata-label">Creator</span>
                    <span class="metadata-value">{{ workbook.properties.creator }}</span>
                </div>
                {% endif %}
                {% if workbook.properties.created %}
                <div class="metadata-item">
                    <span class="metadata-label">Created</span>
                    <span class="metadata-value">{{ workbook.properties.created }}</span>
                </div>
                {% endif %}
                <div class="metadata-item">
                    <span class="metadata-label">Analyzer Version</span>
                    <span class="metadata-value">{{ version }}</span>
                </div>
            </div>
        </div>

        <div class="tabs">
            {% for ws in workbook.worksheets %}
            <button class="tab {% if loop.first %}active{% endif %}" onclick="showWorksheet({{ loop.index0 }})">
                {{ ws.name }}
                <span class="badge">{{ ws.cells | length }}</span>
            </button>
            {% endfor %}
        </div>

        {% for ws in workbook.worksheets %}
        <div class="worksheet-content {% if loop.first %}active{% endif %}" id="worksheet-{{ loop.index0 }}">
            <div class="section">
                <h2 class="section-title">Worksheet Overview</h2>
                <div class="info-grid">
                    <div class="info-item">
                        <div class="info-label">Sheet Name</div>
                        <div class="info-value">{{ ws.name }}</div>
                    </div>
                    <div class="info-item">
                        <div class="info-label">Total Cells</div>
                        <div class="info-value">{{ ws.cells | length }}</div>
                    </div>
                    <div class="info-item">
                        <div class="info-label">Sheet State</div>
                        <div class="info-value">{{ ws.sheet_state }}</div>
                    </div>
                    {% if ws.freeze_panes %}
                    <div class="info-item">
                        <div class="info-label">Freeze Panes</div>
                        <div class="info-value">{{ ws.freeze_panes }}</div>
                    </div>
                    {% endif %}
                    {% if ws.auto_filter %}
                    <div class="info-item">
                        <div class="info-label">Auto Filter</div>
                        <div class="info-value">{{ ws.auto_filter }}</div>
                    </div>
                    {% endif %}
                </div>
            </div>

            {% if ws.merged_cells %}
            <div class="section">
                <h3 class="section-title">Merged Cells ({{ ws.merged_cells | length }})</h3>
                <div>
                    {% for mc in ws.merged_cells %}
                    <div class="list-item">{{ mc }}</div>
                    {% endfor %}
                </div>
            </div>
            {% endif %}

            {% if ws.cells %}
            <div class="section">
                <h3 class="section-title">Cell Data (First 100 cells)</h3>
                <table>
                    <thead>
                        <tr>
                            <th>Cell</th>
                            <th>Type</th>
                            <th>Value</th>
                            <th>Formula</th>
                            <th>Format</th>
                            <th>Colors</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for cell in ws.cells[:100] %}
                        {% set bg_color = '' %}
                        {% set font_color = '' %}
                        {% if cell.formatting %}
                            {% if cell.formatting.fill and cell.formatting.fill.fg_color and cell.formatting.fill.pattern_type != 'none' %}
                                {% if cell.formatting.fill.fg_color.value and cell.formatting.fill.fg_color.value not in ['#000000', '#FFFFFF'] %}
                                    {% set bg_color = cell.formatting.fill.fg_color.value %}
                                {% endif %}
                            {% endif %}
                            {% if cell.formatting.font and cell.formatting.font.color %}
                                {% if cell.formatting.font.color.value and cell.formatting.font.color.value not in ['#000000'] %}
                                    {% set font_color = cell.formatting.font.color.value %}
                                {% endif %}
                            {% endif %}
                        {% endif %}
                        <tr>
                            <td><code>{{ cell.coordinate }}</code></td>
                            <td>{{ cell.data_type }}</td>
                            <td {% if bg_color %}style="background-color: {{ bg_color }};"{% endif %} {% if font_color %}style="color: {{ font_color }}; {% if bg_color %}background-color: {{ bg_color }};{% endif %}"{% endif %}>{{ cell.value | string | truncate(50) }}</td>
                            <td>{% if cell.formula %}<code>{{ cell.formula | truncate(50) }}</code>{% endif %}</td>
                            <td>{{ cell.number_format }}</td>
                            <td>
                                {% if bg_color or font_color %}
                                    {% if bg_color %}<div style="display: inline-block; width: 20px; height: 20px; background-color: {{ bg_color }}; border: 1px solid #ccc; vertical-align: middle; margin-right: 5px;" title="Background: {{ bg_color }}"></div>{% endif %}
                                    {% if font_color %}<div style="display: inline-block; width: 20px; height: 20px; background-color: {{ font_color }}; border: 1px solid #ccc; vertical-align: middle;" title="Font: {{ font_color }}"></div>{% endif %}
                                {% endif %}
                            </td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
                {% if ws.cells | length > 100 %}
                <p><em>... and {{ ws.cells | length - 100 }} more cells</em></p>
                {% endif %}
            </div>
            {% endif %}

            {% if ws.charts %}
            <div class="section">
                <h3 class="section-title">Charts ({{ ws.charts | length }})</h3>
                {% for chart in ws.charts %}
                <div class="chart-info">
                    <div class="chart-title">Chart {{ loop.index }}: {{ chart.chart_type | title }}</div>
                    {% if chart.title %}
                    <p><strong>Title:</strong> {{ chart.title }}</p>
                    {% endif %}
                    <p><strong>Series Count:</strong> {{ chart.series | length }}</p>
                    {% for series in chart.series %}
                    <div style="margin-left: 20px; margin-top: 10px;">
                        <p><strong>Series {{ loop.index }}:</strong></p>
                        {% if series.title %}
                        <p style="margin-left: 20px;">Title: {{ series.title }}</p>
                        {% endif %}
                        {% if series.values %}
                        <p style="margin-left: 20px;">Values: <code>{{ series.values }}</code></p>
                        {% endif %}
                        {% if series.categories %}
                        <p style="margin-left: 20px;">Categories: <code>{{ series.categories }}</code></p>
                        {% endif %}
                    </div>
                    {% endfor %}
                </div>
                {% endfor %}
            </div>
            {% endif %}

            {% if ws.images %}
            <div class="section">
                <h3 class="section-title">Images ({{ ws.images | length }})</h3>
                {% for img in ws.images %}
                <div class="image-info">
                    <div class="image-title">Image {{ loop.index }}</div>
                    <p><strong>Format:</strong> {{ img.format | upper }}</p>
                    <p><strong>Size:</strong> {{ img.width }} x {{ img.height }}</p>
                    <p><strong>Anchor:</strong> <code>{{ img.anchor }}</code></p>
                    {% if img.format in ['png', 'jpeg', 'jpg', 'gif'] %}
                    <p><img src="data:image/{{ img.format }};base64,{{ img.data }}" alt="Image {{ loop.index }}" style="max-width: 400px; margin-top: 10px;"></p>
                    {% endif %}
                </div>
                {% endfor %}
            </div>
            {% endif %}

            {% if ws.data_validations %}
            <div class="section">
                <h3 class="section-title">Data Validations ({{ ws.data_validations | length }})</h3>
                {% for dv in ws.data_validations %}
                <div class="list-item">
                    <strong>Range:</strong> <code>{{ dv.sqref }}</code><br>
                    <strong>Type:</strong> {{ dv.validation_type }}
                    {% if dv.formula1 %}<br><strong>Formula:</strong> <code>{{ dv.formula1 }}</code>{% endif %}
                </div>
                {% endfor %}
            </div>
            {% endif %}

            {% if ws.conditional_formatting %}
            <div class="section">
                <h3 class="section-title">Conditional Formatting ({{ ws.conditional_formatting | length }})</h3>
                {% for cf in ws.conditional_formatting %}
                <div class="list-item">
                    <strong>Range:</strong> <code>{{ cf.sqref }}</code><br>
                    <strong>Type:</strong> {{ cf.rule_type }}<br>
                    <strong>Priority:</strong> {{ cf.priority }}
                </div>
                {% endfor %}
            </div>
            {% endif %}
        </div>
        {% endfor %}
    </div>

    <script>
        function showWorksheet(index) {
            // Hide all worksheets
            const worksheets = document.querySelectorAll('.worksheet-content');
            worksheets.forEach(ws => ws.classList.remove('active'));

            // Remove active class from all tabs
            const tabs = document.querySelectorAll('.tab');
            tabs.forEach(tab => tab.classList.remove('active'));

            // Show selected worksheet
            document.getElementById('worksheet-' + index).classList.add('active');
            tabs[index].classList.add('active');
        }
    </script>
</body>
</html>
"""


class HTMLFormatter(BaseFormatter):
    """Formats workbook analysis as HTML"""

    def format(self, workbook_model: WorkbookModel, output_path: str, verbose: bool = False):
        """
        Generate HTML output.

        Args:
            workbook_model: WorkbookModel to format
            output_path: Output file path
            verbose: If True, log progress

        Returns:
            Output file path
        """
        if verbose:
            logger.info(f"Generating HTML output: {output_path}")

        # Create Jinja2 environment with custom filters
        env = Environment(loader=BaseLoader())

        # Add custom filter for basename
        def basename_filter(path):
            return Path(path).name

        env.filters['basename'] = basename_filter

        # Create template
        template = env.from_string(HTML_TEMPLATE)

        # Render template
        html_content = template.render(
            workbook=workbook_model,
            version=__version__,
        )

        # Write HTML file
        output_path = Path(output_path)
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)

        if verbose:
            logger.info(f"HTML output written to: {output_path}")

        return output_path
