# Excel Analyzer

A comprehensive Python CLI tool for analyzing Excel files (.xlsx and .xls) and extracting all data necessary for programmatic file duplication. This tool provides detailed analysis of cell data, formulas, formatting, charts, images, data validation, conditional formatting, and more.

## Features

- **Comprehensive Analysis**: Extracts all Excel file components:
  - Cell data, formulas, and calculated values
  - Complete formatting (fonts, colors, borders, alignment, number formats)
  - Merged cells, column widths, row heights
  - Data validation rules
  - Conditional formatting rules
  - Charts with full configuration (series, axes, titles, legends)
  - Images and shapes (with base64 encoding)
  - Print settings, headers, and footers
  - Freeze panes, auto filters, and sheet properties

- **Multiple Output Formats**: Generate analysis in various formats:
  - **JSON**: Complete hierarchical data structure
  - **HTML**: Interactive visual report with tabs for each worksheet
  - **Text**: Human-readable text format with tables
  - **CSV**: Cell data export (one file per worksheet)
  - **Excel**: Summary analysis in Excel format

- **Format Support**:
  - ✅ `.xlsx` files (full support)
  - ✅ `.xls` files (legacy format with some limitations)

- **Command-Line Interface**: Easy-to-use CLI for programmers
  - Verbose mode for detailed progress
  - Multiple output formats in a single run
  - Customizable output directory

## Installation

### Option 1: Install from source

```bash
# Clone or navigate to the project directory
cd /path/to/ExcelAnalyzer

# Create virtual environment (recommended)
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Install the package in development mode
pip install -e .
```

### Option 2: Install as a package

```bash
pip install -e /path/to/ExcelAnalyzer
```

## Usage

### Basic Usage

```bash
# Analyze an Excel file and generate JSON output
excel-analyzer myfile.xlsx --json

# Generate multiple output formats
excel-analyzer myfile.xlsx --json --html --text

# Generate all output formats
excel-analyzer myfile.xlsx --json --html --text --csv --excel

# Enable verbose mode for detailed progress
excel-analyzer myfile.xlsx --json --html -v

# Specify output directory
excel-analyzer myfile.xlsx --json -o /path/to/output
```

### Command-Line Options

```
usage: excel-analyzer [-h] [--json] [--html] [--text] [--csv] [--excel]
                     [-o OUTPUT_DIR] [-v] [--version]
                     input_file

positional arguments:
  input_file            Path to Excel file (.xlsx or .xls)

optional arguments:
  -h, --help            show this help message and exit
  --json                Generate JSON output
  --html                Generate HTML output
  --text                Generate text output
  --csv                 Generate CSV output (one file per sheet)
  --excel               Generate Excel summary output
  -o OUTPUT_DIR, --output-dir OUTPUT_DIR
                        Output directory for generated files (default: current directory)
  -v, --verbose         Enable verbose output with detailed progress
  --version             show program's version number and exit
```

### Python Module Usage

You can also use the analyzer as a Python module:

```python
from excel_analyzer.analyzers import XLSXAnalyzer
from excel_analyzer.formatters import JSONFormatter

# Create analyzer
analyzer = XLSXAnalyzer()

# Analyze file
workbook_model = analyzer.analyze('myfile.xlsx', verbose=True)

# Generate output
formatter = JSONFormatter()
formatter.format(workbook_model, 'output.json', verbose=True)

# Access data programmatically
for worksheet in workbook_model.worksheets:
    print(f"Worksheet: {worksheet.name}")
    print(f"  Cells: {len(worksheet.cells)}")
    print(f"  Charts: {len(worksheet.charts)}")
    print(f"  Images: {len(worksheet.images)}")
```

## Output Examples

### JSON Output

```json
{
  "metadata": {
    "analyzer_version": "1.0.0",
    "analyzed_at": "2026-02-08T15:30:00",
    "source_file": "/path/to/file.xlsx"
  },
  "workbook": {
    "file_format": "xlsx",
    "properties": {
      "creator": "John Doe",
      "created": "2026-01-15T10:00:00"
    },
    "worksheets": [
      {
        "name": "Sheet1",
        "index": 0,
        "cells": [
          {
            "coordinate": "A1",
            "value": "Product",
            "data_type": "s",
            "formatting": {
              "font": {
                "name": "Calibri",
                "size": 11,
                "bold": true,
                "color": {"type": "rgb", "value": "#000000"}
              },
              "fill": {
                "pattern_type": "solid",
                "fg_color": {"type": "rgb", "value": "#DDEBF7"}
              }
            }
          }
        ],
        "merged_cells": ["A1:B1"],
        "charts": [...],
        "images": [...]
      }
    ]
  }
}
```

### HTML Output

Interactive HTML report with:
- File metadata and properties
- Tabbed interface for multiple worksheets
- Styled tables showing cell data
- Chart and image information
- Embedded images (base64)
- Color-coded sections

### Text Output

Human-readable text format:
```
================================================================================
EXCEL WORKBOOK ANALYSIS
================================================================================

File: /path/to/file.xlsx
Format: XLSX
Analyzer Version: 1.0.0

--------------------------------------------------------------------------------
WORKBOOK PROPERTIES
--------------------------------------------------------------------------------
Creator: John Doe
Created: 2026-01-15
Total Worksheets: 3

================================================================================
WORKSHEET: Sheet1
================================================================================

Sheet Index: 0
Total Cells: 288
Sheet State: visible

Cell Data (showing first 100 cells):
+------+------+----------+-----------+---------+
| Cell | Type | Value    | Formula   | Format  |
+------+------+----------+-----------+---------+
| A1   | s    | Product  |           | General |
| B1   | s    | Price    |           | General |
+------+------+----------+-----------+---------+
```

### CSV Output

One CSV file per worksheet containing cell data (values only, no formatting).

### Excel Output

Summary Excel file with:
- Overview sheet with file metadata
- One analysis sheet per source worksheet
- Tables with cell data, charts, and images information

## Architecture

The project is organized into modular components:

```
excel_analyzer/
├── analyzers/          # Analysis modules
│   ├── xlsx_analyzer.py    # XLSX file analyzer
│   ├── xls_analyzer.py     # XLS file analyzer
│   ├── cell_analyzer.py    # Cell data extraction
│   ├── format_analyzer.py  # Formatting extraction
│   ├── structure_analyzer.py   # Merged cells, dimensions
│   ├── chart_analyzer.py   # Chart extraction
│   ├── image_analyzer.py   # Image extraction
│   └── ...
├── models/             # Data models
│   ├── workbook.py     # Workbook model
│   ├── worksheet.py    # Worksheet model
│   ├── cell.py         # Cell model
│   ├── formatting.py   # Formatting models
│   ├── chart.py        # Chart models
│   └── image.py        # Image model
├── formatters/         # Output formatters
│   ├── json_formatter.py
│   ├── html_formatter.py
│   ├── text_formatter.py
│   ├── csv_formatter.py
│   └── excel_formatter.py
├── utils/              # Utilities
│   ├── color_utils.py
│   ├── file_utils.py
│   └── logging_utils.py
└── cli.py              # CLI interface
```

## Limitations

### XLS Format (.xls)
The legacy .xls format has some limitations compared to .xlsx:
- Limited chart extraction (charts are difficult to parse in XLS)
- Limited image extraction
- No conditional formatting support
- Basic formatting only

These limitations are due to the xlrd library's capabilities and the binary nature of the XLS format.

## Dependencies

- `openpyxl>=3.1.0` - XLSX file parsing
- `xlrd>=2.0.0` - XLS file parsing
- `Pillow>=10.0.0` - Image handling
- `Jinja2>=3.1.0` - HTML template rendering
- `colorama>=0.4.6` - Colored terminal output
- `tqdm>=4.65.0` - Progress bars
- `tabulate>=0.9.0` - Text table formatting
- `lxml>=4.9.0` - XML parsing

## Use Cases

This tool is ideal for:

1. **File Migration**: Analyze Excel files before migrating to a new format
2. **Documentation**: Generate comprehensive documentation of Excel file structures
3. **Validation**: Verify that Excel files contain expected data and formatting
4. **Backup & Recovery**: Create detailed snapshots of Excel files for recovery purposes
5. **Programmatic Duplication**: Extract all information needed to recreate Excel files programmatically
6. **Auditing**: Review complex Excel files without opening them
7. **Data Extraction**: Extract data from Excel files for processing in other systems

## Development

### Running Tests

```bash
# Install development dependencies
pip install -r requirements-dev.txt

# Run tests (when implemented)
pytest
```

### Code Style

The project follows PEP 8 style guidelines. Format code with:

```bash
black excel_analyzer/
```

## Troubleshooting

### "Error: At least one output format must be specified"
You must specify at least one output format using `--json`, `--html`, `--text`, `--csv`, or `--excel`.

### "Error: File not found or not readable"
Check that the file path is correct and the file is accessible.

### "Error: Unsupported file format"
Only `.xlsx` and `.xls` files are supported. Other formats like `.xlsb` or `.xlsm` may not work correctly.

### Memory Issues with Large Files
For very large Excel files:
- Use `--json` only to reduce memory usage
- Avoid `--html` which embeds images and can be memory-intensive
- Close other applications to free up memory

## Version History

### Version 1.0.0 (2026-02-08)
- Initial release
- Support for .xlsx and .xls formats
- Multiple output formats (JSON, HTML, Text, CSV, Excel)
- Comprehensive extraction of all Excel features
- CLI interface with verbose mode

## License

This project is provided as-is for educational and commercial use.

## Contributing

Contributions are welcome! Areas for improvement:
- Additional output formats (XML, YAML)
- Excel file recreation from JSON
- Diff tool to compare two Excel files
- Web interface for visual analysis
- Support for .xlsb format
- Performance optimizations for large files

## Author

Enginerd-2019

## Acknowledgments

- Built with openpyxl for XLSX parsing
- Uses xlrd for legacy XLS support
- Jinja2 for HTML template rendering
- Inspired by the need for comprehensive Excel file documentation tools
