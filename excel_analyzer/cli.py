"""Command-line interface for Excel Analyzer"""

import sys
import argparse
from pathlib import Path
from .analyzers import XLSXAnalyzer, XLSAnalyzer
from .formatters import (
    JSONFormatter,
    HTMLFormatter,
    TextFormatter,
    CSVFormatter,
    ExcelFormatter,
)
from .utils.file_utils import validate_file, detect_file_format, get_output_filename
from .utils.logging_utils import setup_logging, get_logger
from . import __version__

logger = get_logger(__name__)


def create_parser():
    """Create argument parser"""
    parser = argparse.ArgumentParser(
        prog='excel-analyzer',
        description='Comprehensive Excel file analyzer for programmatic duplication',
        epilog='For programmers who need complete Excel file analysis'
    )

    # Required arguments
    parser.add_argument(
        'input_file',
        help='Path to Excel file (.xlsx or .xls)'
    )

    # Output format flags (multiple allowed)
    parser.add_argument(
        '--json',
        action='store_true',
        help='Generate JSON output'
    )
    parser.add_argument(
        '--html',
        action='store_true',
        help='Generate HTML output'
    )
    parser.add_argument(
        '--text',
        action='store_true',
        help='Generate text output'
    )
    parser.add_argument(
        '--csv',
        action='store_true',
        help='Generate CSV output (one file per sheet)'
    )
    parser.add_argument(
        '--excel',
        action='store_true',
        help='Generate Excel summary output'
    )

    # Output directory
    parser.add_argument(
        '-o', '--output-dir',
        default='.',
        help='Output directory for generated files (default: current directory)'
    )

    # Verbosity
    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Enable verbose output with detailed progress'
    )

    # Version
    parser.add_argument(
        '--version',
        action='version',
        version=f'%(prog)s {__version__}'
    )

    return parser


def main():
    """Main CLI entry point"""
    parser = create_parser()
    args = parser.parse_args()

    # Set up logging
    setup_logging(args.verbose)

    # Validate input file
    if not validate_file(args.input_file):
        print(f"Error: File not found or not readable: {args.input_file}", file=sys.stderr)
        return 1

    # Check if at least one output format is specified
    output_formats = [args.json, args.html, args.text, args.csv, args.excel]
    if not any(output_formats):
        print("Error: At least one output format must be specified", file=sys.stderr)
        print("Use --json, --html, --text, --csv, or --excel", file=sys.stderr)
        return 1

    # Detect file format
    file_format = detect_file_format(args.input_file)
    if file_format not in ['xlsx', 'xls']:
        print(f"Error: Unsupported file format. Only .xlsx and .xls files are supported.", file=sys.stderr)
        return 1

    # Create analyzer
    if file_format == 'xlsx':
        analyzer = XLSXAnalyzer()
    else:
        analyzer = XLSAnalyzer()

    # Analyze file
    try:
        if args.verbose:
            logger.info(f"Analyzing {args.input_file}...")

        workbook_model = analyzer.analyze(args.input_file, args.verbose)

        if args.verbose:
            logger.info("Analysis complete. Generating outputs...")

    except Exception as e:
        print(f"Error analyzing file: {e}", file=sys.stderr)
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1

    # Prepare output directory
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # Get base filename
    base_name = Path(args.input_file).stem

    # Generate outputs
    try:
        if args.json:
            formatter = JSONFormatter()
            output_path = output_dir / f"{base_name}_analysis.json"
            formatter.format(workbook_model, output_path, args.verbose)
            print(f"✓ Generated: {output_path}")

        if args.html:
            formatter = HTMLFormatter()
            output_path = output_dir / f"{base_name}_analysis.html"
            formatter.format(workbook_model, output_path, args.verbose)
            print(f"✓ Generated: {output_path}")

        if args.text:
            formatter = TextFormatter()
            output_path = output_dir / f"{base_name}_analysis.txt"
            formatter.format(workbook_model, output_path, args.verbose)
            print(f"✓ Generated: {output_path}")

        if args.csv:
            formatter = CSVFormatter()
            output_paths = formatter.format(workbook_model, output_dir, args.verbose)
            for path in output_paths:
                print(f"✓ Generated: {path}")

        if args.excel:
            formatter = ExcelFormatter()
            output_path = output_dir / f"{base_name}_analysis.xlsx"
            formatter.format(workbook_model, output_path, args.verbose)
            print(f"✓ Generated: {output_path}")

        if args.verbose:
            logger.info("All outputs generated successfully!")

        return 0

    except Exception as e:
        print(f"Error generating output: {e}", file=sys.stderr)
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1


if __name__ == '__main__':
    sys.exit(main())
