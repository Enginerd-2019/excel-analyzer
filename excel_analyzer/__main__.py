"""Entry point for python -m excel_analyzer"""

from .cli import main
import sys

if __name__ == '__main__':
    sys.exit(main())
