"""Setup configuration for Excel Analyzer"""

from setuptools import setup, find_packages
from pathlib import Path

# Read README
readme_path = Path(__file__).parent / "README.md"
long_description = ""
if readme_path.exists():
    long_description = readme_path.read_text(encoding='utf-8')

setup(
    name='excel-analyzer',
    version='1.0.0',
    description='Comprehensive Excel file analyzer for programmatic duplication',
    long_description=long_description,
    long_description_content_type='text/markdown',
    author='Excel Analyzer Team',
    python_requires='>=3.8',
    packages=find_packages(),
    install_requires=[
        'openpyxl>=3.1.0',
        'xlrd>=2.0.0',
        'Pillow>=10.0.0',
        'Jinja2>=3.1.0',
        'colorama>=0.4.6',
        'tqdm>=4.65.0',
        'tabulate>=0.9.0',
        'lxml>=4.9.0',
    ],
    entry_points={
        'console_scripts': [
            'excel-analyzer=excel_analyzer.cli:main',
        ],
    },
    classifiers=[
        'Development Status :: 4 - Beta',
        'Intended Audience :: Developers',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3.10',
        'Programming Language :: Python :: 3.11',
        'Programming Language :: Python :: 3.12',
    ],
)
