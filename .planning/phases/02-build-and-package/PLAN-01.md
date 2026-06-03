# Plan: Python EXE Build & Package

## Context
Build a standalone Windows .EXE using Python 3.10+, python-pptx, pandas, openpyxl, PyMuPDF, and PyInstaller. The application reads financial data files (Excel + PDF), performs fuzzy name matching across 84 buyer variants, aggregates AR data, and generates one PowerPoint report per client (10 clients total).

**Stack:** Python 3.10+, python-pptx, pandas, openpyxl, PyMuPDF, rapidfuzz, PyInstaller

## Files
- src/main.py
- src/data_loader.py
- src/fuzzy_match.py
- src/ppt_generator.py
- src/charts.py
- src/cli.py
- src/error_handler.py
- requirements.txt
- build.spec
- README.md
- .gitignore

## Objective
Generate the complete source tree for a Python desktop app that:
1. Reads 7 source files from the EXE's directory (via sys._MEIPASS for PyInstaller)
2. Fuzzy-matches 84 buyer name variants across sage_ar_extract.xlsx, fronter_soa.pdf, insurer_*.xlsx
3. Aggregates outstanding balances, ageing buckets (0-30, 31-60, 61-90, 90+), transaction history
4. Generates one formatted .pptx per client in output/ subfolder
5. CLI: --client <name> for single client, no args = all 10
6. Programmatic bar/column chart per PPT (python-pptx chart API)
7. Error handling: missing file → print message + skip client
8. PyInstaller spec: single-folder bundle with _internal/, console=true, one file mode

## Done When
- All .py files created with complete, runnable code
- requirements.txt with pinned versions
- build.spec for PyInstaller
- README.md explaining how to add a client and update data files
- .gitignore excluding __pycache__, build/, dist/, output/