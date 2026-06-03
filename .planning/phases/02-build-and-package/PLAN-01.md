# Plan: Python EXE Build & Package

## Objective
**Project:** Python Finance Report Generator — Standalone Windows .EXE

Build a Python desktop app that:
1. Reads 7 source files from the EXE directory (via sys._MEIPASS for PyInstaller)
2. Fuzzy-matches 84 buyer name variants across sage_ar_extract.xlsx, fronter_soa.pdf, insurer_*.xlsx using rapidfuzz
3. Aggregates outstanding balances, ageing buckets (0-30, 31-60, 61-90, 90+ days), transaction history
4. Generates one formatted .pptx per client in output/ subfolder
5. CLI: --client <name> for single client, no args = all 10 clients
6. Programmatic bar/column chart per PPT via python-pptx chart API (not embedded images)
7. Error handling: missing file → print message + skip that client gracefully
8. PyInstaller: console=true, single-folder bundle with _internal/

## Files
src/main.py
src/data_loader.py
src/fuzzy_match.py
src/ppt_generator.py
src/charts.py
src/cli.py
src/error_handler.py
src/client_config.py
requirements.txt
build.spec
README.md
.gitignore

## Done When
All .py files created with complete, runnable Python code. requirements.txt with pinned versions. build.spec for PyInstaller. README.md explaining how to add a new client and update data files. src/main.py orchestrates full pipeline with --client CLI flag. All modules error-handled gracefully.