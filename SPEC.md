# Job: Python Developer — Automated Finance Report Generator (Packaged .EXE)
# Platform: Upwork | Budget: $300–500 fixed / $20–40/hr
# URL: https://www.upwork.com/jobs/~022062116727469046701

## Business Problem Solved
Client needs a standalone Windows .EXE that automates insurance finance PowerPoint reports. Eliminates manual Excel/PDF processing for 10 clients by building a Python tool that reads source files, fuzzy-matches 84 buyer name variants, and generates one .pptx per client via PyInstaller-packaged .EXE. No Python installation on end-user machines.

## Stack
- Python 3.10+, python-pptx, pandas, openpyxl, pdfplumber/PyMuPDF
- PyInstaller → single .EXE with `_internal/` folder
- sys._MEIPASS for relative data-file paths

## Source Data
- sage_ar_extract.xlsx (AR extract)
- fronter_soa.pdf (statements)
- insurer_*.xlsx (per-client insurer data)

## Pipeline
1. Read 7 source files from EXE directory (relative paths via sys._MEIPASS)
2. Fuzzy-match 84 buyer name variants across sources
3. Aggregate outstanding balances, ageing buckets, transaction history
4. Generate 1 .pptx per client (10 clients) in `output/` subfolder
5. CLI: `--client <name>` for single client; no args = all 10
6. Console summary log on completion

## PPT Requirements
- Slide dimensions, colour palette, fonts, layout matching reference .pptx
- Dynamic content: client name, report date, balance tables, ageing chart
- Programmatic bar/column chart via python-pptx (not embedded images)
- All monetary values: thousand separators + currency symbol

## Error Handling
- Missing source file → clear message + skip that client gracefully

## Deliverables
1. Working .EXE — runs on clean Windows 10/11, no Python installed
2. Source code — private GitHub repo (client added as collaborator)
3. requirements.txt, .spec file, build instructions
4. Google Drive upload for acceptance testing
5. README: how to add a new client + update data files

## Build Quality Gates
- .EXE runs standalone from any Windows path (no hardcoded paths)
- Relative paths via sys._MEIPASS for all data files
- Charts built programmatically (not embedded images)
- Clean commented Python with functions/modules (not one-file script)
- Reproducible: PyInstaller re-run from spec file produces same EXE

## Acceptance Criteria
- Double-click EXE with 7 source files in same folder → 10 .pptx in output/
- EXE works on clean Windows 10/11 with no Python/pip/libraries installed
- Each PPT has at least one programmatic bar/column chart
- --client flag runs single client, no args runs all 10
