# Finance Report Generator — Standalone Windows .EXE

Builds formatted PowerPoint finance reports for 10 insurance clients from Excel and PDF source files. No Python or libraries needed on end-user machine.

## How to Run
1. Place all 7 source files in the same folder as `FinanceReportGenerator.exe`:
   - `sage_ar_extract.xlsx`
   - `fronter_soa.pdf`
   - `insurer_client_01.xlsx` through `insurer_client_10.xlsx`
2. Double-click `FinanceReportGenerator.exe`
3. Reports are generated in `output/` subfolder (one .pptx per client)

## CLI Options
- All clients (default): just run the .EXE
- Single client: `FinanceReportGenerator.exe --client client_01`
- Help: `FinanceReportGenerator.exe --help`

## How to Add a New Client
1. Add client entry to `src/client_config.py` → CLIENTS dict
2. Add buyer's name variants to BUYER_NAME_VARIANTS (8-9 variants recommended)
3. Place the insurer data file (e.g., `insurer_client_11.xlsx`) in the same folder
4. Rebuild the .EXE with `pyinstaller build.spec`

## How to Update Data Files
Replace the Excel/PDF files in the same folder as the .EXE. No rebuild needed.

## System Requirements
- Windows 10 or Windows 11
- No Python, pip, or any libraries required on end-user machine
- Approximately 500MB disk space for the .EXE and its _internal folder

## Build Instructions (for developer)
1. Install dependencies: `pip install -r requirements.txt`
2. Build EXE: `pyinstaller build.spec`
3. Find output in `dist/FinanceReportGenerator/`
