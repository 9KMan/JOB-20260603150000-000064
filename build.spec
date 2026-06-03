# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for Finance Report Generator
# Build (one-folder):   pyinstaller build.spec
# Output:               dist/FinanceReportGenerator/FinanceReportGenerator.exe
#                       dist/FinanceReportGenerator/_internal/...   (deps, data)
#
# Runtime data lookup in main.py should use:
#   if getattr(sys, 'frozen', False):
#       BASE_DIR = sys._MEIPASS          # one-folder -> points to _internal
#   else:
#       BASE_DIR = os.path.dirname(os.path.abspath(__file__))

import sys
import os

block_cipher = None

# ------------------------------------------------------------------
# Analysis
# ------------------------------------------------------------------
a = Analysis(
    ['src/main.py'],
    pathex=[],
    binaries=[],
    datas=[
        # Bundle every data file PyInstaller glob supports.
        # Destination is the *runtime* path inside _internal/.
        ('src/*.xlsx', 'src'),
        ('src/*.pdf',  'src'),
        ('src/*.py',   'src'),
    ],
    hiddenimports=[
        # --- Office / document I/O ---
        'pptx',
        'pptx.chart.data',
        'pptx.enum.chart',
        'pandas',
        'pandas._libs',
        'pandas._libs.tslibs',
        'openpyxl',
        'openpyxl.cell',
        'openpyxl.cell._writer',
        'fitz',                       # PyMuPDF

        # --- Fuzzy matching ---
        'rapidfuzz',
        'rapidfuzz.process',
        'rapidfuzz.distance',

        # --- Scientific / numeric ---
        'scipy',
        'scipy.spatial',
        'scipy.cluster',
        'scipy.optimize',
        'scipy.signal',
        'numpy',
        'numpy.core',
        'numpy.random',
        'numpy.fft',

        # --- ML ---
        'sklearn',
        'sklearn.utils',
        'sklearn.metrics',
        'sklearn.cluster',
        'sklearn.preprocessing',
        'sklearn.decomposition',
        'sklearn.neighbors',
        'sklearn.svm',
        'sklearn.tree',
        'sklearn.ensemble',

        # --- Imaging ---
        'PIL',
        'PIL._imaging',
        'PIL._imagingtk',
        'PIL._imagingft',
        'PIL._imagingmorph',
        'PIL._imagingcms',
        'PIL._imagingmath',
        'PIL._imagingresample',
        'PIL._imagingwebp',
        'PIL._imagingjpeg',
        'PIL._imagingpng',
        'PIL._imagingtiff',

        # --- Common transitive deps that PyInstaller's static
        #     analysis frequently misses ---
        'python_dateutil',
        'pytz',
        'six',
        'cycler',
        'kiwisolver',
        'pyparsing',
        'et_xmlfile',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Trim bloat; add/remove as you measure.
        'tkinter',
        'test',
        'unittest',
        'pydoc',
        'doctest',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# ------------------------------------------------------------------
# PYZ archive (pure-python modules, zipped)
# ------------------------------------------------------------------
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# ------------------------------------------------------------------
# EXE (bootloader + entry script only — one-folder pattern)
# ------------------------------------------------------------------
# exclude_binaries=True is REQUIRED for a one-folder / onedir build:
# it keeps the exe small and lets COLLECT lay out the rest into
#   dist/FinanceReportGenerator/_internal/
exe = EXE(
    pyz,
    a.scripts,
    [],                       # <-- NO binaries/zipfiles/datas here
    exclude_binaries=True,    # <-- one-folder key flag
    name='FinanceReportGenerator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,      # default: %TEMP%\_MEIxxxx for one-file,
                              # same dir as exe for one-folder
    console=True,             # <-- show stdout/stderr console
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,                # set to 'assets/app.ico' if you have one
)

# ------------------------------------------------------------------
# COLLECT (assemble the onedir distribution folder)
# ------------------------------------------------------------------
# Result:
#   dist/FinanceReportGenerator/
#       FinanceReportGenerator.exe
#       _internal/
#           python311.dll, *.pyd, *.so, *.dll
#           *.dist-info/...
#           src/*.xlsx, src/*.pdf, src/*.py   (from datas)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[
        'vcruntime140.dll',
        'vcruntime140_1.dll',
        'ucrtbase.dll',
        'msvcp140.dll',
        'msvcp140_1.dll',
        'msvcp140_2.dll',
        'concrt140.dll',
    ],
    name='FinanceReportGenerator',
)
