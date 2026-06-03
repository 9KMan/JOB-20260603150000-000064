"""Loader for the 7 source files.

When the program is frozen by PyInstaller, the source files live next to
the EXE inside the temporary bundle directory referenced by
``sys._MEIPASS``. When run as a normal Python script, they live in the
project's ``data/`` folder.
"""

from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import pandas as pd


# The 7 source files we expect. They are looked up inside the bundle.
SOURCE_FILES = {
    "buyers_canonical": "buyers_canonical.csv",
    "buyer_variants": "buyer_variants.csv",
    "financial_data": "financial_data.xlsx",
    "quarterly_revenue": "quarterly_revenue.xlsx",
    "expenses": "expenses.xlsx",
    "kpi_metrics": "kpi_metrics.xlsx",
    "client_metadata": "client_metadata.xlsx",
}


@dataclass
class SourceData:
    """Container for the 7 loaded source tables."""

    buyers_canonical: pd.DataFrame
    buyer_variants: pd.DataFrame
    financial_data: pd.DataFrame
    quarterly_revenue: pd.DataFrame
    expenses: pd.DataFrame
    kpi_metrics: pd.DataFrame
    client_metadata: pd.DataFrame


def get_data_dir(override: Optional[str] = None) -> Path:
    """Return the directory holding the 7 source files.

    Resolution order:
      1. ``override`` argument (e.g. ``--data-dir``).
      2. ``sys._MEIPASS`` when the program is frozen by PyInstaller.
      3. ``<project>/data/`` for normal script execution.
    """
    if override:
        p = Path(override).expanduser().resolve()
        if not p.exists():
            raise FileNotFoundError(f"--data-dir does not exist: {p}")
        return p

    if getattr(sys, "frozen", False):
        # PyInstaller one-file / one-folder build: data/ is bundled at the
        # root of the temporary extraction directory.
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))

    # Development: assume data/ sits next to the src/ package.
    return Path(__file__).resolve().parent.parent / "data"


def _read(path: Path, kind: str) -> pd.DataFrame:
    """Read a CSV or Excel file. Raises ``FileNotFoundError`` on missing."""
    if not path.exists():
        raise FileNotFoundError(f"Required source file missing: {path}")
    suffix = path.suffix.lower()
    try:
        if suffix == ".csv":
            df = pd.read_csv(path, dtype=str, keep_default_na=False)
        elif suffix in {".xlsx", ".xls"}:
            df = pd.read_excel(path, sheet_name=0, dtype=str)
        else:
            raise ValueError(f"Unsupported file type for {path}")
    except FileNotFoundError:
        raise
    except Exception as exc:
        raise RuntimeError(f"Failed to read {path}: {exc}") from exc

    # Strip whitespace from string columns for consistent matching
    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].astype(str).str.strip()

    print(f"[DATA] Loaded {kind:20s} rows={len(df):5d}  from {path.name}")
    return df


def _validate_minimal(df: pd.DataFrame, required: list, label: str) -> None:
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            f"Source '{label}' is missing required column(s): {missing}"
        )


def load_all_sources(
    data_dir: Optional[str] = None,
) -> SourceData:
    """Load all 7 source files and return a :class:`SourceData` bundle."""
    base = get_data_dir(data_dir)

    buyers_canonical = _read(
        base / SOURCE_FILES["buyers_canonical"], "buyers_canonical"
    )
    _validate_minimal(
        buyers_canonical, ["canonical_name"], "buyers_canonical"
    )

    buyer_variants = _read(
        base / SOURCE_FILES["buyer_variants"], "buyer_variants"
    )
    _validate_minimal(buyer_variants, ["variant"], "buyer_variants")

    financial_data = _read(
        base / SOURCE_FILES["financial_data"], "financial_data"
    )
    _validate_minimal(
        financial_data, ["buyer", "period", "revenue"], "financial_data"
    )

    quarterly_revenue = _read(
        base / SOURCE_FILES["quarterly_revenue"], "quarterly_revenue"
    )
    _validate_minimal(
        quarterly_revenue, ["buyer", "Q1", "Q2", "Q3", "Q4"],
        "quarterly_revenue",
    )

    expenses = _read(base / SOURCE_FILES["expenses"], "expenses")
    _validate_minimal(
        expenses, ["buyer", "category", "amount"], "expenses"
    )

    kpi_metrics = _read(base / SOURCE_FILES["kpi_metrics"], "kpi_metrics")
    _validate_minimal(
        kpi_metrics, ["buyer", "kpi_name", "value"], "kpi_metrics"
    )

    client_metadata = _read(
        base / SOURCE_FILES["client_metadata"], "client_metadata"
    )
    _validate_minimal(
        client_metadata, ["buyer"], "client_metadata"
    )

    return SourceData(
        buyers_canonical=buyers_canonical,
        buyer_variants=buyer_variants,
        financial_data=financial_data,
        quarterly_revenue=quarterly_revenue,
        expenses=expenses,
        kpi_metrics=kpi_metrics,
        client_metadata=client_metadata,
    )


# ---------------------------------------------------------------------------
# Convenience per-file loaders (also used by the unit-test style imports)
# ---------------------------------------------------------------------------
def load_buyers_canonical(data_dir: Optional[str] = None) -> pd.DataFrame:
    return _read(
        get_data_dir(data_dir) / SOURCE_FILES["buyers_canonical"],
        "buyers_canonical",
    )


def load_buyer_variants(data_dir: Optional[str] = None) -> pd.DataFrame:
    return _read(
        get_data_dir(data_dir) / SOURCE_FILES["buyer_variants"],
        "buyer_variants",
    )


def load_financial_data(data_dir: Optional[str] = None) -> pd.DataFrame:
    return _read(
        get_data_dir(data_dir) / SOURCE_FILES["financial_data"],
        "financial_data",
    )


def load_quarterly_revenue(data_dir: Optional[str] = None) -> pd.DataFrame:
    return _read(
        get_data_dir(data_dir) / SOURCE_FILES["quarterly_revenue"],
        "quarterly_revenue",
    )


def load_expenses(data_dir: Optional[str] = None) -> pd.DataFrame:
    return _read(
        get_data_dir(data_dir) / SOURCE_FILES["expenses"],
        "expenses",
    )


def load_kpi_metrics(data_dir: Optional[str] = None) -> pd.DataFrame:
    return _read(
        get_data_dir(data_dir) / SOURCE_FILES["kpi_metrics"],
        "kpi_metrics",
    )


def load_client_metadata(data_dir: Optional[str] = None) -> pd.DataFrame:
    return _read(
        get_data_dir(data_dir) / SOURCE_FILES["client_metadata"],
        "client_metadata",
    )
