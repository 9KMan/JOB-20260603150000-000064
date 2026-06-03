"""Command-line interface for the Finance Report Generator."""

from __future__ import annotations

import argparse
from typing import List, Optional


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="finance-report",
        description=(
            "Generate per-client .pptx finance reports. "
            "Source files are loaded from the EXE bundle (sys._MEIPASS) "
            "when frozen, or from the project's 'data/' folder otherwise."
        ),
    )
    parser.add_argument(
        "--client",
        type=str,
        default=None,
        help=(
            "Generate a report for a single client. Accepts either a "
            "canonical buyer name or a fuzzy-matched variant. Omit to "
            "process every client found in the financial data."
        ),
    )
    parser.add_argument(
        "--output",
        type=str,
        default=None,
        help=(
            "Directory to write .pptx files into. Defaults to "
            "'output/' next to the EXE (or the project root in dev mode)."
        ),
    )
    parser.add_argument(
        "--data-dir",
        type=str,
        default=None,
        help=(
            "Override the directory containing the 7 source files. By "
            "default the EXE bundle directory (sys._MEIPASS) is used."
        ),
    )
    parser.add_argument(
        "--no-charts",
        action="store_true",
        help="Skip programmatic chart generation (text-only reports).",
    )
    parser.add_argument(
        "--threshold",
        type=int,
        default=80,
        help="Fuzzy-match score threshold (0-100). Default: 80.",
    )
    parser.add_argument(
        "--verbose",
        "-v",
        action="store_true",
        help="Enable verbose logging.",
    )
    return parser


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = build_parser()
    return parser.parse_args(argv)
