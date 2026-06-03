"""Finance Report Generator — entry point.

Reads 7 source files (bundled with the EXE via PyInstaller sys._MEIPASS),
fuzzy-matches 84 buyer-name variants to canonical names, and writes one
formatted .pptx per client into the ``output/`` subfolder next to the EXE.
"""

from __future__ import annotations

import sys
import traceback
from pathlib import Path
from typing import List, Optional

from cli import parse_args
from data_loader import (
    SourceData,
    load_all_sources,
)
from fuzzy_match import build_buyer_map, canonicalize_column
from ppt_generator import generate_report


def resolve_clients_to_process(args, source: SourceData) -> List[str]:
    """Determine which clients to process.

    * If ``--client`` is given, try to resolve the value (variant or
      canonical) to a single canonical name.
    * Otherwise, return every canonical buyer present in the financial data.
    """
    if not args.client:
        # All canonical buyers that appear in financial data
        return sorted(
            source.financial_data["buyer"].dropna().astype(str).unique().tolist()
        )

    # Single-client mode — resolve via the buyer map
    buyer_map = build_buyer_map(source.buyer_variants, source.buyers_canonical)
    requested = args.client.strip()
    if requested in source.buyers_canonical["canonical_name"].values:
        return [requested]

    if requested in buyer_map and buyer_map[requested] is not None:
        return [buyer_map[requested]]

    # Fallback — try direct fuzzy match against canonical list
    from rapidfuzz import process, fuzz

    canonical = source.buyers_canonical["canonical_name"].tolist()
    match = process.extractOne(requested, canonical, scorer=fuzz.WRatio)
    if match and match[1] >= 75:
        print(
            f"[INFO] '--client' value '{requested}' fuzzy-matched to "
            f"'{match[0]}' (score={match[1]})."
        )
        return [match[0]]

    print(
        f"[ERROR] Could not resolve client '{requested}' to any canonical "
        f"buyer. Available: {', '.join(canonical)}"
    )
    return []


def determine_output_dir(args) -> Path:
    """Pick the output directory.

    For a frozen EXE, default to ``<exe-dir>/output``. Otherwise default to
    ``<project-root>/output``.
    """
    if args.output:
        out = Path(args.output).expanduser().resolve()
    elif getattr(sys, "frozen", False):
        out = Path(sys.executable).resolve().parent / "output"
    else:
        out = Path(__file__).resolve().parent.parent / "output"
    out.mkdir(parents=True, exist_ok=True)
    return out


def main(argv: Optional[List[str]] = None) -> int:
    args = parse_args(argv)

    print("=" * 70)
    print("Finance Report Generator")
    print("=" * 70)

    # ------------------------------------------------------------------ data
    try:
        source = load_all_sources(data_dir=args.data_dir)
    except FileNotFoundError as exc:
        print(f"[ERROR] {exc}")
        return 2
    except Exception as exc:  # pragma: no cover - defensive
        print(f"[ERROR] Failed to load source data: {exc}")
        traceback.print_exc()
        return 2

    # Print quick stats
    print(
        f"[INFO] Loaded {len(source.buyers_canonical)} canonical buyers, "
        f"{len(source.buyer_variants)} variants, "
        f"{len(source.financial_data)} financial rows, "
        f"{len(source.quarterly_revenue)} quarterly rows, "
        f"{len(source.expenses)} expense rows, "
        f"{len(source.kpi_metrics)} KPI rows, "
        f"{len(source.client_metadata)} client-meta rows."
    )

    # ---------------------------------------------------------- client list
    clients = resolve_clients_to_process(args, source)
    if not clients:
        print("[ERROR] No clients to process. Exiting.")
        return 1
    print(f"[INFO] Will generate reports for {len(clients)} client(s).")

    # ----------------------------------------------------------- output dir
    out_dir = determine_output_dir(args)
    print(f"[INFO] Output directory: {out_dir}")

    # -------------------------------------------------------- per-client run
    failures = 0
    for client in clients:
        print("-" * 70)
        print(f"[INFO] Generating report for: {client}")
        try:
            path = generate_report(
                client=client,
                source=source,
                output_dir=out_dir,
                include_charts=not args.no_charts,
            )
            print(f"[OK]   Wrote: {path}")
        except FileNotFoundError as exc:
            failures += 1
            print(f"[WARN] {client}: {exc} — skipping.")
        except KeyError as exc:
            failures += 1
            print(f"[WARN] {client}: missing expected column {exc} — skipping.")
        except Exception as exc:  # pragma: no cover - defensive
            failures += 1
            print(f"[ERROR] {client}: {exc}")
            traceback.print_exc()

    print("=" * 70)
    print(
        f"[DONE] {len(clients) - failures} succeeded, {failures} failed. "
        f"Reports in: {out_dir}"
    )
    return 0 if failures == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
