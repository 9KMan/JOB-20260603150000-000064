"""Fuzzy matching of buyer-name variants to canonical names.

The data sources contain 84 raw buyer-name variants (typos, abbreviations,
spelling differences). They are mapped to a small master list of canonical
buyer names using :mod:`rapidfuzz`.
"""

from __future__ import annotations

from typing import Dict, Iterable, Optional

import pandas as pd
from rapidfuzz import fuzz, process


# A few sensible scorers to try in order — the best score wins.
_SCORERS = (
    fuzz.WRatio,
    fuzz.token_set_ratio,
    fuzz.token_sort_ratio,
    fuzz.ratio,
)


def _best_match(query: str, choices: Iterable[str]) -> tuple[Optional[str], int]:
    """Return the best (choice, score) for ``query`` against ``choices``."""
    choices = list(choices)
    if not query or not choices:
        return None, 0
    best_choice: Optional[str] = None
    best_score = 0
    for scorer in _SCORERS:
        match = process.extractOne(query, choices, scorer=scorer)
        if match and match[1] > best_score:
            best_choice, best_score = match[0], match[1]
    return best_choice, best_score


def build_buyer_map(
    variants_df: pd.DataFrame,
    canonical_df: pd.DataFrame,
    threshold: int = 80,
) -> Dict[str, Optional[str]]:
    """Map every variant row to a canonical name.

    Parameters
    ----------
    variants_df
        DataFrame with at least a ``variant`` column.
    canonical_df
        DataFrame with at least a ``canonical_name`` column.
    threshold
        Minimum score (0-100) to accept a match. Lower scores are mapped
        to ``None`` so callers can detect unresolved variants.
    """
    if "variant" not in variants_df.columns:
        raise ValueError("variants_df must contain a 'variant' column")
    if "canonical_name" not in canonical_df.columns:
        raise ValueError("canonical_df must contain a 'canonical_name' column")

    canonical = canonical_df["canonical_name"].astype(str).tolist()
    mapping: Dict[str, Optional[str]] = {}

    unmatched: list = []
    for variant in variants_df["variant"].astype(str).str.strip():
        if not variant or variant.lower() == "nan":
            continue
        match, score = _best_match(variant, canonical)
        if match is not None and score >= threshold:
            mapping[variant] = match
        else:
            mapping[variant] = None
            unmatched.append((variant, match, score))

    if unmatched:
        print(
            f"[FUZZY] {len(unmatched)}/{len(mapping)} variant(s) below "
            f"threshold ({threshold}):"
        )
        for variant, match, score in unmatched[:10]:
            print(
                f"         - '{variant}' "
                f"-> '{match}' (score={score})"
            )
        if len(unmatched) > 10:
            print(f"         ... and {len(unmatched) - 10} more")

    print(
        f"[FUZZY] Mapped {sum(1 for v in mapping.values() if v)}/"
        f"{len(mapping)} variants to canonical names."
    )
    return mapping


def canonicalize_column(
    series: pd.Series,
    buyer_map: Dict[str, Optional[str]],
    canonical_df: pd.DataFrame,
) -> pd.Series:
    """Return a Series where known variants are replaced with canonical names.

    Values already in the canonical list are left untouched. Variants not
    present in the mapping are returned unchanged (so the caller can decide
    how to handle them).
    """
    canonical_set = set(
        canonical_df["canonical_name"].astype(str).str.strip().tolist()
    )
    out = series.astype(str).str.strip()

    def _lookup(v: str) -> str:
        if v in canonical_set:
            return v
        if v in buyer_map and buyer_map[v] is not None:
            return buyer_map[v]
        return v

    return out.map(_lookup)


def resolve_client_name(
    name: str,
    variants_df: pd.DataFrame,
    canonical_df: pd.DataFrame,
    threshold: int = 80,
) -> Optional[str]:
    """Resolve an arbitrary user-supplied client name to a canonical one.

    Tries, in order:
      1. Exact match against canonical names.
      2. Exact match against the variant list (using the cached mapping).
      3. Fuzzy match against the canonical list.
    """
    if not name:
        return None
    name = name.strip()
    canonical = canonical_df["canonical_name"].astype(str).str.strip().tolist()
    if name in canonical:
        return name

    buyer_map = build_buyer_map(variants_df, canonical_df, threshold=threshold)
    if name in buyer_map and buyer_map[name] is not None:
        return buyer_map[name]

    match, score = _best_match(name, canonical)
    if match and score >= threshold:
        return match
    return None
