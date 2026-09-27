"""Cleaning step: parse dates and amounts, drop duplicates, isolate rejected rows.

Nothing is silently discarded. Exact duplicate rows are counted, and rows whose
date or amount cannot be parsed are returned separately so the report can show
them in a "Rejected" tab.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

log = logging.getLogger(__name__)

SOURCE_ROW = "Source Row"
REASON = "Reason"

# ISO-style formats are tried first; the slash format depends on ``dayfirst``.
_ISO_FORMATS = ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d")
_AMOUNT_JUNK = re.compile(r"[^0-9.\-]")


@dataclass(frozen=True)
class CleanResult:
    """Outcome of :func:`clean_data`.

    Attributes:
        clean: Rows with a parsed date (datetime) and numeric amount.
        rejected: Original rows that failed parsing, plus a ``Reason`` column.
        duplicates_dropped: Number of exact duplicate rows removed.
    """

    clean: pd.DataFrame
    rejected: pd.DataFrame
    duplicates_dropped: int


def load_csv(path: str | Path) -> pd.DataFrame:
    """Read a CSV with every column as text so raw values survive for reporting."""
    frame = pd.read_csv(path, dtype=str, keep_default_na=False, skipinitialspace=True)
    log.info("Read %d rows x %d columns from %s", len(frame), frame.shape[1], path)
    return frame


def parse_dates(values: pd.Series, *, dayfirst: bool = False) -> pd.Series:
    """Parse mixed-format date strings; unparseable values become ``NaT``.

    Accepts ISO dates (``2024-03-05``, optional time, ``2024/03/05``) and slash
    dates. Slash dates are read as DD/MM/YYYY when ``dayfirst`` is true and
    MM/DD/YYYY otherwise; a value impossible in the chosen order is rejected
    rather than guessed, so ambiguous dates are never flipped silently.
    """
    text = values.astype(str).str.strip()
    formats = (*_ISO_FORMATS, "%d/%m/%Y" if dayfirst else "%m/%d/%Y")
    parsed = pd.Series(pd.NaT, index=values.index, dtype="datetime64[us]")
    for fmt in formats:
        parsed = parsed.fillna(pd.to_datetime(text, format=fmt, errors="coerce"))
    return parsed


def parse_amounts(values: pd.Series) -> pd.Series:
    """Strip currency symbols and thousands separators; return floats or ``NaN``.

    ``"$1,200.50"`` -> 1200.5, ``"NT$ 300"`` -> 300.0, ``"(45.00)"`` -> -45.0.
    Blank or non-numeric text becomes ``NaN``.
    """
    text = values.astype(str).str.strip()
    negative = text.str.startswith("(") & text.str.endswith(")")
    digits = text.str.replace(_AMOUNT_JUNK, "", regex=True)
    numbers = pd.to_numeric(digits, errors="coerce").astype(float)
    return numbers.where(~negative, -numbers)


def clean_data(
    frame: pd.DataFrame,
    *,
    date_col: str,
    category_col: str,
    amount_col: str,
    dayfirst: bool = False,
) -> CleanResult:
    """Clean a raw text DataFrame into typed rows plus a rejected-rows table.

    Steps: drop exact duplicates (counted), parse ``date_col`` and
    ``amount_col``, fill blank categories with ``"Uncategorized"``, and move any
    row with an unparseable date or amount to ``rejected`` with a reason.
    """
    missing = [c for c in (date_col, category_col, amount_col) if c not in frame.columns]
    if missing:
        raise KeyError(f"Column(s) not found in CSV: {missing}; have {list(frame.columns)}")

    raw = frame.copy()
    raw.insert(0, SOURCE_ROW, range(2, len(raw) + 2))  # 1-based, header is row 1
    dup_mask = raw.duplicated(subset=list(frame.columns), keep="first")
    duplicates = int(dup_mask.sum())
    raw = raw.loc[~dup_mask]

    dates = parse_dates(raw[date_col], dayfirst=dayfirst)
    amounts = parse_amounts(raw[amount_col])
    reasons = pd.Series("", index=raw.index)
    reasons = reasons.where(dates.notna(), "unparseable date")
    reasons = reasons.where(amounts.notna(), (reasons + "; unparseable amount").str.strip("; "))
    bad = reasons != ""

    rejected = raw.loc[bad].copy()
    rejected[REASON] = reasons[bad]

    clean = raw.loc[~bad].drop(columns=SOURCE_ROW).copy()
    clean[date_col] = dates[~bad]
    clean[amount_col] = amounts[~bad]
    category = clean[category_col].astype(str).str.strip()
    clean[category_col] = category.mask(category == "", "Uncategorized")
    clean = clean.reset_index(drop=True)

    log.info(
        "Cleaned: %d kept, %d rejected, %d duplicates dropped",
        len(clean), len(rejected), duplicates,
    )
    return CleanResult(clean=clean, rejected=rejected.reset_index(drop=True), duplicates_dropped=duplicates)
