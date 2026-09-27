"""Unit tests for the cleaning step (no workbook involved)."""

import math

import pandas as pd
import pytest

from csv2xlsx.cleaning import REASON, SOURCE_ROW, clean_data, parse_amounts, parse_dates


def _ts(text: str) -> pd.Timestamp:
    return pd.Timestamp(text)


def test_parse_dates_handles_iso_and_both_slash_orders():
    iso = parse_dates(pd.Series(["2024-03-25", "2024-03-25 14:05:00", "2024/03/25"]))
    assert iso.tolist() == [_ts("2024-03-25"), _ts("2024-03-25 14:05"), _ts("2024-03-25")]
    assert parse_dates(pd.Series(["03/25/2024"]), dayfirst=False).tolist() == [_ts("2024-03-25")]
    assert parse_dates(pd.Series(["25/03/2024"]), dayfirst=True).tolist() == [_ts("2024-03-25")]


def test_parse_dates_ambiguous_value_follows_dayfirst_flag():
    assert parse_dates(pd.Series(["12/03/2024"]), dayfirst=False).tolist() == [_ts("2024-12-03")]
    assert parse_dates(pd.Series(["12/03/2024"]), dayfirst=True).tolist() == [_ts("2024-03-12")]


def test_parse_dates_rejects_impossible_and_garbage_values():
    parsed = parse_dates(pd.Series(["25/03/2024", "2024-02-30", "yesterday", ""]), dayfirst=False)
    assert parsed.isna().all()


def test_parse_amounts_strips_currency_and_separators():
    values = pd.Series(["$1,200.50", "NT$ 3,000", "(45.00)", " 12.5 ", "-7", "", "n/a"])
    parsed = parse_amounts(values).tolist()
    assert parsed[:5] == [1200.5, 3000.0, -45.0, 12.5, -7.0]
    assert all(math.isnan(v) for v in parsed[5:])


def _raw() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "Date": ["2024-01-05", "2024-01-05", "02/10/2024", "bad-date", "2024-03-01", "2024-03-02"],
            "Category": ["Books", "Books", "Toys", "Toys", "", "Apparel"],
            "Amount": ["10.00", "10.00", "$1,000.00", "5.00", "20", ""],
        }
    )


def test_clean_reports_exact_duplicates():
    result = clean_data(_raw(), date_col="Date", category_col="Category", amount_col="Amount")
    assert result.duplicates_dropped == 1
    assert len(result.clean) + len(result.rejected) + result.duplicates_dropped == len(_raw())


def test_clean_moves_bad_rows_to_rejected_with_reason_and_source_row():
    result = clean_data(_raw(), date_col="Date", category_col="Category", amount_col="Amount")
    rejected = result.rejected.set_index(SOURCE_ROW)
    assert rejected.loc[5, REASON] == "unparseable date"
    assert rejected.loc[7, REASON] == "unparseable amount"
    assert rejected.loc[5, "Amount"] == "5.00"  # original text is preserved


def test_clean_types_values_and_fills_blank_category():
    clean = clean_data(_raw(), date_col="Date", category_col="Category", amount_col="Amount").clean
    assert clean["Amount"].tolist() == [10.0, 1000.0, 20.0]
    assert clean["Date"].tolist() == [_ts("2024-01-05"), _ts("2024-02-10"), _ts("2024-03-01")]
    assert clean["Category"].tolist() == ["Books", "Toys", "Uncategorized"]


def test_clean_reports_both_reasons_on_one_row():
    frame = pd.DataFrame({"Date": ["??"], "Category": ["x"], "Amount": ["??"]})
    result = clean_data(frame, date_col="Date", category_col="Category", amount_col="Amount")
    assert result.rejected[REASON].tolist() == ["unparseable date; unparseable amount"]


def test_clean_missing_column_is_a_clear_error():
    with pytest.raises(KeyError, match="Amt"):
        clean_data(_raw(), date_col="Date", category_col="Category", amount_col="Amt")
