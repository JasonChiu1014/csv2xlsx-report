"""Workbook-level tests, including the reconciliation check."""

import math
import subprocess
import sys
from pathlib import Path

import pytest
from openpyxl import load_workbook
from openpyxl.chart import BarChart, LineChart

from csv2xlsx import build_report, clean_data, load_csv

ROOT = Path(__file__).resolve().parents[1]
SAMPLE = ROOT / "examples" / "sample_orders.csv"
COLS = dict(date_col="Date", category_col="Category", amount_col="Amount", dayfirst=True)


@pytest.fixture(scope="module")
def report(tmp_path_factory):
    out = tmp_path_factory.mktemp("report") / "sample_report.xlsx"
    stats = build_report(SAMPLE, out, **COLS)
    return stats, load_workbook(out)


def _column_values(ws, header: str, header_row: int = 1) -> list:
    col = next(c.column for c in ws[header_row] if c.value == header)
    return [ws.cell(r, col).value for r in range(header_row + 1, ws.max_row + 1)
            if ws.cell(r, col).value is not None]


def _pivot_grand_total(summary) -> float:
    header_row = next(r for r in range(1, summary.max_row + 1) if summary.cell(r, 1).value == "Month")
    total_col = next(c.column for c in summary[header_row] if c.value == "Total")
    total_row = next(r for r in range(header_row + 1, summary.max_row + 1)
                     if summary.cell(r, 1).value == "Total")
    return summary.cell(total_row, total_col).value


def test_workbook_has_expected_tabs(report):
    _, wb = report
    assert wb.sheetnames == ["Summary", "By Month", "By Category", "Data", "Rejected"]


def test_charts_exist(report):
    _, wb = report
    assert len(wb["By Month"]._charts) == 1 and isinstance(wb["By Month"]._charts[0], LineChart)
    assert len(wb["By Category"]._charts) == 1 and isinstance(wb["By Category"]._charts[0], BarChart)


def test_pivot_totals_reconcile_with_cleaned_data(report):
    """The key check: every total in the workbook equals the sum of the Data tab."""
    stats, wb = report
    data_total = sum(_column_values(wb["Data"], "Amount"))
    independent = clean_data(load_csv(SAMPLE), **COLS).clean["Amount"].sum()
    for candidate in (
        independent,
        stats.total,
        _pivot_grand_total(wb["Summary"]),
        sum(_column_values(wb["By Month"], "Total")),
        sum(_column_values(wb["By Category"], "Total")),
        wb["Summary"]["B4"].value,  # KPI "Total amount"
    ):
        assert math.isclose(candidate, data_total, abs_tol=0.005)
    assert math.isclose(sum(_column_values(wb["By Category"], "Share")), 1.0, abs_tol=1e-9)


def test_rejected_rows_land_in_rejected_tab(report):
    _, wb = report
    reasons = _column_values(wb["Rejected"], "Reason")
    assert reasons == ["unparseable date", "unparseable amount"] or sorted(reasons) == [
        "unparseable amount", "unparseable date"]
    assert 42 in _column_values(wb["Rejected"], "Source Row")  # the blank-amount line


def test_no_row_disappears(report):
    stats, wb = report
    raw_rows = len(load_csv(SAMPLE))
    data_rows = len(_column_values(wb["Data"], "Order ID"))
    rejected_rows = len(_column_values(wb["Rejected"], "Reason"))
    assert stats.duplicates_dropped == 1
    assert data_rows + rejected_rows + stats.duplicates_dropped == raw_rows


def test_data_tab_is_a_table_with_autofilter(report):
    _, wb = report
    table = wb["Data"].tables["CleanData"]
    assert table.autoFilter.ref == table.ref
    assert wb["Data"].freeze_panes == "A2"


def test_group_by_quarter_renames_period_tab(tmp_path):
    out = tmp_path / "q.xlsx"
    build_report(SAMPLE, out, group_by=("quarter", "category"), **COLS)
    wb = load_workbook(out)
    assert "By Quarter" in wb.sheetnames
    assert _column_values(wb["By Quarter"], "Quarter") == ["2024Q1", "2024Q2"]


def test_bad_group_by_is_rejected(tmp_path):
    with pytest.raises(ValueError, match="group_by"):
        build_report(SAMPLE, tmp_path / "x.xlsx", group_by=("week", "category"), **COLS)


def test_cli_end_to_end(tmp_path):
    out = tmp_path / "cli.xlsx"
    cmd = [sys.executable, "-m", "csv2xlsx", str(SAMPLE), str(out), "--date-col", "Date",
           "--category-col", "Category", "--amount-col", "Amount", "--dayfirst", "--quiet"]
    proc = subprocess.run(cmd, cwd=ROOT, capture_output=True, text=True, check=False)
    assert proc.returncode == 0, proc.stderr
    assert "1 rejected" not in proc.stdout and "2 rejected" in proc.stdout
    assert out.exists()


def test_cli_reports_missing_column_as_exit_code_1(tmp_path):
    cmd = [sys.executable, "-m", "csv2xlsx", str(SAMPLE), str(tmp_path / "x.xlsx"),
           "--date-col", "Date", "--category-col", "Category", "--amount-col", "Nope"]
    proc = subprocess.run(cmd, cwd=ROOT, capture_output=True, text=True, check=False)
    assert proc.returncode == 1 and "Nope" in proc.stderr
