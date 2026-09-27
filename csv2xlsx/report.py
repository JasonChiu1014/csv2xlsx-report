"""Workbook builder: Summary, By Month, By Category, Data and Rejected tabs."""

from __future__ import annotations

import logging
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Sequence

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.worksheet import Worksheet

from .cleaning import clean_data, load_csv

log = logging.getLogger(__name__)

PERIODS: dict[str, tuple[str, str]] = {
    "month": ("M", "Month"),
    "quarter": ("Q", "Quarter"),
    "year": ("Y", "Year"),
}
DATE_FORMAT = "yyyy-mm-dd"
INT_FORMAT = "#,##0"
PCT_FORMAT = "0.0%"
_HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
_HEADER_FONT = Font(bold=True, color="FFFFFF")
_TOTAL_FILL = PatternFill("solid", fgColor="DDEBF7")
_BOLD = Font(bold=True)


@dataclass(frozen=True)
class ReportStats:
    """Headline numbers of a generated report (useful for logs and tests)."""

    out_path: Path
    rows: int
    rejected: int
    duplicates_dropped: int
    total: float


def build_report(
    csv_path: str | Path,
    out_xlsx: str | Path,
    *,
    date_col: str,
    category_col: str,
    amount_col: str,
    group_by: Sequence[str] = ("month", "category"),
    currency_format: str = "#,##0.00",
    dayfirst: bool = False,
) -> ReportStats:
    """Read ``csv_path``, clean it and write a formatted report to ``out_xlsx``.

    Args:
        csv_path: Source CSV export.
        out_xlsx: Destination ``.xlsx`` path (parent directories are created).
        date_col, category_col, amount_col: Column names in the CSV.
        group_by: ``(period, "category")`` where period is ``"month"``,
            ``"quarter"`` or ``"year"``; drives the pivot rows and trend chart.
        currency_format: Excel number format applied to every amount cell.
        dayfirst: Read slash dates as DD/MM/YYYY instead of MM/DD/YYYY.

    Returns:
        :class:`ReportStats` with row counts and the grand total.
    """
    period = _validate_group_by(group_by)
    csv_path, out_xlsx = Path(csv_path), Path(out_xlsx)
    result = clean_data(
        load_csv(csv_path),
        date_col=date_col, category_col=category_col, amount_col=amount_col, dayfirst=dayfirst,
    )
    clean = result.clean
    if clean.empty:
        raise ValueError("No usable rows after cleaning; check column names and --dayfirst.")

    freq, label = PERIODS[period]
    key = label if label not in clean.columns else f"Report {label}"
    clean = clean.assign(**{key: clean[date_col].dt.to_period(freq).astype(str)})

    by_period = clean.groupby(key)[amount_col].agg(Orders="size", Total="sum").reset_index()
    by_cat = (
        clean.groupby(category_col)[amount_col].agg(Orders="size", Total="sum")
        .reset_index().sort_values("Total", ascending=False, ignore_index=True)
    )
    by_cat["Share"] = by_cat["Total"] / by_cat["Total"].sum()
    pivot = _pivot(clean, key, category_col, amount_col)

    wb = Workbook()
    _write_summary(wb.active, csv_path.name, result, by_period, pivot, key, currency_format)
    _write_period_sheet(wb.create_sheet(f"By {label}"), by_period, key, currency_format)
    _write_category_sheet(wb.create_sheet("By Category"), by_cat, category_col, currency_format)
    _write_data_sheet(wb.create_sheet("Data"), clean, date_col, amount_col, currency_format)
    _write_rejected_sheet(wb.create_sheet("Rejected"), result.rejected)

    out_xlsx.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_xlsx)
    total = float(clean[amount_col].sum())
    log.info("Wrote %s (total %.2f over %d rows)", out_xlsx, total, len(clean))
    return ReportStats(out_xlsx, len(clean), len(result.rejected), result.duplicates_dropped, total)


def _validate_group_by(group_by: Sequence[str]) -> str:
    if len(group_by) != 2 or group_by[0] not in PERIODS or group_by[1] != "category":
        raise ValueError(f"group_by must be (one of {sorted(PERIODS)}, 'category'); got {tuple(group_by)}")
    return group_by[0]


def _pivot(clean: pd.DataFrame, key: str, category_col: str, amount_col: str) -> pd.DataFrame:
    """Period x category sum of amounts with a Total column and a Total row."""
    table = pd.pivot_table(
        clean, index=key, columns=category_col, values=amount_col, aggfunc="sum", fill_value=0.0
    )
    table.columns = [str(c) for c in table.columns]
    table["Total"] = table.sum(axis=1)
    table.loc["Total"] = table.sum(axis=0)
    return table.reset_index()


def _write_frame(ws: Worksheet, frame: pd.DataFrame, start_row: int = 1,
                 formats: dict[str, str] | None = None) -> int:
    """Write header + rows starting at ``start_row``; return the last row written."""
    formats = formats or {}
    for col, name in enumerate(frame.columns, 1):
        ws.cell(start_row, col, str(name))
    for row_idx, row in enumerate(frame.itertuples(index=False), start_row + 1):
        for col, value in enumerate(row, 1):
            cell = ws.cell(row_idx, col, value)
            fmt = formats.get(str(frame.columns[col - 1]))
            if fmt:
                cell.number_format = fmt
    _style_header(ws, start_row, len(frame.columns))
    return start_row + len(frame)


def _style_header(ws: Worksheet, row: int, ncols: int) -> None:
    for col in range(1, ncols + 1):
        cell = ws.cell(row, col)
        cell.fill, cell.font = _HEADER_FILL, _HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _autofit(ws: Worksheet, min_width: int = 8, max_width: int = 45, skip_rows: int = 0) -> None:
    """Set each column width from its longest rendered value (ignoring the first ``skip_rows``)."""
    widths: dict[int, int] = {}
    for row in ws.iter_rows(min_row=skip_rows + 1):
        for cell in row:
            if cell.value is None:
                continue
            if isinstance(cell.value, (datetime, date)):
                length = 10
            elif isinstance(cell.value, float):
                length = len(f"{cell.value:,.2f}")
            else:
                length = len(str(cell.value))
            widths[cell.column] = max(widths.get(cell.column, 0), length)
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = min(max(width + 2, min_width), max_width)


def _write_summary(ws: Worksheet, source: str, result, by_period: pd.DataFrame,
                   pivot: pd.DataFrame, key: str, currency_format: str) -> None:
    ws.title = "Summary"
    clean, amounts = result.clean, by_period["Total"]
    best, worst = by_period.loc[amounts.idxmax()], by_period.loc[amounts.idxmin()]
    ws["A1"] = "Report summary"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = f"Source: {source}  |  Generated: {date.today().isoformat()}"
    kpis = [
        ("Total amount", float(amounts.sum()), currency_format),
        ("Rows included", len(clean), INT_FORMAT),
        ("Average amount", float(amounts.sum() / len(clean)), currency_format),
        (f"Highest {key.lower()}", best[key], None),
        (f"Lowest {key.lower()}", worst[key], None),
        ("Duplicates removed", result.duplicates_dropped, INT_FORMAT),
        ("Rejected rows", len(result.rejected), INT_FORMAT),
    ]
    for offset, (name, value, fmt) in enumerate(kpis):
        row = 4 + offset
        ws.cell(row, 1, name).font = _BOLD
        cell = ws.cell(row, 2, value)
        if fmt:
            cell.number_format = fmt
    for row, extreme in ((7, best), (8, worst)):
        ws.cell(row, 3, float(extreme["Total"])).number_format = currency_format

    start = 4 + len(kpis) + 2
    ws.cell(start - 1, 1, f"Amount by {key.lower()} and category").font = _BOLD
    formats = {c: currency_format for c in pivot.columns if c != key}
    last = _write_frame(ws, pivot, start_row=start, formats=formats)
    for col in range(1, len(pivot.columns) + 1):          # totals row
        ws.cell(last, col).font, ws.cell(last, col).fill = _BOLD, _TOTAL_FILL
    for row in range(start + 1, last + 1):                # totals column
        cell = ws.cell(row, len(pivot.columns))
        cell.font, cell.fill = _BOLD, _TOTAL_FILL
    ws.freeze_panes = "B3"
    _autofit(ws, skip_rows=2)  # title and source line may overflow column A


def _write_period_sheet(ws: Worksheet, frame: pd.DataFrame, key: str, currency_format: str) -> None:
    last = _write_frame(ws, frame, formats={"Orders": INT_FORMAT, "Total": currency_format})
    chart = LineChart()
    chart.title, chart.y_axis.title, chart.x_axis.title = f"Total by {key.lower()}", "Amount", key
    _attach_chart(ws, chart, last)


def _write_category_sheet(ws: Worksheet, frame: pd.DataFrame, category_col: str,
                          currency_format: str) -> None:
    formats = {"Orders": INT_FORMAT, "Total": currency_format, "Share": PCT_FORMAT}
    last = _write_frame(ws, frame, formats=formats)
    chart = BarChart()
    chart.type = "col"
    chart.title, chart.y_axis.title, chart.x_axis.title = "Total by category", "Amount", category_col
    _attach_chart(ws, chart, last)


def _attach_chart(ws: Worksheet, chart, last_row: int) -> None:
    """Plot column C (Total) against column A labels; anchor right of the table."""
    chart.add_data(Reference(ws, min_col=3, min_row=1, max_row=last_row), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=last_row))
    chart.x_axis.delete = chart.y_axis.delete = False  # openpyxl >= 3.1 hides axes otherwise
    chart.legend = None
    chart.width, chart.height = 20, 10
    ws.add_chart(chart, "F2")
    ws.freeze_panes = "A2"
    _autofit(ws)


def _write_data_sheet(ws: Worksheet, clean: pd.DataFrame, date_col: str, amount_col: str,
                      currency_format: str) -> None:
    last = _write_frame(ws, clean, formats={date_col: DATE_FORMAT, amount_col: currency_format})
    ref = f"A1:{get_column_letter(len(clean.columns))}{last}"
    table = Table(displayName="CleanData", ref=ref, autoFilter=AutoFilter(ref=ref))
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(table)
    ws.freeze_panes = "A2"
    _autofit(ws)


def _write_rejected_sheet(ws: Worksheet, rejected: pd.DataFrame) -> None:
    _write_frame(ws, rejected)
    if rejected.empty:
        ws.cell(2, 1, "No rows were rejected.")
    ws.freeze_panes = "A2"
    _autofit(ws)
