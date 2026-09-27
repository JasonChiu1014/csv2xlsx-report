"""Command-line entry point: ``python -m csv2xlsx input.csv out.xlsx --date-col ...``."""

from __future__ import annotations

import argparse
import logging
import sys
from typing import Sequence

from .report import PERIODS, build_report


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="python -m csv2xlsx",
        description="Turn a raw CSV export into a formatted Excel report with summary tabs and charts.",
    )
    parser.add_argument("csv_path", help="input CSV file")
    parser.add_argument("out_xlsx", help="output .xlsx file")
    parser.add_argument("--date-col", required=True, help="name of the date column")
    parser.add_argument("--category-col", required=True, help="name of the category column")
    parser.add_argument("--amount-col", required=True, help="name of the amount column")
    parser.add_argument("--dayfirst", action="store_true",
                        help="read slash dates as DD/MM/YYYY (default: MM/DD/YYYY)")
    parser.add_argument("--group-by", choices=sorted(PERIODS), default="month",
                        help="row period for the pivot and trend chart (default: month)")
    parser.add_argument("--currency-format", default="#,##0.00",
                        help="Excel number format for amounts (default: #,##0.00)")
    parser.add_argument("-q", "--quiet", action="store_true", help="only print warnings and errors")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    """Run the CLI; return a process exit code."""
    args = _parser().parse_args(argv)
    logging.basicConfig(level=logging.WARNING if args.quiet else logging.INFO,
                        format="%(levelname)s %(name)s: %(message)s")
    try:
        stats = build_report(
            args.csv_path, args.out_xlsx,
            date_col=args.date_col, category_col=args.category_col, amount_col=args.amount_col,
            group_by=(args.group_by, "category"), currency_format=args.currency_format,
            dayfirst=args.dayfirst,
        )
    except (FileNotFoundError, KeyError, ValueError) as exc:
        logging.getLogger("csv2xlsx").error("%s", exc)
        return 1
    print(f"Wrote {stats.out_path}: {stats.rows} rows, total {stats.total:,.2f}, "
          f"{stats.rejected} rejected, {stats.duplicates_dropped} duplicates removed")
    return 0


if __name__ == "__main__":
    sys.exit(main())
