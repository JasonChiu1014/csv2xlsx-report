#!/usr/bin/env sh
# Rebuild examples/sample_report.xlsx from examples/sample_orders.csv.
# --dayfirst: the one slash date in the sample (12/03/2024) is 12 March, not 3 December.
# Uses $PYTHON if set, otherwise python3 (activate your virtualenv first).
set -eu
cd "$(dirname "$0")/.."
"${PYTHON:-python3}" -m csv2xlsx examples/sample_orders.csv examples/sample_report.xlsx \
    --date-col Date --category-col Category --amount-col Amount --dayfirst
