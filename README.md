# csv2xlsx-report

Turn a raw CSV export into a formatted Excel report: a KPI summary, a month x category
pivot with totals, a trend chart, a category chart, the cleaned data as a filterable
table, and a Rejected tab so that no row ever disappears silently.

Built with pandas and openpyxl. Sample output: `examples/sample_report.xlsx`.
[繁體中文說明在下方](#繁體中文)

## What it does

Input: any CSV with a date column, a category column and an amount column.
Output: one `.xlsx` with five tabs.

| Tab | Content |
| --- | --- |
| Summary | KPI block (total, row count, average, highest/lowest month, duplicates removed, rejected rows) and a month x category pivot with a totals row and column. Styled header, frozen panes, fitted column widths. |
| By Month | Monthly order count and total, with a line chart. (`--group-by quarter` or `year` changes the period and tab name.) |
| By Category | Category count, total and share, sorted by total, with a bar chart. |
| Data | The cleaned rows as an Excel Table with autofilter, typed dates and formatted amounts, plus a Month key column. |
| Rejected | Rows that could not be used, with their original text, source line number and a reason. |

Cleaning rules (`csv2xlsx/cleaning.py`, unit tested):

- Dates: ISO (`2024-03-05`, `2024/03/05`, optional time) always work. Slash dates are read
  as MM/DD/YYYY by default or DD/MM/YYYY with `--dayfirst`. A value that is impossible in
  the chosen order is rejected, never guessed, so `12/03/2024` is never silently flipped.
- Amounts: currency symbols, thousands separators and spaces are stripped (`"$1,200.50"`
  becomes 1200.50; `(45.00)` becomes -45.00). Blank or non-numeric values are rejected.
- Exact duplicate rows are dropped once and the count is shown on the Summary tab.
- Blank categories are labelled `Uncategorized`.

## Usage

```sh
python -m csv2xlsx input.csv out.xlsx --date-col Date --category-col Category --amount-col Amount [--dayfirst]
python -m csv2xlsx --help          # also: --group-by {month,quarter,year}, --currency-format, --quiet
```

From Python:

```python
from csv2xlsx import build_report
stats = build_report("orders.csv", "report.xlsx", date_col="Date", category_col="Category",
                     amount_col="Amount", group_by=("month", "category"), currency_format="#,##0.00")
print(stats.rows, stats.rejected, stats.duplicates_dropped, stats.total)
```

Rebuild the bundled example (61 fictional retail orders, including `"$1,200.50"`, a
`12/03/2024` date, an exact duplicate, a blank amount and an impossible `2024-02-30`):

```sh
sh examples/build_sample.sh     # writes examples/sample_report.xlsx
```

## What "reconciled" means

Every number in the workbook is computed from the same cleaned table, and the tests
prove the totals agree: pivot grand total = By Month total = By Category total =
Summary KPI = sum of the Data tab. Rows are conserved as well:

    rows in CSV = rows in Data + rows in Rejected + duplicates removed

Rejected rows are never dropped; they keep their original text and source line so the
client can fix the export and rerun. If a total looks wrong, the answer is always
visible in Data or Rejected, not hidden inside the cleaning code.

## Run tests

```sh
python -m pytest -q       # 19 tests, about one second
```

Requirements: Python 3.10+, pandas 2.x/3.x, openpyxl 3.1+, pytest.

---

# 繁體中文

把原始 CSV 匯出檔轉成排版好的 Excel 報表：KPI 摘要、月份 x 類別的樞紐表（含合計列與合計欄）、
趨勢折線圖、類別長條圖、可篩選的資料表，以及一個「Rejected」分頁，確保沒有任何一列資料被悄悄丟掉。

## 功能

輸入：任何含有日期欄、類別欄、金額欄的 CSV。輸出：一個五分頁的 `.xlsx`。

- **Summary**：KPI 區塊（總額、筆數、平均、最高／最低月份、移除的重複列數、被拒列數）與月份 x 類別樞紐表，含合計列與合計欄、表頭樣式、凍結窗格、自動欄寬。
- **By Month**：每月筆數與總額，附折線圖（`--group-by quarter` 或 `year` 可改為季／年）。
- **By Category**：各類別筆數、總額與占比，依總額排序，附長條圖。
- **Data**：清理後的資料，做成 Excel 表格（自動篩選），日期與金額都是真正的型別並套用格式。
- **Rejected**：無法使用的列，保留原始文字、來源行號與原因。

清理規則（`csv2xlsx/cleaning.py`，有單元測試）：ISO 日期一律可解析；斜線日期預設讀作 MM/DD/YYYY，
加 `--dayfirst` 則讀作 DD/MM/YYYY，在所選順序下不可能的日期會被拒絕而不是猜測；
金額會去掉貨幣符號、千分位與空白（`"$1,200.50"` 變 1200.50），空白或非數字會被拒絕；
完全重複的列只保留一筆並在 Summary 顯示移除數量；空白類別標為 `Uncategorized`。

## 使用方式

```sh
python -m csv2xlsx input.csv out.xlsx --date-col Date --category-col Category --amount-col Amount [--dayfirst]
sh examples/build_sample.sh     # 重建範例報表 examples/sample_report.xlsx
```

## 「對得起來」是什麼意思

工作簿裡每個數字都來自同一張清理後的表，測試會證明：樞紐表總計 = By Month 總計 = By Category 總計
= Summary KPI = Data 分頁金額總和。列數也守恆：CSV 列數 = Data 列數 + Rejected 列數 + 移除的重複列數。
被拒的列不會消失，會保留原文與來源行號，客戶修正匯出檔後即可重跑。

## 執行測試

```sh
python -m pytest -q
```
