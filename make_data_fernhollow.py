"""Generate a fake, deliberately messy KPI workbook for "Fernhollow Systems" (the distressed company).

Output: data/fernhollow.xlsx (all money figures in $K).

This file holds only Fernhollow's numbers and mess settings. The checking and
writing code is shared with the other companies in make_data_common.py.

The story: a company in trouble. In Q2 2026, 7 of 9 flags trip.
- ARR stalls, then shrinks (net new ARR turns negative in Q1 2026), so the
  burn multiple is infinite ("ARR shrank").
- Churn and contraction climb every quarter: NRR ~78%, GRR ~75%.
- Burn keeps rising, ~20% over budget; runway is 6 months.
- New sales dry up, so CAC payback is over 10 years; the budget still assumes growth.
- Pipeline FALLS too, so the "NRR falling while pipeline rising" combo passes:
  this is a sales problem AND a retention problem, not retention alone.
- Q2 2025 is blank. That is exactly 4 quarters before Q2 2026, so the Rule of 40
  (which needs YoY revenue growth) returns "cannot evaluate - data missing" in the
  latest quarter. Northwind's blank quarter never reaches its latest-quarter flags.

Mess that differs from Northwind:
- a title line and an empty row above the header row
- a different column order (each budget column sits next to its actual)
"""

from pathlib import Path

from make_data_common import save_workbook

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Where the workbook is saved, relative to this script's folder.
OUTPUT_PATH = Path(__file__).parent / "data" / "fernhollow.xlsx"

# The 8 actual quarters, oldest first (same quarters as Northwind).
QUARTERS = [
    "Q3 2024", "Q4 2024", "Q1 2025", "Q2 2025",
    "Q3 2025", "Q4 2025", "Q1 2026", "Q2 2026",
]

# This quarter's row is written with its label only - every value left blank.
BLANK_QUARTER = "Q2 2025"

# Label for the extra forecast row after the actuals. "Bud" is one of clean.py's budget words.
BUDGET_ONLY_LABEL = "Q3 2026 - Bud"

# Line written above the table.
TITLE = "Fernhollow Systems - Quarterly KPIs ($K)"

# ---------------------------------------------------------------------------
# The clean, correct data ($K unless noted). One list per column,
# one value per quarter, in the same order as QUARTERS.
# Q2 2025 has real values here so the chains tie out, but they never reach the file.
# ---------------------------------------------------------------------------

TRUE_DATA = {
    "starting_arr":    [6000, 6490, 6920, 7270, 7520, 7670, 7700, 7590],
    "new_arr":         [520, 500, 460, 420, 370, 310, 240, 180],
    "expansion_arr":   [140, 130, 120, 100, 90, 80, 70, 60],
    "contraction_arr": [50, 60, 70, 80, 90, 110, 130, 150],
    "churned_arr":     [120, 140, 160, 190, 220, 250, 290, 330],
    "revenue":         [1480, 1590, 1690, 1760, 1800, 1830, 1820, 1770],
    "gross_profit":    [950, 1000, 1050, 1070, 1080, 1080, 1060, 1010],
    "net_burn":        [1100, 1180, 1250, 1320, 1400, 1480, 1560, 1650],
    "ending_cash":     [13140, 11960, 10710, 9390, 7990, 6510, 4950, 3300],
    "sm_spend":        [1000, 1020, 1040, 1060, 1080, 1100, 1120, 1150],
    "new_customers":   [18, 17, 15, 14, 12, 10, 8, 6],          # count, not $K
    "headcount":       [95, 98, 100, 101, 99, 96, 90, 84],      # count, not $K
    "pipeline":        [4200, 4100, 3900, 3700, 3400, 3100, 2800, 2500],
    "budget_new_arr":  [550, 560, 580, 600, 620, 640, 660, 680],
    "budget_arr":      [6480, 6950, 7430, 7910, 8390, 8870, 9350, 9830],
    "budget_net_burn": [1100, 1120, 1150, 1180, 1220, 1260, 1320, 1380],
}

# Budget for the quarter after the latest actual quarter (the budget-only row).
NEXT_QUARTER_BUDGET = {
    "budget_new_arr": 700,
    "budget_arr": 10310,
    "budget_net_burn": 1300,
}

# ---------------------------------------------------------------------------
# The mess
# ---------------------------------------------------------------------------

# Clean column name -> the inconsistent header written in the workbook.
# The order here is the column order in the workbook.
HEADER_NAMES = {
    "starting_arr":    "Beginning ARR",
    "new_arr":         "New ARR ($K)",
    "budget_new_arr":  "New ARR - Budget",
    "expansion_arr":   "EXPANSION ARR",
    "contraction_arr": "Contraction ARR",
    "churned_arr":     "churned arr",
    "budget_arr":      "Budget_ARR",
    "revenue":         "Revenue",
    "gross_profit":    "Gross profit",
    "net_burn":        "Net Burn",
    "budget_net_burn": "Net Burn (Bud.)",
    "ending_cash":     "Cash - End of Qtr",
    "sm_spend":        "S&M Spend",
    "new_customers":   "New Logos",
    "headcount":       "Headcount (FTE)",
    "pipeline":        "Qualified Pipeline",
}

# (column, quarter) -> text style: "M" -> "$13.14M", "K" -> "4200K", "comma" -> "1,830".
TEXT_CELLS = {
    ("ending_cash", "Q3 2024"): "M",
    ("ending_cash", "Q4 2024"): "M",
    ("ending_cash", "Q1 2025"): "M",
    ("ending_cash", "Q3 2025"): "M",
    ("ending_cash", "Q4 2025"): "M",
    ("ending_cash", "Q1 2026"): "M",
    ("ending_cash", "Q2 2026"): "M",
    ("pipeline", "Q3 2024"): "K",
    ("pipeline", "Q4 2024"): "K",
    ("pipeline", "Q1 2025"): "K",
    ("contraction_arr", "Q1 2026"): "K",
    ("contraction_arr", "Q2 2026"): "K",
    ("revenue", "Q4 2025"): "comma",
    ("revenue", "Q1 2026"): "comma",
    ("budget_arr", "Q2 2026"): "M",
    ("starting_arr", "Q3 2024"): "comma",
}

# Junk notes tab: cell address -> content.
NOTES = {
    "A1": "Fernhollow Systems - notes",
    "A3": "Q2 2025 close never finished (finance system migration)",
    "A4": "Two largest customers sent non-renewal notices",
    "A5": "Bridge financing conversation - see email",
    "A7": "runway?",
    "B7": 3300,
    "C7": 1650,
    "D7": "~6 mo",
    "A9": "Budget not re-forecast since January",
}


def main():
    save_workbook({
        "output_path": OUTPUT_PATH,
        "quarters": QUARTERS,
        "blank_quarter": BLANK_QUARTER,
        "budget_only_label": BUDGET_ONLY_LABEL,
        "true_data": TRUE_DATA,
        "next_quarter_budget": NEXT_QUARTER_BUDGET,
        "header_names": HEADER_NAMES,
        "text_cells": TEXT_CELLS,
        "notes": NOTES,
        "title": TITLE,
    })


if __name__ == "__main__":
    main()
