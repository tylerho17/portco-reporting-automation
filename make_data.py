"""Generate a fake, deliberately messy KPI workbook for "Northwind Software".

Output: data/northwind.xlsx (all money figures in $K).

This file holds only Northwind's numbers and mess settings. The checking and
writing code is shared with the other companies in make_data_common.py.

The numbers are chosen to tell the Northwind story in CLAUDE.md:
strong ARR growth, NRR sliding 108% -> 97%, rising pipeline,
burn ~20% over budget and ~11 months of runway.
"""

from pathlib import Path

from make_data_common import save_workbook

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Where the workbook is saved, relative to this script's folder.
OUTPUT_PATH = Path(__file__).parent / "data" / "northwind.xlsx"

# The 8 actual quarters, oldest first.
QUARTERS = [
    "Q3 2024", "Q4 2024", "Q1 2025", "Q2 2025",
    "Q3 2025", "Q4 2025", "Q1 2026", "Q2 2026",
]

# This quarter's row is written with its label only - every value left blank.
BLANK_QUARTER = "Q1 2025"

# Label for the extra forecast row after the actuals (budget columns only).
BUDGET_ONLY_LABEL = "Q3 2026 (Budget)"

# ---------------------------------------------------------------------------
# The clean, correct data ($K unless noted). One list per column,
# one value per quarter, in the same order as QUARTERS.
# Q1 2025 has real values here so the chains tie out, but they never reach the file.
# ---------------------------------------------------------------------------

TRUE_DATA = {
    "starting_arr":    [12000, 13580, 15320, 17190, 19230, 21460, 23790, 25810],
    "new_arr":         [1400, 1500, 1600, 1700, 1800, 1900, 1900, 1850],
    "expansion_arr":   [360, 440, 540, 640, 770, 860, 710, 580],
    "contraction_arr": [60, 70, 90, 100, 110, 140, 200, 260],
    "churned_arr":     [120, 130, 180, 200, 230, 290, 390, 510],
    "revenue":         [3200, 3610, 4060, 4550, 5090, 5660, 6200, 6660],
    "gross_profit":    [2300, 2600, 2960, 3320, 3770, 4190, 4650, 5000],
    "net_burn":        [2400, 2600, 2800, 3000, 3200, 3400, 3650, 3900],
    "ending_cash":     [36850, 34250, 31450, 28450, 25250, 21850, 18200, 14300],
    "sm_spend":        [1700, 1800, 1900, 2000, 2100, 2250, 2350, 2400],
    "new_customers":   [38, 41, 44, 47, 50, 53, 52, 50],            # count, not $K
    "headcount":       [110, 118, 126, 134, 142, 150, 158, 165],    # count, not $K
    "pipeline":        [6000, 6800, 7500, 8300, 9200, 10100, 11200, 12500],
    "budget_new_arr":  [1400, 1500, 1650, 1750, 1850, 1950, 2000, 2000],
    "budget_arr":      [13500, 15150, 16900, 18750, 20700, 22700, 24750, 26800],
    "budget_net_burn": [2400, 2520, 2670, 2780, 2910, 2980, 3120, 3250],
}

# Budget for the quarter after the latest actual quarter (the budget-only row).
NEXT_QUARTER_BUDGET = {
    "budget_new_arr": 2050,
    "budget_arr": 28900,
    "budget_net_burn": 3300,
}

# ---------------------------------------------------------------------------
# The mess
# ---------------------------------------------------------------------------

# Clean column name -> the inconsistent header written in the workbook.
HEADER_NAMES = {
    "starting_arr":    "Beginning ARR",
    "new_arr":         "New ARR ($K)",
    "expansion_arr":   "expansion arr",
    "contraction_arr": "Contraction_ARR",
    "churned_arr":     " Churned ARR ",      # stray spaces on purpose
    "revenue":         "Revenue",
    "gross_profit":    "GROSS PROFIT",
    "net_burn":        "Net Burn",
    "ending_cash":     "Cash - End of Qtr",
    "sm_spend":        "S&M Spend",
    "new_customers":   "New Logos",
    "headcount":       "Headcount (FTE)",
    "pipeline":        "Qualified Pipeline",
    "budget_new_arr":  "New ARR - Budget",
    "budget_arr":      "Budget ARR",
    "budget_net_burn": "Net Burn (Bud.)",
}

# (column, quarter) -> text style. These cells are written as text instead of numbers.
#   "M"     -> "$14.3M"  (millions with a dollar sign)
#   "K"     -> "120K"    (thousands with a K)
#   "comma" -> "5,090"   (number with a thousands comma)
TEXT_CELLS = {
    ("pipeline", "Q3 2024"): "M",
    ("pipeline", "Q4 2024"): "M",
    ("pipeline", "Q2 2025"): "M",
    ("pipeline", "Q3 2025"): "M",
    ("pipeline", "Q4 2025"): "M",
    ("pipeline", "Q1 2026"): "M",
    ("pipeline", "Q2 2026"): "M",
    ("ending_cash", "Q4 2025"): "M",
    ("ending_cash", "Q1 2026"): "M",
    ("ending_cash", "Q2 2026"): "M",
    ("churned_arr", "Q3 2024"): "K",
    ("churned_arr", "Q4 2024"): "K",
    ("contraction_arr", "Q2 2026"): "K",
    ("revenue", "Q3 2025"): "comma",
}


# Junk notes tab: cell address -> content.
NOTES = {
    "A1": "Northwind Software - misc notes",
    "A3": "Q1 2025 actuals still with finance, do not use",
    "A4": "Pipeline from CRM export, pulled first week of quarter",
    "A5": "Board meeting moved - check calendar",
    "A7": "scratch",
    "B7": 3900,
    "C7": 3250,
    "D7": "20% over??",
    "A9": "TODO: update headcount plan",
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
    })


if __name__ == "__main__":
    main()
