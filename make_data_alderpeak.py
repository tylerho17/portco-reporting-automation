"""Generate a fake, deliberately messy KPI workbook for "Alderpeak Software" (the healthy company).

Output: data/alderpeak.xlsx (all money figures in $K).

This file holds only Alderpeak's numbers and mess settings. The checking and
writing code is shared with the other companies in make_data_common.py.

The story: a well-run company that should trip NO flags, in any quarter.
- ARR grows ~47% a year; NRR holds around 110% and GRR around 95%.
- Burn shrinks every quarter and stays under budget, so runway is long.
- Rule of 40 stays above 40% once a year of history exists.
- NRR moves up and down (never falls at every step) while pipeline rises,
  so the "NRR falling while pipeline rising" combo passes.

Mess that differs from Northwind (so cleaning isn't tuned to one file):
- different header spellings and capitalization
- the junk Notes tab comes BEFORE the KPI tab
- no blank quarter (a clean company should show zero data gaps)
"""

from pathlib import Path

from make_data_common import save_workbook

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Where the workbook is saved, relative to this script's folder.
OUTPUT_PATH = Path(__file__).parent / "data" / "alderpeak.xlsx"

# The 8 actual quarters, oldest first (same quarters as Northwind).
QUARTERS = [
    "Q3 2024", "Q4 2024", "Q1 2025", "Q2 2025",
    "Q3 2025", "Q4 2025", "Q1 2026", "Q2 2026",
]

# No blank quarter for the healthy company.
BLANK_QUARTER = None

# Label for the extra forecast row after the actuals. "Plan" is one of clean.py's budget words.
BUDGET_ONLY_LABEL = "Q3 2026 Plan"

# ---------------------------------------------------------------------------
# The clean, correct data ($K unless noted). One list per column,
# one value per quarter, in the same order as QUARTERS.
# ---------------------------------------------------------------------------

TRUE_DATA = {
    "starting_arr":    [8000, 9000, 10090, 11250, 12510, 13850, 15290, 16820],
    "new_arr":         [800, 860, 920, 980, 1040, 1100, 1160, 1220],
    "expansion_arr":   [300, 340, 360, 410, 440, 490, 520, 580],
    "contraction_arr": [40, 40, 50, 50, 60, 60, 60, 70],
    "churned_arr":     [60, 70, 70, 80, 80, 90, 90, 100],
    "revenue":         [2060, 2310, 2580, 2870, 3180, 3510, 3850, 4230],
    "gross_profit":    [1610, 1800, 2010, 2240, 2480, 2740, 3000, 3300],
    "net_burn":        [700, 620, 540, 460, 380, 310, 250, 200],
    "ending_cash":     [9960, 9340, 8800, 8340, 7960, 7650, 7400, 7200],
    "sm_spend":        [900, 980, 1050, 1120, 1190, 1260, 1330, 1400],
    "new_customers":   [22, 24, 25, 27, 28, 30, 31, 33],       # count, not $K
    "headcount":       [64, 68, 72, 76, 80, 84, 88, 92],       # count, not $K
    "pipeline":        [3400, 3600, 3850, 4100, 4400, 4700, 5000, 5350],
    "budget_new_arr":  [780, 850, 900, 960, 1020, 1080, 1140, 1200],
    "budget_arr":      [8950, 9980, 11100, 12320, 13620, 14990, 16450, 17990],
    "budget_net_burn": [740, 660, 580, 500, 420, 340, 270, 220],
}

# Budget for the quarter after the latest actual quarter (the budget-only row).
NEXT_QUARTER_BUDGET = {
    "budget_new_arr": 1260,
    "budget_arr": 19650,
    "budget_net_burn": 150,
}

# ---------------------------------------------------------------------------
# The mess
# ---------------------------------------------------------------------------

# Clean column name -> the inconsistent header written in the workbook.
# Every one of these normalizes to a name clean.py already knows.
HEADER_NAMES = {
    "starting_arr":    "Starting ARR",
    "new_arr":         "NEW ARR",
    "expansion_arr":   "Expansion ARR",
    "contraction_arr": "contraction-arr",
    "churned_arr":     "Churned_ARR",
    "revenue":         "revenue ",           # trailing space on purpose
    "gross_profit":    "Gross Profit",
    "net_burn":        "NET BURN",
    "ending_cash":     "Ending Cash",
    "sm_spend":        "S & M spend",
    "new_customers":   "new customers",
    "headcount":       "HEADCOUNT",
    "pipeline":        "Pipeline",
    "budget_new_arr":  "Budget - New ARR",
    "budget_arr":      "BUDGET ARR",
    "budget_net_burn": "Budget Net Burn",
}

# (column, quarter) -> text style: "M" -> "$3.85M", "K" -> "90K", "comma" -> "3,300".
TEXT_CELLS = {
    ("pipeline", "Q3 2024"): "M",
    ("pipeline", "Q4 2024"): "M",
    ("pipeline", "Q1 2025"): "M",
    ("pipeline", "Q2 2025"): "M",
    ("ending_cash", "Q3 2024"): "M",
    ("ending_cash", "Q4 2024"): "M",
    ("ending_cash", "Q1 2025"): "M",
    ("starting_arr", "Q2 2026"): "M",
    ("churned_arr", "Q4 2025"): "K",
    ("churned_arr", "Q1 2026"): "K",
    ("sm_spend", "Q1 2025"): "comma",
    ("revenue", "Q3 2025"): "comma",
    ("gross_profit", "Q2 2026"): "comma",
}

# Junk notes tab: cell address -> content.
NOTES = {
    "A1": "Alderpeak Software - board prep scratchpad",
    "A3": "Q2 2026 numbers final per CFO",
    "A4": "H2 plan assumes 2 more AEs",
    "A6": "check",
    "B6": 4230,
    "C6": "rev ok?",
    "A8": "ignore this tab",
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
        "notes_first": True,
    })


if __name__ == "__main__":
    main()
