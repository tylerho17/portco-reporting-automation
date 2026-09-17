"""Unit tests for excel_output.py: the labels written when a value is missing or not meaningful.

Expected values are worked out by hand (shown in comments), not copied from the code.
Run from the project folder:  pytest
"""

import math

import pandas as pd
import pytest

from clean import BUDGET_COLUMNS, STANDARD_COLUMNS
from excel_output import build_workbook

NAN = math.nan

# Thresholds written out here, so tuning config.yaml never breaks these tests.
TEST_CONFIG = {
    "nrr_min": 1.00, "grr_min": 0.85, "burn_multiple_max": 2.0, "burn_over_budget_max": 0.15,
    "runway_min_months": 12, "cac_payback_max_months": 24, "net_new_arr_vs_budget_min": -0.20,
    "rule_of_40_min": 0.40, "nrr_falling_pipeline_rising_flag": True, "combo_lookback_quarters": 3,
    "combo_min_nrr_drop": 0.01,
}


def two_quarters(**changes):
    """Two complete quarters, every input 100 (ending cash 1200), with `changes` applied to Q4 2024."""
    actuals = pd.DataFrame(100.0, index=["Q3 2024", "Q4 2024"], columns=STANDARD_COLUMNS)
    actuals["ending_cash"] = 1200.0
    for column, value in changes.items():
        actuals.loc["Q4 2024", column] = value
    return actuals


def budget_row(budget_net_burn):
    """The budget-only forecast row, as clean.py returns it."""
    return pd.Series({"budget_new_arr": 100.0, "budget_arr": 300.0, "budget_net_burn": budget_net_burn},
                     index=BUDGET_COLUMNS, name="Q1 2025 (Budget)")


def runway_context_cell(actuals, next_budget):
    """The value in the 'Runway at next quarter's budgeted burn' row (last row of Flags, column C)."""
    sheet = build_workbook(actuals, next_budget, TEST_CONFIG)["Flags"]
    assert sheet.cell(row=sheet.max_row, column=1).value.startswith("Runway at next quarter's budgeted burn")
    return sheet.cell(row=sheet.max_row, column=3).value


@pytest.mark.parametrize("actuals, next_budget, expected, why", [
    (two_quarters(), budget_row(300.0), 12.0, "1200 / (300 / 3) = 12.0 months"),
    (two_quarters(), None, "n/a (no budget row)", "the workbook has no budget-only row"),
    (two_quarters(), budget_row(0.0), "∞ (budget not burning)", "budgeted burn 0: never runs out"),
    (two_quarters(), budget_row(NAN), "data missing", "budget row exists but its burn cell is blank"),
    (two_quarters(ending_cash=NAN), budget_row(300.0), "data missing", "latest quarter's cash is blank"),
    (two_quarters(ending_cash=NAN), budget_row(0.0), "data missing",
     "cash blank while the budget isn't burning: a blank wins over the ∞ edge case"),
])
def test_runway_context_label(actuals, next_budget, expected, why):
    # A blank input must say "data missing", never "no budget row" (CLAUDE.md: missing never looks
    # like not meaningful). Bug found in the Task 7 review: both blank cases said "n/a (no budget row)".
    assert runway_context_cell(actuals, next_budget) == expected, why


# ---------------------------------------------------------------------------
# Each reason has its own words and fill (decisions A, C, J)
# ---------------------------------------------------------------------------

GRAY = "D9D9D9"


def fill(cell):
    """The cell's fill as 6-digit hex, or None."""
    return None if cell.fill.fill_type is None else cell.fill.fgColor.rgb[-6:]


def metrics_cell(sheet, quarter, label):
    """The Metrics sheet cell for one quarter (row) and one metric label (header)."""
    column = [cell.value for cell in sheet[1]].index(label) + 1
    row = next(r for r in range(2, sheet.max_row + 1) if sheet.cell(row=r, column=1).value == quarter)
    return sheet.cell(row=row, column=column)


def reason_workbook():
    """Q4 2024: S&M blank (missing input) and budgeted burn 0 (not meaningful). Q3 2024 has no prior quarter."""
    return build_workbook(two_quarters(sm_spend=NAN, budget_net_burn=0.0), budget_row(300.0), TEST_CONFIG)


def test_metrics_sheet_words_and_fill_for_each_reason():
    sheet = reason_workbook()["Metrics"]
    no_prior = metrics_cell(sheet, "Q3 2024", "ARR growth QoQ")
    assert (no_prior.value, fill(no_prior)) == ("n/a (no prior period)", None)
    missing = metrics_cell(sheet, "Q4 2024", "CAC payback")
    assert (missing.value, fill(missing)) == ("data missing", GRAY)  # only missing data is gray
    not_meaningful = metrics_cell(sheet, "Q4 2024", "Net burn vs budget")
    # every input is 100, so net burn is 100; the budgeted burn was set to 0
    assert (not_meaningful.value, fill(not_meaningful)) == ("n/m: net burn 100 vs budget 0 ($K)", None)


def test_flags_sheet_status_says_why_a_flag_cannot_be_evaluated():
    sheet = reason_workbook()["Flags"]
    statuses = {sheet.cell(row=r, column=1).value: sheet.cell(row=r, column=6).value for r in range(2, 11)}
    assert statuses["CAC payback"] == "Cannot evaluate — missing input"
    assert statuses["Net burn vs budget"] == "Cannot evaluate — not meaningful"
    assert statuses["NRR falling while pipeline rising"] == "Cannot evaluate — no prior period"  # 2 quarters < 3


def test_data_gaps_sheet_lists_only_missing_input():
    rows = [(a.value, b.value) for a, b in reason_workbook()["Data gaps"].iter_rows(min_row=2, max_col=2)]
    assert rows == [("CAC payback", "Q4 2024"), ("Flag: CAC payback", "Q4 2024")]
