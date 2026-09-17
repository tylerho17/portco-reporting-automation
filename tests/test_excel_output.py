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
])
def test_runway_context_label(actuals, next_budget, expected, why):
    # A blank input must say "data missing", never "no budget row" (CLAUDE.md: missing never looks
    # like not meaningful). Bug found in the Task 7 review: both blank cases said "n/a (no budget row)".
    assert runway_context_cell(actuals, next_budget) == expected, why
