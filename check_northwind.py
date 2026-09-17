"""Automated proof for build step 2, using Northwind's answer key from make_data.py.

Checks three things:
1. Cleaning is lossless: clean.py recovers TRUE_DATA exactly (Q1 2025 all blank).
2. Key Q2 2026 metrics match the hand formulas, and flags match the story.
3. Data gaps land exactly where the CLAUDE.md rules say.

Run: python check_northwind.py  -> prints "All checks passed" or stops at the first failure.
"""

import math

from clean import clean_workbook
from make_data import BLANK_QUARTER, NEXT_QUARTER_BUDGET, OUTPUT_PATH, QUARTERS, TRUE_DATA
from metrics import (PASS, TRIP, compute_metrics, data_gaps, evaluate_flags,
                     load_config, runway_at_next_budget)

LATEST = "Q2 2026"

# Q2 2026 metrics written as the same hand formulas you'd type into Excel.
EXPECTED_LATEST = {
    "nrr": 1 + 4 * (580 - 260 - 510) / 25810,
    "grr": 1 - 4 * (260 + 510) / 25810,
    "burn_multiple": 3900 / (1850 + 580 - 260 - 510),
    "burn_vs_budget": 3900 / 3250 - 1,
    "runway_months": 14300 / (3900 / 3),
    "rule_of_40": 6660 / 4550 - 1 + (-3900 / 6660),
    "cac_payback_months": 2400 / (1850 * 5000 / 6660) * 12,
    "net_new_arr_vs_budget": 1660 / (26800 - 24750) - 1,
    "arr_yoy": 27470 / 19230 - 1,
}
EXPECTED_RUNWAY_AT_BUDGET = 14300 / (3300 / 3)

# The Northwind story: 6 flags trip, 3 pass.
EXPECTED_FLAGS = {
    "NRR (annualized)": TRIP,
    "GRR (annualized)": PASS,
    "Burn multiple": TRIP,
    "Burn vs budget": TRIP,
    "Runway (months)": TRIP,
    "CAC payback (months)": PASS,
    "Net new ARR vs budget": PASS,
    "Rule of 40": TRIP,
    "NRR falling while pipeline rising": TRIP,
}

# Metrics that look back 1 quarter (net new ARR vs budget uses last quarter's budget_arr).
QOQ_METRICS = {"arr_qoq", "revenue_qoq", "pipeline_qoq", "net_new_arr_vs_budget"}
YOY_METRICS = {"arr_yoy", "revenue_yoy", "rule_of_40"}


def check_cleaning(actuals, next_budget):
    """Every cleaned value equals the answer key; the blank quarter is entirely NaN."""
    assert list(actuals.index) == QUARTERS, f"Quarters wrong: {list(actuals.index)}"
    assert set(actuals.columns) == set(TRUE_DATA), "Columns don't match TRUE_DATA"
    for column, values in TRUE_DATA.items():
        for quarter, expected in zip(QUARTERS, values):
            got = actuals.loc[quarter, column]
            if quarter == BLANK_QUARTER:
                assert math.isnan(got), f"{quarter} {column} should be blank, got {got}"
            else:
                assert got == expected, f"{quarter} {column}: expected {expected}, got {got}"
    assert next_budget.to_dict() == NEXT_QUARTER_BUDGET, f"Budget row wrong: {next_budget.to_dict()}"


def check_latest_metrics(metrics, runway_budget):
    """Q2 2026 metrics match the hand formulas (to 9 significant digits)."""
    for column, expected in EXPECTED_LATEST.items():
        got = metrics.loc[LATEST, column]
        assert math.isclose(got, expected, rel_tol=1e-9), f"{column}: expected {expected}, got {got}"
    assert math.isclose(runway_budget, EXPECTED_RUNWAY_AT_BUDGET, rel_tol=1e-9), \
        f"Runway at budget: expected {EXPECTED_RUNWAY_AT_BUDGET}, got {runway_budget}"


def check_flags(flags):
    """Each flag's status matches the story."""
    statuses = {flag["flag"]: flag["status"] for flag in flags}
    assert statuses == EXPECTED_FLAGS, f"Flags wrong: {statuses}"


def check_gaps(gaps, metrics):
    """QoQ gaps: blank quarter + next. YoY gaps: blank quarter + 4 later. Everything else: blank quarter."""
    for column in metrics.columns:
        if column in QOQ_METRICS:
            expected = ["Q1 2025", "Q2 2025"]
        elif column in YOY_METRICS:
            expected = ["Q1 2025", "Q1 2026"]
        else:
            expected = ["Q1 2025"]
        assert gaps.get(column) == expected, f"{column} gaps: expected {expected}, got {gaps.get(column)}"
    assert set(gaps) == set(metrics.columns), f"Unexpected gaps: {set(gaps) - set(metrics.columns)}"


def main():
    actuals, next_budget = clean_workbook(OUTPUT_PATH)
    metrics = compute_metrics(actuals)
    flags = evaluate_flags(metrics, load_config())

    check_cleaning(actuals, next_budget)
    print("✓ Cleaning is lossless (all values match TRUE_DATA, Q1 2025 blank, budget row read)")
    check_latest_metrics(metrics, runway_at_next_budget(actuals, next_budget))
    print("✓ Q2 2026 metrics match the hand formulas")
    check_flags(flags)
    print("✓ Flags match the story (6 trip, 3 pass)")
    check_gaps(data_gaps(actuals, metrics, flags), metrics)
    print("✓ Data gaps are exactly where the rules say")
    print("All checks passed")


if __name__ == "__main__":
    main()
