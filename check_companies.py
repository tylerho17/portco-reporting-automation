"""Automated proof for build step 5: clean.py and metrics.py work for all three companies.

Each company's answer key is its make_data script (TRUE_DATA, blank quarter, budget row).
For every company this checks:
1. Cleaning is lossless: clean.py recovers TRUE_DATA exactly (blank quarter all NaN).
2. Key Q2 2026 metrics match hand formulas typed from the answer key.
3. Every flag status matches the company's story.
4. Data gaps land exactly where the CLAUDE.md rules say.
Plus, for the healthy company: no flag trips in ANY quarter.

Northwind's expected values are reused from check_northwind.py so they live in one place.

Run: python check_companies.py  -> prints "All checks passed" or stops at the first failure.
"""

import math

import check_northwind
import make_data as northwind
import make_data_alderpeak as alderpeak
import make_data_fernhollow as fernhollow
from check_northwind import QOQ_METRICS, YOY_METRICS
from clean import clean_workbook
from metrics import (FLAG_RULES, COMBO_FLAG_NAME, MISSING, PASS, TRIP, compute_metrics,
                     data_gaps, evaluate_flags, load_config, runway_at_next_budget)

LATEST = "Q2 2026"
ALL_FLAG_NAMES = [rule[0] for rule in FLAG_RULES] + [COMBO_FLAG_NAME]

# ---------------------------------------------------------------------------
# Alderpeak (healthy): Q2 2026 hand formulas, same style as check_northwind.py.
# Q2 2026 = last column of TRUE_DATA; Q2 2025 (a year earlier) = 4th column.
# ---------------------------------------------------------------------------

ALDERPEAK_LATEST = {
    "nrr": 1 + 4 * (580 - 70 - 100) / 16820,
    "grr": 1 - 4 * (70 + 100) / 16820,
    "burn_multiple": 200 / (1220 + 580 - 70 - 100),
    "burn_vs_budget": 200 / 220 - 1,
    "runway_months": 7200 / (200 / 3),
    "rule_of_40": 4230 / 2870 - 1 + (-200 / 4230),
    "cac_payback_months": 1400 / (1220 * 3300 / 4230) * 12,
    "net_new_arr_vs_budget": 1630 / (17990 - 16450) - 1,
    "arr_yoy": 18450 / 12510 - 1,
    "arr_vs_budget": 18450 / 17990 - 1,
}
ALDERPEAK_RUNWAY_AT_BUDGET = 7200 / (150 / 3)

# Healthy: every flag passes.
ALDERPEAK_FLAGS = {name: PASS for name in ALL_FLAG_NAMES}

# ---------------------------------------------------------------------------
# Fernhollow (distressed)
# ---------------------------------------------------------------------------

FERNHOLLOW_LATEST = {
    "nrr": 1 + 4 * (60 - 150 - 330) / 7590,
    "grr": 1 - 4 * (150 + 330) / 7590,
    "burn_multiple": math.inf,             # net new ARR = 180 + 60 - 150 - 330 = -240 while burning
    "burn_vs_budget": 1650 / 1380 - 1,
    "runway_months": 3300 / (1650 / 3),
    "rule_of_40": math.nan,                # needs Q2 2025 revenue, which is blank
    "cac_payback_months": 1150 / (180 * 1010 / 1770) * 12,
    "net_new_arr_vs_budget": -240 / (9830 - 9350) - 1,
    "arr_yoy": math.nan,                   # needs Q2 2025 ARR, which is blank
    "arr_vs_budget": 7350 / 9830 - 1,
}
FERNHOLLOW_RUNWAY_AT_BUDGET = 3300 / (1300 / 3)

# Distressed: 7 trip. Rule of 40 can't be evaluated; the combo passes because pipeline is falling.
FERNHOLLOW_FLAGS = {
    "NRR (annualized)": TRIP,
    "GRR (annualized)": TRIP,
    "Burn multiple": TRIP,
    "Burn vs budget": TRIP,
    "Runway (months)": TRIP,
    "CAC payback (months)": TRIP,
    "Net new ARR vs budget": TRIP,
    "Rule of 40": MISSING,
    "NRR falling while pipeline rising": PASS,
}

# ---------------------------------------------------------------------------
# One entry per company. "blank_gaps" = (blank quarter, the quarter after, 4 quarters after),
# or None if the company has no blank quarter. "flag_gaps" = flags that can't be evaluated.
# ---------------------------------------------------------------------------

COMPANIES = [
    {
        "name": "Northwind",
        "answer_key": northwind,
        "expected_latest": check_northwind.EXPECTED_LATEST,
        "expected_runway_at_budget": check_northwind.EXPECTED_RUNWAY_AT_BUDGET,
        "expected_flags": check_northwind.EXPECTED_FLAGS,
        "blank_gaps": ("Q1 2025", "Q2 2025", "Q1 2026"),
        "flag_gaps": {},
    },
    {
        "name": "Alderpeak",
        "answer_key": alderpeak,
        "expected_latest": ALDERPEAK_LATEST,
        "expected_runway_at_budget": ALDERPEAK_RUNWAY_AT_BUDGET,
        "expected_flags": ALDERPEAK_FLAGS,
        "blank_gaps": None,
        "flag_gaps": {},
    },
    {
        "name": "Fernhollow",
        "answer_key": fernhollow,
        "expected_latest": FERNHOLLOW_LATEST,
        "expected_runway_at_budget": FERNHOLLOW_RUNWAY_AT_BUDGET,
        "expected_flags": FERNHOLLOW_FLAGS,
        "blank_gaps": ("Q2 2025", "Q3 2025", "Q2 2026"),
        "flag_gaps": {"flag: Rule of 40": [LATEST]},
    },
]


# ---------------------------------------------------------------------------
# Checks
# ---------------------------------------------------------------------------

def same_number(got, expected):
    """True if two metric values match: both NaN, or equal to 9 significant digits (inf == inf)."""
    if math.isnan(expected):
        return math.isnan(got)
    return math.isclose(got, expected, rel_tol=1e-9)


def check_cleaning(answer_key, actuals, next_budget):
    """Every cleaned value equals the answer key; the blank quarter (if any) is entirely NaN."""
    assert list(actuals.index) == answer_key.QUARTERS, f"Quarters wrong: {list(actuals.index)}"
    assert set(actuals.columns) == set(answer_key.TRUE_DATA), "Columns don't match TRUE_DATA"
    for column, values in answer_key.TRUE_DATA.items():
        for quarter, expected in zip(answer_key.QUARTERS, values):
            got = actuals.loc[quarter, column]
            if quarter == answer_key.BLANK_QUARTER:
                assert math.isnan(got), f"{quarter} {column} should be blank, got {got}"
            else:
                assert got == expected, f"{quarter} {column}: expected {expected}, got {got}"
    assert next_budget.to_dict() == answer_key.NEXT_QUARTER_BUDGET, \
        f"Budget row wrong: {next_budget.to_dict()}"


def check_latest_metrics(company, metrics, runway_budget):
    """Q2 2026 metrics match the hand formulas."""
    for column, expected in company["expected_latest"].items():
        got = metrics.loc[LATEST, column]
        assert same_number(got, expected), f"{column}: expected {expected}, got {got}"
    expected = company["expected_runway_at_budget"]
    assert same_number(runway_budget, expected), f"Runway at budget: expected {expected}, got {runway_budget}"


def check_flags(company, flags):
    """Each flag's status matches the company's story."""
    statuses = {flag["flag"]: flag["status"] for flag in flags}
    assert statuses == company["expected_flags"], f"Flags wrong: {statuses}"


def expected_gaps(company, metric_columns):
    """Build the gap list the CLAUDE.md rules predict for this company.

    Every metric misses the blank quarter; QoQ metrics also miss the quarter after;
    YoY metrics also miss the quarter 4 later. Then add any flags that can't be evaluated.
    """
    gaps = {}
    if company["blank_gaps"] is not None:
        blank, next_quarter, year_later = company["blank_gaps"]
        for column in metric_columns:
            if column in QOQ_METRICS:
                gaps[column] = [blank, next_quarter]
            elif column in YOY_METRICS:
                gaps[column] = [blank, year_later]
            else:
                gaps[column] = [blank]
    gaps.update(company["flag_gaps"])
    return gaps


def check_gaps(company, gaps, metric_columns):
    """Data gaps are exactly the predicted ones - no more, no fewer."""
    expected = expected_gaps(company, metric_columns)
    assert gaps == expected, f"Gaps wrong.\nExpected: {expected}\nGot:      {gaps}"


def check_never_trips(metrics, config):
    """No flag trips in any quarter (used for the healthy company)."""
    for quarter in metrics.index:
        tripped = [f["flag"] for f in evaluate_flags(metrics, config, quarter) if f["status"] == TRIP]
        assert not tripped, f"{quarter}: healthy company tripped {tripped}"


def count_statuses(flags):
    """Short summary like '7 trip, 1 pass, 1 cannot evaluate'."""
    counts = [sum(f["status"] == status for f in flags) for status in (TRIP, PASS, MISSING)]
    return f"{counts[0]} trip, {counts[1]} pass, {counts[2]} cannot evaluate"


def check_company(company, config):
    """Run every check for one company and print one line per check."""
    name, answer_key = company["name"], company["answer_key"]
    actuals, next_budget = clean_workbook(answer_key.OUTPUT_PATH)
    metrics = compute_metrics(actuals)
    flags = evaluate_flags(metrics, config)
    gaps = data_gaps(actuals, metrics, flags)

    check_cleaning(answer_key, actuals, next_budget)
    print(f"✓ {name}: cleaning is lossless (blank quarter: {answer_key.BLANK_QUARTER or 'none'})")
    check_latest_metrics(company, metrics, runway_at_next_budget(actuals, next_budget))
    print(f"✓ {name}: {LATEST} metrics match the hand formulas")
    check_flags(company, flags)
    print(f"✓ {name}: flags match the story ({count_statuses(flags)})")
    check_gaps(company, gaps, metrics.columns)
    print(f"✓ {name}: data gaps are exactly where the rules say ({len(gaps)} metrics/flags affected)")


def main():
    config = load_config()
    for company in COMPANIES:
        check_company(company, config)

    # The healthy company should be clean in every quarter, not just the latest.
    actuals, _ = clean_workbook(alderpeak.OUTPUT_PATH)
    check_never_trips(compute_metrics(actuals), config)
    print("✓ Alderpeak: no flag trips in any quarter")
    print("All checks passed")


if __name__ == "__main__":
    main()
