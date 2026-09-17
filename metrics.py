"""Compute board metrics, threshold flags and data gaps from the clean table.

Rules (see CLAUDE.md):
- Python does all the math here; Claude only interprets the results later.
- Ratios are decimals (0.97 = 97%). Formatting to % happens only at output.
- Missing inputs are NaN. Any math with NaN gives NaN, so a gap automatically
  spreads to every metric that uses it. Nothing is ever filled in.

Run directly to print everything:  python metrics.py data/northwind.xlsx
"""

import math
import sys
from pathlib import Path

import pandas as pd
import yaml

from clean import clean_workbook

CONFIG_PATH = Path(__file__).parent / "config.yaml"

# Flag statuses.
TRIP = "trip"
PASS = "pass"
MISSING = "cannot evaluate — data missing"


# ---------------------------------------------------------------------------
# Metric definitions. Each takes the clean table (df) and returns one value per quarter.
# ---------------------------------------------------------------------------

def net_new_arr(df):
    """Net new ARR = new + expansion - contraction - churn."""
    return df["new_arr"] + df["expansion_arr"] - df["contraction_arr"] - df["churned_arr"]


def ending_arr(df):
    """Ending ARR = starting_arr + net new ARR."""
    return df["starting_arr"] + net_new_arr(df)


def nrr(df):
    """NRR (annualized) = 1 + 4 * (expansion - contraction - churn) / starting_arr."""
    return 1 + 4 * (df["expansion_arr"] - df["contraction_arr"] - df["churned_arr"]) / df["starting_arr"]


def grr(df):
    """GRR (annualized) = 1 - 4 * (contraction + churn) / starting_arr."""
    return 1 - 4 * (df["contraction_arr"] + df["churned_arr"]) / df["starting_arr"]


def gross_margin(df):
    """Gross margin = gross_profit / revenue."""
    return df["gross_profit"] / df["revenue"]


def growth(series, quarters_back):
    """Growth vs N quarters earlier: 1 = QoQ, 4 = YoY.

    .shift(N) lines each quarter up with the value N rows above it.
    The first N quarters have nothing to compare with, so they come out NaN.
    """
    return series / series.shift(quarters_back) - 1


def fcf_margin(df):
    """FCF margin = -net_burn / revenue (burning cash = negative margin)."""
    return -df["net_burn"] / df["revenue"]


def rule_of_40(df):
    """Rule of 40 = YoY revenue growth + FCF margin (as decimals: 0.40 = 40)."""
    return growth(df["revenue"], 4) + fcf_margin(df)


def burn_multiple(df):
    """Burn multiple = net_burn / net new ARR.

    Edge cases:
    - not burning (net_burn <= 0) -> 0, passes
    - burning while ARR shrank or stood still (net new ARR <= 0) -> infinite, trips
    """
    new = net_new_arr(df)
    result = df["net_burn"] / new
    result = result.mask(df["net_burn"] <= 0, 0.0)                    # .mask(condition, x): use x where condition is true
    result = result.mask((df["net_burn"] > 0) & (new <= 0), math.inf)
    return result


def burn_vs_budget(df):
    """Burn vs budget = net_burn / budget_net_burn - 1 (0.20 = 20% over budget)."""
    return df["net_burn"] / df["budget_net_burn"] - 1


def budget_net_new_arr(df):
    """Budgeted net new ARR = budget_arr this quarter - budget_arr last quarter.

    Net of expansion and churn, same as actual net new ARR. Needs the prior quarter,
    so the first quarter (and the quarter after a blank one) comes out NaN.
    """
    return df["budget_arr"] - df["budget_arr"].shift(1)


def net_new_arr_vs_budget(df):
    """Net new ARR vs budget = net new ARR / (budget_arr - prior quarter's budget_arr) - 1."""
    return net_new_arr(df) / budget_net_new_arr(df) - 1


def arr_vs_budget(df):
    """ARR vs budget = ending ARR / budget_arr - 1 (budget_arr is budgeted ending ARR)."""
    return ending_arr(df) / df["budget_arr"] - 1


def cac_payback_months(df):
    """CAC payback (months) = sm_spend / (new_arr * gross margin) * 12.

    Edge case: new_arr * gross margin <= 0 -> infinite (never pays back), trips.
    """
    gross_profit_on_new_arr = df["new_arr"] * gross_margin(df)
    result = df["sm_spend"] / gross_profit_on_new_arr * 12
    return result.mask(gross_profit_on_new_arr <= 0, math.inf)


def runway_months(df):
    """Runway (months) = ending_cash / (net_burn / 3), at current burn.

    Edge case: not burning (net_burn <= 0) -> infinite, passes.
    """
    result = df["ending_cash"] / (df["net_burn"] / 3)
    return result.mask(df["net_burn"] <= 0, math.inf)


def runway_at_next_budget(actuals, next_budget):
    """Runway (months) at next quarter's budgeted burn = latest cash / (budgeted burn / 3).

    NaN if there's no budget-only row; infinite if the budget has no burn.
    Shown as context on the deck, not flagged.
    """
    if next_budget is None:
        return math.nan
    cash = actuals["ending_cash"].iloc[-1]  # .iloc[-1] = last row = latest quarter
    burn = next_budget["budget_net_burn"]
    if burn <= 0:
        return math.inf
    return cash / (burn / 3)


def compute_metrics(actuals):
    """Build one table: a row per quarter, a column per metric."""
    arr = ending_arr(actuals)
    return pd.DataFrame({
        "ending_arr": arr,
        "net_new_arr": net_new_arr(actuals),
        "nrr": nrr(actuals),
        "grr": grr(actuals),
        "gross_margin": gross_margin(actuals),
        "arr_qoq": growth(arr, 1),
        "arr_yoy": growth(arr, 4),
        "revenue_qoq": growth(actuals["revenue"], 1),
        "revenue_yoy": growth(actuals["revenue"], 4),
        "pipeline": actuals["pipeline"],  # carried over for the combo rule
        "pipeline_qoq": growth(actuals["pipeline"], 1),
        "fcf_margin": fcf_margin(actuals),
        "rule_of_40": rule_of_40(actuals),
        "burn_multiple": burn_multiple(actuals),
        "burn_vs_budget": burn_vs_budget(actuals),
        "net_new_arr_vs_budget": net_new_arr_vs_budget(actuals),
        "arr_vs_budget": arr_vs_budget(actuals),
        "cac_payback_months": cac_payback_months(actuals),
        "runway_months": runway_months(actuals),
    })


# How many quarters back each metric looks. Anything not listed only uses its own quarter (0).
METRIC_LOOKBACK = {
    "arr_qoq": 1, "revenue_qoq": 1, "pipeline_qoq": 1, "net_new_arr_vs_budget": 1,
    "arr_yoy": 4, "revenue_yoy": 4, "rule_of_40": 4,
}


# ---------------------------------------------------------------------------
# Flags
# ---------------------------------------------------------------------------

# (flag name, metric column, config.yaml key, "min" = trips below / "max" = trips above)
FLAG_RULES = [
    ("NRR (annualized)", "nrr", "nrr_min", "min"),
    ("GRR (annualized)", "grr", "grr_min", "min"),
    ("Burn multiple", "burn_multiple", "burn_multiple_max", "max"),
    ("Burn vs budget", "burn_vs_budget", "burn_over_budget_max", "max"),
    ("Runway (months)", "runway_months", "runway_min_months", "min"),
    ("CAC payback (months)", "cac_payback_months", "cac_payback_max_months", "max"),
    ("Net new ARR vs budget", "net_new_arr_vs_budget", "net_new_arr_vs_budget_min", "min"),
    ("Rule of 40", "rule_of_40", "rule_of_40_min", "min"),
]
COMBO_FLAG_NAME = "NRR falling while pipeline rising"


def load_config(path=CONFIG_PATH):
    """Read the thresholds from config.yaml into a dictionary."""
    with open(path) as file:
        return yaml.safe_load(file)


def check_threshold(value, threshold, kind):
    """Return TRIP, PASS or MISSING for one value against one threshold.

    "min": trips below the threshold. "max": trips above it. Exactly at the threshold passes.
    """
    if math.isnan(value):
        return MISSING
    value = round(value, 6)  # removes float noise, e.g. 0.19999999999999996 -> 0.2
    if kind == "min":
        return TRIP if value < threshold else PASS
    return TRIP if value > threshold else PASS


def check_combo(metrics, config, quarter):
    """Combo rule: NRR fell AND pipeline rose at every step in the window ending at `quarter`.

    Window = last `combo_lookback_quarters` quarters, including `quarter`.
    Any missing value, or not enough history for a full window -> MISSING (never PASS).
    """
    size = config["combo_lookback_quarters"]
    end = metrics.index.get_loc(quarter) + 1  # row position just after `quarter`
    if end < size:
        return MISSING
    window = metrics.iloc[end - size:end]
    if window["nrr"].isna().any() or window["pipeline"].isna().any():
        return MISSING

    # .diff() = this quarter minus the previous one; the first row has no previous, so skip it.
    nrr_steps = window["nrr"].round(6).diff().iloc[1:]
    pipeline_steps = window["pipeline"].diff().iloc[1:]
    nrr_falling = (nrr_steps < 0).all()
    pipeline_rising = (pipeline_steps > 0).all()
    return TRIP if nrr_falling and pipeline_rising else PASS


def evaluate_flags(metrics, config, quarter=None):
    """Check every flag for one quarter (default: the latest). Returns a list of dicts."""
    if quarter is None:
        quarter = metrics.index[-1]

    flags = []
    for name, column, config_key, kind in FLAG_RULES:
        value = metrics.loc[quarter, column]
        threshold = config[config_key]
        flags.append({
            "flag": name, "metric": column, "quarter": quarter,
            "value": value, "threshold": threshold,
            "status": check_threshold(value, threshold, kind),
        })

    if config["nrr_falling_pipeline_rising_flag"]:
        flags.append({
            "flag": COMBO_FLAG_NAME, "metric": None, "quarter": quarter,
            "value": None, "threshold": None,
            "status": check_combo(metrics, config, quarter),
        })
    return flags


# ---------------------------------------------------------------------------
# Data gaps
# ---------------------------------------------------------------------------

def data_gaps(actuals, metrics, flags):
    """List every metric and flag that can't be shown because input data is missing.

    Returns {metric or flag: [quarters]}.
    A NaN counts as a gap if the quarter has enough history for that metric, or if
    the quarter's own inputs are incomplete. So YoY in Q3 2024 is "no prior year",
    not a gap, but YoY in a blank quarter is a gap.
    """
    position = pd.Series(range(len(metrics)), index=metrics.index)  # 0, 1, 2, ... per quarter
    incomplete = actuals.isna().any(axis=1)                          # quarter has any blank input

    gaps = {}
    for column in metrics.columns:
        has_history = position >= METRIC_LOOKBACK.get(column, 0)
        is_gap = metrics[column].isna() & (has_history | incomplete)
        if is_gap.any():
            gaps[column] = list(metrics.index[is_gap])

    for flag in flags:
        if flag["status"] == MISSING:
            gaps[f"flag: {flag['flag']}"] = [flag["quarter"]]
    return gaps


# ---------------------------------------------------------------------------
# Printing (formatting to % happens only here)
# ---------------------------------------------------------------------------

DOLLAR_COLUMNS = {"ending_arr", "net_new_arr", "pipeline", "net_burn", "ending_cash"}
MONTH_COLUMNS = {"cac_payback_months", "runway_months"}


def format_value(column, value):
    """Format one value for display: $K with commas, months, multiple, or %."""
    if value is None or pd.isna(value):
        return "n/a"
    if math.isinf(value):
        return "∞"
    if column in DOLLAR_COLUMNS:
        return f"{value:,.0f}"
    if column in MONTH_COLUMNS:
        return f"{value:.1f} mo"
    if column == "burn_multiple":
        return f"{value:.2f}x"
    return f"{value:.1%}"


def print_report(actuals, next_budget, config):
    """Print the metrics table, runway at budget, flags and data gaps."""
    metrics = compute_metrics(actuals)
    flags = evaluate_flags(metrics, config)

    formatted = pd.DataFrame({c: [format_value(c, v) for v in metrics[c]] for c in metrics.columns},
                             index=metrics.index)
    print("=== Metrics ($K, ratios as %) ===")
    print(formatted.T.to_string())

    runway_budget = runway_at_next_budget(actuals, next_budget)
    print(f"\nRunway at next quarter's budgeted burn: {format_value('runway_months', runway_budget)}")

    print(f"\n=== Flags ({flags[0]['quarter']}) ===")
    for flag in flags:
        detail = ""
        if flag["metric"]:
            detail = (f"{format_value(flag['metric'], flag['value'])} "
                      f"(threshold {format_value(flag['metric'], flag['threshold'])})")
        print(f"{flag['status'].upper():<32} {flag['flag']:<36} {detail}")

    print("\n=== Data gaps ===")
    for name, quarters in data_gaps(actuals, metrics, flags).items():
        print(f"{name:<24} {', '.join(quarters)}")


if __name__ == "__main__":
    workbook_path = sys.argv[1] if len(sys.argv) > 1 else "data/northwind.xlsx"
    actuals, next_budget = clean_workbook(workbook_path)
    print_report(actuals, next_budget, load_config())
