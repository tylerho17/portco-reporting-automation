"""Compute board metrics, threshold flags and data gaps from the clean table.

Rules (see CLAUDE.md):
- Python does all the math here; Claude only interprets the results later.
- Ratios are decimals (0.97 = 97%). Formatting to % happens only at output.
- Missing inputs are NaN. Any math with NaN gives NaN, so a gap automatically
  spreads to every metric that uses it. Nothing is ever filled in.
- When a metric has no number, there are exactly three possible reasons, and they must
  never look alike: missing input, no prior period, not meaningful.

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
CANNOT_EVALUATE = "cannot evaluate"

# Why a metric (or flag) has no usable number.
MISSING_INPUT = "missing input"        # an input cell the metric uses is blank: a data gap
NO_PRIOR_PERIOD = "no prior period"    # it needs an earlier quarter the workbook doesn't have (e.g. YoY in year 1)
NOT_MEANINGFUL = "not meaningful"      # every input is there, but the math is undefined (e.g. 0 / 0)

# How each reason is shown in every output (printout, Excel, Claude's payload).
REASON_DISPLAY = {
    MISSING_INPUT: "data missing",
    NO_PRIOR_PERIOD: "n/a (no prior period)",
    NOT_MEANINGFUL: "n/m (not meaningful)",
}

# ---------------------------------------------------------------------------
# Labels: the one set of display names used by every output
# ---------------------------------------------------------------------------

# Raw inputs shown next to the metrics.
INPUT_LABELS = {
    "net_burn": "Net burn ($K)",
    "ending_cash": "Ending cash ($K)",
}

# Metric column -> display name. Flag names are taken from here too, so they always match.
METRIC_LABELS = {
    "ending_arr": "Ending ARR ($K)",
    "net_new_arr": "Net new ARR ($K)",
    "nrr": "NRR (annualized)",
    "grr": "GRR (annualized)",
    "gross_margin": "Gross margin",
    "arr_qoq": "ARR growth QoQ",
    "arr_yoy": "ARR growth YoY",
    "revenue_qoq": "Revenue growth QoQ",
    "revenue_yoy": "Revenue growth YoY",
    "pipeline": "Pipeline ($K)",
    "pipeline_qoq": "Pipeline growth QoQ",
    "fcf_margin": "FCF margin",
    "rule_of_40": "Rule of 40",
    "burn_multiple": "Burn multiple",
    "burn_vs_budget": "Net burn vs budget",
    "net_new_arr_vs_budget": "Net new ARR vs budget",
    "arr_vs_budget": "Ending ARR vs budget",
    "cac_payback_months": "CAC payback",
    "runway_months": "Runway at current burn",
}


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


def all_present(df, columns):
    """True for each quarter where every one of `columns` has a value (none blank)."""
    return df[columns].notna().all(axis=1)


def burn_multiple(df):
    """Burn multiple = net_burn / net new ARR.

    Edge cases (only when every input is there - a blank must stay blank):
    - not burning (net_burn <= 0) -> 0, passes
    - burning while ARR shrank or stood still (net new ARR <= 0) -> infinite, trips
    """
    new = net_new_arr(df)
    present = all_present(df, ["net_burn", "new_arr", "expansion_arr", "contraction_arr", "churned_arr"])
    result = df["net_burn"] / new
    result = result.mask(present & (df["net_burn"] <= 0), 0.0)        # .mask(condition, x): use x where true
    result = result.mask(present & (df["net_burn"] > 0) & (new <= 0), math.inf)
    return result


def burn_vs_budget(df):
    """Burn vs budget = net_burn / budget_net_burn - 1 (0.20 = 20% over budget).

    Not meaningful (NaN) when budgeted burn is 0 or less: "% over a budget of no burn" means nothing.
    """
    result = df["net_burn"] / df["budget_net_burn"] - 1
    return result.mask(df["budget_net_burn"] <= 0)  # .mask with no value -> NaN


def budget_net_new_arr(df):
    """Budgeted net new ARR = budget_arr this quarter - budget_arr last quarter.

    Net of expansion and churn, same as actual net new ARR. Needs the prior quarter,
    so the first quarter (and the quarter after a blank one) comes out NaN.
    """
    return df["budget_arr"] - df["budget_arr"].shift(1)


def net_new_arr_vs_budget(df):
    """Net new ARR vs budget = net new ARR / (budget_arr - prior quarter's budget_arr) - 1.

    Not meaningful (NaN) when budgeted net new ARR is 0 or less: a % of a flat or shrinking plan means nothing.
    """
    budget = budget_net_new_arr(df)
    result = net_new_arr(df) / budget - 1
    return result.mask(budget <= 0)


def arr_vs_budget(df):
    """ARR vs budget = ending ARR / budget_arr - 1 (budget_arr is budgeted ending ARR)."""
    return ending_arr(df) / df["budget_arr"] - 1


def cac_payback_months(df):
    """CAC payback (months) = sm_spend / (new_arr * gross margin) * 12.

    Edge case (only when every input is there): new_arr * gross margin <= 0 -> infinite (never pays back), trips.
    """
    gross_profit_on_new_arr = df["new_arr"] * gross_margin(df)
    present = all_present(df, ["sm_spend", "new_arr", "gross_profit", "revenue"])
    result = df["sm_spend"] / gross_profit_on_new_arr * 12
    return result.mask(present & (gross_profit_on_new_arr <= 0), math.inf)


def runway_months(df):
    """Runway (months) = ending_cash / (net_burn / 3), at current burn.

    Edge case (only when every input is there): not burning (net_burn <= 0) -> infinite, passes.
    """
    present = all_present(df, ["ending_cash", "net_burn"])
    result = df["ending_cash"] / (df["net_burn"] / 3)
    return result.mask(present & (df["net_burn"] <= 0), math.inf)


def runway_at_next_budget(actuals, next_budget):
    """Runway (months) at next quarter's budgeted burn = latest cash / (budgeted burn / 3).

    NaN if there's no budget-only row or an input is blank; infinite if the budget has no burn.
    Shown as context on the deck, not flagged.
    """
    if next_budget is None:
        return math.nan
    cash = actuals["ending_cash"].iloc[-1]  # .iloc[-1] = last row = latest quarter
    burn = next_budget["budget_net_burn"]
    if pd.isna(cash) or pd.isna(burn):
        return math.nan  # a blank stays blank: the "not burning" edge case needs every input
    if burn <= 0:
        return math.inf
    return cash / (burn / 3)


# ---------------------------------------------------------------------------
# Which inputs each metric uses, and why a value can be missing
# ---------------------------------------------------------------------------

ARR_FLOWS = ["new_arr", "expansion_arr", "contraction_arr", "churned_arr"]


def now(*columns):
    """Inputs from the same quarter: [(column, 0), ...]."""
    return [(column, 0) for column in columns]


def back(quarters, *columns):
    """Inputs from `quarters` quarters earlier: [(column, quarters), ...]."""
    return [(column, quarters) for column in columns]


# Metric -> every (input column, quarters back) it uses. data_gaps and the reasons come from here.
METRIC_INPUTS = {
    "ending_arr": now("starting_arr", *ARR_FLOWS),
    "net_new_arr": now(*ARR_FLOWS),
    "nrr": now("starting_arr", "expansion_arr", "contraction_arr", "churned_arr"),
    "grr": now("starting_arr", "contraction_arr", "churned_arr"),
    "gross_margin": now("gross_profit", "revenue"),
    "arr_qoq": now("starting_arr", *ARR_FLOWS) + back(1, "starting_arr", *ARR_FLOWS),
    "arr_yoy": now("starting_arr", *ARR_FLOWS) + back(4, "starting_arr", *ARR_FLOWS),
    "revenue_qoq": now("revenue") + back(1, "revenue"),
    "revenue_yoy": now("revenue") + back(4, "revenue"),
    "pipeline": now("pipeline"),
    "pipeline_qoq": now("pipeline") + back(1, "pipeline"),
    "fcf_margin": now("net_burn", "revenue"),
    "rule_of_40": now("revenue", "net_burn") + back(4, "revenue"),
    "burn_multiple": now("net_burn", *ARR_FLOWS),
    "burn_vs_budget": now("net_burn", "budget_net_burn"),
    "net_new_arr_vs_budget": now(*ARR_FLOWS, "budget_arr") + back(1, "budget_arr"),
    "arr_vs_budget": now("starting_arr", *ARR_FLOWS, "budget_arr"),
    "cac_payback_months": now("sm_spend", "new_arr", "gross_profit", "revenue"),
    "runway_months": now("ending_cash", "net_burn"),
}

# How many quarters back each metric looks (the furthest input). 0 = only its own quarter.
METRIC_LOOKBACK = {metric: max(quarters for _, quarters in inputs) for metric, inputs in METRIC_INPUTS.items()}

# The CLAUDE.md edge cases that are allowed to be infinite. Any other infinity is "not meaningful".
INFINITE_ALLOWED = {"burn_multiple", "runway_months", "cac_payback_months"}


def input_reason(actuals, metric, position):
    """Why the metric can't be computed from its inputs at row `position`: MISSING_INPUT, NO_PRIOR_PERIOD or None.

    A blank input wins over a missing earlier quarter: a blank quarter's own YoY is still a data gap.
    """
    inputs = METRIC_INPUTS[metric]
    for column, quarters_back in inputs:
        row = position - quarters_back
        if row >= 0 and pd.isna(actuals[column].iloc[row]):
            return MISSING_INPUT
    if any(position - quarters_back < 0 for _, quarters_back in inputs):
        return NO_PRIOR_PERIOD
    return None


def compute_metrics(actuals):
    """Build one table: a row per quarter, a column per metric.

    Two safety rules on top of the formulas:
    - wherever an input is missing (or an earlier quarter doesn't exist), the value is NaN;
    - an infinity that isn't a CLAUDE.md edge case (e.g. growth from a zero base) becomes NaN.
    """
    arr = ending_arr(actuals)
    metrics = pd.DataFrame({
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
    for metric in metrics.columns:
        if metric not in INFINITE_ALLOWED:
            metrics[metric] = metrics[metric].mask(metrics[metric].abs() == math.inf)
        for position, quarter in enumerate(metrics.index):
            if input_reason(actuals, metric, position) is not None:
                metrics.loc[quarter, metric] = math.nan
    return metrics


def metric_reasons(actuals, metrics):
    """A table shaped like `metrics`: for each value, why there is no number (a reason), or None."""
    reasons = {}
    for metric in metrics.columns:
        column = []
        for position, quarter in enumerate(metrics.index):
            reason = input_reason(actuals, metric, position)
            if reason is None and pd.isna(metrics.loc[quarter, metric]):
                reason = NOT_MEANINGFUL  # every input is there, yet no number: the math is undefined
            column.append(reason)
        reasons[metric] = column
    return pd.DataFrame(reasons, index=metrics.index, dtype=object)  # dtype=object keeps None as None


def not_meaningful_text(actuals, metric, quarter):
    """What to show instead of a not-meaningful value: the $K figures behind it where they help."""
    row = actuals.loc[quarter]
    if metric == "burn_vs_budget":
        return f"n/m: net burn {row['net_burn']:,.0f} vs budget {row['budget_net_burn']:,.0f} ($K)"
    if metric == "net_new_arr_vs_budget":
        actual = net_new_arr(actuals).loc[quarter]
        budget = budget_net_new_arr(actuals).loc[quarter]
        return f"n/m: net new ARR {actual:,.0f} vs budget {budget:,.0f} ($K)"
    return REASON_DISPLAY[NOT_MEANINGFUL]


def reason_text(actuals, metric, quarter, reason):
    """The words shown for a value that has no number."""
    if reason == NOT_MEANINGFUL:
        return not_meaningful_text(actuals, metric, quarter)
    return REASON_DISPLAY[reason]


# ---------------------------------------------------------------------------
# Flags
# ---------------------------------------------------------------------------

# (metric column, config.yaml key, "min" = trips below / "max" = trips above)
FLAG_THRESHOLDS = [
    ("nrr", "nrr_min", "min"),
    ("grr", "grr_min", "min"),
    ("burn_multiple", "burn_multiple_max", "max"),
    ("burn_vs_budget", "burn_over_budget_max", "max"),
    ("runway_months", "runway_min_months", "min"),
    ("cac_payback_months", "cac_payback_max_months", "max"),
    ("net_new_arr_vs_budget", "net_new_arr_vs_budget_min", "min"),
    ("rule_of_40", "rule_of_40_min", "min"),
]

# (flag name, metric column, config key, kind). The flag name IS the metric's label.
FLAG_RULES = [(METRIC_LABELS[column], column, key, kind) for column, key, kind in FLAG_THRESHOLDS]
COMBO_FLAG_NAME = "NRR falling while pipeline rising"


def validate_config(config):
    """Stop with a clear message if the combo settings can't work."""
    lookback = config.get("combo_lookback_quarters")
    if not isinstance(lookback, int) or isinstance(lookback, bool) or lookback < 2:
        # With 1 quarter there are no steps to compare, and "every step fell" would be true of nothing.
        raise ValueError(f"config.yaml: combo_lookback_quarters must be a whole number of at least 2 "
                         f"(got {lookback!r}) - the combo rule needs at least one quarter-to-quarter step")
    drop = config.get("combo_min_nrr_drop")
    if not isinstance(drop, (int, float)) or isinstance(drop, bool) or drop < 0:
        raise ValueError(f"config.yaml: combo_min_nrr_drop must be a number of 0 or more "
                         f"(got {drop!r}), e.g. 0.01 for 1 point")


def load_config(path=CONFIG_PATH):
    """Read the thresholds from config.yaml into a dictionary, and check the combo settings."""
    with open(path) as file:
        config = yaml.safe_load(file)
    validate_config(config)
    return config


def check_threshold(value, threshold, kind):
    """Return TRIP, PASS or CANNOT_EVALUATE for one value against one threshold.

    "min": trips below the threshold. "max": trips above it. Exactly at the threshold passes.
    """
    if math.isnan(value):
        return CANNOT_EVALUATE
    value = round(value, 6)  # removes float noise, e.g. 0.19999999999999996 -> 0.2
    if kind == "min":
        return TRIP if value < threshold else PASS
    return TRIP if value > threshold else PASS


def check_combo(metrics, reasons, config, quarter):
    """Combo rule: NRR fell by at least combo_min_nrr_drop AND pipeline rose, at every step. Returns (status, reason).

    Window = last `combo_lookback_quarters` quarters, including `quarter`.
    Reasons, in the same order as a metric's: missing input (a blank wins, even with too little
    history), then no prior period (not enough history), then the value's own reason.
    Never PASS on missing data.
    """
    validate_config(config)
    size = config["combo_lookback_quarters"]
    end = metrics.index.get_loc(quarter) + 1  # row position just after `quarter`
    start = max(end - size, 0)                # a short history still gets its blanks checked
    window = metrics.iloc[start:end]
    window_reasons = [r for r in reasons.iloc[start:end][["nrr", "pipeline"]].to_numpy().ravel()
                      if isinstance(r, str)]
    if MISSING_INPUT in window_reasons:
        return CANNOT_EVALUATE, MISSING_INPUT
    if end < size:
        return CANNOT_EVALUATE, NO_PRIOR_PERIOD
    if window_reasons:
        return CANNOT_EVALUATE, window_reasons[0]

    # .diff() = this quarter minus the previous one; the first row has no previous, so skip it.
    # Round after subtracting, so 1.07 - 1.08 = -0.010000000000000009 counts as exactly 1 point.
    nrr_steps = window["nrr"].diff().iloc[1:].round(6)
    pipeline_steps = window["pipeline"].diff().iloc[1:]
    nrr_falling = (nrr_steps <= -config["combo_min_nrr_drop"]).all()
    pipeline_rising = (pipeline_steps > 0).all()
    return (TRIP if nrr_falling and pipeline_rising else PASS), None


def evaluate_flags(metrics, reasons, config, quarter=None):
    """Check every flag for one quarter (default: the latest). Returns a list of dicts.

    Each dict: flag, metric, quarter, value, threshold, status, and reason (None unless it can't be evaluated).
    """
    validate_config(config)
    if quarter is None:
        quarter = metrics.index[-1]

    flags = []
    for name, column, config_key, kind in FLAG_RULES:
        value = metrics.loc[quarter, column]
        status = check_threshold(value, config[config_key], kind)
        flags.append({
            "flag": name, "metric": column, "quarter": quarter,
            "value": value, "threshold": config[config_key], "status": status,
            "reason": reasons.loc[quarter, column] if status == CANNOT_EVALUATE else None,
        })

    if config["nrr_falling_pipeline_rising_flag"]:
        status, reason = check_combo(metrics, reasons, config, quarter)
        flags.append({
            "flag": COMBO_FLAG_NAME, "metric": None, "quarter": quarter,
            "value": None, "threshold": None, "status": status, "reason": reason,
        })
    return flags


def flag_status_text(flag):
    """'trip', 'pass', or 'cannot evaluate — missing input' (the reason is part of the status)."""
    if flag["status"] == CANNOT_EVALUATE:
        return f"{CANNOT_EVALUATE} — {flag['reason']}"
    return flag["status"]


# ---------------------------------------------------------------------------
# Data gaps
# ---------------------------------------------------------------------------

def data_gaps(actuals, metrics, flags):
    """List every metric and flag that can't be shown because an input it uses is blank.

    Returns {metric or "flag: <name>": [quarters]}. Only MISSING INPUT counts: "no prior period"
    (e.g. YoY in year 1) and "not meaningful" (e.g. 0 / 0) are not data gaps.
    """
    reasons = metric_reasons(actuals, metrics)
    gaps = {}
    for column in metrics.columns:
        quarters = [quarter for quarter in metrics.index if reasons.loc[quarter, column] == MISSING_INPUT]
        if quarters:
            gaps[column] = quarters

    for flag in flags:
        if flag["status"] == CANNOT_EVALUATE and flag.get("reason") == MISSING_INPUT:
            gaps[f"flag: {flag['flag']}"] = [flag["quarter"]]
    return gaps


# ---------------------------------------------------------------------------
# Printing (formatting to % happens only here and in the other outputs)
# ---------------------------------------------------------------------------

DOLLAR_COLUMNS = {"ending_arr", "net_new_arr", "pipeline", "net_burn", "ending_cash"}
MONTH_COLUMNS = {"cac_payback_months", "runway_months"}

NO_BUDGET_ROW = "n/a (no budget row)"
BUDGET_NOT_BURNING = "∞ (budget not burning)"


def format_value(column, value):
    """Format one real number for display: $K with commas, months, multiple, or %. (No NaN: see reason_text.)"""
    if math.isinf(value):
        return "∞"
    if column in DOLLAR_COLUMNS:
        return f"{value:,.0f}"
    if column in MONTH_COLUMNS:
        return f"{value:.1f} mo"
    if column == "burn_multiple":
        return f"{value:.2f}x"
    return f"{value:.1%}"


def display_value(actuals, metrics, reasons, metric, quarter):
    """One metric value as text: the formatted number, or the words for why there isn't one."""
    reason = reasons.loc[quarter, metric]
    if isinstance(reason, str):
        return reason_text(actuals, metric, quarter, reason)
    return format_value(metric, metrics.loc[quarter, metric])


def runway_context_label(runway, has_budget_row):
    """Runway at next quarter's budgeted burn: None if it's a normal number, else the words to show.

    NaN has two causes: no budget row at all, or a blank input (latest cash or budgeted burn).
    """
    if not has_budget_row:
        return NO_BUDGET_ROW
    if math.isnan(runway):
        return REASON_DISPLAY[MISSING_INPUT]
    if math.isinf(runway):
        return BUDGET_NOT_BURNING
    return None


def print_report(actuals, next_budget, config):
    """Print the metrics table, runway at budget, flags and data gaps."""
    metrics = compute_metrics(actuals)
    reasons = metric_reasons(actuals, metrics)
    flags = evaluate_flags(metrics, reasons, config)

    formatted = pd.DataFrame(
        {METRIC_LABELS[m]: [display_value(actuals, metrics, reasons, m, q) for q in metrics.index]
         for m in metrics.columns},
        index=metrics.index)
    print("=== Metrics ($K, ratios as %) ===")
    print(formatted.T.to_string())

    runway = runway_at_next_budget(actuals, next_budget)
    runway_text = runway_context_label(runway, next_budget is not None) or format_value("runway_months", runway)
    print(f"\nRunway at next quarter's budgeted burn: {runway_text}")

    print(f"\n=== Flags ({flags[0]['quarter']}) ===")
    for flag in flags:
        detail = ""
        if flag["metric"]:
            value = display_value(actuals, metrics, reasons, flag["metric"], flag["quarter"])
            detail = f"{value} (threshold {format_value(flag['metric'], flag['threshold'])})"
        print(f"{flag_status_text(flag).upper():<34} {flag['flag']:<36} {detail}")

    print("\n=== Data gaps ===")
    for name, quarters in data_gaps(actuals, metrics, flags).items():
        label = name if name.startswith("flag: ") else METRIC_LABELS[name]
        print(f"{label:<36} {', '.join(quarters)}")


if __name__ == "__main__":
    workbook_path = sys.argv[1] if len(sys.argv) > 1 else "data/northwind.xlsx"
    actuals, next_budget = clean_workbook(workbook_path)
    print_report(actuals, next_budget, load_config())
