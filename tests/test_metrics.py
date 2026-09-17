"""Unit tests for metrics.py: every metric, its edge cases, flags and data gaps.

Expected values are worked out by hand from the CLAUDE.md definitions (shown in comments),
not copied from the code. Run from the project folder:  pytest
"""

import math

import numpy as np
import pandas as pd
import pytest

from clean import STANDARD_COLUMNS
from metrics import (COMBO_FLAG_NAME, FLAG_RULES, METRIC_LOOKBACK, MISSING, PASS, TRIP,
                     arr_vs_budget, budget_net_new_arr, burn_multiple, burn_vs_budget,
                     cac_payback_months, check_combo, check_threshold, compute_metrics, data_gaps,
                     ending_arr, evaluate_flags, fcf_margin, gross_margin, growth, grr, load_config,
                     net_new_arr, net_new_arr_vs_budget, nrr, rule_of_40, runway_at_next_budget,
                     runway_months)

NAN = math.nan
INF = math.inf

# Thresholds written out here, so tuning config.yaml never breaks these tests.
TEST_CONFIG = {
    "nrr_min": 1.00, "grr_min": 0.85, "burn_multiple_max": 2.0, "burn_over_budget_max": 0.15,
    "runway_min_months": 12, "cac_payback_max_months": 24, "net_new_arr_vs_budget_min": -0.20,
    "rule_of_40_min": 0.40, "nrr_falling_pipeline_rising_flag": True, "combo_lookback_quarters": 3,
}

EIGHT_QUARTERS = ["Q3 2024", "Q4 2024", "Q1 2025", "Q2 2025",
                  "Q3 2025", "Q4 2025", "Q1 2026", "Q2 2026"]


def table(**columns):
    """Small table with only the columns a test needs, one row per quarter (Q3 2024 onward)."""
    size = len(next(iter(columns.values())))
    return pd.DataFrame(columns, index=EIGHT_QUARTERS[:size], dtype=float)


def values(series):
    """A Series as a plain list, for comparing with pytest.approx (NaN allowed)."""
    return pytest.approx(series.tolist(), nan_ok=True)


# ---------------------------------------------------------------------------
# ARR and retention
# ---------------------------------------------------------------------------

def test_net_new_arr_and_ending_arr():
    df = table(starting_arr=[10000, 10000], new_arr=[1200, 100], expansion_arr=[300, 50],
               contraction_arr=[100, 200], churned_arr=[400, 350])
    # 1200 + 300 - 100 - 400 = 1000 ;  100 + 50 - 200 - 350 = -400 (ARR shrank)
    assert values(net_new_arr(df)) == [1000, -400]
    assert values(ending_arr(df)) == [11000, 9600]


def test_nrr_annualized():
    df = table(starting_arr=[10000, 10000], expansion_arr=[500, 100],
               contraction_arr=[100, 200], churned_arr=[150, 300])
    # 1 + 4 * (500 - 100 - 150) / 10000 = 1.10 ;  1 + 4 * (100 - 200 - 300) / 10000 = 0.84
    assert values(nrr(df)) == [1.10, 0.84]


def test_grr_annualized():
    df = table(starting_arr=[10000], contraction_arr=[100], churned_arr=[150])
    # 1 - 4 * (100 + 150) / 10000 = 0.90
    assert values(grr(df)) == [0.90]


def test_gross_margin():
    df = table(gross_profit=[750], revenue=[1000])
    assert values(gross_margin(df)) == [0.75]


def test_missing_input_gives_nan_not_zero():
    # A blank input must spread as NaN, never be treated as 0.
    df = table(starting_arr=[10000], new_arr=[NAN], expansion_arr=[500],
               contraction_arr=[100], churned_arr=[150])
    assert math.isnan(net_new_arr(df).iloc[0])
    assert math.isnan(ending_arr(df).iloc[0])


# ---------------------------------------------------------------------------
# Growth and Rule of 40
# ---------------------------------------------------------------------------

def test_growth_qoq_first_quarter_has_no_prior():
    series = pd.Series([100.0, 110.0, 99.0])
    # 110 / 100 - 1 = 0.10 ;  99 / 110 - 1 = -0.10
    assert values(growth(series, 1)) == [NAN, 0.10, -0.10]


def test_growth_yoy_needs_four_quarters_back():
    series = pd.Series([100.0, 105.0, 110.0, 115.0, 150.0, 160.0])
    # 150 / 100 - 1 = 0.50 ;  160 / 105 - 1
    assert values(growth(series, 4)) == [NAN, NAN, NAN, NAN, 0.50, 160 / 105 - 1]


def test_growth_blank_quarter_spreads_to_next_quarter():
    series = pd.Series([100.0, NAN, 120.0])
    assert values(growth(series, 1)) == [NAN, NAN, NAN]


def test_fcf_margin_sign():
    df = table(net_burn=[200, -100], revenue=[1000, 1000])
    # burning 200 -> -20% ;  generating 100 of cash -> +10%
    assert values(fcf_margin(df)) == [-0.20, 0.10]


def test_rule_of_40():
    df = table(revenue=[1000, 1050, 1100, 1150, 1300], net_burn=[100, 100, 100, 100, 130])
    # Q3 2025: YoY growth 1300 / 1000 - 1 = 0.30, FCF margin -130 / 1300 = -0.10 -> 0.20
    assert values(rule_of_40(df)) == [NAN, NAN, NAN, NAN, 0.20]


# ---------------------------------------------------------------------------
# Burn
# ---------------------------------------------------------------------------

def burn_table(net_burn, new_arr):
    """Table for burn multiple: net new ARR = new_arr (expansion, contraction, churn are 0)."""
    zeros = [0] * len(net_burn)
    return table(net_burn=net_burn, new_arr=new_arr, expansion_arr=zeros,
                 contraction_arr=zeros, churned_arr=zeros)


def test_burn_multiple_normal():
    # 600 / 400 = 1.5
    assert values(burn_multiple(burn_table([600], [400]))) == [1.5]


@pytest.mark.parametrize("net_burn, new_arr, expected, why", [
    (600, 0, INF, "burning while net new ARR is zero -> infinite, trips"),
    (600, -400, INF, "burning while ARR shrank -> infinite, trips"),
    (0, 400, 0.0, "zero burn -> 0, passes"),
    (-300, 400, 0.0, "generating cash -> 0, passes"),
    (0, 0, 0.0, "zero burn AND zero net new ARR -> 0 (not 0/0 = NaN)"),
    (-300, -400, 0.0, "generating cash while ARR shrank -> 0, not a positive multiple"),
])
def test_burn_multiple_edge_cases(net_burn, new_arr, expected, why):
    assert burn_multiple(burn_table([net_burn], [new_arr])).iloc[0] == expected, why


@pytest.mark.parametrize("net_burn, new_arr", [(NAN, 400), (600, NAN), (NAN, NAN)])
def test_burn_multiple_missing_input_is_nan(net_burn, new_arr):
    # "Data missing" must never turn into infinite or 0.
    assert math.isnan(burn_multiple(burn_table([net_burn], [new_arr])).iloc[0])


def test_burn_vs_budget():
    df = table(net_burn=[1200, 900], budget_net_burn=[1000, 1000])
    # 1200 / 1000 - 1 = +20% over ;  900 / 1000 - 1 = 10% under
    assert values(burn_vs_budget(df)) == [0.20, -0.10]


def test_runway_months():
    df = table(ending_cash=[4000, 4000, 4000, NAN], net_burn=[1000, 0, -500, 1000])
    # 4000 / (1000 / 3) = 12 ;  zero burn and cash-generating -> infinite ;  no cash figure -> NaN
    assert values(runway_months(df)) == [12, INF, INF, NAN]


def test_runway_months_missing_burn_is_nan():
    df = table(ending_cash=[4000], net_burn=[NAN])
    assert math.isnan(runway_months(df).iloc[0])


def test_runway_at_next_budget_uses_latest_cash():
    actuals = table(ending_cash=[9000, 4000])
    next_budget = pd.Series({"budget_new_arr": 1000, "budget_arr": 20000, "budget_net_burn": 600})
    # latest cash 4000 / (600 / 3) = 20 months (the older 9000 is ignored)
    assert runway_at_next_budget(actuals, next_budget) == pytest.approx(20)


@pytest.mark.filterwarnings("error")  # a divide-by-zero warning fails the test: zero burn must be handled on purpose
@pytest.mark.parametrize("budget_burn, expected", [(0, INF), (-200, INF)])
def test_runway_at_next_budget_not_burning(budget_burn, expected):
    actuals = table(ending_cash=[4000])
    next_budget = pd.Series({"budget_net_burn": budget_burn})
    assert runway_at_next_budget(actuals, next_budget) == expected


def test_runway_at_next_budget_missing():
    actuals = table(ending_cash=[4000])
    assert math.isnan(runway_at_next_budget(actuals, None))                                # no budget row
    assert math.isnan(runway_at_next_budget(actuals, pd.Series({"budget_net_burn": NAN})))  # budget burn blank
    blank_cash = table(ending_cash=[NAN])
    assert math.isnan(runway_at_next_budget(blank_cash, pd.Series({"budget_net_burn": 600})))


# ---------------------------------------------------------------------------
# Budget variance and CAC payback
# ---------------------------------------------------------------------------

def test_net_new_arr_vs_budget():
    df = table(budget_arr=[10000, 11000], new_arr=[900, 900], expansion_arr=[0, 0],
               contraction_arr=[0, 0], churned_arr=[0, 100])
    # Q4 2024: budgeted net new ARR = 11000 - 10000 = 1000; actual 900 - 100 = 800 -> 800 / 1000 - 1 = -20%
    # Q3 2024 has no prior quarter's budget -> NaN
    assert values(budget_net_new_arr(df)) == [NAN, 1000]
    assert values(net_new_arr_vs_budget(df)) == [NAN, -0.20]


def test_arr_vs_budget():
    df = table(starting_arr=[10000], new_arr=[1000], expansion_arr=[0], contraction_arr=[0],
               churned_arr=[0], budget_arr=[10000])
    # ending ARR 11000 / budgeted ending ARR 10000 - 1 = +10%
    assert values(arr_vs_budget(df)) == [0.10]


def cac_table(sm_spend, new_arr, gross_profit, revenue=1000):
    return table(sm_spend=[sm_spend], new_arr=[new_arr], gross_profit=[gross_profit], revenue=[revenue])


def test_cac_payback_months():
    # 600 / (400 * 750 / 1000) * 12 = 600 / 300 * 12 = 24 months
    assert cac_payback_months(cac_table(600, 400, 750)).iloc[0] == pytest.approx(24)


@pytest.mark.parametrize("new_arr, gross_profit, why", [
    (0, 750, "no new ARR -> never pays back"),
    (400, 0, "zero gross margin -> never pays back"),
    (400, -100, "negative gross margin -> never pays back"),
    (0, -100, "no new ARR and negative margin: 0 * negative = -0.0, must still be +infinite"),
])
def test_cac_payback_infinite(new_arr, gross_profit, why):
    assert cac_payback_months(cac_table(600, new_arr, gross_profit)).iloc[0] == INF, why


@pytest.mark.parametrize("sm_spend, new_arr, gross_profit", [(NAN, 400, 750), (600, NAN, 750), (600, 400, NAN)])
def test_cac_payback_missing_input_is_nan(sm_spend, new_arr, gross_profit):
    assert math.isnan(cac_payback_months(cac_table(sm_spend, new_arr, gross_profit)).iloc[0])


# ---------------------------------------------------------------------------
# compute_metrics and config wiring
# ---------------------------------------------------------------------------

def full_actuals(blank=(), blank_cells=()):
    """8 realistic quarters with every input column; nothing divides by zero.

    blank: quarters left entirely blank. blank_cells: (quarter, column) pairs left blank.
    ARR rolls forward: net new ARR is 1000 each quarter, so ending ARR = next starting ARR.
    """
    rows = {}
    for i, quarter in enumerate(EIGHT_QUARTERS):
        revenue = 2500 + 250 * i
        rows[quarter] = {
            "starting_arr": 10000 + 1000 * i, "new_arr": 1200, "expansion_arr": 300,
            "contraction_arr": 100, "churned_arr": 400, "revenue": revenue,
            "gross_profit": 0.75 * revenue, "net_burn": 900, "ending_cash": 20000 - 900 * i,
            "sm_spend": 1500, "new_customers": 20, "headcount": 100 + 5 * i,
            "pipeline": 8000 + 500 * i, "budget_new_arr": 1100, "budget_arr": 11000 + 1000 * i,
            "budget_net_burn": 800,
        }
    df = pd.DataFrame.from_dict(rows, orient="index", columns=STANDARD_COLUMNS, dtype=float)
    for quarter in blank:
        df.loc[quarter] = NAN
    for quarter, column in blank_cells:
        df.loc[quarter, column] = NAN
    return df


def test_compute_metrics_has_every_lookback_and_flag_column():
    # Catches a typo between compute_metrics, METRIC_LOOKBACK and FLAG_RULES.
    columns = set(compute_metrics(full_actuals()).columns)
    assert set(METRIC_LOOKBACK) <= columns
    assert {rule[1] for rule in FLAG_RULES} <= columns


def test_config_yaml_has_every_threshold():
    # Reads config.yaml (never edits it): every flag rule needs its threshold there.
    config = load_config()
    assert {rule[2] for rule in FLAG_RULES} <= set(config)
    assert set(TEST_CONFIG) <= set(config)


# ---------------------------------------------------------------------------
# check_threshold
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("value, threshold, kind, expected", [
    (0.99, 1.00, "min", TRIP),   # below a minimum trips
    (1.00, 1.00, "min", PASS),   # exactly at the threshold passes
    (1.01, 1.00, "min", PASS),
    (0.16, 0.15, "max", TRIP),   # above a maximum trips
    (0.15, 0.15, "max", PASS),   # exactly at the threshold passes
    (0.14, 0.15, "max", PASS),
    (12, 12, "min", PASS),       # runway exactly 12 months passes
    (-0.20, -0.20, "min", PASS),
])
def test_check_threshold_basic(value, threshold, kind, expected):
    assert check_threshold(value, threshold, kind) == expected


@pytest.mark.parametrize("value, threshold, kind", [
    (math.nextafter(0.15, 1), 0.15, "max"),   # the next float above 0.15: pure noise
    (math.nextafter(0.40, 0), 0.40, "min"),   # the next float below 0.40
    (math.nextafter(-0.20, -1), -0.20, "min"),
    (math.nextafter(24.0, 99), 24, "max"),
])
def test_check_threshold_float_noise_at_threshold_passes(value, threshold, kind):
    # Without rounding every one of these would trip on a difference of ~1e-16.
    assert value != threshold
    assert check_threshold(value, threshold, kind) == PASS


def test_check_threshold_rule_of_40_exactly_40_from_real_math():
    # Revenue 600 -> 1020 is +70% YoY, burn 306 on 1020 revenue is -30% FCF margin: exactly 40.
    df = table(revenue=[600, 700, 800, 900, 1020], net_burn=[0, 0, 0, 0, 306])
    value = rule_of_40(df).iloc[-1]
    assert value < 0.40  # Python gets 0.39999999999999997
    assert check_threshold(value, 0.40, "min") == PASS


def test_check_threshold_real_misses_still_trip():
    # Rounding is to 6 decimals, so a difference of 0.000001 is still real and still trips.
    assert check_threshold(0.150001, 0.15, "max") == TRIP
    assert check_threshold(0.399999, 0.40, "min") == TRIP


def test_check_threshold_missing_and_infinite():
    assert check_threshold(NAN, 1.00, "min") == MISSING
    assert check_threshold(np.float64("nan"), 0.15, "max") == MISSING  # pandas hands over numpy floats
    assert check_threshold(INF, 12, "min") == PASS    # runway when not burning
    assert check_threshold(INF, 2.0, "max") == TRIP   # burn multiple when ARR shrank
    assert check_threshold(INF, 24, "max") == TRIP    # CAC payback that never pays back


# ---------------------------------------------------------------------------
# check_combo: NRR falling while pipeline rising
# ---------------------------------------------------------------------------

def combo_metrics(nrr_values, pipeline_values):
    return table(nrr=nrr_values, pipeline=pipeline_values)


def latest(metrics):
    return metrics.index[-1]


@pytest.mark.parametrize("nrr_values, pipeline_values, expected", [
    ([1.08, 1.03, 0.97], [9000, 10000, 11000], TRIP),   # falling every step, rising every step
    ([1.08, 1.08, 0.97], [9000, 10000, 11000], PASS),   # NRR flat in one step: not "falling"
    ([1.00, 1.03, 0.97], [9000, 10000, 11000], PASS),   # NRR rose once
    ([1.08, 1.03, 0.97], [9000, 9000, 11000], PASS),    # pipeline flat in one step: not "rising"
    ([1.08, 1.03, 0.97], [9000, 12000, 11000], PASS),   # pipeline fell once
    ([1.10, 1.20, 1.08, 1.03, 0.97], [12000, 9000, 9500, 10000, 11000], TRIP),  # older quarters ignored
])
def test_check_combo_pass_and_trip(nrr_values, pipeline_values, expected):
    metrics = combo_metrics(nrr_values, pipeline_values)
    assert check_combo(metrics, TEST_CONFIG, latest(metrics)) == expected


def test_check_combo_nrr_float_noise_counts_as_flat():
    # 1.0 then 1.0 minus 1e-12 is noise, not a fall.
    metrics = combo_metrics([1.08, 1.0, 1.0 - 1e-12], [9000, 10000, 11000])
    assert check_combo(metrics, TEST_CONFIG, latest(metrics)) == PASS


@pytest.mark.parametrize("nrr_values, pipeline_values", [
    ([NAN, 1.03, 0.97], [9000, 10000, 11000]),   # NRR missing in the first quarter of the window
    ([1.08, NAN, 0.97], [9000, 10000, 11000]),   # ... in the middle
    ([1.08, 1.03, NAN], [9000, 10000, 11000]),   # ... in the latest quarter
    ([1.08, 1.03, 0.97], [9000, NAN, 11000]),    # pipeline missing
    ([1.08, NAN, 1.20], [9000, NAN, 8000]),      # missing, and the rest would PASS: still not PASS
])
def test_check_combo_missing_value_cannot_evaluate(nrr_values, pipeline_values):
    metrics = combo_metrics(nrr_values, pipeline_values)
    assert check_combo(metrics, TEST_CONFIG, latest(metrics)) == MISSING


def test_check_combo_not_enough_history_cannot_evaluate():
    # Only 2 quarters exist but the window needs 3.
    metrics = combo_metrics([1.08, 0.97], [9000, 11000])
    assert check_combo(metrics, TEST_CONFIG, latest(metrics)) == MISSING


def test_check_combo_gap_outside_window_is_ignored():
    # Q3 2024 is blank, but the window for Q2 2025 is Q4 2024 - Q2 2025.
    metrics = combo_metrics([NAN, 1.08, 1.03, 0.97], [NAN, 9000, 10000, 11000])
    assert check_combo(metrics, TEST_CONFIG, "Q2 2025") == TRIP


def test_check_combo_earlier_quarter():
    # Evaluating Q1 2025 uses the window ending there, not the latest quarter.
    metrics = combo_metrics([1.08, 1.03, 0.97, 1.10], [9000, 10000, 11000, 12000])
    assert check_combo(metrics, TEST_CONFIG, "Q1 2025") == TRIP
    assert check_combo(metrics, TEST_CONFIG, "Q2 2025") == PASS
    assert check_combo(metrics, TEST_CONFIG, "Q4 2024") == MISSING  # only 2 quarters of history


# ---------------------------------------------------------------------------
# evaluate_flags
# ---------------------------------------------------------------------------

def test_evaluate_flags_latest_quarter_by_default():
    metrics = compute_metrics(full_actuals())
    flags = evaluate_flags(metrics, TEST_CONFIG)
    assert [f["flag"] for f in flags] == [rule[0] for rule in FLAG_RULES] + [COMBO_FLAG_NAME]
    assert {f["quarter"] for f in flags} == {"Q2 2026"}
    assert all(f["status"] in (TRIP, PASS) for f in flags)  # full data: nothing is missing


def test_evaluate_flags_combo_can_be_switched_off():
    config = {**TEST_CONFIG, "nrr_falling_pipeline_rising_flag": False}
    flags = evaluate_flags(compute_metrics(full_actuals()), config)
    assert COMBO_FLAG_NAME not in [f["flag"] for f in flags]


# ---------------------------------------------------------------------------
# data_gaps
# ---------------------------------------------------------------------------

SAME_QUARTER_METRICS = ["ending_arr", "net_new_arr", "nrr", "grr", "gross_margin", "pipeline",
                        "fcf_margin", "burn_multiple", "burn_vs_budget", "arr_vs_budget",
                        "cac_payback_months", "runway_months"]
QOQ_METRICS = ["arr_qoq", "revenue_qoq", "pipeline_qoq", "net_new_arr_vs_budget"]
YOY_METRICS = ["arr_yoy", "revenue_yoy", "rule_of_40"]


def gaps_for(actuals):
    """Run the real pipeline (metrics -> flags -> gaps) on a test table."""
    metrics = compute_metrics(actuals)
    return data_gaps(actuals, metrics, evaluate_flags(metrics, TEST_CONFIG))


def test_data_gaps_none_when_nothing_is_blank():
    # The first quarters have no prior quarter or prior year (NaN), but that is not missing data.
    actuals = full_actuals()
    assert math.isnan(compute_metrics(actuals).loc["Q3 2024", "arr_yoy"])
    assert gaps_for(actuals) == {}


def test_data_gaps_blank_quarter_follows_the_rules():
    # Q1 2025 blank (like Northwind). CLAUDE.md: QoQ -> blank + next quarter; YoY -> blank + 4 later.
    gaps = gaps_for(full_actuals(blank=["Q1 2025"]))
    expected = {}
    expected.update({m: ["Q1 2025"] for m in SAME_QUARTER_METRICS})
    expected.update({m: ["Q1 2025", "Q2 2025"] for m in QOQ_METRICS})
    expected.update({m: ["Q1 2025", "Q1 2026"] for m in YOY_METRICS})
    assert gaps == expected  # flags for Q2 2026 don't touch Q1 2025, so no flag gaps


def test_data_gaps_blank_first_quarter():
    # The blank quarter's own QoQ/YoY count as gaps even though it has no history either.
    gaps = gaps_for(full_actuals(blank=["Q3 2024"]))
    assert gaps["nrr"] == ["Q3 2024"]
    assert gaps["arr_qoq"] == ["Q3 2024", "Q4 2024"]
    assert gaps["arr_yoy"] == ["Q3 2024", "Q3 2025"]


def test_data_gaps_blank_recent_quarter_includes_flags():
    # Q1 2026 blank: the Q2 2026 QoQ-based flag and the combo window can't be evaluated.
    gaps = gaps_for(full_actuals(blank=["Q1 2026"]))
    assert gaps["arr_qoq"] == ["Q1 2026", "Q2 2026"]
    assert gaps["arr_yoy"] == ["Q1 2026"]  # 4 quarters later is outside the data
    assert gaps["flag: Net new ARR vs budget"] == ["Q2 2026"]
    assert gaps[f"flag: {COMBO_FLAG_NAME}"] == ["Q2 2026"]
    assert "flag: Rule of 40" not in gaps  # Rule of 40 looks back to Q2 2025, which is fine
    assert "flag: Runway (months)" not in gaps


def test_data_gaps_blank_year_ago_quarter_trips_rule_of_40_gap():
    gaps = gaps_for(full_actuals(blank=["Q2 2025"]))
    assert gaps["rule_of_40"] == ["Q2 2025", "Q2 2026"]
    assert gaps["flag: Rule of 40"] == ["Q2 2026"]


def test_data_gaps_one_blank_cell_spreads_only_to_metrics_that_use_it():
    # Only pipeline is blank in Q4 2025: pipeline, pipeline QoQ and the combo rule are affected.
    gaps = gaps_for(full_actuals(blank_cells=[("Q4 2025", "pipeline")]))
    assert gaps == {
        "pipeline": ["Q4 2025"],
        "pipeline_qoq": ["Q4 2025", "Q1 2026"],
        f"flag: {COMBO_FLAG_NAME}": ["Q2 2026"],  # window Q4 2025 - Q2 2026 includes the blank
    }


def test_data_gaps_lists_missing_flags_given_directly():
    # A flag marked MISSING is listed under "flag: <name>" for its quarter; others are not.
    actuals = full_actuals()
    metrics = compute_metrics(actuals)
    flags = [
        {"flag": "Rule of 40", "quarter": "Q2 2026", "status": MISSING},
        {"flag": "Burn multiple", "quarter": "Q2 2026", "status": TRIP},
        {"flag": "NRR (annualized)", "quarter": "Q2 2026", "status": PASS},
    ]
    assert data_gaps(actuals, metrics, flags) == {"flag: Rule of 40": ["Q2 2026"]}
