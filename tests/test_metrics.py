"""Unit tests for metrics.py: every metric, its edge cases, flags and data gaps.

Expected values are worked out by hand from the CLAUDE.md definitions (shown in comments),
not copied from the code. Run from the project folder:  pytest
"""

import math
import re

import numpy as np
import pandas as pd
import pytest

from clean import STANDARD_COLUMNS
from metrics import (CANNOT_EVALUATE, COMBO_FLAG_NAME, FLAG_RULES, INPUT_LABELS, METRIC_INPUTS, METRIC_LABELS,
                     METRIC_LOOKBACK, MISSING_INPUT, NO_PRIOR_PERIOD, NOT_MEANINGFUL, PASS, TRIP,
                     arr_vs_budget, budget_net_new_arr, burn_multiple, burn_vs_budget,
                     cac_payback_months, check_combo, check_threshold, compute_metrics, data_gaps,
                     display_value, ending_arr, evaluate_flags, fcf_margin, gross_margin, growth, grr, load_config,
                     metric_reasons, net_new_arr, net_new_arr_vs_budget, not_meaningful_text, nrr,
                     print_report, rule_of_40, runway_at_next_budget, runway_months, validate_config)

NAN = math.nan
INF = math.inf

# Thresholds written out here, so tuning config.yaml never breaks these tests.
TEST_CONFIG = {
    "nrr_min": 1.00, "grr_min": 0.85, "burn_multiple_max": 2.0, "burn_over_budget_max": 0.15,
    "runway_min_months": 12, "cac_payback_max_months": 24, "net_new_arr_vs_budget_min": -0.20,
    "rule_of_40_min": 0.40, "nrr_falling_pipeline_rising_flag": True, "combo_lookback_quarters": 3,
    "combo_min_nrr_drop": 0.01,
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


@pytest.mark.parametrize("net_burn, new_arr", [(NAN, 400), (600, NAN), (NAN, NAN), (0, NAN), (-300, NAN)])
def test_burn_multiple_missing_input_is_nan(net_burn, new_arr):
    # "Data missing" must never turn into infinite or 0.
    # (0, NaN) and (-300, NaN): the "not burning -> 0" rule must not hide a blank ARR input (decision B).
    assert math.isnan(burn_multiple(burn_table([net_burn], [new_arr])).iloc[0])


def test_burn_vs_budget():
    df = table(net_burn=[1200, 900], budget_net_burn=[1000, 1000])
    # 1200 / 1000 - 1 = +20% over ;  900 / 1000 - 1 = 10% under
    assert values(burn_vs_budget(df)) == [0.20, -0.10]


def test_burn_vs_budget_not_meaningful_when_budget_not_burning():
    # Decision C: a budget of 0 or less can't be "over budget" by a %. Old code: -200 / -150 - 1 = +33%, trips.
    df = table(net_burn=[200, -200, 200], budget_net_burn=[0, -150, NAN])
    assert values(burn_vs_budget(df)) == [NAN, NAN, NAN]


def test_runway_months():
    df = table(ending_cash=[4000, 4000, 4000, NAN], net_burn=[1000, 0, -500, 1000])
    # 4000 / (1000 / 3) = 12 ;  zero burn and cash-generating -> infinite ;  no cash figure -> NaN
    assert values(runway_months(df)) == [12, INF, INF, NAN]


def test_runway_months_missing_burn_is_nan():
    df = table(ending_cash=[4000], net_burn=[NAN])
    assert math.isnan(runway_months(df).iloc[0])


@pytest.mark.parametrize("net_burn", [0, -500])
def test_runway_months_blank_cash_while_not_burning_is_nan(net_burn):
    # Decision B: the "not burning -> infinite" rule must not turn a blank cash cell into a pass.
    df = table(ending_cash=[NAN], net_burn=[net_burn])
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


@pytest.mark.parametrize("budget_burn", [0, -200])
def test_runway_at_next_budget_blank_cash_while_budget_not_burning_is_nan(budget_burn):
    # Edge cases apply only when every input is there: a blank cash cell must not become "∞ (budget not burning)".
    blank_cash = table(ending_cash=[NAN])
    assert math.isnan(runway_at_next_budget(blank_cash, pd.Series({"budget_net_burn": budget_burn})))


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


@pytest.mark.parametrize("next_budget_arr, why", [(10000, "budget flat: 0 net new ARR"),
                                                   (9500, "budget shrinks: -500 net new ARR")])
def test_net_new_arr_vs_budget_not_meaningful_when_budget_not_growing(next_budget_arr, why):
    # Decision C: "% vs a budget of 0 or less" means nothing (800 / 0 would be infinite).
    df = table(budget_arr=[10000, next_budget_arr], new_arr=[900, 900], expansion_arr=[0, 0],
               contraction_arr=[0, 0], churned_arr=[0, 100])
    assert math.isnan(net_new_arr_vs_budget(df).iloc[1]), why


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


@pytest.mark.parametrize("sm_spend, new_arr, gross_profit", [
    (NAN, 400, 750), (600, NAN, 750), (600, 400, NAN),
    (NAN, 0, 750),     # decision B: blank S&M with no new ARR used to be infinite -> a red flag on missing data
    (NAN, 400, -100),  # ... and blank S&M with negative margin, the same way
])
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
    assert check_threshold(NAN, 1.00, "min") == CANNOT_EVALUATE
    assert check_threshold(np.float64("nan"), 0.15, "max") == CANNOT_EVALUATE  # pandas hands over numpy floats
    assert check_threshold(INF, 12, "min") == PASS    # runway when not burning
    assert check_threshold(INF, 2.0, "max") == TRIP   # burn multiple when ARR shrank
    assert check_threshold(INF, 24, "max") == TRIP    # CAC payback that never pays back


# ---------------------------------------------------------------------------
# check_combo: NRR falling while pipeline rising
# ---------------------------------------------------------------------------

def combo_metrics(nrr_values, pipeline_values):
    return table(nrr=nrr_values, pipeline=pipeline_values)


def reasons_for(metrics, reason=MISSING_INPUT):
    """A reasons table for a hand-made metrics table: `reason` wherever a value is NaN, else None."""
    return pd.DataFrame({column: [reason if pd.isna(v) else None for v in metrics[column]]
                         for column in metrics.columns}, index=metrics.index, dtype=object)


def combo(metrics, quarter=None, config=TEST_CONFIG, reason=MISSING_INPUT):
    """check_combo on a hand-made table. Returns (status, reason)."""
    quarter = metrics.index[-1] if quarter is None else quarter
    return check_combo(metrics, reasons_for(metrics, reason), config, quarter)


@pytest.mark.parametrize("nrr_values, pipeline_values, expected", [
    ([1.08, 1.03, 0.97], [9000, 10000, 11000], TRIP),   # falling every step, rising every step
    ([1.08, 1.08, 0.97], [9000, 10000, 11000], PASS),   # NRR flat in one step: not "falling"
    ([1.00, 1.03, 0.97], [9000, 10000, 11000], PASS),   # NRR rose once
    ([1.08, 1.03, 0.97], [9000, 9000, 11000], PASS),    # pipeline flat in one step: not "rising"
    ([1.08, 1.03, 0.97], [9000, 12000, 11000], PASS),   # pipeline fell once
    ([1.10, 1.20, 1.08, 1.03, 0.97], [12000, 9000, 9500, 10000, 11000], TRIP),  # older quarters ignored
])
def test_check_combo_pass_and_trip(nrr_values, pipeline_values, expected):
    assert combo(combo_metrics(nrr_values, pipeline_values)) == (expected, None)


def test_check_combo_nrr_float_noise_counts_as_flat():
    # 1.0 then 1.0 minus 1e-12 is noise, not a fall.
    assert combo(combo_metrics([1.08, 1.0, 1.0 - 1e-12], [9000, 10000, 11000])) == (PASS, None)


@pytest.mark.parametrize("nrr_values, expected, why", [
    ([1.08, 1.03, 1.029], PASS, "decision D: a 0.1-point dip is noise, not a fall"),
    ([1.08, 1.03, 1.0205], PASS, "0.95 points: still under the 1-point minimum"),
    ([1.08, 1.07, 1.06], TRIP, "exactly 1.0 point at each step counts as falling"),
    ([1.0802, 1.0202, 0.9706], TRIP, "Northwind: -6.0 and -5.0 points still trips"),
])
def test_check_combo_nrr_must_fall_at_least_the_minimum(nrr_values, expected, why):
    assert combo(combo_metrics(nrr_values, [9000, 10000, 11000]))[0] == expected, why


@pytest.mark.parametrize("nrr_values, pipeline_values", [
    ([NAN, 1.03, 0.97], [9000, 10000, 11000]),   # NRR missing in the first quarter of the window
    ([1.08, NAN, 0.97], [9000, 10000, 11000]),   # ... in the middle
    ([1.08, 1.03, NAN], [9000, 10000, 11000]),   # ... in the latest quarter
    ([1.08, 1.03, 0.97], [9000, NAN, 11000]),    # pipeline missing
    ([1.08, NAN, 1.20], [9000, NAN, 8000]),      # missing, and the rest would PASS: still not PASS
])
def test_check_combo_missing_value_cannot_evaluate(nrr_values, pipeline_values):
    metrics = combo_metrics(nrr_values, pipeline_values)
    assert combo(metrics) == (CANNOT_EVALUATE, MISSING_INPUT)


def test_check_combo_not_meaningful_nrr_says_so():
    # NRR that can't be computed for a reason other than a blank (e.g. starting ARR of 0).
    metrics = combo_metrics([1.08, NAN, 0.97], [9000, 10000, 11000])
    assert combo(metrics, reason=NOT_MEANINGFUL) == (CANNOT_EVALUATE, NOT_MEANINGFUL)


def test_check_combo_not_enough_history_cannot_evaluate():
    # Only 2 quarters exist but the window needs 3.
    assert combo(combo_metrics([1.08, 0.97], [9000, 11000])) == (CANNOT_EVALUATE, NO_PRIOR_PERIOD)


@pytest.mark.parametrize("nrr_values, pipeline_values", [
    ([1.08, NAN], [9000, 11000]),   # latest NRR blank
    ([1.08, 0.97], [NAN, 11000]),   # earlier pipeline blank
])
def test_check_combo_blank_input_wins_over_not_enough_history(nrr_values, pipeline_values):
    # CLAUDE.md: a blank input wins. Too little history AND a blank is "missing input", so it is a data gap.
    # Old code checked history first and said "no prior period", hiding the blank.
    assert combo(combo_metrics(nrr_values, pipeline_values)) == (CANNOT_EVALUATE, MISSING_INPUT)


def test_check_combo_not_enough_history_beats_not_meaningful():
    # Same order as a metric's own reasons: missing input, then no prior period, then not meaningful.
    metrics = combo_metrics([1.08, NAN], [9000, 11000])
    assert combo(metrics, reason=NOT_MEANINGFUL) == (CANNOT_EVALUATE, NO_PRIOR_PERIOD)


def test_check_combo_gap_outside_window_is_ignored():
    # Q3 2024 is blank, but the window for Q2 2025 is Q4 2024 - Q2 2025.
    metrics = combo_metrics([NAN, 1.08, 1.03, 0.97], [NAN, 9000, 10000, 11000])
    assert combo(metrics, "Q2 2025") == (TRIP, None)


def test_check_combo_earlier_quarter():
    # Evaluating Q1 2025 uses the window ending there, not the latest quarter.
    metrics = combo_metrics([1.08, 1.03, 0.97, 1.10], [9000, 10000, 11000, 12000])
    assert combo(metrics, "Q1 2025") == (TRIP, None)
    assert combo(metrics, "Q2 2025") == (PASS, None)
    assert combo(metrics, "Q4 2024") == (CANNOT_EVALUATE, NO_PRIOR_PERIOD)  # only 2 quarters of history


# ---------------------------------------------------------------------------
# Config validation (decision E)
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("change, expected_in_message", [
    ({"combo_lookback_quarters": 1}, "combo_lookback_quarters must be a whole number of at least 2"),
    ({"combo_lookback_quarters": 0}, "combo_lookback_quarters must be a whole number of at least 2"),
    ({"combo_lookback_quarters": 2.5}, "combo_lookback_quarters must be a whole number of at least 2"),
    ({"combo_min_nrr_drop": -0.01}, "combo_min_nrr_drop must be a decimal from 0 to 1"),
    ({"combo_min_nrr_drop": None}, "combo_min_nrr_drop must be a decimal from 0 to 1"),
])
def test_validate_config_stops_on_bad_combo_settings(change, expected_in_message):
    # Old code: a lookback of 1 has no steps to compare, and all() of nothing is True, so the combo always tripped.
    with pytest.raises(ValueError, match=expected_in_message):
        validate_config({**TEST_CONFIG, **change})


def test_validate_config_stops_on_missing_min_drop():
    config = {key: value for key, value in TEST_CONFIG.items() if key != "combo_min_nrr_drop"}
    with pytest.raises(ValueError, match="combo_min_nrr_drop"):
        validate_config(config)


def test_bad_lookback_stops_flags_and_config_file(tmp_path):
    with pytest.raises(ValueError, match="combo_lookback_quarters"):
        evaluate_flags(compute_metrics(full_actuals()), metric_reasons(full_actuals(), compute_metrics(full_actuals())),
                       {**TEST_CONFIG, "combo_lookback_quarters": 1})
    bad_file = tmp_path / "config.yaml"
    bad_file.write_text("combo_lookback_quarters: 1\ncombo_min_nrr_drop: 0.01\n")
    with pytest.raises(ValueError, match="config.yaml: combo_lookback_quarters"):
        load_config(bad_file)


def test_validate_config_accepts_the_real_config():
    validate_config(TEST_CONFIG)
    validate_config({**TEST_CONFIG, "combo_lookback_quarters": 2, "combo_min_nrr_drop": 0})


# ---------------------------------------------------------------------------
# evaluate_flags
# ---------------------------------------------------------------------------

def flags_for(actuals, config=TEST_CONFIG, quarter=None):
    """The real pipeline for a test table: metrics -> reasons -> flags."""
    metrics = compute_metrics(actuals)
    return evaluate_flags(metrics, metric_reasons(actuals, metrics), config, quarter)


def flag_named(flags, name):
    return next(flag for flag in flags if flag["flag"] == name)


def test_evaluate_flags_latest_quarter_by_default():
    flags = flags_for(full_actuals())
    assert [f["flag"] for f in flags] == [rule[0] for rule in FLAG_RULES] + [COMBO_FLAG_NAME]
    assert {f["quarter"] for f in flags} == {"Q2 2026"}
    assert all(f["status"] in (TRIP, PASS) for f in flags)  # full data: nothing is missing
    assert all(f["reason"] is None for f in flags)          # a reason only when a flag can't be evaluated


def test_evaluate_flags_combo_can_be_switched_off():
    flags = flags_for(full_actuals(), {**TEST_CONFIG, "nrr_falling_pipeline_rising_flag": False})
    assert COMBO_FLAG_NAME not in [f["flag"] for f in flags]


def test_flag_names_are_the_metric_labels():
    # Decision I: one label set, so the Flags and Metrics sheets use the same words.
    assert [rule[0] for rule in FLAG_RULES] == [
        "NRR (annualized)", "GRR (annualized)", "Burn multiple", "Net burn vs budget",
        "Runway at current burn", "CAC payback", "Net new ARR vs budget", "Rule of 40"]
    assert all(rule[0] == METRIC_LABELS[rule[1]] for rule in FLAG_RULES)
    assert set(METRIC_LABELS) == set(compute_metrics(full_actuals()).columns)


def test_net_burn_is_shown_next_to_the_other_dollar_rows_under_one_label():
    # The memo's AI questions quote net burn in dollars, and every number in the memo must be one the metrics
    # workbook shows, so the workbook shows it. One label set: the label lives in METRIC_LABELS only.
    labels = list(METRIC_LABELS)
    assert METRIC_LABELS["net_burn"] == "Net burn ($K)"
    assert labels[:3] == ["ending_arr", "net_new_arr", "net_burn"]   # with ending ARR and net new ARR
    assert "net_burn" not in INPUT_LABELS
    assert set(INPUT_LABELS.values()).isdisjoint(METRIC_LABELS.values())


def test_net_burn_is_an_input_not_a_flag():
    # Shown, never judged: no flag reads it, and its inputs are just itself (no other quarter).
    assert "net_burn" not in {rule[1] for rule in FLAG_RULES}
    assert "Net burn ($K)" not in [rule[0] for rule in FLAG_RULES]
    assert METRIC_INPUTS["net_burn"] == [("net_burn", 0)]
    assert METRIC_LOOKBACK["net_burn"] == 0


def test_net_burn_column_is_the_input_as_it_came():
    # full_actuals: net burn 900 in every quarter. Nothing is calculated, so nothing can differ from the workbook.
    actuals = full_actuals()
    metrics = compute_metrics(actuals)
    assert list(metrics["net_burn"]) == [900.0] * 8
    assert display_value(actuals, metrics, metric_reasons(actuals, metrics), "net_burn", "Q2 2026") == "900"


def test_a_negative_or_zero_net_burn_is_shown_as_it_is():
    # Not burning is a real number (cash generative), not a gap and not "not meaningful".
    actuals = full_actuals()
    actuals.loc["Q1 2026", "net_burn"] = -250.0
    actuals.loc["Q2 2026", "net_burn"] = 0.0
    metrics = compute_metrics(actuals)
    reasons = metric_reasons(actuals, metrics)
    assert [display_value(actuals, metrics, reasons, "net_burn", q) for q in ("Q1 2026", "Q2 2026")] == ["-250", "0"]
    assert reasons.loc[["Q1 2026", "Q2 2026"], "net_burn"].isna().all()   # no reason: there is a number


def test_a_blank_net_burn_is_a_data_gap_of_its_own_quarter_only():
    # A blank input stays blank and is a gap, but net burn has no other quarter in it, so the gap doesn't spread.
    actuals = full_actuals(blank_cells=[("Q4 2025", "net_burn")])
    metrics = compute_metrics(actuals)
    assert math.isnan(metrics.loc["Q4 2025", "net_burn"])
    assert metric_reasons(actuals, metrics).loc["Q4 2025", "net_burn"] == MISSING_INPUT
    assert gaps_for(actuals)["net_burn"] == ["Q4 2025"]
    assert metric_reasons(actuals, metrics).loc["Q1 2026", "net_burn"] is None


def test_adding_net_burn_leaves_every_other_metric_where_it_was():
    # The first values of the two metrics that read net burn, worked out by hand from full_actuals:
    # burn multiple 900 / net new ARR 1000 = 0.9; FCF margin -900 / revenue 2500 = -0.36.
    metrics = compute_metrics(full_actuals())
    assert metrics.loc["Q3 2024", "burn_multiple"] == pytest.approx(0.9)
    assert metrics.loc["Q3 2024", "fcf_margin"] == pytest.approx(-0.36)
    assert len(metrics.columns) == 20


def test_cac_payback_blank_sm_spend_cannot_evaluate_never_trips():
    # Decision B: blank S&M with no new ARR was infinite payback -> a red flag built on missing data.
    actuals = full_actuals(blank_cells=[("Q2 2026", "sm_spend")])
    actuals.loc["Q2 2026", "new_arr"] = 0
    flag = flag_named(flags_for(actuals), "CAC payback")
    assert (flag["status"], flag["reason"]) == (CANNOT_EVALUATE, MISSING_INPUT)
    assert gaps_for(actuals)["flag: CAC payback"] == ["Q2 2026"]


def test_burn_vs_budget_flag_not_meaningful_does_not_trip():
    # Decision C: budgeted burn of 0 -> not meaningful; the flag doesn't trip and isn't a data gap.
    actuals = full_actuals()
    actuals.loc["Q2 2026", "budget_net_burn"] = 0
    flag = flag_named(flags_for(actuals), "Net burn vs budget")
    assert (flag["status"], flag["reason"]) == (CANNOT_EVALUATE, NOT_MEANINGFUL)
    assert gaps_for(actuals) == {}
    # full_actuals: net burn 900 in every quarter.
    assert not_meaningful_text(actuals, "burn_vs_budget", "Q2 2026") == "n/m: net burn 900 vs budget 0 ($K)"


def test_net_new_arr_vs_budget_flag_not_meaningful_does_not_trip():
    # full_actuals: Q1 2026 budget_arr = 11000 + 1000 * 6 = 17000; net new ARR = 1200 + 300 - 100 - 400 = 1000.
    actuals = full_actuals()
    actuals.loc["Q2 2026", "budget_arr"] = 16500  # budgeted net new ARR = 16500 - 17000 = -500
    flag = flag_named(flags_for(actuals), "Net new ARR vs budget")
    assert (flag["status"], flag["reason"]) == (CANNOT_EVALUATE, NOT_MEANINGFUL)
    assert gaps_for(actuals) == {}
    assert (not_meaningful_text(actuals, "net_new_arr_vs_budget", "Q2 2026")
            == "n/m: net new ARR 1,000 vs budget -500 ($K)")


def test_not_meaningful_text_for_other_metrics():
    assert not_meaningful_text(full_actuals(), "gross_margin", "Q2 2026") == "n/m (not meaningful)"


# ---------------------------------------------------------------------------
# data_gaps
# ---------------------------------------------------------------------------

SAME_QUARTER_METRICS = ["ending_arr", "net_new_arr", "net_burn", "nrr", "grr", "gross_margin", "pipeline",
                        "fcf_margin", "burn_multiple", "burn_vs_budget", "arr_vs_budget",
                        "cac_payback_months", "runway_months"]
QOQ_METRICS = ["arr_qoq", "revenue_qoq", "pipeline_qoq", "net_new_arr_vs_budget"]
YOY_METRICS = ["arr_yoy", "revenue_yoy", "rule_of_40"]


def gaps_for(actuals):
    """Run the real pipeline (metrics -> reasons -> flags -> gaps) on a test table."""
    metrics = compute_metrics(actuals)
    return data_gaps(actuals, metrics, evaluate_flags(metrics, metric_reasons(actuals, metrics), TEST_CONFIG))


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
    assert "flag: Runway at current burn" not in gaps


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
    # Only a flag that can't be evaluated because of a MISSING INPUT is a data gap (decision A).
    actuals = full_actuals()
    metrics = compute_metrics(actuals)
    flags = [
        {"flag": "Rule of 40", "quarter": "Q2 2026", "status": CANNOT_EVALUATE, "reason": MISSING_INPUT},
        {"flag": "Net burn vs budget", "quarter": "Q2 2026", "status": CANNOT_EVALUATE, "reason": NOT_MEANINGFUL},
        {"flag": COMBO_FLAG_NAME, "quarter": "Q2 2026", "status": CANNOT_EVALUATE, "reason": NO_PRIOR_PERIOD},
        {"flag": "Burn multiple", "quarter": "Q2 2026", "status": TRIP, "reason": None},
        {"flag": "NRR (annualized)", "quarter": "Q2 2026", "status": PASS, "reason": None},
    ]
    assert data_gaps(actuals, metrics, flags) == {"flag: Rule of 40": ["Q2 2026"]}


# ---------------------------------------------------------------------------
# Reasons: missing input vs no prior period vs not meaningful (decision A)
# ---------------------------------------------------------------------------

def reasons_of(actuals):
    return metric_reasons(actuals, compute_metrics(actuals))


def test_partly_blank_quarter_is_not_a_gap_for_metrics_that_dont_use_the_blank():
    # Old code: any blank cell in Q3 2024 made every NaN in that quarter a gap, so ARR YoY
    # ("no prior year") was reported as data missing. Headcount feeds no metric at all.
    actuals = full_actuals(blank_cells=[("Q3 2024", "headcount")])
    assert gaps_for(actuals) == {}
    assert reasons_of(actuals).loc["Q3 2024", "arr_yoy"] == NO_PRIOR_PERIOD


def test_partly_blank_quarter_only_hits_metrics_that_use_it():
    # Q3 2024 revenue blank: revenue metrics are missing input; ARR YoY there is still "no prior period".
    actuals = full_actuals(blank_cells=[("Q3 2024", "revenue")])
    reasons = reasons_of(actuals)
    assert reasons.loc["Q3 2024", "gross_margin"] == MISSING_INPUT
    assert reasons.loc["Q3 2024", "revenue_yoy"] == MISSING_INPUT  # its own input is blank: a gap
    assert reasons.loc["Q3 2024", "arr_yoy"] == NO_PRIOR_PERIOD    # doesn't use revenue: not a gap
    assert reasons.loc["Q3 2025", "revenue_yoy"] == MISSING_INPUT  # a year later needs the blank
    assert reasons.loc["Q2 2026", "revenue_yoy"] is None           # a normal value


def test_zero_divided_by_zero_is_not_meaningful_not_missing():
    # Revenue 0 and gross profit 0: gross margin 0 / 0 is undefined, but no data is missing.
    actuals = full_actuals()
    actuals.loc["Q2 2026", ["revenue", "gross_profit"]] = 0
    metrics = compute_metrics(actuals)
    reasons = metric_reasons(actuals, metrics)
    assert math.isnan(metrics.loc["Q2 2026", "gross_margin"])
    assert reasons.loc["Q2 2026", "gross_margin"] == NOT_MEANINGFUL
    assert math.isnan(metrics.loc["Q2 2026", "fcf_margin"])        # -900 / 0: infinite is not a real margin
    assert reasons.loc["Q2 2026", "fcf_margin"] == NOT_MEANINGFUL
    assert gaps_for(actuals) == {}


def test_metric_inputs_name_real_columns_for_every_metric():
    assert set(METRIC_INPUTS) == set(compute_metrics(full_actuals()).columns)
    for metric, inputs in METRIC_INPUTS.items():
        for column, quarters_back in inputs:
            assert column in STANDARD_COLUMNS, f"{metric}: unknown input {column}"
            assert quarters_back in (0, 1, 4), f"{metric}: odd lookback {quarters_back}"
    assert METRIC_LOOKBACK == {m: max(back for _, back in inputs) for m, inputs in METRIC_INPUTS.items()}


@pytest.mark.parametrize("column", STANDARD_COLUMNS)
@pytest.mark.parametrize("quarters_back", [0, 1, 4])
def test_every_blank_input_is_reported_as_missing_input(column, quarters_back):
    # Blank one cell, then every Q2 2026 metric that lost its value must say MISSING INPUT.
    # Catches an input a metric uses but METRIC_INPUTS forgot (it would wrongly say "not meaningful").
    blank_quarter = EIGHT_QUARTERS[-1 - quarters_back]
    before = compute_metrics(full_actuals()).loc["Q2 2026"]
    actuals = full_actuals(blank_cells=[(blank_quarter, column)])
    metrics = compute_metrics(actuals)
    reasons = metric_reasons(actuals, metrics)
    for metric in metrics.columns:
        declared = (column, quarters_back) in METRIC_INPUTS[metric]
        lost = not math.isnan(before[metric]) and math.isnan(metrics.loc["Q2 2026", metric])
        assert declared == (reasons.loc["Q2 2026", metric] == MISSING_INPUT), f"{metric}: {column} -{quarters_back}"
        assert lost == declared, f"{metric}: blank {column} {quarters_back} quarters back, value lost={lost}"


def test_print_report_says_data_missing_or_no_prior_period(capsys):
    # Decision J: the printout used to show a bare "n/a" for both.
    actuals = full_actuals(blank=["Q1 2025"])
    budget = pd.Series({"budget_new_arr": 1100.0, "budget_arr": 19000.0, "budget_net_burn": 800.0}, name="Q3 2026 (Budget)")
    print_report(actuals, budget, TEST_CONFIG)
    output = capsys.readouterr().out
    arr_yoy_line = next(line for line in output.splitlines() if line.startswith("ARR growth YoY"))
    assert "n/a (no prior period)" in arr_yoy_line  # Q3 2024 - Q4 2024
    assert "data missing" in arr_yoy_line           # Q1 2025 and Q1 2026
    assert re.search(r"n/a(?! \()", output) is None, "a bare 'n/a' is still printed"
