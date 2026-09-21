"""Automated proof for build step 2, using Northwind's answer key from make_data.py.

Checks three things:
1. Cleaning is lossless: clean.py recovers TRUE_DATA exactly (Q1 2025 all blank).
2. Key Q2 2026 metrics match the hand formulas, and flags match the story.
3. Data gaps land exactly where the CLAUDE.md rules say.

Run: python check_northwind.py  -> prints "All checks passed" or stops at the first failure.
"""

import math

from analyze import BoardSummary, Point, Question, build_payload, payload_to_text, validate_summary
from clean import clean_workbook
from compare_models import HAIKU, SONNET, parse_scores, recommend
from make_data import BLANK_QUARTER, NEXT_QUARTER_BUDGET, OUTPUT_PATH, QUARTERS, TRUE_DATA
from metrics import (PASS, TRIP, compute_metrics, data_gaps, evaluate_flags,
                     load_config, metric_reasons, runway_at_next_budget)

LATEST = "Q2 2026"

# Q2 2026 metrics written as the same hand formulas you'd type into Excel.
EXPECTED_LATEST = {
    "net_burn": 3900,   # an input shown as it came: the workbook's Q2 2026 cell (the memo's questions quote it)
    "nrr": 1 + 4 * (580 - 260 - 510) / 25810,
    "grr": 1 - 4 * (260 + 510) / 25810,
    "burn_multiple": 3900 / (1850 + 580 - 260 - 510),
    "burn_vs_budget": 3900 / 3250 - 1,
    "runway_months": 14300 / (3900 / 3),
    "rule_of_40": 6660 / 4550 - 1 + (-3900 / 6660),
    "cac_payback_months": 2400 / (1850 * 5000 / 6660) * 12,
    "net_new_arr_vs_budget": 1660 / (26800 - 24750) - 1,
    "arr_yoy": 27470 / 19230 - 1,
    "arr_vs_budget": 27470 / 26800 - 1,
}
EXPECTED_RUNWAY_AT_BUDGET = 14300 / (3300 / 3)

# The Northwind story: 6 flags trip, 3 pass. Names typed by hand (the Metrics sheet labels).
EXPECTED_FLAGS = {
    "NRR (annualized)": TRIP,
    "GRR (annualized)": PASS,
    "Burn multiple": TRIP,
    "Net burn vs budget": TRIP,
    "Runway at current burn": TRIP,
    "CAC payback": PASS,
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


def good_summary():
    """A hand-written v5 answer that quotes only numbers from the payload, exactly as shown."""
    return BoardSummary(
        headline="ARR reached $27,470K but NRR fell to 97.1% and runway is 11.0 mo, tripping 6 of 9 flags.",
        diagnosis=(
            "Ending ARR reached $27,470K in Q2 2026 with ARR growth YoY of 42.8%, but NRR (annualized) is "
            "97.1% and GRR (annualized) 88.1%, both under Q3 2025's 108.9% and 92.9%. Net burn vs budget is "
            "20.0%, the burn multiple is 2.35x and runway at current burn is 11.0 mo. The pattern is "
            "consistent with new sales carrying a weakening installed base while spend stayed at plan."),
        risks=[
            Point(title="Retention slipping", detail="NRR went from 108.0% to 97.1% while pipeline rose."),
            Point(title="Burn over plan", detail="Net burn is 20.0% over budget with a 2.35x burn multiple."),
            Point(title="Short runway", detail="Runway is 11.0 mo at current burn, 13.0 mo if burn returns to plan."),
        ],
        questions=[
            Question(theme="Retention decomposition",
                     question="How much of the NRR (annualized) fall to 97.1% sits in GRR (annualized) at 88.1%?"),
            Question(theme="Retention decomposition",
                     question="Which accounts make up the churn behind GRR (annualized) at 88.1%?"),
            Question(theme="Burn and budget variance",
                     question="Which cost lines carry net burn vs budget at 20.0%, against 10.0% in Q3 2025?"),
            Question(theme="Burn and budget variance",
                     question="How much of the 2.35x burn multiple is headcount added since Q4 2025's 1.46x?"),
            Question(theme="Liquidity and runway",
                     question="How far must net burn vs budget fall for runway at current burn to clear 12.0 mo?"),
            Question(theme="Liquidity and runway",
                     question="Which cost actions stand behind the 13.0 mo runway if burn returns to plan?"),
            Question(theme="Pipeline and sales efficiency",
                     question="Which segments make up the $12,500K pipeline behind the 20.7 mo CAC payback?"),
            Question(theme="Pipeline and sales efficiency",
                     question="How much of net new ARR vs budget at -19.0% is deals slipping rather than lost?"),
            Question(theme="Definitions and assumptions",
                     question="Which definition gives NRR (annualized) at 97.1%, and was Q1 2025 ever restated?"),
        ],
    )


def replace_question(summary, position, theme, text):
    """The same answer with one question swapped, so one rule at a time is broken."""
    summary.questions[position] = Question(theme=theme, question=text)
    return summary


def check_analysis_validation(payload_text):
    """validate_summary passes a good answer and catches every way a v5 answer can go wrong."""
    assert validate_summary(good_summary(), payload_text) == [], \
        f"Good summary should pass: {validate_summary(good_summary(), payload_text)}"

    calculated = good_summary()
    calculated.risks[0].detail = "NRR fell 11.8 points in two quarters."  # 108.9 - 97.1: Claude did math
    problems = validate_summary(calculated, payload_text)
    assert any("11.8" in p for p in problems), f"Calculated number not caught: {problems}"

    rounded = good_summary()
    rounded.headline = "NRR fell to 97% as burn ran over budget."  # 97.1% rounded to 97%
    problems = validate_summary(rounded, payload_text)
    assert any("97" in p for p in problems), f"Rounded number not caught: {problems}"

    too_long = good_summary()
    too_long.risks[1].detail = ("Net burn is 20.0% over budget. The burn multiple is 2.35x. "
                                "Runway is 11.0 mo.")  # 3 sentences
    problems = validate_summary(too_long, payload_text)
    assert any("risks #2 detail" in p for p in problems), f"3-sentence detail not caught: {problems}"

    two_risks = good_summary()
    two_risks.risks = two_risks.risks[:2]
    problems = validate_summary(two_risks, payload_text)
    assert any("risks must have exactly 3" in p for p in problems), f"Wrong risk count not caught: {problems}"

    short_diagnosis = good_summary()
    short_diagnosis.diagnosis = "Retention is falling and burn is over budget."
    problems = validate_summary(short_diagnosis, payload_text)
    assert any("diagnosis has 8 words" in p for p in problems), f"Short diagnosis not caught: {problems}"

    few_questions = good_summary()
    few_questions.questions = few_questions.questions[:7]
    problems = validate_summary(few_questions, payload_text)
    assert any("questions must have 8 to 10 items" in p for p in problems), f"7 questions not caught: {problems}"

    explaining = replace_question(good_summary(), 2, "Burn and budget variance",
                                  "What is driving net burn vs budget to 20.0% this quarter?")
    problems = validate_summary(explaining, payload_text)
    assert any("asks for an explanation" in p for p in problems), f"Explanation question not caught: {problems}"

    no_value = replace_question(good_summary(), 2, "Burn and budget variance",
                                "Which cost lines carry the net burn vs budget variance?")
    problems = validate_summary(no_value, payload_text)
    assert any("must name a metric and quote its value" in p for p in problems), \
        f"Question without a value not caught: {problems}"

    no_divergence = replace_question(good_summary(), 0, "Retention decomposition",
                                     "Which accounts make up the NRR (annualized) fall to 97.1%?")
    problems = validate_summary(no_divergence, payload_text)
    assert any("divergence test" in p for p in problems), f"Missing gross versus net test not caught: {problems}"

    no_definitions = replace_question(good_summary(), 8, "Pipeline and sales efficiency",
                                      "Which segments carry the $12,500K pipeline into next quarter?")
    problems = validate_summary(no_definitions, payload_text)
    assert any("interrogates a definition" in p for p in problems), f"Missing definitions question: {problems}"

    wrong_theme = replace_question(good_summary(), 1, "Retention",
                                   "Which accounts make up the churn behind GRR (annualized) at 88.1%?")
    problems = validate_summary(wrong_theme, payload_text)
    assert any("has the theme 'Retention'" in p for p in problems), f"Unknown theme not caught: {problems}"

    assumption_only = replace_question(good_summary(), 8, "Definitions and assumptions",
                                       "Which assumption sits behind runway at current burn of 11.0 mo?")
    problems = validate_summary(assumption_only, payload_text)
    assert any("interrogates a definition" in p for p in problems), f"An assumption stood in for a definition: {problems}"

    industry_standard = good_summary()
    industry_standard.headline = "NRR fell to 97.1%, below the industry standard of 100.0%."
    problems = validate_summary(industry_standard, payload_text)
    assert any("outside standard" in p for p in problems), f"An industry standard was not caught: {problems}"


def check_recommendation_rule():
    """Haiku only with avg score >= 4.0 AND 100% pass rate; scores must be complete and 1-5."""
    assert recommend({"avg_score": 4.0, "pass_rate": 1.0}) == HAIKU, "4.0 and 100% should pick Haiku"
    assert recommend({"avg_score": 3.9, "pass_rate": 1.0}) == SONNET, "3.9 should pick Sonnet"
    assert recommend({"avg_score": 4.7, "pass_rate": 2 / 3}) == SONNET, "67% pass rate should pick Sonnet"
    assert recommend({"avg_score": None, "pass_rate": 0.0}) == SONNET, "no passing runs should pick Sonnet"

    runs = {"A": {"passed": True}, "B": {"passed": False}}
    assert parse_scores(["A=4"], runs) == {"A": 4}
    for bad_args in (["A=6"], ["A=0"], ["A=four"], [], ["B=3", "A=4"], ["Z=4"]):
        try:
            parse_scores(bad_args, runs)
        except ValueError:
            continue
        raise AssertionError(f"parse_scores should reject {bad_args}")


def main():
    actuals, next_budget = clean_workbook(OUTPUT_PATH)
    config = load_config()
    metrics = compute_metrics(actuals)
    flags = evaluate_flags(metrics, metric_reasons(actuals, metrics), config)

    check_cleaning(actuals, next_budget)
    print("✓ Cleaning is lossless (all values match TRUE_DATA, Q1 2025 blank, budget row read)")
    check_latest_metrics(metrics, runway_at_next_budget(actuals, next_budget))
    print("✓ Q2 2026 metrics match the hand formulas")
    check_flags(flags)
    print("✓ Flags match the story (6 trip, 3 pass)")
    check_gaps(data_gaps(actuals, metrics, flags), metrics)
    print("✓ Data gaps are exactly where the rules say")
    check_analysis_validation(payload_to_text(build_payload("Northwind", actuals, next_budget, config)))
    print("✓ Analysis validation passes a good answer, catches calculated/rounded numbers, wrong counts, a short "
          "diagnosis, and questions that explain, quote no value, skip the gross versus net test, leave out "
          "definitions (an assumption is not one), use an unknown theme, or cite an industry standard")
    check_recommendation_rule()
    print("✓ Model recommendation rule and score entry behave as agreed")
    print("All checks passed")


if __name__ == "__main__":
    main()
