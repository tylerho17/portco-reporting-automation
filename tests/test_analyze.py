"""Unit tests for analyze.py that need no API call: the number check and the payload.

Expected values are worked out by hand (shown in comments), not copied from the code.
Run from the project folder:  pytest
"""

import json

import pytest

from analyze import BoardSummary, Point, build_payload, find_ungrounded_numbers, numbers_in, save_analysis
from build_deck import load_analysis
from clean import clean_workbook
from make_data import OUTPUT_PATH as NORTHWIND
from metrics import load_config


# ---------------------------------------------------------------------------
# The number check reads minus signs (decision O)
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("text, expected", [
    ("-19.0%", {-19.0}),                      # hyphen as a minus sign
    ("−12.2%", {-12.2}),                      # true minus sign (U+2212)
    ("–3.8%", {-3.8}),                        # en dash used as a minus sign
    ("(-19.0%)", {-19.0}),                    # after a bracket
    ("-$240K", {-240.0}),                     # minus before a dollar sign
    ("fell to -12.2% from -3.8%", {-12.2, -3.8}),
    ("Q4 2025-Q2 2026", {4.0, 2025.0, 2.0, 2026.0}),  # a range between quarters, not negative 2
    ("2025–2026", {2025.0, 2026.0}),          # en dash between numbers is a range
    ("74–75%", {74.0, 75.0}),
    ("1,660 - 2,000", {1660.0, 2000.0}),      # a spaced dash is not attached to the number
    ("$27,470K and 97.1%", {27470.0, 97.1}),  # plain numbers still work
])
def test_numbers_in_reads_signs(text, expected):
    assert numbers_in(text) == expected


def summary_saying(text):
    """A BoardSummary whose only numbers are in `text` (the headline)."""
    point = Point(title="Title", detail="Detail.")
    return BoardSummary(headline=text, wins=[point] * 3, risks=[point] * 3, questions=["Why?"] * 3)


PAYLOAD_TEXT = '{"Net new ARR vs budget": "-19.0%", "Rule of 40": "-12.2%"}'


@pytest.mark.parametrize("headline, ungrounded", [
    ("Net new ARR vs budget was -19.0%.", []),
    ("Net new ARR vs budget was −19.0%.", []),        # typographic minus reads the same
    ("Net new ARR missed budget by 19.0%.", [19.0]),  # the sign was dropped: old code let this through
    ("Rule of 40 is 12.2%.", [12.2]),
])
def test_positive_number_is_not_grounded_by_a_negative_one(headline, ungrounded):
    assert find_ungrounded_numbers(summary_saying(headline), PAYLOAD_TEXT) == ungrounded


# ---------------------------------------------------------------------------
# The payload uses the shared labels and reasons (decisions A, C, I, J)
# ---------------------------------------------------------------------------

@pytest.fixture(scope="module")
def northwind():
    return clean_workbook(NORTHWIND)


def test_payload_says_data_missing_or_no_prior_period(northwind):
    actuals, next_budget = northwind
    trends = build_payload("Northwind", actuals, next_budget, load_config())["trends_by_quarter"]
    assert trends["ARR growth YoY"]["Q3 2024"] == "n/a (no prior period)"  # nothing a year earlier
    assert trends["ARR growth YoY"]["Q1 2025"] == "data missing"           # the blank quarter
    assert trends["ARR growth YoY"]["Q1 2026"] == "data missing"           # a year after the blank
    assert trends["Net burn ($K)"]["Q1 2025"] == "data missing"


def test_payload_flag_names_are_the_metric_labels(northwind):
    actuals, next_budget = northwind
    flags = build_payload("Northwind", actuals, next_budget, load_config())["flags_latest_quarter"]
    assert [flag["flag"] for flag in flags] == [
        "NRR (annualized)", "GRR (annualized)", "Burn multiple", "Net burn vs budget",
        "Runway at current burn", "CAC payback", "Net new ARR vs budget", "Rule of 40",
        "NRR falling while pipeline rising"]


def test_payload_not_meaningful_flag_shows_the_k_figures(northwind):
    # Decision C: Northwind Q2 2026 net burn is 3,900; with a budgeted burn of 0 the % means nothing.
    actuals, next_budget = northwind
    actuals = actuals.copy()
    actuals.loc["Q2 2026", "budget_net_burn"] = 0
    payload = build_payload("Northwind", actuals, next_budget, load_config())
    flag = next(f for f in payload["flags_latest_quarter"] if f["flag"] == "Net burn vs budget")
    assert flag["status"] == "cannot evaluate — not meaningful"
    assert flag["value"] == "n/m: net burn 3,900 vs budget 0 ($K)"
    assert payload["flags_tripped"] == 5 and payload["flags_cannot_evaluate"] == 1
    # Not meaningful is not missing data. (Q1 2025 is still a gap: Northwind's blank quarter.)
    assert payload["data_gaps"]["Net burn vs budget"] == ["Q1 2025"]
    assert "flag: Net burn vs budget" not in payload["data_gaps"]


def test_save_analysis_passed_answer_loads_on_the_deck(northwind, tmp_path):
    # analyze.py and main.py save through save_analysis, so build_deck.py can read what either wrote.
    actuals, next_budget = northwind
    payload = build_payload("Northwind", actuals, next_budget, load_config())
    point = {"title": "Steady base", "detail": "Customers stayed."}
    answer = BoardSummary.model_validate({"headline": "Retention is the question.", "wins": [point] * 3,
                                          "risks": [point] * 3, "questions": ["Why?", "Where?", "How?"]})
    path = save_analysis(tmp_path / "sub" / "northwind_analysis.json", payload, answer, {"passed": True})
    saved = json.loads(path.read_text())
    assert saved == {"summary": answer.model_dump(), "run_info": {"passed": True}, "error": None, "payload": payload}
    assert load_analysis(path, payload) == (answer, None)


def test_save_analysis_failure_keeps_the_reason_and_gives_the_placeholder(northwind, tmp_path):
    actuals, next_budget = northwind
    payload = build_payload("Northwind", actuals, next_budget, load_config())
    path = save_analysis(tmp_path / "northwind_analysis.json", payload, error="AnthropicError: outage")
    saved = json.loads(path.read_text())
    assert saved["summary"] is None and saved["run_info"] is None and saved["error"] == "AnthropicError: outage"
    summary, why = load_analysis(path, payload)
    assert summary is None and "failed validation when it was made" in why
