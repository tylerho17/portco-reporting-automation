"""Unit tests for build_deck.py: slide text, the AI placeholder, status colors, footer, and loud failures.

Small made-up tables stand in for a workbook. Expected text is written out by hand.
Run from the project folder:  pytest
"""

import ast
import datetime
import json
import math
from pathlib import Path

import pandas as pd
import pytest

from analyze import build_payload
from build_deck import (FICTIONAL_NOTE, PLACEHOLDER_TEXT, build_presentation, collect_deck_data, deck_path,
                        flag_count_text, gaps_text, load_analysis, threshold_text)
from clean import STANDARD_COLUMNS
from metrics import CANNOT_EVALUATE, MISSING_INPUT, PASS, TRIP
from text_fit import TextDoesNotFitError

NAN = math.nan
PROJECT_DIR = Path(__file__).parent.parent
RUN_DATE = datetime.date(2026, 9, 17)

# Thresholds written out here, so tuning config.yaml never breaks these tests.
TEST_CONFIG = {
    "nrr_min": 1.00, "grr_min": 0.85, "burn_multiple_max": 2.0, "burn_over_budget_max": 0.15,
    "runway_min_months": 12, "cac_payback_max_months": 24, "net_new_arr_vs_budget_min": -0.20,
    "rule_of_40_min": 0.40, "nrr_falling_pipeline_rising_flag": True, "combo_lookback_quarters": 3,
    "combo_min_nrr_drop": 0.01,
}

# Colors from excel_output.STATUS_COLORS (fill): red, green, gray.
RED, GREEN, GRAY = "FFC7CE", "C6EFCE", "D9D9D9"


def flag(name, status, reason=None, metric="nrr", threshold=1.0):
    return {"flag": name, "metric": metric, "quarter": "Q2 2026", "value": 0.97, "threshold": threshold,
            "status": status, "reason": reason}


def three_quarters(blank=None):
    """Three complete quarters, every input 100 (ending cash 1200), with one quarter optionally blank."""
    actuals = pd.DataFrame(100.0, index=["Q4 2025", "Q1 2026", "Q2 2026"], columns=STANDARD_COLUMNS)
    actuals["ending_cash"] = 1200.0
    if blank:
        actuals.loc[blank] = NAN
    return actuals


def deck_data(company="Testco", blank=None):
    return collect_deck_data(company, "testco.xlsx", three_quarters(blank), None, TEST_CONFIG)


def summary_dict():
    """A valid summary with no numbers in it, so the number check has nothing to reject."""
    point = {"title": "Steady base", "detail": "Customers stayed."}
    return {"headline": "Retention is the main question for the board.",
            "wins": [point] * 3, "risks": [point] * 3,
            "questions": ["What drives churn?", "Where is pipeline coming from?", "How is hiring going?"]}


def shape(slide, name):
    matches = [s for s in slide.shapes if s.name == name]
    assert len(matches) == 1, f"expected one shape named {name!r}, found {len(matches)}"
    return matches[0]


# ---------------------------------------------------------------------------
# Text pieces
# ---------------------------------------------------------------------------

def test_flag_count_text():
    flags = [flag("a", TRIP)] * 6 + [flag("b", PASS)] * 3
    assert flag_count_text(flags) == "6 of 9 flags tripped"


def test_flag_count_text_mentions_flags_that_cannot_be_evaluated():
    flags = [flag("a", TRIP)] * 7 + [flag("b", PASS), flag("c", CANNOT_EVALUATE, MISSING_INPUT)]
    assert flag_count_text(flags) == "7 of 9 flags tripped, 1 cannot evaluate"


def test_threshold_text_says_which_way_the_flag_trips():
    assert threshold_text(flag("NRR (annualized)", TRIP, metric="nrr", threshold=1.0)) == "trips below 100.0%"
    assert threshold_text(flag("Burn multiple", TRIP, metric="burn_multiple", threshold=2.0)) == "trips above 2.00x"


def test_gaps_text_none():
    assert gaps_text({}) == "None — every metric and flag has the data it needs"


def test_gaps_text_groups_metrics_by_the_quarters_they_miss():
    gaps = {"ending_arr": ["Q1 2025"], "nrr": ["Q1 2025"], "arr_qoq": ["Q1 2025", "Q2 2025"],
            "flag: Rule of 40": ["Q2 2026"]}
    assert gaps_text(gaps) == ("Q1 2025: Ending ARR ($K), NRR (annualized); "
                               "Q1 2025 + Q2 2025: ARR growth QoQ; "
                               "Q2 2026: Flag: Rule of 40")


# ---------------------------------------------------------------------------
# Loading the AI analysis: anything not valid today shows the placeholder
# ---------------------------------------------------------------------------

@pytest.fixture
def payload():
    return build_payload("Testco", three_quarters(), None, TEST_CONFIG)


def write_analysis(tmp_path, summary, company="Testco", quarter="Q2 2026"):
    path = tmp_path / "testco_analysis.json"
    path.write_text(json.dumps({"summary": summary, "payload": {"company": company, "latest_quarter": quarter}}))
    return path


def test_valid_analysis_loads(tmp_path, payload):
    summary, reason = load_analysis(write_analysis(tmp_path, summary_dict()), payload)
    assert reason is None and summary.headline == summary_dict()["headline"]


def test_missing_analysis_file(tmp_path, payload):
    summary, reason = load_analysis(tmp_path / "nothing.json", payload)
    assert summary is None and "no analysis file" in reason


def test_analysis_that_failed_validation_when_made(tmp_path, payload):
    # analyze.py saves "summary": null when both attempts failed.
    summary, reason = load_analysis(write_analysis(tmp_path, None), payload)
    assert summary is None and "failed validation" in reason


def test_analysis_that_is_not_json(tmp_path, payload):
    path = tmp_path / "bad.json"
    path.write_text("{not json")
    summary, reason = load_analysis(path, payload)
    assert summary is None and "not valid JSON" in reason


def test_analysis_with_the_wrong_shape(tmp_path, payload):
    broken = summary_dict()
    del broken["questions"]
    summary, reason = load_analysis(write_analysis(tmp_path, broken), payload)
    assert summary is None and "shape" in reason


def test_analysis_for_another_quarter_is_stale(tmp_path, payload):
    summary, reason = load_analysis(write_analysis(tmp_path, summary_dict(), quarter="Q1 2026"), payload)
    assert summary is None and "Q1 2026" in reason


def test_analysis_for_another_company(tmp_path, payload):
    summary, reason = load_analysis(write_analysis(tmp_path, summary_dict(), company="Othercorp"), payload)
    assert summary is None and "Othercorp" in reason


def test_analysis_with_a_number_not_in_todays_data(tmp_path, payload):
    invented = summary_dict()
    invented["headline"] = "ARR grew 12345% this quarter."
    summary, reason = load_analysis(write_analysis(tmp_path, invented), payload)
    assert summary is None and "12345" in reason


def test_analysis_with_two_questions(tmp_path, payload):
    short = summary_dict()
    short["questions"] = short["questions"][:2]
    summary, reason = load_analysis(write_analysis(tmp_path, short), payload)
    assert summary is None and "exactly 3" in reason


# ---------------------------------------------------------------------------
# The whole deck
# ---------------------------------------------------------------------------

def all_text(slide):
    return " ".join(s.text_frame.text for s in slide.shapes if s.has_text_frame)


def build(tmp_path, summary=None, **data_options):
    from analyze import BoardSummary
    parsed = BoardSummary.model_validate(summary) if summary else None
    return build_presentation(deck_data(**data_options), parsed, RUN_DATE, tmp_path)


def test_five_slides_in_order(tmp_path):
    slides = build(tmp_path).slides
    assert len(slides) == 5
    assert [slide.shapes.title.text.split(":")[0].split(" — ")[0] for slide in slides] == [
        "Testco", "Key metrics", "ARR and cash", "Risks and flags", "Questions for management"]


def test_placeholder_on_slides_1_and_5_when_there_is_no_analysis(tmp_path):
    slides = build(tmp_path).slides
    assert shape(slides[0], "Headline").text_frame.text == PLACEHOLDER_TEXT
    assert PLACEHOLDER_TEXT in shape(slides[4], "Questions").text_frame.text
    # The Python numbers are still there: the flag count doesn't depend on the AI.
    assert "flags tripped" in shape(slides[0], "Flag count").text_frame.text


def test_analysis_text_on_slides_1_and_5(tmp_path):
    slides = build(tmp_path, summary=summary_dict()).slides
    assert shape(slides[0], "Headline").text_frame.text == summary_dict()["headline"]
    assert "Steady base" in shape(slides[0], "Wins").text_frame.text
    assert "Steady base" in shape(slides[0], "Risks").text_frame.text
    questions = shape(slides[4], "Questions").text_frame.text
    assert all(question in questions for question in summary_dict()["questions"])
    assert PLACEHOLDER_TEXT not in all_text(slides[0]) + all_text(slides[4])


def test_footer_on_every_slide(tmp_path):
    for slide in build(tmp_path).slides:
        footer = shape(slide, "Footer").text_frame.text
        assert FICTIONAL_NOTE in footer and "testco.xlsx" in footer and "2026-09-17" in footer


def status_fills(table):
    """{row label: status cell fill hex} for rows that are flags."""
    fills = {}
    for row in list(table.rows)[1:]:
        cell = row.cells[len(row.cells) - 1]
        if cell.fill.type is not None and cell.text != "—":
            fills[row.cells[0].text] = str(cell.fill.fore_color.rgb)
    return fills


def test_status_cells_are_red_green_or_gray(tmp_path):
    # Every input 100: NRR = 1 + 4 * (100 - 100 - 100) / 100 = -300% -> trips.
    # GRR = 1 - 4 * 200 / 100 = -700% -> trips. Net new ARR = 0 while burning -> burn multiple ∞ -> trips.
    # Runway = 1200 / (100 / 3) = 36 months -> passes. 3 quarters: Rule of 40 needs 4 back -> can't evaluate.
    table = shape(build(tmp_path).slides[1], "KPI table").table
    fills = status_fills(table)
    assert fills["NRR (annualized)"] == RED
    assert fills["Burn multiple"] == RED
    assert fills["Runway at current burn"] == GREEN
    assert fills["Rule of 40"] == GRAY


def test_kpi_table_shows_why_a_value_is_missing(tmp_path):
    # Q1 2026 blank: the prior-quarter column says "data missing"; Rule of 40 says no prior period.
    table = shape(build(tmp_path, blank="Q1 2026").slides[1], "KPI table").table
    rows = {row.cells[0].text: [cell.text for cell in row.cells] for row in table.rows}
    assert rows["NRR (annualized)"][2] == "data missing"
    assert rows["Rule of 40"][1] == "n/a (no prior period)"
    assert rows["Burn multiple"][1] == "∞ (ARR shrank)"
    assert rows["Rule of 40"][4] == "Cannot evaluate — no prior period"


def test_risks_slide_lists_tripped_flags_combo_and_data_gaps(tmp_path):
    slide = build(tmp_path, blank="Q1 2026").slides[3]
    text = shape(slide, "Risks and flags").text_frame.text
    assert "NRR (annualized): -300.0% (trips below 100.0%)" in text
    assert "NRR falling while pipeline rising: Cannot evaluate — missing input" in text
    gaps = shape(slide, "Data gaps").text_frame.text
    assert gaps.startswith("Data gaps") and "Q1 2026 + Q2 2026: ARR growth QoQ" in gaps


def test_charts_slide_has_two_pictures(tmp_path):
    slide = build(tmp_path).slides[2]
    assert shape(slide, "ARR chart").shape_type == 13   # MSO_SHAPE_TYPE.PICTURE
    assert shape(slide, "Cash chart").shape_type == 13


def test_text_too_long_for_its_box_fails_loudly(tmp_path):
    with pytest.raises(TextDoesNotFitError, match="Slide 1"):
        build(tmp_path, company="Testco " * 60)


def test_deck_path():
    assert deck_path("data/northwind.xlsx", "output") == Path("output/northwind_board_pack.pptx")


# ---------------------------------------------------------------------------
# No number typed by hand
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("file_name", ["build_deck.py", "charts.py"])
def test_no_digit_in_any_text_written_in_the_code(file_name):
    # Every number a reader sees comes from metrics, flags, config or the analysis - never from a
    # string typed into the code. (Layout sizes like 0.2 inches are code numbers, not text.)
    tree = ast.parse((PROJECT_DIR / file_name).read_text())
    docstrings = {id(node.body[0].value) for node in ast.walk(tree)
                  if isinstance(node, (ast.Module, ast.FunctionDef, ast.ClassDef)) and node.body
                  and isinstance(node.body[0], ast.Expr) and isinstance(node.body[0].value, ast.Constant)}
    typed = [node.value for node in ast.walk(tree)
             if isinstance(node, ast.Constant) and isinstance(node.value, str) and id(node) not in docstrings]
    with_digits = [text for text in typed if any(character.isdigit() for character in text)]
    assert with_digits == []
