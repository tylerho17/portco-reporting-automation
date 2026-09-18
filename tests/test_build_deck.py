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
from pptx import Presentation
from pptx.util import Emu

from analyze import build_payload
from build_deck import (AI_DRAFTED_LINE, FICTIONAL_NOTE, NO_AI_MODEL, PLACEHOLDER_NOTE, PLACEHOLDER_TEXT,
                        build_presentation, collect_deck_data, deck_path, flag_count_text, gaps_text, load_analysis,
                        save_deck, shorten_middle, threshold_text)
from clean import STANDARD_COLUMNS
from metrics import CANNOT_EVALUATE, MISSING_INPUT, PASS, TRIP
from provenance import (NOT_REVIEWED, build_manifest, file_sha256, git_commit, manifest_path, read_manifest,
                        save_manifest)
from text_fit import MIN_FONT_PT, TextDoesNotFitError, text_width_pt

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


def test_analysis_too_long_for_the_slides_gives_the_placeholder(tmp_path, payload):
    # Decision K: the numbers are valid, so the deck is still built - it just can't carry this text
    # (review finding 1a). Before this, the build stopped and the company had no deck at all.
    too_long = summary_dict()
    too_long["risks"] = [{"title": "Steady base", "detail": "customers " * 45}] * 3
    summary, reason = load_analysis(write_analysis(tmp_path, too_long), payload)
    assert summary is None and "does not fit slide 4" in reason


def test_long_wins_do_not_cost_the_deck_its_ai_text(tmp_path, payload):
    # Wins aren't on the deck any more (Task 2), so their length can't make the text "not fit".
    long_wins = summary_dict()
    long_wins["wins"] = [{"title": "Steady base", "detail": "customers " * 45}] * 3
    summary, reason = load_analysis(write_analysis(tmp_path, long_wins), payload)
    assert reason is None and summary is not None


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


def build(tmp_path, summary=None, approval=None, model=None, draft=False, **data_options):
    from analyze import BoardSummary
    parsed = BoardSummary.model_validate(summary) if summary else None
    return build_presentation(deck_data(**data_options), parsed, RUN_DATE, tmp_path, approval=approval, model=model,
                              draft=draft)


def test_four_slides_in_order(tmp_path):
    slides = build(tmp_path).slides
    assert [slide.shapes.title.text for slide in slides] == [
        "Testco: key metrics — Q2 2026 vs Q1 2026", "ARR and cash — Q4 2025 to Q2 2026",
        "Risks and flags — Q2 2026", "AI commentary — Q2 2026"]


def test_placeholder_on_slide_4_when_there_is_no_analysis(tmp_path):
    slide = build(tmp_path).slides[3]
    assert shape(slide, "Headline").text_frame.text == PLACEHOLDER_TEXT
    assert shape(slide, "AI note").text_frame.text == PLACEHOLDER_NOTE
    # Nothing on the slide is AI-drafted, so the "AI-drafted ... review before use" line isn't shown.
    assert AI_DRAFTED_LINE not in all_text(slide)


def test_analysis_text_on_slide_4(tmp_path):
    slide = build(tmp_path, summary=summary_dict()).slides[3]
    assert shape(slide, "Headline").text_frame.text == summary_dict()["headline"]
    assert "Steady base" in shape(slide, "Risks").text_frame.text
    questions = shape(slide, "Questions").text_frame.text
    assert all(question in questions for question in summary_dict()["questions"])
    assert PLACEHOLDER_TEXT not in all_text(slide)


def test_the_ai_line_sits_under_the_title_and_above_the_headline(tmp_path):
    slide = build(tmp_path, summary=summary_dict()).slides[3]
    line = shape(slide, "AI-drafted line")
    assert line.text_frame.text == "AI-drafted from computed metrics - review before use"
    assert slide.shapes.title.top + slide.shapes.title.height <= line.top
    assert line.top + line.height <= shape(slide, "Headline").top


def test_wins_are_not_on_the_deck(tmp_path):
    summary = summary_dict()
    summary["wins"] = [{"title": "A win nobody sees", "detail": "Customers stayed."}] * 3
    slides = build(tmp_path, summary=summary).slides
    assert all("A win nobody sees" not in all_text(slide) and "Wins" not in all_text(slide) for slide in slides)


def test_the_flag_count_is_on_the_risks_slide(tmp_path):
    # It was on the Summary slide, which is gone. Python counts it, so it shows with or without AI text.
    text = shape(build(tmp_path).slides[2], "Risks and flags").text_frame.text
    assert "Tripped flags (" in text and "flags tripped" in text


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
    table = shape(build(tmp_path).slides[0], "KPI table").table
    fills = status_fills(table)
    assert fills["NRR (annualized)"] == RED
    assert fills["Burn multiple"] == RED
    assert fills["Runway at current burn"] == GREEN
    assert fills["Rule of 40"] == GRAY


def test_kpi_table_shows_why_a_value_is_missing(tmp_path):
    # Q1 2026 blank: the prior-quarter column says "data missing"; Rule of 40 says no prior period.
    table = shape(build(tmp_path, blank="Q1 2026").slides[0], "KPI table").table
    rows = {row.cells[0].text: [cell.text for cell in row.cells] for row in table.rows}
    assert rows["NRR (annualized)"][2] == "data missing"
    assert rows["Rule of 40"][1] == "n/a (no prior period)"
    assert rows["Burn multiple"][1] == "∞ (ARR shrank)"
    assert rows["Rule of 40"][4] == "Cannot evaluate — no prior period"


def test_risks_slide_lists_tripped_flags_combo_and_data_gaps(tmp_path):
    slide = build(tmp_path, blank="Q1 2026").slides[2]
    text = shape(slide, "Risks and flags").text_frame.text
    assert "NRR (annualized): -300.0% (trips below 100.0%)" in text
    assert "NRR falling while pipeline rising: Cannot evaluate — missing input" in text
    gaps = shape(slide, "Data gaps").text_frame.text
    assert gaps.startswith("Data gaps") and "Q1 2026 + Q2 2026: ARR growth QoQ" in gaps


def run_sizes(text_shape):
    return {run.font.size.pt for item in text_shape.text_frame.paragraphs for run in item.runs}


def test_side_by_side_columns_use_the_same_font_sizes(tmp_path):
    # Risks are long, questions are short: the questions column must shrink along with the risks column.
    summary = summary_dict()
    summary["risks"] = [{"title": "Steady base", "detail": " ".join(["Customers stayed with the product."] * 5)}] * 3
    slide = build(tmp_path, summary=summary).slides[3]
    risks, questions = shape(slide, "Risks"), shape(slide, "Questions")
    assert max(run_sizes(risks)) < 18      # the risks heading had to shrink from 18 pt...
    assert run_sizes(risks) == run_sizes(questions)  # ...and the questions column shrank with it


def test_charts_slide_has_two_pictures(tmp_path):
    slide = build(tmp_path).slides[1]
    assert shape(slide, "ARR chart").shape_type == 13   # MSO_SHAPE_TYPE.PICTURE
    assert shape(slide, "Cash chart").shape_type == 13


def test_text_too_long_for_its_box_fails_loudly(tmp_path):
    with pytest.raises(TextDoesNotFitError, match="Slide 1"):
        build(tmp_path, company="Testco " * 60)


# ---------------------------------------------------------------------------
# Provenance and review status in the footer; the DRAFT watermark only with --draft
# ---------------------------------------------------------------------------

APPROVED = {"reviewer": "Tyler Ho", "approved_at": "2026-09-17T15:00:00",
            "input_sha256": "input-hash", "config_sha256": "config-hash"}


def test_the_footer_carries_the_commit_and_the_model(tmp_path):
    footer = shape(build(tmp_path, model="claude-sonnet-5").slides[0], "Footer").text_frame.text
    assert git_commit()["commit"] in footer
    assert "claude-sonnet-5" in footer
    assert FICTIONAL_NOTE in footer and "testco.xlsx" in footer and RUN_DATE.isoformat() in footer


def test_the_footer_says_when_there_is_no_ai_text(tmp_path):
    assert NO_AI_MODEL in shape(build(tmp_path).slides[0], "Footer").text_frame.text


def footer(presentation):
    return shape(presentation.slides[0], "Footer").text_frame.text


def watermarks(presentation):
    return [item.name for slide in presentation.slides for item in slide.shapes if item.name == "Watermark"]


def test_no_watermark_by_default(tmp_path):
    # The watermark is opt-in (--draft): the footer already says whether anyone has reviewed the deck.
    assert watermarks(build(tmp_path)) == []


def test_draft_watermarks_every_slide_of_an_unreviewed_deck(tmp_path):
    for slide in build(tmp_path, draft=True).slides:
        assert shape(slide, "Watermark").text_frame.text == NOT_REVIEWED


def test_the_watermark_is_drawn_over_the_content_not_under_it(tmp_path):
    # Slides 1 and 2 are covered by an opaque table and two chart images, so a watermark added
    # first would be invisible on exactly the slides that carry the numbers.
    for slide in build(tmp_path, draft=True).slides:
        assert slide.shapes[-1].name == "Watermark"


def test_draft_never_stamps_an_approved_deck(tmp_path):
    # The stamp says NOT REVIEWED; on a deck someone approved, that would be false.
    assert watermarks(build(tmp_path, approval=APPROVED, draft=True)) == []


def test_the_footer_says_an_unreviewed_deck_is_not_reviewed(tmp_path):
    assert footer(build(tmp_path, model="claude-sonnet-5")).endswith(" | claude-sonnet-5 | AI-drafted | not reviewed")


def test_the_footer_names_the_reviewer_and_the_day(tmp_path):
    # approved_at is to the second; the footer shows the day only.
    assert footer(build(tmp_path, approval=APPROVED)).endswith(" | AI-drafted | reviewed by Tyler Ho on 2026-09-17")


def test_the_footer_is_on_every_slide_with_the_same_review_status(tmp_path):
    texts = {shape(slide, "Footer").text_frame.text for slide in build(tmp_path, approval=APPROVED).slides}
    assert len(texts) == 1


def footer_room_pt(presentation):
    """Width the footer text has on one line: the footer box less its two side margins."""
    box = shape(presentation.slides[0], "Footer")
    return Emu(box.width - box.text_frame.margin_left - box.text_frame.margin_right).pt


def test_the_footer_fits_one_line_with_the_longest_file_name_and_a_reviewer(tmp_path):
    # Worst case today: the longest workbook name, the default model, and a reviewer (plus the "*"
    # for uncommitted code whenever these tests run mid-edit). Measured on width, not only height.
    presentation = build_presentation(
        collect_deck_data("Fernhollow", "fernhollow.xlsx", three_quarters(), None, TEST_CONFIG), None, RUN_DATE,
        tmp_path, approval=APPROVED, model="claude-sonnet-5")
    text = footer(presentation)
    assert "reviewed by Tyler Ho on 2026-09-17" in text
    assert text_width_pt(text, MIN_FONT_PT) <= footer_room_pt(presentation)


def test_a_reviewer_name_too_long_for_the_footer_leaves_the_name_to_the_manifest(tmp_path):
    # A name is typed by a person, so it can be any length. The footer keeps the review day on one
    # line; who reviewed stays in the manifest, which approve.py wrote.
    long_name = {**APPROVED, "reviewer": "Alexandra Richardson-Whitfield of the Portfolio Operations Team"}
    presentation = build(tmp_path, approval=long_name, model="claude-sonnet-5")
    text = footer(presentation)
    assert text.endswith(" | AI-drafted | reviewed on 2026-09-17")
    assert text_width_pt(text, MIN_FONT_PT) <= footer_room_pt(presentation)


def test_a_file_name_too_long_for_the_footer_is_shortened_not_a_failed_deck(tmp_path):
    # The web page takes whatever name the file has. Before the fix, this name stopped the whole
    # deck with "Slide 1, Footer: text doesn't fit". The footer keeps the start and the extension.
    long_file = "Northwind KPI workbook Q2 2026 final version for the board.xlsx"
    presentation = build_presentation(collect_deck_data("Testco", long_file, three_quarters(), None, TEST_CONFIG),
                                      None, RUN_DATE, tmp_path, approval=APPROVED, model="claude-sonnet-5")
    text = footer(presentation)
    assert text.startswith(f"{FICTIONAL_NOTE} | North") and "….xlsx |" in text
    assert text.endswith(" | AI-drafted | reviewed by Tyler Ho on 2026-09-17")
    assert text_width_pt(text, MIN_FONT_PT) <= footer_room_pt(presentation)


def test_a_file_name_that_fits_is_never_shortened():
    assert shorten_middle("fernhollow.xlsx", lambda text: True) == "fernhollow.xlsx"


def northwind_manifest(tmp_path, input_sha256=None, config_sha256=None):
    """A manifest for Northwind in tmp_path, approved against the hashes given (or today's)."""
    workbook, config = PROJECT_DIR / "data" / "northwind.xlsx", PROJECT_DIR / "config.yaml"
    approval = {"reviewer": "Tyler Ho", "approved_at": "2026-09-17T15:00:00",
                "input_sha256": input_sha256 or file_sha256(workbook),
                "config_sha256": config_sha256 or file_sha256(config)}
    manifest = build_manifest("Northwind", workbook, config, "northwind_board_pack.pptx", {},
                              ai_text=False, approval=approval)
    save_manifest(manifest_path(workbook, tmp_path), manifest)
    return workbook


def watermarks_in(path):
    return watermarks(Presentation(path))


def test_rebuilding_an_approved_deck_drops_the_watermark(tmp_path):
    workbook = northwind_manifest(tmp_path)
    path, _ = save_deck(workbook, TEST_CONFIG, None, run_date=RUN_DATE, output_dir=tmp_path, draft=True)
    assert watermarks_in(path) == []


def test_the_footer_reads_the_review_status_from_the_manifest(tmp_path):
    workbook = northwind_manifest(tmp_path)
    path, _ = save_deck(workbook, TEST_CONFIG, None, run_date=RUN_DATE, output_dir=tmp_path)
    assert footer(Presentation(path)).endswith(" | AI-drafted | reviewed by Tyler Ho on 2026-09-17")


def test_without_a_manifest_the_footer_says_not_reviewed_and_there_is_no_watermark(tmp_path):
    path, _ = save_deck(PROJECT_DIR / "data" / "northwind.xlsx", TEST_CONFIG, None, run_date=RUN_DATE,
                        output_dir=tmp_path)
    assert footer(Presentation(path)).endswith(" | AI-drafted | not reviewed")
    assert watermarks_in(path) == []


def test_a_rebuilt_deck_updates_what_the_manifest_says_about_it(tmp_path):
    # The manifest has to describe the deck on disk, or an audit trail is worse than none.
    workbook = northwind_manifest(tmp_path, input_sha256="the-hash-of-an-older-workbook")
    save_deck(workbook, TEST_CONFIG, None, run_date=RUN_DATE, output_dir=tmp_path)
    saved = read_manifest(manifest_path(workbook, tmp_path))
    assert saved["deck"]["status"] == NOT_REVIEWED      # it was "approved by Tyler Ho ..." before
    assert saved["approval"]["reviewer"] == "Tyler Ho"  # the record of the approval is kept


def test_a_workbook_changed_since_approval_goes_back_to_draft(tmp_path):
    # Nobody has reviewed a deck built from numbers that arrived after the approval.
    workbook = northwind_manifest(tmp_path, input_sha256="the-hash-of-an-older-workbook")
    path, _ = save_deck(workbook, TEST_CONFIG, None, run_date=RUN_DATE, output_dir=tmp_path, draft=True)
    assert len(watermarks_in(path)) == 4
    assert footer(Presentation(path)).endswith(" | AI-drafted | not reviewed")


def test_changed_thresholds_send_an_approved_deck_back_to_draft(tmp_path):
    workbook = northwind_manifest(tmp_path, config_sha256="the-hash-of-older-thresholds")
    path, _ = save_deck(workbook, TEST_CONFIG, None, run_date=RUN_DATE, output_dir=tmp_path, draft=True)
    assert len(watermarks_in(path)) == 4
    assert footer(Presentation(path)).endswith(" | AI-drafted | not reviewed")


def test_command_line_draft_flag(tmp_path):
    # build_deck.py --draft turns the watermark on; without it the deck has none.
    import build_deck
    workbook = str(PROJECT_DIR / "data" / "northwind.xlsx")
    build_deck.main([workbook, "--no-analysis", "--draft"], output_dir=tmp_path)
    assert len(watermarks_in(tmp_path / "northwind_board_pack.pptx")) == 4
    build_deck.main([workbook, "--no-analysis"], output_dir=tmp_path)
    assert watermarks_in(tmp_path / "northwind_board_pack.pptx") == []


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
