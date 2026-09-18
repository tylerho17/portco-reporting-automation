"""Smoke tests for app.py (the Streamlit app): its helper functions, run without a browser.

Every build goes into a temporary folder, and a guard makes creating a real Anthropic client
fail, so no test can reach the API. The saved analysis the app may reuse is written by
analyze.save_analysis into a temporary folder too, never read from the real output/.
Run from the project folder:  pytest
"""

import io
import json
from pathlib import Path
from types import SimpleNamespace

import pytest
from openpyxl import Workbook, load_workbook
from pptx import Presentation

import analyze
import app
from analyze import BoardSummary, build_payload, save_analysis
from clean import clean_workbook
from excel_output import SHEET_NAMES, STATUS_COLORS
from metrics import CANNOT_EVALUATE, PASS, TRIP, load_config

PROJECT_DIR = Path(__file__).parent.parent
NORTHWIND = PROJECT_DIR / "data" / "northwind.xlsx"


class FakeClient:
    """Stands in for anthropic.Anthropic() (same as tests/test_main.py): one fixed answer, calls counted."""

    def __init__(self, summary=None, error=None):
        self.summary, self.error, self.calls = summary, error, 0
        self.messages = self  # analyze.py calls client.messages.parse(...)

    def parse(self, **kwargs):
        self.calls += 1
        if self.error:
            raise self.error
        return SimpleNamespace(content=[SimpleNamespace(type="text", text=self.summary.model_dump_json())],
                               parsed_output=self.summary, stop_reason="end_turn",
                               usage=SimpleNamespace(input_tokens=100, output_tokens=50))


def summary(headline):
    """A valid answer with no numbers in it, so the number check has nothing to reject."""
    point = {"title": "Steady base", "detail": "Customers stayed."}
    return BoardSummary.model_validate({
        "headline": headline, "wins": [point] * 3, "risks": [point] * 3,
        "questions": ["What drives churn?", "Where is pipeline coming from?", "How is hiring going?"]})


@pytest.fixture(autouse=True)
def no_real_client(monkeypatch):
    """Creating a real Anthropic client fails the test, so no test here can reach the API."""
    def refuse(*args, **kwargs):
        raise AssertionError("a test tried to create a real Anthropic client")
    monkeypatch.setattr(analyze.anthropic, "Anthropic", refuse)


def build_northwind(tmp_path, include_ai=False, client=None, file_name="northwind.xlsx"):
    """build_outputs on the Northwind workbook's bytes, as if it had been dragged in."""
    return app.build_outputs(file_name, NORTHWIND.read_bytes(), include_ai,
                             saved_dir=tmp_path / "saved", client=client)


def save_northwind_analysis(folder, answer):
    """A saved analysis for today's Northwind numbers, where main.py would have put it."""
    actuals, next_budget = clean_workbook(NORTHWIND)
    payload = build_payload("Northwind", actuals, next_budget, load_config())
    return save_analysis(Path(folder) / "northwind_analysis.json", payload, answer, {"model": "claude-sonnet-5"})


def headline_in(deck_bytes):
    """The Headline box on slide 4 (AI commentary)."""
    slide = Presentation(io.BytesIO(deck_bytes)).slides[-1]
    return next(shape.text_frame.text for shape in slide.shapes if shape.name == "Headline")


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------

def test_the_ai_checkbox_label_names_the_typical_cost():
    label = app.ai_checkbox_label()
    assert "AI commentary" in label and "$0.09" in label


def test_an_upload_is_saved_under_its_own_name_only(tmp_path):
    path = app.save_upload("../../elsewhere/Northwind.xlsx", b"abc", tmp_path)
    assert path == tmp_path / "Northwind.xlsx" and path.read_bytes() == b"abc"


def test_status_colors_are_the_excel_ones():
    for status in (TRIP, PASS, CANNOT_EVALUATE):
        fill, text = STATUS_COLORS[status]
        assert app.status_css(status) == f"background-color: #{fill}; color: #{text}"


# ---------------------------------------------------------------------------
# Errors: a plain message, never a traceback
# ---------------------------------------------------------------------------

def test_a_clean_py_error_is_shown_word_for_word(tmp_path):
    book = Workbook()
    book.active.append(["Quarter", "ARR"])       # a table clean.py can't read with certainty
    book.active.append(["Q1 2025", 100])
    buffer = io.BytesIO()
    book.save(buffer)
    with pytest.raises(ValueError) as expected:  # what clean.py itself says about this file
        clean_workbook(app.save_upload("bad.xlsx", buffer.getvalue(), tmp_path))
    result = app.build_outputs("bad.xlsx", buffer.getvalue(), False, saved_dir=tmp_path)
    assert result["error"] == str(expected.value)


def test_a_file_that_is_not_a_workbook_gets_a_plain_message(tmp_path):
    result = app.build_outputs("notes.xlsx", b"just some text", False, saved_dir=tmp_path)
    assert result["error"] == app.NOT_A_WORKBOOK
    assert "Traceback" not in result["error"]


def test_an_unexpected_error_is_named_without_a_traceback(tmp_path, monkeypatch):
    def broken(*args, **kwargs):
        raise KeyError("oops")
    monkeypatch.setattr(app, "collect_deck_data", broken)
    result = build_northwind(tmp_path)
    assert result["error"].startswith(app.UNEXPECTED_ERROR_START) and "KeyError" in result["error"]
    assert "Traceback" not in result["error"]


# ---------------------------------------------------------------------------
# A good workbook, AI off
# ---------------------------------------------------------------------------

def test_northwind_builds_a_deck_and_a_metrics_workbook(tmp_path):
    result = build_northwind(tmp_path)
    assert result["error"] is None
    assert result["company"] == "Northwind" and result["quarter"] == "Q2 2026"
    assert result["deck_name"] == "northwind_board_pack.pptx"
    assert result["excel_name"] == "northwind_metrics.xlsx"
    assert len(Presentation(io.BytesIO(result["deck_bytes"])).slides) == 4
    assert load_workbook(io.BytesIO(result["excel_bytes"])).sheetnames == SHEET_NAMES
    assert headline_in(result["deck_bytes"]) == app.PLACEHOLDER_TEXT
    assert result["ai_note"] == app.AI_OFF_NOTE


def test_the_flag_table_matches_northwinds_story(tmp_path):
    result = build_northwind(tmp_path)
    flags, colors = result["flags"], result["flags_colors"]
    assert list(flags.columns) == app.FLAG_COLUMNS and len(flags) == 9
    assert result["flag_count"] == "6 of 9 flags tripped"
    tripped_rows = flags["Status"] == "Tripped"
    assert tripped_rows.sum() == 6
    assert (colors[tripped_rows] == app.status_css(TRIP)).all().all()      # whole row red
    assert (colors[~tripped_rows] == app.status_css(PASS)).all().all()     # the rest green


def test_the_metrics_table_is_colored_like_the_excel_sheet(tmp_path):
    result = build_northwind(tmp_path)
    table, colors = result["metrics"], result["metrics_colors"]
    assert table.shape == colors.shape
    assert list(table.columns)[-1] == "Q2 2026"
    assert table.loc["NRR (annualized)", "Q1 2025"] == "data missing"                  # the blank quarter
    assert colors.loc["NRR (annualized)", "Q1 2025"] == app.status_css(CANNOT_EVALUATE)  # gray
    assert colors.loc["NRR (annualized)", "Q2 2026"] == app.status_css(TRIP)             # below 100%: red
    assert colors.loc["Ending ARR ($K)", "Q2 2026"] == ""                                # no flag: no color


def test_a_long_file_name_still_gets_a_deck(tmp_path):
    # People upload files named anything. Before the fix, this name made the footer too wide and
    # the page showed "Slide 1, Footer: text doesn't fit" instead of a deck.
    result = build_northwind(tmp_path, file_name="Northwind - Q2 2026 KPI pack (final version for board).xlsx")
    assert result["error"] is None
    assert len(Presentation(io.BytesIO(result["deck_bytes"])).slides) == 4


def test_the_data_gaps_lines_name_the_blank_quarter(tmp_path):
    result = build_northwind(tmp_path)
    assert any(line.startswith("Q1 2025") for line in result["gaps"])


# ---------------------------------------------------------------------------
# AI on
# ---------------------------------------------------------------------------

def test_a_saved_analysis_for_the_same_numbers_is_reused_without_an_api_call(tmp_path):
    save_northwind_analysis(tmp_path / "saved", summary("Reused headline."))
    client = FakeClient(summary("Fresh headline."))
    result = build_northwind(tmp_path, include_ai=True, client=client)
    assert client.calls == 0
    assert headline_in(result["deck_bytes"]) == "Reused headline."
    assert result["ai_note"].startswith(app.AI_REUSED_START)


def test_a_saved_analysis_for_other_numbers_is_not_reused(tmp_path):
    path = save_northwind_analysis(tmp_path / "saved", summary("Old headline."))
    saved = json.loads(path.read_text())
    saved["payload"]["company"] = "Northwind"                   # same company and quarter...
    saved["payload"]["flags"] = []                              # ...but different facts
    path.write_text(json.dumps(saved))
    assert app.reusable_analysis(NORTHWIND, tmp_path / "saved", load_config()) is None


def test_with_no_saved_analysis_claude_is_asked(tmp_path):
    client = FakeClient(summary("Fresh headline."))
    result = build_northwind(tmp_path, include_ai=True, client=client)
    assert client.calls == 1
    assert headline_in(result["deck_bytes"]) == "Fresh headline."
    assert result["ai_note"] == app.AI_NEW_NOTE


def test_no_api_key_still_builds_the_deck_with_the_placeholder(tmp_path, monkeypatch):
    monkeypatch.setattr(app, "api_key_problem", lambda: "no key here")
    result = build_northwind(tmp_path, include_ai=True)
    assert result["error"] is None
    assert headline_in(result["deck_bytes"]) == app.PLACEHOLDER_TEXT
    assert "no key here" in result["ai_note"]


def test_an_ai_failure_still_builds_the_deck_and_says_why(tmp_path):
    client = FakeClient(error=analyze.anthropic.AnthropicError("simulated outage"))
    result = build_northwind(tmp_path, include_ai=True, client=client)
    assert result["error"] is None
    assert headline_in(result["deck_bytes"]) == app.PLACEHOLDER_TEXT
    assert result["ai_note"].startswith(app.PLACEHOLDER_TEXT)


# ---------------------------------------------------------------------------
# The page itself
# ---------------------------------------------------------------------------

def test_the_page_loads_without_an_error():
    from streamlit.testing.v1 import AppTest
    page = AppTest.from_file(str(PROJECT_DIR / "app.py")).run()
    assert not page.exception
    assert page.checkbox[0].label == app.ai_checkbox_label()


def test_the_mac_launcher_is_double_clickable_and_starts_this_app():
    launcher = PROJECT_DIR / "run_app.command"
    assert launcher.stat().st_mode & 0o111                     # executable, so Finder will run it
    assert "streamlit run app.py" in launcher.read_text()
    assert "streamlit" in (PROJECT_DIR / "requirements.txt").read_text().split()


def test_the_launcher_installs_packages_when_the_venv_has_no_streamlit():
    # A .venv made before the web page existed has Python but no streamlit. Checking only for
    # .venv/bin/python skipped the install, and the double-click ended in "streamlit: command not found".
    assert "[ ! -x .venv/bin/streamlit ]" in (PROJECT_DIR / "run_app.command").read_text()


def render_northwind():
    """Draws Northwind's results the way the page does after an upload (AppTest can't upload files)."""
    import app
    app.show_result(app.build_outputs("northwind.xlsx", app.Path("data/northwind.xlsx").read_bytes(), False))


def render_bad_file():
    import app
    app.show_result(app.build_outputs("notes.xlsx", b"just some text", False))


def test_the_results_draw_two_colored_tables_and_two_downloads(monkeypatch):
    from streamlit.testing.v1 import AppTest
    monkeypatch.chdir(PROJECT_DIR)
    page = AppTest.from_function(render_northwind).run(timeout=60)
    assert not page.exception and not page.error
    assert len(page.dataframe) == 2                          # flags, then metrics
    assert page.subheader[0].value == "6 of 9 flags tripped"
    assert len(page.get("download_button")) == 2


def test_a_bad_file_draws_one_plain_error_and_nothing_else():
    from streamlit.testing.v1 import AppTest
    page = AppTest.from_function(render_bad_file).run()
    assert not page.exception
    assert [error.value for error in page.error] == [app.NOT_A_WORKBOOK]
    assert len(page.dataframe) == 0
