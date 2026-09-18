"""Task 16: every stop says what is wrong, where, and what to do next, in plain words.

The audit read every message raised in clean.py, metrics.py, analyze.py, build_deck.py, memo.py,
mapping.py and main.py (metrics.py and memo.py raise none: a value with no number gets a reason,
never a stop). These tests pin the wording of the ones that fell short, so a later edit can't quietly
drop the "what to do next" half. Each expected text is typed out by hand, not copied from a run.

No test calls the Anthropic API: the API errors are built by hand, and the one test that reaches
analyze.py's command line stops before any call.

Run from the project folder:  python -m pytest -q
"""

import sys
from types import SimpleNamespace

import anthropic
import httpx2 as httpx   # the HTTP library the anthropic SDK uses here (as in test_batch.py)
import pytest

import analyze
import build_deck
import main
import mapping
import run_log
from clean import clean_workbook
from run_log import error_text
from test_bad_inputs import SHEET, drop_column, error_from, good_table, write_workbook

API_REQUEST = httpx.Request("POST", "https://api.anthropic.com/v1/messages")


def status_error(error_class, status):
    """The error the SDK raises for an HTTP status (401, 429, 500, ...), as it would arrive."""
    return error_class("the API's own words", response=httpx.Response(status, request=API_REQUEST), body=None)


# ---------------------------------------------------------------------------
# clean.py: the workbook
# ---------------------------------------------------------------------------

def test_a_missing_workbook_names_the_file_and_says_to_check_the_path(tmp_path):
    with pytest.raises(FileNotFoundError) as caught:
        clean_workbook(tmp_path / "nope.xlsx")
    assert str(caught.value) == f"Can't find {tmp_path / 'nope.xlsx'}: check the file name and folder, then run again"


@pytest.mark.parametrize("content", [b"hello", b"PK\x03\x04 not really a zip"])
def test_a_file_that_is_not_an_excel_workbook_says_how_to_fix_it(tmp_path, content):
    # Before: pandas' "Excel file format cannot be determined, you must specify an engine manually".
    path = tmp_path / "notes.xlsx"
    path.write_bytes(content)
    with pytest.raises(ValueError) as caught:
        clean_workbook(path)
    assert str(caught.value) == ("notes.xlsx isn't a readable Excel workbook: open it in Excel, save it as an "
                                 "Excel Workbook (.xlsx), then run again")


def test_no_kpi_tab_lists_the_tabs_in_plain_words_and_says_what_to_add(tmp_path):
    message = error_from(tmp_path, drop_column(good_table(), "Quarter"))
    assert message == ("No tab in broken.xlsx has a 'Quarter' header in its first 10 rows (tabs checked: 'Notes', "
                       "'KPI Tracker'): add a 'Quarter' header above the column of quarter labels, like 'Q2 2026'")


def test_three_kpi_tabs_are_all_named_with_the_right_grammar(tmp_path):
    from openpyxl import load_workbook
    path = write_workbook(tmp_path / "tabs.xlsx", good_table())
    workbook = load_workbook(path)
    for title in ("KPI (old)", "KPI (older)"):
        workbook.copy_worksheet(workbook[SHEET]).title = title
    workbook.save(path)
    with pytest.raises(ValueError) as caught:
        clean_workbook(path)
    assert str(caught.value).startswith("Tabs 'KPI Tracker', 'KPI (old)' and 'KPI (older)' all have a 'Quarter' "
                                        "header - keep one KPI tab")


def test_a_note_under_the_table_says_to_move_it(tmp_path):
    rows = good_table() + [["Source: FP&A, unaudited"] + [None] * (len(good_table()[0]) - 1)]
    assert error_from(tmp_path, rows) == (
        "Sheet 'KPI Tracker', row 9: Can't read quarter label 'Source: FP&A, unaudited' (expected e.g. "
        "'Q2 2026') - write the label like that, or if the row is a note, move it to another tab")


def test_a_missing_column_says_to_add_it(tmp_path):
    assert error_from(tmp_path, drop_column(good_table(), "pipeline")) == (
        "Sheet 'KPI Tracker', row 3 (header) is missing columns: pipeline - add a column headed with each "
        "name (its cells can stay empty if there's no data)")


def test_a_budget_row_with_actuals_says_what_to_do(tmp_path):
    rows = good_table()
    rows[-1][rows[0].index("revenue")] = 5
    assert error_from(tmp_path, rows).endswith(
        "a budget-only row may fill only budget_new_arr, budget_arr, budget_net_burn: move the actuals to their "
        "own quarter row, or if this is an actual quarter, take 'budget', 'bud' or 'plan' out of its label")


def test_a_header_with_nothing_under_it_says_to_add_the_quarters(tmp_path):
    assert error_from(tmp_path, [good_table()[0]]) == (
        "Sheet 'KPI Tracker', row 3 (header) has no quarter rows under it: add one row per quarter under the "
        "header, labelled like 'Q2 2026'")


# ---------------------------------------------------------------------------
# main.py: how a failed company reads (the Result column, the batch CSV, the run log)
# ---------------------------------------------------------------------------

def test_an_input_error_reads_without_the_python_error_name():
    assert error_text(ValueError("Sheet 'KPI', row 3 (header) is missing columns: pipeline")) == \
        "Sheet 'KPI', row 3 (header) is missing columns: pipeline"


def test_a_file_that_cannot_be_opened_says_to_close_it():
    error = PermissionError(13, "Permission denied", "output/northwind_board_pack.pptx")
    assert error_text(error) == ("can't open or save output/northwind_board_pack.pptx: if it's open in Excel, "
                                 "PowerPoint or Word, close it, then run again")


def test_a_missing_file_names_it():
    error = FileNotFoundError(2, "No such file or directory", "data/nope.xlsx")
    assert error_text(error) == "can't find data/nope.xlsx: check the file name and folder, then run again"


def test_a_bug_still_shows_its_python_name_so_it_can_be_fixed():
    assert error_text(TypeError("unsupported operand")) == (
        "unexpected problem, probably a bug in this tool rather than the workbook (TypeError: unsupported "
        "operand): the Terminal window shows where it happened")


def test_a_failed_company_in_the_batch_reads_in_plain_words(tmp_path, capsys):
    broken = write_workbook(tmp_path / "broken.xlsx", drop_column(good_table(), "pipeline"))
    [result] = main.run_batch([broken], main.load_config(), True, output_dir=tmp_path / "out")
    assert main.result_text(result).startswith("FAILED: Sheet 'KPI Tracker', row 3 (header) is missing columns")
    assert "  ✗ FAILED: Sheet 'KPI Tracker', row 3 (header)" in capsys.readouterr().out


def test_a_step_that_fails_is_logged_in_plain_words(tmp_path):
    log = run_log.RunLog(tmp_path, run_log.COMMAND_LINE)
    with pytest.raises(ValueError):
        with log.company("Broken").step(run_log.CLEAN):
            raise ValueError("Sheet 'KPIs', row 7: can't read '12..5'")
    assert run_log.read_lines(log.path)[0][0]["error"] == "Sheet 'KPIs', row 7: can't read '12..5'"


@pytest.mark.parametrize("error, expected_start", [
    (status_error(anthropic.AuthenticationError, 401),
     "Claude's API didn't accept the key (ANTHROPIC_API_KEY in .env): check the key, or build without AI text "
     "(--skip-ai)"),
    (status_error(anthropic.PermissionDeniedError, 403),
     "Claude's API says this key isn't allowed to do that: check the key's account in the Anthropic console, or "
     "build without AI text (--skip-ai)"),
    (status_error(anthropic.RateLimitError, 429),
     "Claude's API was still turning requests away (rate limit) after the waits: run again in a few minutes"),
    (anthropic.APIConnectionError(request=API_REQUEST),
     "couldn't reach Claude's API (no internet connection, or no answer in time): check the connection, then "
     "run again"),
    (status_error(anthropic.InternalServerError, 500),
     "Claude's API had a problem on its side: run again in a few minutes"),
    (status_error(anthropic.BadRequestError, 400),
     "Claude's API refused the request (the details say why, e.g. a low credit balance): fix that, then run again"),
    (anthropic.AnthropicError("simulated outage"),
     "the call to Claude's API failed: run again, or build without AI text (--skip-ai)"),
])
def test_api_errors_say_what_happened_and_what_to_do(error, expected_start):
    text = analyze.api_error_text(error)
    assert text.startswith(expected_start + ". The numbers, deck and memo are still built, without AI text. "
                                            "Details: ")
    assert "Error:" not in text.split("Details: ")[0]   # no Python error name in the plain part


def test_failed_validation_says_what_to_do_next():
    error = analyze.AnalysisError(["headline: too long", "risk 2: says 'grew' but ARR fell"], run_info={})
    assert str(error) == ("Claude's answer failed validation twice, so it isn't used: run again for a fresh "
                          "answer, or build without AI text (--skip-ai). What was wrong:\n"
                          "- headline: too long\n- risk 2: says 'grew' but ARR fell")


@pytest.mark.parametrize("argv, expected", [
    (["--all", "--max-cost", "abc"], "argument --max-cost: 'abc' is not a number: give one above 0, e.g. 2.5"),
    (["--all", "--timeout", "0"], "argument --timeout: must be more than 0, not 0: give a number above 0, e.g. 2.5"),
    (["--all", "--workers", "two"], "argument --workers: 'two' is not a whole number: give one of 1 or more, e.g. 4"),
    (["--all", "--workers", "0"], "argument --workers: must be 1 or more, not 0: give a whole number, e.g. 4"),
])
def test_a_bad_option_value_says_what_to_give_instead(argv, expected, capsys):
    with pytest.raises(SystemExit):
        main.parse_args(argv)
    assert capsys.readouterr().err.strip().endswith(f"error: {expected}")


# ---------------------------------------------------------------------------
# analyze.py's own command line: plain words, never a traceback
# ---------------------------------------------------------------------------

def run_analyze(monkeypatch, workbook, key=None):
    """analyze.py's main() on `workbook` with the key given (None = no key). The API can't be called."""
    monkeypatch.setattr(sys, "argv", ["analyze.py", str(workbook)])
    monkeypatch.setattr(analyze, "load_dotenv", lambda: None)   # the project's .env is never read
    monkeypatch.setattr(analyze.anthropic, "Anthropic", lambda: pytest.fail("the API must not be called"))
    if key:
        monkeypatch.setenv("ANTHROPIC_API_KEY", key)
    else:
        monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
    with pytest.raises(SystemExit) as caught:
        analyze.main()
    return caught.value.code


def test_analyze_on_a_broken_workbook_prints_the_stop_not_a_traceback(tmp_path, monkeypatch, capsys):
    broken = write_workbook(tmp_path / "broken.xlsx", drop_column(good_table(), "pipeline"))
    assert run_analyze(monkeypatch, broken, key="not-a-real-key") == 1
    assert capsys.readouterr().out.startswith("Sheet 'KPI Tracker', row 3 (header) is missing columns: pipeline")


def test_analyze_without_a_key_says_where_to_put_it(tmp_path, monkeypatch, capsys):
    good = write_workbook(tmp_path / "good.xlsx", good_table())
    assert run_analyze(monkeypatch, good) == 1
    assert capsys.readouterr().out.strip() == "ANTHROPIC_API_KEY isn't set: add it to .env, then run again"


# ---------------------------------------------------------------------------
# mapping.py: the saved mappings/<company>.yaml
# ---------------------------------------------------------------------------

def saved_file_error(tmp_path, text):
    """saved_columns' stop for a mappings/acme.yaml holding `text`."""
    folder = tmp_path / "mappings"
    folder.mkdir()
    (folder / "acme.yaml").write_text(text)
    with pytest.raises(ValueError) as caught:
        mapping.saved_columns(tmp_path / "acme.xlsx", folder)
    return str(caught.value)


def test_a_saved_mapping_to_an_unknown_column_says_how_to_fix_the_file(tmp_path):
    message = saved_file_error(tmp_path, "columns:\n  Opening ARR: arr_start\n")
    assert message.endswith(": fix that line in the file, or delete it and confirm the header again with "
                            f"python mapping.py {tmp_path / 'acme.xlsx'} --confirm")


def test_a_file_without_a_columns_list_says_what_it_needs(tmp_path):
    assert saved_file_error(tmp_path, "company: acme\n").endswith(
        "acme.yaml isn't a mapping file this tool can read: it needs a 'columns:' line, then one "
        "'workbook header: input column' line per mapping. Fix it, or delete the file and confirm the headers "
        f"again with python mapping.py {tmp_path / 'acme.xlsx'} --confirm")


def test_a_file_with_broken_yaml_names_the_line_not_a_yaml_error(tmp_path):
    message = saved_file_error(tmp_path, "columns:\n  Opening ARR: starting_arr\n  GP: gross: profit\n")
    assert "acme.yaml, line 3: this line isn't in 'workbook header: input column' form" in message
    words = message.replace(str(tmp_path), "").replace("acme.yaml", "")   # the folder name has "yaml" in it
    assert "python mapping.py" in words and "yaml" not in words.lower()   # no YAML library error text


# ---------------------------------------------------------------------------
# build_deck.py: the template
# ---------------------------------------------------------------------------

def test_a_missing_template_says_how_to_build_it(tmp_path, monkeypatch):
    monkeypatch.setattr(build_deck, "TEMPLATE_PATH", tmp_path / "base.pptx")
    with pytest.raises(ValueError) as caught:
        build_deck.open_template()
    assert str(caught.value) == ("The slide template base.pptx is missing: build it with python make_template.py, "
                                 "then run again")


def test_a_template_layout_without_its_text_box_names_the_layout(monkeypatch):
    layout = SimpleNamespace(name="Title and Content", placeholders=[])
    with pytest.raises(ValueError) as caught:
        build_deck.layout_box(layout, build_deck.BODY_TYPES)
    assert str(caught.value) == ("The slide template base.pptx: layout 'Title and Content' has no text box of the "
                                 "kind this slide needs: rebuild the template with python make_template.py, then "
                                 "run again")


def test_a_template_without_the_content_layout_says_how_to_rebuild_it():
    presentation = SimpleNamespace(slide_layouts=SimpleNamespace(get_by_name=lambda name: None))
    with pytest.raises(ValueError) as caught:
        build_deck.find_content_layout(presentation)
    assert str(caught.value) == ("The slide template base.pptx has no 'Title and Content' layout: rebuild the "
                                 "template with python make_template.py, then run again")
