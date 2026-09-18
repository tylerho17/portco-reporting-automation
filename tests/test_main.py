"""Unit tests for main.py: the batch outputs (summary CSV, quarter warning) and the AI step (Task 3).

The summary tests use results written out by hand, so no workbook is needed. The AI-step tests run
data/northwind.xlsx through run_batch into a temporary folder with a fake Claude client, so the
real output/ folder is untouched and no API call can happen (a guard makes creating a real client fail).
Run from the project folder:  pytest
"""

import csv
import json
from pathlib import Path
from types import SimpleNamespace

import anthropic
import pytest
from docx import Document
from pptx import Presentation

import analyze
import main
import make_data
from analyze import BoardSummary
from build_deck import PLACEHOLDER_TEXT
from check_diff import last_quarter_workbook
from diff_runs import HEADING, NO_EARLIER_RUN
from main import quarter_mismatch_warning, write_summary_csv
from provenance import NOT_REVIEWED, file_sha256, git_commit, manifest_path, read_manifest, save_manifest
from text_fit import TextDoesNotFitError

PROJECT_DIR = Path(__file__).parent.parent
NORTHWIND = PROJECT_DIR / "data" / "northwind.xlsx"


def ok(company, quarter, tripped=(), cannot_evaluate=(), gap_count=0, blank_quarters=(), ai="skipped"):
    """A successful company's result."""
    return {"company": company, "quarter": quarter, "flags_total": 9, "tripped": list(tripped),
            "cannot_evaluate": list(cannot_evaluate), "gap_count": gap_count,
            "blank_quarters": list(blank_quarters), "ai": ai, "error": None}


def failed(company, error):
    return {"company": company, "error": error}


# ---------------------------------------------------------------------------
# Summary table and CSV
# ---------------------------------------------------------------------------

def test_summary_csv_has_one_row_per_company(tmp_path):
    results = [
        ok("Northwind", "Q2 2026", tripped=["a"] * 6, gap_count=19, blank_quarters=["Q1 2025"], ai="ok"),
        ok("Fernhollow", "Q2 2026", tripped=["a"] * 7, cannot_evaluate=["Rule of 40"], gap_count=20,
           blank_quarters=["Q2 2025"], ai="failed"),
        ok("Alderpeak", "Q2 2026", ai="skipped"),
        failed("Broken", "ValueError: row 3 (header) is missing columns: pipeline"),
    ]
    path = write_summary_csv(results, tmp_path / "batch_summary.csv")
    with open(path, newline="") as file:
        rows = list(csv.reader(file))
    assert rows == [
        ["Company", "Latest quarter", "Flags tripped", "Flags total", "Cannot evaluate", "Data gaps",
         "Blank quarters", "Result", "AI cost (USD)", "Notes"],
        ["Northwind", "Q2 2026", "6", "9", "0", "19", "Q1 2025", "OK", "", ""],
        ["Fernhollow", "Q2 2026", "7", "9", "1", "20", "Q2 2025", "OK (AI failed)", "", ""],
        ["Alderpeak", "Q2 2026", "0", "9", "0", "0", "", "OK (AI skipped)", "", ""],
        ["Broken", "", "", "", "", "", "", "FAILED: ValueError: row 3 (header) is missing columns: pipeline", "",
         ""],
    ]


def test_no_warning_when_every_company_ends_on_the_same_quarter():
    results = [ok("Alderpeak", "Q2 2026"), ok("Northwind", "Q2 2026"), failed("Broken", "ValueError: x")]
    assert quarter_mismatch_warning(results) is None


def test_no_warning_for_a_single_company():
    assert quarter_mismatch_warning([ok("Northwind", "Q2 2026")]) is None


def test_warning_names_each_quarter_and_its_companies():
    results = [ok("Alderpeak", "Q2 2026"), ok("Northwind", "Q1 2026"), ok("Fernhollow", "Q2 2026"),
               failed("Broken", "ValueError: x")]  # a failed company has no quarter and is left out
    assert quarter_mismatch_warning(results) == (
        "⚠ Companies end on different quarters (Q2 2026: Alderpeak, Fernhollow; Q1 2026: Northwind) "
        "- compare them with care")


def test_result_text_says_what_happened_to_the_ai_summary():
    assert main.result_text(ok("Northwind", "Q2 2026", ai="ok")) == "OK"
    assert main.result_text(ok("Northwind", "Q2 2026", ai="skipped")) == "OK (AI skipped)"
    assert main.result_text(ok("Northwind", "Q2 2026", ai="failed")) == "OK (AI failed)"
    assert main.result_text(failed("Broken", "ValueError: x")) == "FAILED: ValueError: x"


def test_ai_failed_warning_names_the_companies():
    results = [ok("Northwind", "Q2 2026", ai="failed"), ok("Alderpeak", "Q2 2026", ai="ok"),
               ok("Fernhollow", "Q2 2026", ai="failed"), failed("Broken", "ValueError: x")]
    assert main.ai_failed_warning(results) == (
        "⚠ AI summary unavailable for 2 companies (Northwind, Fernhollow): their decks show the placeholder; "
        "the reasons are in output/<company>_analysis.json")
    assert main.ai_failed_warning([ok("Northwind", "Q2 2026", ai="ok")]) is None


# ---------------------------------------------------------------------------
# The AI step, with a fake Claude client
# ---------------------------------------------------------------------------

@pytest.fixture(autouse=True)
def no_real_client(monkeypatch):
    """Creating a real Anthropic client fails the test, so no test here can reach the API."""
    def refuse(*args, **kwargs):
        raise AssertionError("a test tried to create a real Anthropic client")
    monkeypatch.setattr(analyze.anthropic, "Anthropic", refuse)


class FakeClient:
    """Stands in for anthropic.Anthropic(): every call returns the same answer (or raises), and calls are counted."""

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


def summary(headline="Retention is the main question for the board."):
    """A valid answer with no numbers in it, so the number check has nothing to reject."""
    point = {"title": "Steady base", "detail": "Customers stayed."}
    return BoardSummary.model_validate({
        "headline": headline, "wins": [point] * 3, "risks": [point] * 3,
        "questions": ["What drives churn?", "Where is pipeline coming from?", "How is hiring going?"]})


def run_northwind(tmp_path, skip_ai=False, client=None, draft=False):
    """run_batch on Northwind, saving into tmp_path. Returns Northwind's result."""
    return main.run_batch([NORTHWIND], main.load_config(), skip_ai, client=client, output_dir=tmp_path,
                          draft=draft)[0]


def headline_on_deck(tmp_path):
    """The Headline box on the last slide, AI commentary (slide 4)."""
    slide = Presentation(tmp_path / "northwind_board_pack.pptx").slides[-1]
    return next(shape.text_frame.text for shape in slide.shapes if shape.name == "Headline")


def saved_analysis(tmp_path):
    return json.loads((tmp_path / "northwind_analysis.json").read_text())


# ---------------------------------------------------------------------------
# Reusing a saved analysis (the web page's Generate button, Task 2 of the final run)
# ---------------------------------------------------------------------------

def run_northwind_reusing(tmp_path, skip_ai=True, client=None):
    """run_company on Northwind with reuse_saved=True, saving into tmp_path."""
    return main.run_company(NORTHWIND, main.load_config(), skip_ai, client=client, output_dir=tmp_path,
                            reuse_saved=True)


def test_a_saved_analysis_of_the_same_numbers_is_reused_with_no_api_call(tmp_path):
    run_northwind(tmp_path, client=FakeClient(summary("First headline.")))
    client = FakeClient(summary("Second headline."))
    result = run_northwind_reusing(tmp_path, skip_ai=False, client=client)
    assert client.calls == 0
    assert headline_on_deck(tmp_path) == "First headline."
    assert main.RESULT_TEXTS[result["ai"]] == "OK"
    ai = read_manifest(manifest_path(NORTHWIND, tmp_path))["ai"]
    assert ai["validation"] == main.AI_REUSED
    assert ai["cost_usd"] is None and ai["input_tokens"] is None   # nothing was spent in this run


def test_reuse_with_ai_skipped_still_puts_the_saved_text_on_the_deck(tmp_path):
    run_northwind(tmp_path, client=FakeClient(summary("First headline.")))
    result = run_northwind_reusing(tmp_path, skip_ai=True)
    assert headline_on_deck(tmp_path) == "First headline."
    assert main.RESULT_TEXTS[result["ai"]] == "OK"


def test_reuse_with_no_saved_analysis_and_ai_skipped_gives_the_placeholder(tmp_path):
    result = run_northwind_reusing(tmp_path, skip_ai=True)
    assert headline_on_deck(tmp_path) == PLACEHOLDER_TEXT
    assert main.RESULT_TEXTS[result["ai"]] == "OK (AI skipped)"


def test_a_saved_analysis_of_other_numbers_is_not_reused(tmp_path):
    run_northwind(tmp_path, client=FakeClient(summary("Old headline.")))
    path = tmp_path / "northwind_analysis.json"
    saved = json.loads(path.read_text())
    saved["payload"]["flags"] = []                     # same company and quarter, different facts
    path.write_text(json.dumps(saved))
    client = FakeClient(summary("New headline."))
    run_northwind_reusing(tmp_path, skip_ai=False, client=client)
    assert client.calls == 1 and headline_on_deck(tmp_path) == "New headline."


def test_ai_that_passes_goes_on_the_deck_and_is_saved(tmp_path):
    client = FakeClient(summary())
    result = run_northwind(tmp_path, client=client)
    assert result["error"] is None and main.result_text(result) == "OK"
    assert client.calls == 1
    assert headline_on_deck(tmp_path) == "Retention is the main question for the board."
    saved = saved_analysis(tmp_path)
    assert saved["summary"]["headline"] == "Retention is the main question for the board."
    assert saved["run_info"]["passed"] is True and saved["error"] is None
    assert saved["payload"]["company"] == "Northwind"


def test_ai_that_fails_validation_twice_still_builds_the_deck(tmp_path):
    # An old, good analysis is already there: it must be replaced, never put on this run's deck.
    run_northwind(tmp_path, client=FakeClient(summary()))
    assert saved_analysis(tmp_path)["summary"] is not None

    client = FakeClient(summary(headline="ARR grew 555.5% this quarter."))  # a number not in the data
    result = run_northwind(tmp_path, client=client)
    assert result["error"] is None and main.result_text(result) == "OK (AI failed)"
    assert client.calls == analyze.MAX_ATTEMPTS  # first try + one retry
    assert headline_on_deck(tmp_path) == PLACEHOLDER_TEXT
    saved = saved_analysis(tmp_path)
    assert saved["summary"] is None and saved["run_info"]["attempts"] == analyze.MAX_ATTEMPTS
    assert "555.5" in saved["error"]


def test_api_error_is_ai_failed_not_a_failed_company(tmp_path):
    client = FakeClient(error=anthropic.AnthropicError("simulated outage"))
    result = run_northwind(tmp_path, client=client)
    assert result["error"] is None and main.result_text(result) == "OK (AI failed)"
    assert headline_on_deck(tmp_path) == PLACEHOLDER_TEXT
    saved = saved_analysis(tmp_path)
    assert saved["summary"] is None and saved["error"] == analyze.api_error_text(anthropic.AnthropicError("simulated outage"))


def test_unexpected_error_in_the_ai_step_fails_the_company(tmp_path):
    # A bug (not a validation failure or an API error) must not hide behind "AI failed".
    run_northwind(tmp_path, client=FakeClient(summary()))  # an old, good analysis from an earlier run
    old_analysis = (tmp_path / "northwind_analysis.json").read_bytes()
    result = run_northwind(tmp_path, client=FakeClient(error=KeyError("simulated bug")))
    assert result["error"] == main.error_text(KeyError("simulated bug"))   # "unexpected problem ... (KeyError: ...)"
    # Task 7: a failed company's files never reach output/, so the old analysis stays beside the old
    # deck it belongs to (before, it was deleted and the old deck stayed without it).
    assert (tmp_path / "northwind_analysis.json").read_bytes() == old_analysis


def test_ai_text_too_long_for_the_slides_is_ai_failed_not_a_failed_company(tmp_path):
    # Review finding 1a, the last line of defence: an analysis that passed when it was made but
    # doesn't fit the slides still gets a deck, with the placeholder, and the run reads OK (AI failed).
    long_point = {"title": "Steady base", "detail": "customers " * 45}
    saved = {"summary": {"headline": "Retention is the main question for the board.",
                         "wins": [long_point] * 3, "risks": [long_point] * 3,
                         "questions": ["What drives churn?", "Where is pipeline from?", "How is hiring?"]},
             "payload": {"company": "Northwind", "latest_quarter": "Q2 2026"}}
    path = tmp_path / "northwind_analysis.json"
    path.write_text(json.dumps(saved))

    why_unavailable = main.deck_step(NORTHWIND, main.load_config(), path, tmp_path)
    assert "does not fit slide 4" in why_unavailable
    assert headline_on_deck(tmp_path) == PLACEHOLDER_TEXT
    assert main.RESULT_TEXTS[main.ai_status(False, why_unavailable)] == "OK (AI failed)"


def test_a_text_that_does_not_fit_prints_one_clear_line_not_a_traceback(capsys):
    # Review finding 2: the message already names the slide and the box, so a traceback (which means
    # "this is a bug in the code") only makes the real problem harder to see.
    main.describe_error(TextDoesNotFitError("Slide 3, Risks and flags: text doesn't fit"))
    printed = capsys.readouterr()
    assert "Slide 3, Risks and flags" in printed.out
    assert "Traceback" not in printed.out + printed.err


def test_skip_ai_builds_the_placeholder_deck_and_never_calls_claude(tmp_path, monkeypatch):
    def refuse(*args, **kwargs):
        raise AssertionError("--skip-ai called analyze()")
    monkeypatch.setattr(main, "analyze", refuse)
    old = tmp_path / "northwind_analysis.json"
    old.write_text("an analysis from an earlier run")

    result = run_northwind(tmp_path, skip_ai=True)
    assert result["error"] is None and main.result_text(result) == "OK (AI skipped)"
    assert headline_on_deck(tmp_path) == PLACEHOLDER_TEXT
    assert old.read_text() == "an analysis from an earlier run"  # --skip-ai leaves an earlier analysis alone


# ---------------------------------------------------------------------------
# The memo (final Task 1): built beside the deck, from the same analysis
# ---------------------------------------------------------------------------

def memo_text(tmp_path):
    """Every paragraph of the saved Word memo."""
    return "\n".join(item.text for item in Document(tmp_path / "northwind_board_memo.docx").paragraphs)


def test_a_run_writes_the_memo_beside_the_deck_with_the_ai_text(tmp_path, capsys):
    result = run_northwind(tmp_path, client=FakeClient(summary()))
    assert result["error"] is None and main.result_text(result) == "OK"
    assert (tmp_path / "northwind_board_memo.docx").exists() and (tmp_path / "northwind_board_memo.pdf").exists()
    assert "Retention is the main question for the board." in memo_text(tmp_path)
    assert "✓ Memo: " in capsys.readouterr().out


def test_skip_ai_builds_the_memo_with_ai_commentary_unavailable(tmp_path):
    run_northwind(tmp_path, skip_ai=True)
    text = memo_text(tmp_path)
    assert "AI commentary unavailable" in text and "Key metrics" in text


def test_ai_failed_builds_the_memo_with_ai_commentary_unavailable(tmp_path):
    run_northwind(tmp_path, client=FakeClient(error=anthropic.AnthropicError("simulated outage")))
    assert "AI commentary unavailable" in memo_text(tmp_path)


def test_ai_text_quoting_a_number_the_metrics_workbook_lacks_stays_off_the_memo_only(tmp_path, capsys):
    # 14,300 is Northwind's latest ending cash: in Claude's payload (so analyze.py and the deck accept it),
    # but not in the metrics workbook, and every number in the memo must be.
    answer = summary()
    answer.questions[0] = "How long will the ending cash of $14,300K last at the current burn?"
    result = run_northwind(tmp_path, client=FakeClient(answer))
    assert main.result_text(result) == "OK"   # the deck has the AI text
    assert headline_on_deck(tmp_path) == "Retention is the main question for the board."
    assert "AI commentary unavailable" in memo_text(tmp_path)
    printed = capsys.readouterr().out
    assert "AI commentary unavailable" in printed and "14,300" in printed


def test_the_manifest_records_the_memo(tmp_path):
    run_northwind(tmp_path, client=FakeClient(summary()))
    assert northwind_manifest(tmp_path)["memo"] == {
        "files": ["northwind_board_memo.docx", "northwind_board_memo.pdf"], "ai_text": True}
    run_northwind(tmp_path, skip_ai=True)
    assert northwind_manifest(tmp_path)["memo"]["ai_text"] is False


# ---------------------------------------------------------------------------
# The manifest: where each deck came from (provenance)
# ---------------------------------------------------------------------------

CONFIG_FILE = PROJECT_DIR / "config.yaml"


def northwind_manifest(tmp_path):
    return read_manifest(manifest_path(NORTHWIND, tmp_path))


def test_a_run_writes_a_manifest_for_the_company(tmp_path):
    client = FakeClient(summary())
    run_northwind(tmp_path, client=client)
    saved = northwind_manifest(tmp_path)

    assert saved["company"] == "Northwind"
    assert saved["input"] == {"file": "northwind.xlsx", "sha256": file_sha256(NORTHWIND)}
    assert saved["config"] == {"file": "config.yaml", "sha256": file_sha256(CONFIG_FILE)}
    assert saved["code"] == git_commit()
    assert saved["ai"]["model"] == analyze.DEFAULT_MODEL
    assert saved["ai"]["prompt_version"] == analyze.PROMPT_VERSION
    assert saved["ai"]["validation"] == "passed" and saved["ai"]["attempts"] == 1
    assert saved["ai"]["input_tokens"] == 100 and saved["ai"]["output_tokens"] == 50
    assert saved["ai"]["cost_usd"] > 0
    assert saved["deck"] == {"file": "northwind_board_pack.pptx", "ai_text": True, "status": NOT_REVIEWED,
                             "draft": False, "appendix": False}
    assert saved["approval"] is None  # nobody has reviewed it yet
    assert saved["run_at"]


def test_the_manifest_records_a_skipped_ai_step(tmp_path):
    run_northwind(tmp_path, skip_ai=True)
    saved = northwind_manifest(tmp_path)
    assert saved["ai"]["validation"] == "skipped"
    assert saved["ai"]["model"] is None and saved["ai"]["cost_usd"] is None
    assert saved["deck"]["ai_text"] is False


def test_the_manifest_records_an_ai_failure_and_that_the_deck_has_the_placeholder(tmp_path):
    run_northwind(tmp_path, client=FakeClient(summary(headline="ARR grew 555.5% this quarter.")))
    saved = northwind_manifest(tmp_path)
    assert saved["ai"]["validation"].startswith("failed") and "555.5" in saved["ai"]["validation"]
    assert saved["ai"]["attempts"] == analyze.MAX_ATTEMPTS and saved["ai"]["cost_usd"] > 0
    assert saved["deck"]["ai_text"] is False


def test_the_manifest_hash_follows_the_workbook(tmp_path, monkeypatch):
    # The hash is of the file that was read, so a different workbook can't pass for this one.
    run_northwind(tmp_path, skip_ai=True)
    first = northwind_manifest(tmp_path)["input"]["sha256"]

    copy = tmp_path / "northwind.xlsx"
    copy.write_bytes(NORTHWIND.read_bytes() + b"a change")
    main.run_batch([copy], main.load_config(), True, output_dir=tmp_path)
    assert northwind_manifest(tmp_path)["input"]["sha256"] != first


def test_a_re_run_keeps_an_approval_that_is_still_valid(tmp_path):
    run_northwind(tmp_path, skip_ai=True)
    saved = northwind_manifest(tmp_path)
    saved["approval"] = {"reviewer": "Tyler Ho", "approved_at": "2026-09-17T15:00:00",
                         "input_sha256": file_sha256(NORTHWIND), "config_sha256": file_sha256(CONFIG_FILE)}
    save_manifest(manifest_path(NORTHWIND, tmp_path), saved)

    run_northwind(tmp_path, skip_ai=True)  # same workbook, same thresholds
    after = northwind_manifest(tmp_path)
    assert after["approval"]["reviewer"] == "Tyler Ho"
    assert after["deck"]["status"] == "approved by Tyler Ho on 2026-09-17T15:00:00"


def test_a_run_after_the_workbook_changed_drops_back_to_draft(tmp_path):
    run_northwind(tmp_path, skip_ai=True)
    saved = northwind_manifest(tmp_path)
    saved["approval"] = {"reviewer": "Tyler Ho", "approved_at": "2026-09-17T15:00:00",
                         "input_sha256": "the-hash-of-an-older-workbook", "config_sha256": file_sha256(CONFIG_FILE)}
    save_manifest(manifest_path(NORTHWIND, tmp_path), saved)

    run_northwind(tmp_path, skip_ai=True)
    assert northwind_manifest(tmp_path)["deck"]["status"] == NOT_REVIEWED


# ---------------------------------------------------------------------------
# What changed since the last run (Task 10)
# ---------------------------------------------------------------------------

def run_last_quarter_then_today(tmp_path):
    """Northwind as it stood at Q1 2026, then today's workbook, into the same output folder (tmp_path/out)."""
    earlier = last_quarter_workbook(make_data, tmp_path)
    main.run_batch([earlier], main.load_config(), True, output_dir=tmp_path / "out")
    run_northwind(tmp_path / "out", skip_ai=True)
    return northwind_manifest(tmp_path / "out")


def test_a_first_run_records_its_results_and_has_nothing_to_compare_with(tmp_path, capsys):
    run_northwind(tmp_path, skip_ai=True)
    saved = northwind_manifest(tmp_path)
    assert saved["results"]["quarter"] == "Q2 2026"
    assert saved["results"]["flags"]["Runway at current burn"] == "Tripped"
    assert saved["previous_run"] is None
    assert HEADING not in memo_text(tmp_path)
    assert f"✓ {HEADING}: {NO_EARLIER_RUN}" in capsys.readouterr().out


def test_a_run_a_quarter_later_is_compared_with_the_last_one(tmp_path, capsys):
    saved = run_last_quarter_then_today(tmp_path)
    assert saved["previous_run"]["results"]["quarter"] == "Q1 2026"
    memo = memo_text(tmp_path / "out")
    assert HEADING in memo and "Runway at current burn: Tripped (was Passed)" in memo
    assert "whose latest quarter was Q1 2026 (now Q2 2026)" in memo
    assert f"✓ {HEADING}: 5 flags flipped" in capsys.readouterr().out


def test_a_rebuild_from_the_same_numbers_keeps_last_quarter_s_comparison(tmp_path):
    # approve.py, then a rebuild: the memo must still say what changed since last quarter.
    first = run_last_quarter_then_today(tmp_path)
    run_northwind(tmp_path / "out", skip_ai=True)
    again = northwind_manifest(tmp_path / "out")
    assert again["previous_run"] == first["previous_run"]
    assert again["run_at"] >= first["run_at"]
    assert "Runway at current burn: Tripped (was Passed)" in memo_text(tmp_path / "out")


def test_run_company_straight_into_output_compares_too(tmp_path):
    # The batch above builds in a private folder; the web page's Generate calls run_company on output/ itself.
    earlier = last_quarter_workbook(make_data, tmp_path)
    main.run_company(earlier, main.load_config(), True, output_dir=tmp_path / "out", reuse_saved=True)
    main.run_company(NORTHWIND, main.load_config(), True, output_dir=tmp_path / "out", reuse_saved=True)
    assert northwind_manifest(tmp_path / "out")["previous_run"]["results"]["quarter"] == "Q1 2026"
    assert "Runway at current burn: Tripped (was Passed)" in memo_text(tmp_path / "out")


# ---------------------------------------------------------------------------
# The DRAFT watermark is opt-in (--draft); the footer always says whether the deck was reviewed
# ---------------------------------------------------------------------------

def deck_watermarks(tmp_path):
    deck = Presentation(tmp_path / "northwind_board_pack.pptx")
    return [item.name for slide in deck.slides for item in slide.shapes if item.name == "Watermark"]


def deck_footer(tmp_path):
    deck = Presentation(tmp_path / "northwind_board_pack.pptx")
    return [item for item in deck.slides[0].shapes if item.name == "Footer"][0].text_frame.text


def test_a_run_has_no_watermark_by_default_and_the_footer_says_not_reviewed(tmp_path):
    run_northwind(tmp_path, skip_ai=True)
    assert deck_watermarks(tmp_path) == []
    assert deck_footer(tmp_path).endswith(" | AI-drafted | not reviewed")


def test_draft_watermarks_every_slide_of_an_unreviewed_deck(tmp_path):
    run_northwind(tmp_path, skip_ai=True, draft=True)
    assert len(deck_watermarks(tmp_path)) == 4


def test_the_printout_names_the_slide_the_ai_text_is_on(tmp_path, capsys):
    # The deck has 4 slides and the AI text is only on slide 4 (AI commentary).
    run_northwind(tmp_path, client=FakeClient(summary()))
    assert "(AI text on slide 4)" in capsys.readouterr().out
    run_northwind(tmp_path, skip_ai=True)
    assert f"({PLACEHOLDER_TEXT} on slide 4)" in capsys.readouterr().out


def test_the_footer_of_a_re_run_names_a_reviewer_whose_approval_still_counts(tmp_path):
    run_northwind(tmp_path, skip_ai=True)
    saved = northwind_manifest(tmp_path)
    saved["approval"] = {"reviewer": "Tyler Ho", "approved_at": "2026-09-17T15:00:00",
                         "input_sha256": file_sha256(NORTHWIND), "config_sha256": file_sha256(CONFIG_FILE)}
    save_manifest(manifest_path(NORTHWIND, tmp_path), saved)

    run_northwind(tmp_path, skip_ai=True, draft=True)
    assert deck_footer(tmp_path).endswith(" | AI-drafted | reviewed by Tyler Ho on 2026-09-17")
    assert deck_watermarks(tmp_path) == []  # --draft never stamps an approved deck


def test_draft_is_passed_from_the_command_line_to_the_batch(monkeypatch):
    seen = {}
    monkeypatch.setattr(main, "run_batch", lambda *args, **kwargs: seen.update(kwargs) or [])
    monkeypatch.setattr(main, "write_summary_csv", lambda results: PROJECT_DIR / "output" / "batch_summary.csv")
    monkeypatch.setattr(main, "save_batch_manifest", lambda *args: PROJECT_DIR / "output" / "batch_manifest.json")
    main.main(["--all", "--skip-ai", "--draft"])
    assert seen["draft"] is True
    main.main(["--all", "--skip-ai"])
    assert seen["draft"] is False


def test_appendix_is_passed_from_the_command_line_to_the_batch(monkeypatch):
    seen = {}
    monkeypatch.setattr(main, "run_batch", lambda *args, **kwargs: seen.update(kwargs) or [])
    monkeypatch.setattr(main, "write_summary_csv", lambda results: PROJECT_DIR / "output" / "batch_summary.csv")
    monkeypatch.setattr(main, "save_batch_manifest", lambda *args: PROJECT_DIR / "output" / "batch_manifest.json")
    main.main(["--all", "--skip-ai", "--appendix"])
    assert seen["appendix"] is True
    main.main(["--all", "--skip-ai"])
    assert seen["appendix"] is False


def test_appendix_adds_the_metric_table_slide_and_the_manifest_says_so(tmp_path):
    main.run_batch([NORTHWIND], main.load_config(), True, output_dir=tmp_path, appendix=True)
    slides = Presentation(tmp_path / "northwind_board_pack.pptx").slides
    assert len(slides) == 5 and slides[-1].shapes.title.text.startswith("Appendix: every metric")
    assert read_manifest(manifest_path(NORTHWIND, tmp_path))["deck"]["appendix"] is True
    run_northwind(tmp_path, skip_ai=True)   # off by default
    assert len(Presentation(tmp_path / "northwind_board_pack.pptx").slides) == 4
    assert read_manifest(manifest_path(NORTHWIND, tmp_path))["deck"]["appendix"] is False


# ---------------------------------------------------------------------------
# The API key is checked before any company runs
# ---------------------------------------------------------------------------

def test_api_key_problem(monkeypatch):
    monkeypatch.setattr(main, "load_dotenv", lambda: None)  # don't let .env fill the key back in
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
    assert main.api_key_problem() == main.NO_KEY_MESSAGE
    monkeypatch.setenv("ANTHROPIC_API_KEY", "   ")
    assert main.api_key_problem() == main.NO_KEY_MESSAGE
    monkeypatch.setenv("ANTHROPIC_API_KEY", "sk-test")
    assert main.api_key_problem() is None


def test_no_api_key_stops_before_any_company(monkeypatch, capsys):
    monkeypatch.setattr(main, "load_dotenv", lambda: None)
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
    monkeypatch.setattr(main, "run_batch", lambda *args, **kwargs: pytest.fail("the batch ran without a key"))
    assert main.main(["--all"]) == 1
    assert main.NO_KEY_MESSAGE in capsys.readouterr().out


def test_skip_ai_never_needs_a_key(monkeypatch):
    monkeypatch.setattr(main, "api_key_problem", lambda: pytest.fail("--skip-ai looked for an API key"))
    monkeypatch.setattr(main, "run_batch", lambda *args, **kwargs: [])
    monkeypatch.setattr(main, "write_summary_csv", lambda results: PROJECT_DIR / "output" / "batch_summary.csv")
    monkeypatch.setattr(main, "save_batch_manifest", lambda *args: PROJECT_DIR / "output" / "batch_manifest.json")
    assert main.main(["--all", "--skip-ai"]) == 0
