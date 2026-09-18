"""Tests for the run log (Task 15): output/logs/run_<timestamp>.jsonl, one line per step per company.

Three parts:
- run_log.py on its own: the file name, what a line holds, a step that works, fails or is stopped,
  and reading the logs back for the web page's Recent runs panel (a broken line never stops it)
- main.run_batch writing the log: every step of every company, then one "whole company" line with
  its outcome (built, failed, skipped, stopped, timed out)
- the web page: Generate writes a log, and the portfolio's Recent runs card shows it

Every run goes into a temporary folder, with a fake Claude client or none (a guard makes creating a
real client fail), so the real output/ folder is untouched and no API call can happen.
Run from the project folder:  pytest
"""

import json
import threading
from datetime import datetime
from pathlib import Path

import pytest

import analyze
import main
import portfolio
import run_log
from resilience import Cancelled
from run_log import RunLog
from test_batch import HangingClient, ScriptedClient, answer, TIMEOUT_SECONDS

PROJECT_DIR = Path(__file__).parent.parent
DATA_DIR = PROJECT_DIR / "data"
NORTHWIND, ALDERPEAK, FERNHOLLOW = (DATA_DIR / f"{name}.xlsx" for name in ("northwind", "alderpeak", "fernhollow"))
EM_DASH = chr(0x2014)   # by its Unicode number, so this file shows none

# The steps run_company logs for one company, in order, then the batch's line for the whole company.
STEPS = [run_log.CLEAN, run_log.METRICS, run_log.CHANGES, run_log.EXCEL, run_log.AI, run_log.DECK, run_log.MEMO,
         run_log.MANIFEST]


@pytest.fixture(autouse=True)
def no_real_client(monkeypatch):
    """Creating a real Anthropic client fails the test, so no test here can reach the API."""
    def refuse(*args, **kwargs):
        raise AssertionError("a test tried to create a real Anthropic client")
    monkeypatch.setattr(analyze.anthropic, "Anthropic", refuse)


class FakeClock:
    """A stopwatch the test moves by hand: each reading is `step` seconds after the last."""

    def __init__(self, step=1.5):
        self.now, self.step = 100.0, step

    def __call__(self):
        self.now += self.step
        return self.now


def at(text):
    """A fixed 'now' for a log: at('2026-09-18 14:15:03') -> a function returning that time."""
    return lambda: datetime.fromisoformat(text)


def lines_of(path):
    """Every line of a log file, as dicts."""
    return [json.loads(line) for line in Path(path).read_text().splitlines()]


def only_log(output_dir):
    """The one log file a run wrote into output_dir/logs (fails if there are more or none)."""
    files = sorted((Path(output_dir) / "logs").glob("*.jsonl"))
    assert len(files) == 1, files
    return files[0]


def batch(tmp_path, paths=(NORTHWIND,), skip_ai=True, client=None, **options):
    """main.run_batch into tmp_path."""
    return main.run_batch(list(paths), main.load_config(), skip_ai, client=client, output_dir=tmp_path, **options)


def steps_for(lines, company):
    """[(step, result), ...] for one company, in the order written."""
    return [(line["step"], line["result"]) for line in lines if line["company"] == company]


# ---------------------------------------------------------------------------
# Writing: the file and its lines
# ---------------------------------------------------------------------------

def test_the_log_is_named_after_the_time_the_run_started_and_lives_in_logs(tmp_path):
    log = RunLog(tmp_path, run_log.COMMAND_LINE, now=at("2026-09-18 14:15:03.250"))
    log.write("Northwind", run_log.CLEAN, 0.25, run_log.OK)   # started 0.25 s before now
    assert log.path == tmp_path / "logs" / "run_20260918-141503.jsonl"
    assert log.path.exists()


def test_nothing_is_written_until_the_first_line(tmp_path):
    log = RunLog(tmp_path, run_log.COMMAND_LINE)
    assert log.path is None
    assert not (tmp_path / "logs").exists()


def test_a_line_holds_the_step_its_duration_its_result_and_its_error(tmp_path):
    log = RunLog(tmp_path, run_log.WEB_PAGE)
    log.write("Fernhollow", run_log.EXCEL, 1.23456, run_log.FAILED, "ValueError: bad cell",
              started=datetime.fromisoformat("2026-09-18 14:15:03.250"))
    assert lines_of(log.path) == [{
        "run": "run_20260918-141503", "source": run_log.WEB_PAGE, "company": "Fernhollow", "step": run_log.EXCEL,
        "started_at": "2026-09-18T14:15:03.250", "seconds": 1.235, "result": run_log.FAILED,
        "error": "ValueError: bad cell"}]


def test_two_runs_in_the_same_second_get_two_files(tmp_path):
    first, second = (RunLog(tmp_path, run_log.COMMAND_LINE, now=at("2026-09-18 14:15:03")) for _ in range(2))
    first.write("Northwind", run_log.CLEAN, 0.1, run_log.OK)
    second.write("Northwind", run_log.CLEAN, 0.1, run_log.OK)
    assert first.path != second.path
    assert len(lines_of(first.path)) == len(lines_of(second.path)) == 1


def test_every_line_is_whole_when_many_threads_write_at_once(tmp_path):
    log = RunLog(tmp_path, run_log.COMMAND_LINE)
    writers = [threading.Thread(target=lambda n=n: [log.write(f"Company {n}", run_log.DECK, 0.1, run_log.OK)
                                                    for _ in range(50)]) for n in range(8)]
    for writer in writers:
        writer.start()
    for writer in writers:
        writer.join()
    assert len(lines_of(log.path)) == 400   # json.loads fails on any half-written or joined line


def test_a_log_that_cannot_be_written_warns_once_and_never_stops_the_run(tmp_path, capsys):
    (tmp_path / "logs").write_text("a file where the folder should be")
    log = RunLog(tmp_path, run_log.COMMAND_LINE)
    log.write("Northwind", run_log.CLEAN, 0.1, run_log.OK)
    log.write("Northwind", run_log.METRICS, 0.1, run_log.OK)
    assert capsys.readouterr().out.count("Run log not written") == 1


# ---------------------------------------------------------------------------
# Timing a step
# ---------------------------------------------------------------------------

def test_a_step_that_works_is_ok_and_timed(tmp_path):
    log = RunLog(tmp_path, run_log.COMMAND_LINE, clock=FakeClock(step=1.5))
    with log.company("Northwind").step(run_log.DECK):
        pass
    [line] = lines_of(log.path)
    assert (line["step"], line["seconds"], line["result"], line["error"]) == (run_log.DECK, 1.5, run_log.OK, None)


def test_a_step_that_raises_is_failed_with_the_error_and_the_error_still_stops_the_company(tmp_path):
    log = RunLog(tmp_path, run_log.COMMAND_LINE)
    with pytest.raises(ValueError):
        with log.company("Northwind").step(run_log.EXCEL):
            raise ValueError("Sheet 'KPIs', row 7: can't read '12..5'")
    [line] = lines_of(log.path)
    assert (line["result"], line["error"]) == (run_log.FAILED, "ValueError: Sheet 'KPIs', row 7: can't read '12..5'")


def test_a_step_the_batch_gave_up_on_is_stopped_not_failed(tmp_path):
    log = RunLog(tmp_path, run_log.COMMAND_LINE)
    with pytest.raises(Cancelled):
        with log.company("Northwind").step(run_log.AI):
            raise Cancelled()
    [line] = lines_of(log.path)
    assert (line["result"], line["error"]) == (run_log.STOPPED, run_log.GIVEN_UP)


def test_a_step_can_set_its_own_result(tmp_path):
    log = RunLog(tmp_path, run_log.COMMAND_LINE)
    with log.company("Northwind").step(run_log.AI) as step:
        step.result, step.error = run_log.FAILED, "failed validation after the retry"
    [line] = lines_of(log.path)
    assert (line["result"], line["error"]) == (run_log.FAILED, "failed validation after the retry")


def test_no_log_writes_nothing_anywhere(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    log = run_log.no_log()
    log.write("Northwind", run_log.CLEAN, 0.1, run_log.OK)
    with log.company("Northwind").step(run_log.DECK):
        pass
    assert log.path is None and list(tmp_path.iterdir()) == []


def test_a_company_without_a_run_log_writes_nothing(tmp_path):
    with run_log.CompanyLog(None, "Northwind").step(run_log.CLEAN):
        pass
    assert list(tmp_path.iterdir()) == []


# ---------------------------------------------------------------------------
# The batch writes the log
# ---------------------------------------------------------------------------

def test_a_batch_logs_every_step_of_every_company_then_the_whole_company(tmp_path):
    batch(tmp_path, paths=(NORTHWIND, ALDERPEAK))
    lines = lines_of(only_log(tmp_path))
    expected = [(step, run_log.OK) for step in STEPS]
    expected[STEPS.index(run_log.AI)] = (run_log.AI, run_log.SKIPPED)   # --skip-ai
    for company in ("Northwind", "Alderpeak"):
        assert steps_for(lines, company) == expected + [(run_log.WHOLE_COMPANY, main.BUILT)]
    assert {line["source"] for line in lines} == {run_log.COMMAND_LINE}
    assert all(line["seconds"] >= 0 and line["error"] is None for line in lines)


def test_the_whole_company_line_takes_at_least_as_long_as_its_steps(tmp_path):
    batch(tmp_path)
    lines = lines_of(only_log(tmp_path))
    whole = next(line["seconds"] for line in lines if line["step"] == run_log.WHOLE_COMPANY)
    assert whole >= sum(line["seconds"] for line in lines if line["step"] != run_log.WHOLE_COMPANY) - 0.01


def test_the_whole_company_line_starts_when_the_company_started_not_when_it_ended(tmp_path):
    # Found on the first real run: its started_at was the moment the line was written.
    batch(tmp_path)
    lines = lines_of(only_log(tmp_path))
    first_step = datetime.fromisoformat(lines[0]["started_at"])
    whole = next(line for line in lines if line["step"] == run_log.WHOLE_COMPANY)
    assert datetime.fromisoformat(whole["started_at"]) <= first_step


def test_a_line_written_without_a_start_time_starts_its_seconds_before_now(tmp_path):
    log = RunLog(tmp_path, run_log.COMMAND_LINE, now=at("2026-09-18 14:15:03"))
    log.write("Northwind", run_log.WHOLE_COMPANY, 2.5, main.BUILT)
    assert lines_of(log.path)[0]["started_at"] == "2026-09-18T14:15:00.500"


def test_the_ai_step_is_ok_when_claude_answers(tmp_path):
    batch(tmp_path, skip_ai=False, client=ScriptedClient(answer()))
    lines = lines_of(only_log(tmp_path))
    assert (run_log.AI, run_log.OK) in steps_for(lines, "Northwind")


def test_a_failed_ai_step_logs_why_but_the_company_is_still_built(tmp_path, monkeypatch):
    def always_invalid(text, client=None, **kwargs):
        raise analyze.AnalysisError(["headline too long for slide 4"], run_info=None)
    monkeypatch.setattr(main, "analyze", always_invalid)
    batch(tmp_path, skip_ai=False, client=ScriptedClient(answer()))
    lines = lines_of(only_log(tmp_path))
    ai = next(line for line in lines if line["step"] == run_log.AI)
    assert ai["result"] == run_log.FAILED and "headline too long for slide 4" in ai["error"]
    assert steps_for(lines, "Northwind")[-1] == (run_log.WHOLE_COMPANY, main.BUILT)


def test_a_reused_analysis_is_logged_as_reused(tmp_path):
    batch(tmp_path, skip_ai=False, client=ScriptedClient(answer()))
    (tmp_path / "logs").rename(tmp_path / "first_logs")
    batch(tmp_path, skip_ai=True, resume=True, draft=True)   # --draft forces a rebuild, and the analysis is reused
    assert (run_log.AI, run_log.REUSED) in steps_for(lines_of(only_log(tmp_path)), "Northwind")


def test_a_failing_step_is_logged_with_its_error_and_no_later_step_runs(tmp_path, monkeypatch):
    def broken(*args, **kwargs):
        raise ValueError("the metrics workbook can't be saved")
    monkeypatch.setattr(main, "save_metrics_workbook", broken)
    batch(tmp_path)
    lines = lines_of(only_log(tmp_path))
    assert steps_for(lines, "Northwind") == [(run_log.CLEAN, run_log.OK), (run_log.METRICS, run_log.OK),
                                             (run_log.CHANGES, run_log.OK), (run_log.EXCEL, run_log.FAILED),
                                             (run_log.WHOLE_COMPANY, main.FAILED)]
    errors = [line["error"] for line in lines if line["result"] == run_log.FAILED or line["result"] == main.FAILED]
    assert errors == ["ValueError: the metrics workbook can't be saved"] * 2


def test_a_workbook_that_cannot_be_read_fails_at_the_clean_step(tmp_path):
    bad = tmp_path / "broken.xlsx"
    bad.write_text("not a workbook")
    batch(tmp_path / "out", paths=(bad, NORTHWIND))
    lines = lines_of(only_log(tmp_path / "out"))
    assert steps_for(lines, "Broken") == [(run_log.CLEAN, run_log.FAILED), (run_log.WHOLE_COMPANY, main.FAILED)]
    assert steps_for(lines, "Northwind")[-1] == (run_log.WHOLE_COMPANY, main.BUILT)   # the batch carried on


def test_a_resume_skip_is_one_whole_company_line(tmp_path):
    batch(tmp_path)
    (tmp_path / "logs").rename(tmp_path / "first_logs")
    batch(tmp_path, resume=True)
    assert steps_for(lines_of(only_log(tmp_path)), "Northwind") == [(run_log.WHOLE_COMPANY, main.SKIPPED)]


def test_a_company_stopped_by_the_spend_ceiling_is_logged_with_why(tmp_path):
    batch(tmp_path, paths=(NORTHWIND, ALDERPEAK), skip_ai=False, client=ScriptedClient(answer(10**6, 10**5)),
          max_cost=0.01)
    lines = lines_of(only_log(tmp_path))
    [stopped] = [line for line in lines if line["company"] == "Alderpeak"]
    assert (stopped["step"], stopped["result"]) == (run_log.WHOLE_COMPANY, main.STOPPED)
    assert "--max-cost" in stopped["error"]


def test_a_timed_out_company_is_logged_as_timed_out(tmp_path):
    client = HangingClient("Northwind")
    try:
        batch(tmp_path, paths=(NORTHWIND, ALDERPEAK), skip_ai=False, client=client, timeout=TIMEOUT_SECONDS)
    finally:
        client.release.set()
    whole = {line["company"]: line for line in lines_of(only_log(tmp_path)) if line["step"] == run_log.WHOLE_COMPANY}
    assert whole["Northwind"]["result"] == main.TIMED_OUT
    assert f"{TIMEOUT_SECONDS:g} s" in whole["Northwind"]["error"]
    assert whole["Northwind"]["seconds"] >= TIMEOUT_SECONDS
    assert whole["Alderpeak"]["result"] == main.BUILT


def test_side_by_side_companies_share_one_log_of_whole_lines(tmp_path):
    batch(tmp_path, paths=(NORTHWIND, ALDERPEAK, FERNHOLLOW), workers=3)
    lines = lines_of(only_log(tmp_path))
    assert len(lines) == 3 * (len(STEPS) + 1)
    assert {line["run"] for line in lines} == {only_log(tmp_path).stem}


def test_main_says_where_the_run_log_is(tmp_path, monkeypatch, capsys):
    monkeypatch.setattr(main, "OUTPUT_DIR", tmp_path)
    monkeypatch.setattr(main, "write_summary_csv", lambda results: tmp_path / "batch_summary.csv")
    monkeypatch.setattr(main, "save_batch_manifest", lambda *args, **kwargs: tmp_path / "batch_manifest.json")
    assert main.main([str(NORTHWIND), "--skip-ai"]) == 0
    assert f"Run log saved: {only_log(tmp_path)}" in capsys.readouterr().out


# ---------------------------------------------------------------------------
# Reading the logs back (the web page's Recent runs panel)
# ---------------------------------------------------------------------------

def write_run(output_dir, when, rows, source=run_log.COMMAND_LINE):
    """A log file with the given (company, step, seconds, result, error) rows, all started at `when`."""
    log = RunLog(output_dir, source)
    for row in rows:
        log.write(*row, started=datetime.fromisoformat(when))
    return log.path


def test_no_logs_means_no_recent_runs(tmp_path):
    assert run_log.recent_runs(tmp_path) == []


def test_recent_runs_are_newest_first_and_at_most_the_limit(tmp_path):
    for minute in range(7):
        write_run(tmp_path, f"2026-09-18 14:{minute:02d}:00", [("Northwind", run_log.CLEAN, 0.1, run_log.OK)])
    runs = run_log.recent_runs(tmp_path, limit=5)
    assert [run["started"] for run in runs] == [f"2026-09-18 14:{minute:02d}" for minute in (6, 5, 4, 3, 2)]


def test_the_tenth_run_in_one_second_is_newer_than_the_second(tmp_path):
    # Found by a planted bug: sorted as text, "_10" comes before "_2".
    for _ in range(10):
        write_run(tmp_path, "2026-09-18 14:15:03", [("Northwind", run_log.CLEAN, 0.1, run_log.OK)])
    newest = run_log.recent_runs(tmp_path, limit=2)
    assert [run["run"] for run in newest] == ["run_20260918-141503_10", "run_20260918-141503_9"]


def test_a_run_summary_counts_the_outcomes_and_lists_the_problems(tmp_path):
    write_run(tmp_path, "2026-09-18 14:15:03", [
        ("Northwind", run_log.CLEAN, 0.5, run_log.OK), ("Northwind", run_log.WHOLE_COMPANY, 4.0, main.BUILT),
        ("Fernhollow", run_log.AI, 2.0, run_log.FAILED, "failed validation"),
        ("Fernhollow", run_log.WHOLE_COMPANY, 6.0, main.BUILT),
        ("Broken", run_log.CLEAN, 0.1, run_log.FAILED, "ValueError: no Quarter header"),
        ("Broken", run_log.WHOLE_COMPANY, 0.1, main.FAILED, "ValueError: no Quarter header")])
    [run] = run_log.recent_runs(tmp_path)
    assert run["companies"] == 3
    assert run["outcomes"] == "2 built, 1 failed"
    assert run["problems"] == ["Fernhollow, AI commentary: failed validation",
                               "Broken, clean: ValueError: no Quarter header"]
    assert run["seconds"] == 6.0
    assert run_log.run_label(run) == "2026-09-18 14:15 · command line · 3 companies: 2 built, 1 failed · 6.0 s"


def test_a_company_with_no_whole_company_line_is_unfinished(tmp_path):
    write_run(tmp_path, "2026-09-18 14:15:03", [("Northwind", run_log.CLEAN, 0.5, run_log.OK)])   # Ctrl+C, say
    [run] = run_log.recent_runs(tmp_path)
    assert run["outcomes"] == "1 unfinished"


def test_a_broken_line_is_counted_and_the_rest_are_still_read(tmp_path):
    path = write_run(tmp_path, "2026-09-18 14:15:03", [("Northwind", run_log.WHOLE_COMPANY, 1.0, main.BUILT)])
    with open(path, "a") as file:
        file.write('{"company": "Alder\n[1, 2]\n{"step": "deck"}\n')   # cut off, not a dict, missing fields
    [run] = run_log.recent_runs(tmp_path)
    assert (run["outcomes"], run["unreadable"]) == ("1 built", 3)
    assert "3 lines of this log couldn't be read" in run_log.run_label(run)


def test_a_log_with_no_readable_line_is_left_out(tmp_path):
    (tmp_path / "logs").mkdir()
    (tmp_path / "logs" / "run_20260918-141503.jsonl").write_text("not json\n")
    (tmp_path / "logs" / "notes.txt").write_text("not a log")
    assert run_log.recent_runs(tmp_path) == []


def test_the_panel_groups_interleaved_companies_and_keeps_each_one_s_step_order(tmp_path):
    write_run(tmp_path, "2026-09-18 14:15:03", [
        ("Northwind", run_log.CLEAN, 0.1, run_log.OK), ("Alderpeak", run_log.CLEAN, 0.1, run_log.OK),
        ("Northwind", run_log.METRICS, 0.1, run_log.OK), ("Alderpeak", run_log.METRICS, 0.1, run_log.OK)])
    [run] = run_log.recent_runs(tmp_path)
    assert [row[:2] for row in run_log.step_rows(run)] == [
        ["Northwind", run_log.CLEAN], ["Northwind", run_log.METRICS],
        ["Alderpeak", run_log.CLEAN], ["Alderpeak", run_log.METRICS]]


def test_the_panel_rows_are_one_per_line_in_plain_words(tmp_path):
    write_run(tmp_path, "2026-09-18 14:15:03", [("Northwind", run_log.EXCEL, 1.234, run_log.FAILED,
                                                  f"ValueError: cell B7 {EM_DASH} unreadable")])
    [run] = run_log.recent_runs(tmp_path)
    assert run_log.step_rows(run) == [["Northwind", run_log.EXCEL, "1.2", run_log.FAILED,
                                       "ValueError: cell B7: unreadable"]]


# ---------------------------------------------------------------------------
# The web page
# ---------------------------------------------------------------------------

@pytest.fixture
def folders(tmp_path):
    """(data folder with a copy of Northwind's workbook, empty output folder)."""
    data_dir, output_dir = tmp_path / "data", tmp_path / "output"
    data_dir.mkdir()
    (data_dir / "northwind.xlsx").write_bytes(NORTHWIND.read_bytes())
    return data_dir, output_dir


def test_generate_on_the_page_writes_a_web_page_log(folders):
    data_dir, output_dir = folders
    outcome = portfolio.generate_company(data_dir / "northwind.xlsx", main.load_config(), False, output_dir)
    assert outcome["ok"]
    lines = lines_of(only_log(output_dir))
    assert {line["source"] for line in lines} == {run_log.WEB_PAGE}
    assert steps_for(lines, "Northwind")[-1] == (run_log.WHOLE_COMPANY, main.BUILT)


def test_generate_all_on_the_page_is_one_run(folders):
    data_dir, output_dir = folders
    (data_dir / "alderpeak.xlsx").write_bytes(ALDERPEAK.read_bytes())
    (data_dir / "broken.xlsx").write_text("not a workbook")
    portfolio.generate_all(main.load_config(), False, data_dir, output_dir)
    [run] = run_log.recent_runs(output_dir)
    assert (run["companies"], run["outcomes"]) == (3, "2 built, 1 failed")
