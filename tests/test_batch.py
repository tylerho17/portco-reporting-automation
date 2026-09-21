"""Tests for batch resilience (final Task 7): --resume, --max-cost, --timeout, --workers, and rate-limit retries.

Every test runs real workbooks from data/ through main.run_batch into a temporary folder, with a fake
Claude client, so the real output/ folder is untouched and no API call can happen (a guard makes
creating a real client fail). The fake clients:
- ScriptedClient: answers (or raises) in the order it's given, e.g. two rate limits then an answer
- HangingClient: never answers for one company until the test lets it go (for --timeout)
- CountingClient: records how many calls are in flight at once (for --workers)
Waits are made short by setting resilience.py's wait lengths in the test, never by sleeping for real.

Run from the project folder:  pytest
"""

import csv
import json
import threading
import time
from pathlib import Path
from types import SimpleNamespace

import anthropic
import httpx2
import pytest

import analyze
import main
import resilience
from analyze import BoardSummary
from provenance import file_sha256, manifest_path, read_manifest, save_manifest
from valid_answer import summary_dict

PROJECT_DIR = Path(__file__).parent.parent
DATA_DIR = PROJECT_DIR / "data"
NORTHWIND, ALDERPEAK, FERNHOLLOW = (DATA_DIR / f"{name}.xlsx" for name in ("northwind", "alderpeak", "fernhollow"))
TIMEOUT_SECONDS = 5   # one company takes about 1 s here, so 5 s only runs out on a company that hangs


@pytest.fixture(autouse=True)
def no_real_client(monkeypatch):
    """Creating a real Anthropic client fails the test, so no test here can reach the API."""
    def refuse(*args, **kwargs):
        raise AssertionError("a test tried to create a real Anthropic client")
    monkeypatch.setattr(analyze.anthropic, "Anthropic", refuse)


@pytest.fixture(autouse=True)
def short_waits(monkeypatch):
    """Rate-limit waits of hundredths of a second instead of 5, 10, 20, 40 s."""
    monkeypatch.setattr(resilience, "FIRST_WAIT_SECONDS", 0.01)
    monkeypatch.setattr(resilience, "MAX_WAIT_SECONDS", 0.05)


# ---------------------------------------------------------------------------
# Fake clients
# ---------------------------------------------------------------------------

def summary():
    """A valid answer whose only numbers are thresholds every company's payload holds (valid_answer.py)."""
    return BoardSummary.model_validate(summary_dict())


def answer(input_tokens=100, output_tokens=50):
    """What client.messages.parse returns: the valid summary, with the token counts given."""
    return SimpleNamespace(content=[SimpleNamespace(type="text", text=summary().model_dump_json())],
                           parsed_output=summary(), stop_reason="end_turn",
                           usage=SimpleNamespace(input_tokens=input_tokens, output_tokens=output_tokens))


def rate_limit_error(retry_after=None):
    """The error the SDK raises for HTTP 429, with the retry-after header the API may send."""
    headers = {} if retry_after is None else {"retry-after": str(retry_after)}
    request = httpx2.Request("POST", "https://api.anthropic.com/v1/messages")
    return anthropic.RateLimitError("rate limited", response=httpx2.Response(429, headers=headers, request=request),
                                    body=None)


class ScriptedClient:
    """Each call takes the next step: an exception is raised, anything else is returned. The last step repeats."""

    def __init__(self, *steps):
        self.steps, self.calls = list(steps), 0
        self.messages = self   # analyze.py calls client.messages.parse(...)

    def parse(self, **kwargs):
        step = self.steps[min(self.calls, len(self.steps) - 1)]
        self.calls += 1
        if isinstance(step, Exception):
            raise step
        return step


class HangingClient:
    """Answers at once, except for one company: that call waits until the test sets `release`.

    release_when: another company whose call sets `release` itself, so the hung call wakes while the
    batch is still running.
    """

    def __init__(self, company, release_when=None):
        self.company, self.release_when, self.release, self.calls = company, release_when, threading.Event(), []
        self.messages = self

    def parse(self, **kwargs):
        company = json.loads(kwargs["messages"][0]["content"].split("\n\n", 1)[1])["company"]
        self.calls.append(company)
        if company == self.company:
            self.release.wait(60)
        elif company == self.release_when:
            self.release.set()
            time.sleep(1)   # give the woken company time to finish its work before this one does
        return answer()


class CountingClient:
    """Records the most calls in flight at once. With `together`, each call waits until that many have arrived."""

    def __init__(self, together=None):
        self.lock, self.in_flight, self.most = threading.Lock(), 0, 0
        self.barrier = threading.Barrier(together, timeout=20) if together else None
        self.messages = self

    def parse(self, **kwargs):
        with self.lock:
            self.in_flight += 1
            self.most = max(self.most, self.in_flight)
        try:
            if self.barrier:
                self.barrier.wait()   # raises BrokenBarrierError if they never all arrive together
            else:
                time.sleep(0.2)       # long enough that an overlapping call would be seen
            return answer()
        finally:
            with self.lock:
                self.in_flight -= 1


def batch(tmp_path, paths=(NORTHWIND,), skip_ai=False, client=None, **options):
    """main.run_batch into tmp_path."""
    return main.run_batch(list(paths), main.load_config(), skip_ai, client=client, output_dir=tmp_path, **options)


def events(tmp_path, workbook=NORTHWIND):
    """The batch events recorded in a company's manifest: [(event, detail), ...]."""
    manifest = read_manifest(manifest_path(workbook, tmp_path))
    return [(item["event"], item["detail"]) for item in manifest.get("batch_events", [])]


def output_files(tmp_path, stem="northwind"):
    """{file name: bytes} of a company's outputs, to prove they were left alone."""
    return {path.name: path.read_bytes() for path in tmp_path.glob(f"{stem}_*") if path.is_file()}


# ---------------------------------------------------------------------------
# Rate limits: wait and try again (resilience.RateLimitRetry)
# ---------------------------------------------------------------------------

def test_rate_limit_waits_double_each_time_and_then_the_answer_comes_back(monkeypatch):
    monkeypatch.setattr(resilience, "FIRST_WAIT_SECONDS", 5)
    monkeypatch.setattr(resilience, "MAX_WAIT_SECONDS", 60)
    waited = []
    client = resilience.RateLimitRetry(ScriptedClient(rate_limit_error(), rate_limit_error(), rate_limit_error(),
                                                      answer()), wait=waited.append)
    assert client.messages.parse(model="m").usage.input_tokens == 100
    assert waited == [5, 10, 20] and client.waits == [5, 10, 20]


def test_rate_limit_uses_the_retry_after_the_api_sent_but_never_waits_past_the_cap(monkeypatch):
    monkeypatch.setattr(resilience, "MAX_WAIT_SECONDS", 60)
    waited = []
    client = resilience.RateLimitRetry(ScriptedClient(rate_limit_error(3), rate_limit_error(999), answer()),
                                       wait=waited.append)
    client.messages.parse(model="m")
    assert waited == [3, 60]


def test_rate_limit_gives_up_after_the_last_retry():
    waited = []
    fake = ScriptedClient(rate_limit_error())
    client = resilience.RateLimitRetry(fake, wait=waited.append)
    with pytest.raises(anthropic.RateLimitError):
        client.messages.parse(model="m")
    assert len(waited) == resilience.MAX_RATE_LIMIT_RETRIES
    assert fake.calls == resilience.MAX_RATE_LIMIT_RETRIES + 1   # the first try and every retry


def server_error():
    """The error the SDK raises for HTTP 500 (after its own quick retries)."""
    request = httpx2.Request("POST", "https://api.anthropic.com/v1/messages")
    return anthropic.InternalServerError("server error", response=httpx2.Response(500, request=request), body=None)


@pytest.mark.parametrize("error", [anthropic.AnthropicError("simulated outage"), server_error()])
def test_other_api_errors_are_not_retried(error):
    waited = []
    fake = ScriptedClient(error)
    with pytest.raises(type(error)):
        resilience.RateLimitRetry(fake, wait=waited.append).messages.parse(model="m")
    assert waited == [] and fake.calls == 1


def test_a_rate_limited_company_still_gets_its_ai_text_and_the_retries_are_recorded(tmp_path):
    client = ScriptedClient(rate_limit_error(), rate_limit_error(), answer())
    result = batch(tmp_path, client=client)[0]
    assert main.result_text(result) == "OK"
    assert result["notes"] == ["2 rate-limit retries (waited 0.03 s)"]
    ai = read_manifest(manifest_path(NORTHWIND, tmp_path))["ai"]
    assert ai["rate_limit_waits_s"] == [0.01, 0.02] and ai["validation"] == "passed"


def test_rate_limits_that_never_clear_give_the_placeholder_not_a_failed_company(tmp_path):
    result = batch(tmp_path, client=ScriptedClient(rate_limit_error()))[0]
    assert main.result_text(result) == "OK (AI failed)"
    # Waits of 0.01, 0.02, 0.04, then 0.05 (the cap in these tests): 0.12 s in all.
    assert result["notes"] == [f"{resilience.MAX_RATE_LIMIT_RETRIES} rate-limit retries (waited 0.12 s)"]
    assert read_manifest(manifest_path(NORTHWIND, tmp_path))["ai"]["validation"].startswith(
        "failed: Claude's API was still turning requests away (rate limit)")


def test_a_skipped_ai_step_records_no_rate_limit_waits(tmp_path):
    batch(tmp_path, skip_ai=True)
    assert read_manifest(manifest_path(NORTHWIND, tmp_path))["ai"]["rate_limit_waits_s"] is None


# ---------------------------------------------------------------------------
# --resume: skip a company whose outputs were built from today's workbook and thresholds
# ---------------------------------------------------------------------------

def test_resume_skips_a_company_whose_outputs_are_up_to_date(tmp_path):
    batch(tmp_path, client=ScriptedClient(answer()))
    before = output_files(tmp_path)
    client = ScriptedClient(answer())
    result = batch(tmp_path, client=client, resume=True)[0]
    assert main.result_text(result) == "SKIPPED (--resume): outputs match the workbook and config.yaml"
    assert client.calls == 0
    assert result["tripped"] and result["quarter"] == "Q2 2026"   # the table still shows the flags
    after = output_files(tmp_path)
    assert events(tmp_path) == [("skipped", "outputs match the workbook and config.yaml (--resume)")]
    after.pop("northwind_manifest.json"), before.pop("northwind_manifest.json")   # only the event was added
    assert after == before


def test_every_skip_is_kept_not_only_the_last(tmp_path):
    batch(tmp_path, skip_ai=True)
    batch(tmp_path, skip_ai=True, resume=True)
    batch(tmp_path, skip_ai=True, resume=True, timeout=TIMEOUT_SECONDS)   # skipped before any timeout applies
    assert [event for event, _ in events(tmp_path)] == ["skipped", "skipped"]


def test_the_batch_events_are_kept_when_the_company_is_built_again(tmp_path):
    batch(tmp_path, skip_ai=True)
    batch(tmp_path, skip_ai=True, resume=True)       # skipped: recorded
    batch(tmp_path, skip_ai=True)                    # built again: the record of the skip stays
    assert [event for event, _ in events(tmp_path)] == ["skipped"]


def test_without_resume_an_up_to_date_company_is_built_again(tmp_path):
    batch(tmp_path, skip_ai=True)
    result = batch(tmp_path, skip_ai=True)[0]
    assert main.result_text(result) == "OK (AI skipped)" and events(tmp_path) == []


def test_resume_rebuilds_when_the_workbook_changed(tmp_path):
    batch(tmp_path, skip_ai=True)
    manifest = read_manifest(manifest_path(NORTHWIND, tmp_path))
    manifest["input"]["sha256"] = "0" * 64   # as if the deck was built from another version of the workbook
    save_manifest(manifest_path(NORTHWIND, tmp_path), manifest)
    result = batch(tmp_path, skip_ai=True, resume=True)[0]
    assert main.result_text(result) == "OK (AI skipped)"
    assert result["notes"] == ["rebuilt (--resume): the workbook has changed"]
    assert read_manifest(manifest_path(NORTHWIND, tmp_path))["input"]["sha256"] == file_sha256(NORTHWIND)


@pytest.mark.parametrize("change, reason", [
    (lambda m: m["config"].update(sha256="0" * 64), "config.yaml has changed"),
    (lambda m: m.update(mapping={"file": "mappings/northwind.yaml", "sha256": "0" * 64}),
     "the column mapping has changed"),
    (lambda m: m["deck"].update(draft=True), "the deck was built with a different --draft setting"),
    (lambda m: m["deck"].update(appendix=True), "the deck was built with a different --appendix setting"),
    (lambda m: m.update(approval={"reviewer": "Tyler Ho", "approved_at": "2999-01-01T00:00:00"}),
     "approved after the deck was built, so its footer still says not reviewed"),
    (lambda m: m.pop("deck"), "the manifest has no deck"),
])
def test_resume_says_why_a_company_has_to_be_rebuilt(tmp_path, change, reason):
    batch(tmp_path, skip_ai=True)
    manifest = read_manifest(manifest_path(NORTHWIND, tmp_path))
    change(manifest)
    save_manifest(manifest_path(NORTHWIND, tmp_path), manifest)
    problem = resilience.resume_problem(NORTHWIND, tmp_path, main.CONFIG_PATH, want_ai=False, draft=False)
    assert problem == reason


def test_resume_rebuilds_when_an_output_file_is_missing(tmp_path):
    batch(tmp_path, skip_ai=True)
    (tmp_path / "northwind_board_memo.pdf").unlink()
    problem = resilience.resume_problem(NORTHWIND, tmp_path, main.CONFIG_PATH, want_ai=False, draft=False)
    assert problem == "northwind_board_memo.pdf is missing"


def test_resume_with_no_earlier_run_builds_the_company(tmp_path):
    assert resilience.resume_problem(NORTHWIND, tmp_path, main.CONFIG_PATH, False, False) == "no earlier run"
    result = batch(tmp_path, skip_ai=True, resume=True)[0]
    assert main.result_text(result) == "OK (AI skipped)" and result["notes"] == ["rebuilt (--resume): no earlier run"]


def test_resume_asking_for_ai_rebuilds_a_deck_that_has_no_ai_text(tmp_path):
    batch(tmp_path, skip_ai=True)
    client = ScriptedClient(answer())
    result = batch(tmp_path, client=client, resume=True)[0]
    assert main.result_text(result) == "OK" and client.calls == 1
    assert result["notes"] == ["rebuilt (--resume): the deck has no AI text"]


def test_resume_with_skip_ai_keeps_a_deck_that_has_ai_text(tmp_path):
    batch(tmp_path, client=ScriptedClient(answer()))
    result = batch(tmp_path, skip_ai=True, resume=True)[0]
    assert result["outcome"] == main.SKIPPED


def test_a_resume_rebuild_reuses_a_saved_analysis_of_the_same_numbers(tmp_path):
    batch(tmp_path, client=ScriptedClient(answer()))
    client = ScriptedClient(answer())
    result = batch(tmp_path, client=client, resume=True, draft=True)[0]   # only the watermark setting changed
    assert main.result_text(result) == "OK" and client.calls == 0
    assert read_manifest(manifest_path(NORTHWIND, tmp_path))["deck"]["draft"] is True


# ---------------------------------------------------------------------------
# --max-cost: stop starting companies once the AI spend reaches the ceiling
# ---------------------------------------------------------------------------

def test_the_batch_stops_once_the_spend_reaches_the_ceiling(tmp_path):
    # 100,000 tokens in and 50,000 out at $2 and $10 per million: $0.70 a company.
    client = ScriptedClient(answer(input_tokens=100_000, output_tokens=50_000))
    batch(tmp_path, paths=[FERNHOLLOW], skip_ai=True)   # an earlier run, so Fernhollow has a manifest
    results = batch(tmp_path, paths=[NORTHWIND, ALDERPEAK, FERNHOLLOW], client=client, max_cost=1.00)
    assert [result["outcome"] for result in results] == [main.BUILT, main.BUILT, main.STOPPED]
    assert client.calls == 2
    assert [result["cost_usd"] for result in results[:2]] == [0.7, 0.7]
    stop = "AI spend $1.40 reached the --max-cost ceiling of $1.00"
    assert main.result_text(results[2]) == f"STOPPED: {stop}"
    assert events(tmp_path, FERNHOLLOW) == [("stopped", stop)]


def test_a_spend_exactly_at_the_ceiling_stops_the_batch(tmp_path):
    client = ScriptedClient(answer(input_tokens=100_000, output_tokens=50_000))
    results = batch(tmp_path, paths=[NORTHWIND, ALDERPEAK, FERNHOLLOW], client=client, max_cost=1.40)
    assert [result["outcome"] for result in results] == [main.BUILT, main.BUILT, main.STOPPED]


def test_the_spend_below_the_ceiling_lets_every_company_run(tmp_path):
    client = ScriptedClient(answer(input_tokens=100_000, output_tokens=50_000))
    results = batch(tmp_path, paths=[NORTHWIND, ALDERPEAK], client=client, max_cost=1.50)
    assert [result["outcome"] for result in results] == [main.BUILT, main.BUILT]


def test_skip_ai_spends_nothing_so_the_ceiling_never_stops_it(tmp_path):
    results = batch(tmp_path, paths=[NORTHWIND, ALDERPEAK], skip_ai=True, max_cost=0.01)
    assert [result["outcome"] for result in results] == [main.BUILT, main.BUILT]


def test_a_skipped_company_is_not_stopped_by_the_ceiling(tmp_path):
    batch(tmp_path, paths=[ALDERPEAK], client=ScriptedClient(answer()))   # Alderpeak is up to date, AI text and all
    client = ScriptedClient(answer(input_tokens=1_000_000, output_tokens=1))   # $2.00 on Northwind
    results = batch(tmp_path, paths=[NORTHWIND, ALDERPEAK], client=client, max_cost=1.00, resume=True)
    assert [result["outcome"] for result in results] == [main.BUILT, main.SKIPPED]
    assert results[0]["cost_usd"] == 2.0


def test_failed_validation_counts_toward_the_spend(tmp_path):
    bad = answer(input_tokens=100_000, output_tokens=50_000)
    bad.parsed_output = summary().model_copy(update={"headline": "ARR grew 555.5% this quarter."})
    results = batch(tmp_path, paths=[NORTHWIND, ALDERPEAK], client=ScriptedClient(bad), max_cost=1.00)
    assert main.result_text(results[0]) == "OK (AI failed)" and results[0]["cost_usd"] == 1.4   # two attempts
    assert results[1]["outcome"] == main.STOPPED


# ---------------------------------------------------------------------------
# --timeout: a company that runs too long is given up on; the batch moves on
# ---------------------------------------------------------------------------

def test_a_company_past_its_timeout_is_given_up_and_the_next_one_still_runs(tmp_path):
    batch(tmp_path, skip_ai=True)                    # Northwind's earlier outputs
    before = output_files(tmp_path)
    client = HangingClient("Northwind")
    start = time.perf_counter()
    results = batch(tmp_path, paths=[NORTHWIND, ALDERPEAK], client=client, timeout=TIMEOUT_SECONDS)
    assert time.perf_counter() - start < TIMEOUT_SECONDS + 10
    assert results[0]["outcome"] == main.TIMED_OUT
    assert results[0]["error"] == f"timed out after {TIMEOUT_SECONDS} s"   # so the batch exits 1
    assert main.result_text(results[0]) == (f"TIMED OUT after {TIMEOUT_SECONDS} s: earlier outputs left as they "
                                            f"were")
    assert main.result_text(results[1]) == "OK"
    assert events(tmp_path) == [("timed out", f"gave up after {TIMEOUT_SECONDS} s")]

    client.release.set()                             # the hung call returns: nothing it builds may land
    wait_for_empty_staging(tmp_path)
    after = output_files(tmp_path)
    after.pop("northwind_manifest.json"), before.pop("northwind_manifest.json")
    assert after == before


def wait_for_empty_staging(tmp_path, seconds=20):
    """Wait until the given-up company's thread has stopped and its private folder is gone."""
    staging = tmp_path / resilience.STAGING_FOLDER
    deadline = time.perf_counter() + seconds
    while time.perf_counter() < deadline:
        leftovers = [path for path in staging.rglob("*") if path.is_file()] if staging.exists() else []
        if not leftovers:
            return
        time.sleep(0.1)
    raise AssertionError(f"files still in {staging}")


def test_a_company_that_wakes_after_its_timeout_while_the_batch_runs_on_lands_nothing(tmp_path):
    batch(tmp_path, skip_ai=True)
    before = output_files(tmp_path)
    client = HangingClient("Northwind", release_when="Alderpeak")
    results = batch(tmp_path, paths=[NORTHWIND, ALDERPEAK, FERNHOLLOW], client=client, timeout=TIMEOUT_SECONDS)
    assert [result["outcome"] for result in results] == [main.TIMED_OUT, main.BUILT, main.BUILT]
    wait_for_empty_staging(tmp_path)
    after = output_files(tmp_path)
    after.pop("northwind_manifest.json"), before.pop("northwind_manifest.json")
    assert after == before


def test_the_manifest_is_moved_into_the_output_folder_last(tmp_path, monkeypatch):
    stage = tmp_path / "stage"
    stage.mkdir()
    for name in ("northwind_manifest.json", "northwind_board_pack.pptx", "northwind_metrics.xlsx"):
        (stage / name).write_text(name)
    moved = []
    real_replace = resilience.os.replace
    monkeypatch.setattr(resilience.os, "replace", lambda source, target: moved.append(Path(target).name)
                        or real_replace(source, target))
    resilience.commit_stage(stage, tmp_path / "out")
    assert len(moved) == 3 and moved[-1] == "northwind_manifest.json"
    assert not stage.exists()


def test_a_timed_out_company_with_no_earlier_run_leaves_no_outputs(tmp_path):
    client = HangingClient("Northwind")
    results = batch(tmp_path, client=client, timeout=TIMEOUT_SECONDS)
    client.release.set()
    wait_for_empty_staging(tmp_path)
    assert results[0]["outcome"] == main.TIMED_OUT
    assert output_files(tmp_path) == {}
    assert main.save_batch_manifest(results, {}, 0, tmp_path / "b.json").exists()   # its only record


def test_the_next_batch_clears_what_a_killed_batch_left_in_staging(tmp_path):
    leftover = tmp_path / resilience.STAGING_FOLDER / "northwind_abc" / "northwind_board_pack.pptx"
    leftover.parent.mkdir(parents=True)
    leftover.write_text("half a deck")
    batch(tmp_path, skip_ai=True)
    assert not (tmp_path / resilience.STAGING_FOLDER).exists()


def test_a_cancelled_company_stops_at_its_next_step():
    control = resilience.CompanyControl()
    control.checkpoint()                             # not cancelled: carries on
    control.cancel.set()
    with pytest.raises(resilience.Cancelled):
        control.checkpoint()
    with pytest.raises(resilience.Cancelled):
        control.wait(30)                             # a rate-limit wait ends at once


# ---------------------------------------------------------------------------
# --workers: companies side by side
# ---------------------------------------------------------------------------

def test_three_workers_run_three_companies_at_the_same_time(tmp_path):
    client = CountingClient(together=3)
    results = batch(tmp_path, paths=[NORTHWIND, ALDERPEAK, FERNHOLLOW], client=client, workers=3)
    assert client.most == 3
    assert [result["company"] for result in results] == ["Northwind", "Alderpeak", "Fernhollow"]
    assert [main.result_text(result) for result in results] == ["OK"] * 3


def test_one_worker_by_default_runs_one_company_at_a_time(tmp_path):
    client = CountingClient()
    batch(tmp_path, paths=[NORTHWIND, ALDERPEAK], client=client)
    assert client.most == 1


def test_each_company_prints_its_lines_together_even_side_by_side(tmp_path, capsys):
    batch(tmp_path, paths=[NORTHWIND, ALDERPEAK, FERNHOLLOW], client=CountingClient(together=3), workers=3)
    printed = capsys.readouterr().out.split("=== Summary ===")[0]
    blocks = printed.split("\n=== ")[1:]              # one block per company, from its "=== name.xlsx ===" line
    assert sorted(block.split(".xlsx", 1)[0] for block in blocks) == ["alderpeak", "fernhollow", "northwind"]
    for block in blocks:
        stem = block.split(".xlsx", 1)[0]
        file_lines = [line for line in block.splitlines() if str(tmp_path) in line]
        assert len(file_lines) >= 5, block             # Excel, AI commentary, Deck, Memo, Manifest
        assert all(f"{stem}_" in line for line in file_lines), block   # no other company's line in the block


def test_printed_paths_name_the_output_folder_not_the_private_one(tmp_path, capsys):
    batch(tmp_path, skip_ai=True)
    printed = capsys.readouterr().out
    assert f"✓ Deck: {tmp_path / 'northwind_board_pack.pptx'}" in printed
    assert resilience.STAGING_FOLDER not in printed


def test_a_failed_company_leaves_its_earlier_outputs_as_they_were(tmp_path, monkeypatch):
    batch(tmp_path, client=ScriptedClient(answer()))
    before = output_files(tmp_path)
    result = batch(tmp_path, client=ScriptedClient(KeyError("simulated bug")))[0]
    assert result["error"] == main.error_text(KeyError("simulated bug"))   # "unexpected problem ... (KeyError: ...)"
    after = output_files(tmp_path)
    assert events(tmp_path) == [("failed", main.error_text(KeyError("simulated bug")))]
    after.pop("northwind_manifest.json"), before.pop("northwind_manifest.json")
    assert after == before


# ---------------------------------------------------------------------------
# The summary table, the CSV and the batch manifest
# ---------------------------------------------------------------------------

def test_the_summary_names_every_skip_timeout_and_stop(tmp_path, capsys):
    results = [
        {"company": "Alderpeak", "outcome": main.SKIPPED, "error": None, "quarter": "Q2 2026", "flags_total": 9,
         "tripped": [], "cannot_evaluate": [], "gap_count": 0, "blank_quarters": [], "ai": None, "cost_usd": None,
         "notes": []},
        {"company": "Northwind", "outcome": main.TIMED_OUT, "error": "timed out after 300 s", "timeout": 300,
         "cost_usd": None, "notes": []},
        {"company": "Fernhollow", "outcome": main.STOPPED, "error": "stopped",
         "detail": "AI spend $1.40 reached the --max-cost ceiling of $1.00", "cost_usd": None, "notes": []},
    ]
    main.print_summary(results)
    printed = capsys.readouterr().out
    assert "SKIPPED (--resume): outputs match the workbook and config.yaml" in printed
    assert "TIMED OUT after 300 s: earlier outputs left as they were" in printed
    assert "STOPPED: AI spend $1.40 reached the --max-cost ceiling of $1.00" in printed
    assert "1 skipped (--resume), 1 timed out, 1 stopped" in printed


def test_the_csv_has_the_ai_cost_and_the_notes(tmp_path):
    client = ScriptedClient(rate_limit_error(), answer(input_tokens=100_000, output_tokens=50_000))
    results = batch(tmp_path, client=client)
    path = main.write_summary_csv(results, tmp_path / "batch_summary.csv")
    with open(path, newline="") as file:
        row = list(csv.DictReader(file))[0]
    assert row["AI cost (USD)"] == "0.7" and row["Notes"] == "1 rate-limit retry (waited 0.01 s)"


def test_the_batch_manifest_records_every_outcome_and_the_spend(tmp_path):
    results = [
        {"company": "Northwind", "outcome": main.BUILT, "error": None, "ai": "ok", "cost_usd": 0.7,
         "notes": ["1 rate-limit retry (waited 5 s)"], "quarter": "Q2 2026", "tripped": [], "flags_total": 9,
         "cannot_evaluate": [], "gap_count": 0, "blank_quarters": []},
        {"company": "Fernhollow", "outcome": main.STOPPED, "error": "stopped",
         "detail": "AI spend $1.40 reached the --max-cost ceiling of $1.00", "cost_usd": None, "notes": []},
    ]
    options = {"resume": False, "max_cost": 1.0, "timeout": None, "workers": 1, "skip_ai": False, "draft": False}
    path = main.save_batch_manifest(results, options, spend=1.4, path=tmp_path / "batch_manifest.json")
    saved = json.loads(path.read_text())
    assert saved["options"] == options and saved["ai_spend_usd"] == 1.4
    assert saved["companies"] == [
        {"company": "Northwind", "outcome": "built", "result": "OK", "cost_usd": 0.7,
         "notes": ["1 rate-limit retry (waited 5 s)"]},
        {"company": "Fernhollow", "outcome": "stopped",
         "result": "STOPPED: AI spend $1.40 reached the --max-cost ceiling of $1.00", "cost_usd": None, "notes": []},
    ]
    assert saved["code"]["commit"]


# ---------------------------------------------------------------------------
# Command line
# ---------------------------------------------------------------------------

def run_main_capturing(monkeypatch, argv, results=()):
    """main.main with run_batch, the CSV and the batch manifest replaced. Returns (exit code, run_batch's kwargs)."""
    seen = {}
    monkeypatch.setattr(main, "run_batch", lambda *args, **kwargs: seen.update(kwargs) or list(results))
    monkeypatch.setattr(main, "write_summary_csv", lambda results: PROJECT_DIR / "output" / "batch_summary.csv")
    monkeypatch.setattr(main, "save_batch_manifest", lambda *args, **kwargs: PROJECT_DIR / "output" / "b.json")
    return main.main(argv), seen


def test_the_new_options_reach_the_batch(monkeypatch):
    code, seen = run_main_capturing(monkeypatch, ["--all", "--skip-ai", "--resume", "--max-cost", "2.5",
                                                  "--timeout", "300", "--workers", "4"])
    assert code == 0
    assert (seen["resume"], seen["max_cost"], seen["timeout"], seen["workers"]) == (True, 2.5, 300, 4)


def test_the_defaults_are_one_worker_no_timeout_no_ceiling_no_resume(monkeypatch):
    _, seen = run_main_capturing(monkeypatch, ["--all", "--skip-ai"])
    assert (seen["resume"], seen["max_cost"], seen["timeout"], seen["workers"]) == (False, None, None, 1)


@pytest.mark.parametrize("option, value", [("--workers", "0"), ("--max-cost", "0"), ("--max-cost", "-1"),
                                           ("--timeout", "0"), ("--workers", "two")])
def test_a_nonsense_option_value_stops_with_a_usage_error(option, value, capsys):
    with pytest.raises(SystemExit) as stop:
        main.parse_args(["--all", option, value])
    assert stop.value.code == 2
    assert option in capsys.readouterr().err


def test_exit_code_is_1_after_a_timeout_or_a_stop_and_0_after_skips(monkeypatch):
    skipped = {"company": "Alderpeak", "outcome": main.SKIPPED, "error": None, "quarter": "Q2 2026",
               "flags_total": 9, "tripped": [], "cannot_evaluate": [], "gap_count": 0, "blank_quarters": [],
               "ai": None, "cost_usd": None, "notes": []}
    timed_out = {"company": "Northwind", "outcome": main.TIMED_OUT, "error": "timed out after 5 s", "timeout": 5,
                 "cost_usd": None, "notes": []}
    stopped = {"company": "Fernhollow", "outcome": main.STOPPED, "error": "stopped", "detail": "x",
               "cost_usd": None, "notes": []}
    assert run_main_capturing(monkeypatch, ["--all", "--skip-ai"], [skipped])[0] == 0
    assert run_main_capturing(monkeypatch, ["--all", "--skip-ai"], [skipped, timed_out])[0] == 1
    assert run_main_capturing(monkeypatch, ["--all", "--skip-ai"], [skipped, stopped])[0] == 1
