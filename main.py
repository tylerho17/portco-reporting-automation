"""Batch runner: turn one KPI workbook, or every workbook in data/, into board-pack outputs.

Steps for each company:
1. Clean the workbook (clean.py)
2. Compute metrics, flags and data gaps (metrics.py)
3. Save output/<company>_metrics.xlsx (excel_output.py)
4. AI commentary (analyze.py): output/<company>_analysis.json, saved whether it passed or failed
5. Deck (build_deck.py): output/<company>_board_pack.pptx
6. Memo (memo.py): output/<company>_board_memo.docx and .pdf, from the same numbers and analysis
7. Manifest (provenance.py): output/<company>_manifest.json, with this run's results and the
   earlier run they were compared with (diff_runs.py, Task 10: what changed since the last run)

The deck and the memo are always built, because their numbers come from Python (CLAUDE.md decisions K and L):
- AI passed validation      -> the AI text is on slide 4 (AI commentary), result "OK"
- AI failed (validation failed after the retry, or an API error)
                            -> "AI summary unavailable" on slide 4, result "OK (AI failed)"
- --skip-ai                 -> the same placeholder with no API call, result "OK (AI skipped)"
Without --skip-ai, the API key is checked before any company runs.

Every deck's footer says whether a person has approved it ("AI-drafted | not reviewed", or
"reviewed by NAME on DATE" after approve.py). --draft also stamps "DRAFT - NOT REVIEWED" across
every slide of a deck nobody has approved; it's off by default.

One company failing never stops the batch. The run ends with a summary table
(company, flags tripped, data gaps, result, notes), also saved as output/batch_summary.csv, a warning
if companies end on different quarters or any AI summary is unavailable, and exit code 1 if any
company failed, timed out or was stopped. "OK (AI failed)" is not a failed company: its outputs were all built.

Long batches (final Task 7, the parts are in resilience.py):
- --resume         skip a company whose outputs were built from today's workbook and config.yaml
                   (by hash); a company that is rebuilt anyway says why in the Notes column
- --max-cost 5     start no more companies once the AI spend reaches $5
- --timeout 300    give up on a company after 300 s; its earlier outputs are left as they were
- --workers 4      run 4 companies side by side (default 1)
- a rate-limited API call waits and tries again (resilience.RateLimitRetry)
Every skip, timeout, stop and failure goes in the company's manifest, the summary table, and
output/batch_manifest.json (with the options and the batch's total AI spend).

Each company is built in a private folder and moved into output/ only when it succeeds, so a
company that fails or times out never leaves half-replaced files.

Run: python main.py data/northwind.xlsx
     python main.py --all --skip-ai
     python main.py --all --draft
     python main.py --all --resume --max-cost 5 --timeout 300 --workers 4
"""

import argparse
import csv
import json
import os
import queue
import sys
import threading
import time
import traceback
from pathlib import Path
from types import SimpleNamespace

import anthropic
from dotenv import load_dotenv

from analyze import PROMPT_VERSION, AnalysisError, analyze, build_payload, payload_to_text, save_analysis
from build_deck import (PLACEHOLDER_TEXT, analysis_details, analysis_path, collect_deck_data, commentary_slide,
                        deck_path, load_analysis, save_deck, slide_number)
from clean import clean_workbook
from compare_models import run_cost
from diff_runs import HEADING, baseline, move_settings, report_for, run_results, summary_text
from excel_output import save_metrics_workbook
from mapping import mapping_record
from memo import MEMO_UNAVAILABLE, save_memo
from metrics import (CANNOT_EVALUATE, CONFIG_PATH, TRIP, compute_metrics, data_gaps, evaluate_flags, load_config,
                     metric_reasons)
from provenance import build_manifest, git_commit, manifest_path, read_manifest, save_manifest, timestamp
from resilience import (UP_TO_DATE, Cancelled, CompanyControl, RateLimitRetry, SpendMeter, clear_old_stages,
                        commit_stage, discard_stage, make_stage, rate_limit_note, record_batch_event, resume_problem,
                        routed_stdout)
from text_fit import TextDoesNotFitError

PROJECT_DIR = Path(__file__).parent
DATA_DIR = PROJECT_DIR / "data"
OUTPUT_DIR = PROJECT_DIR / "output"
SUMMARY_CSV_PATH = OUTPUT_DIR / "batch_summary.csv"
BATCH_MANIFEST_PATH = OUTPUT_DIR / "batch_manifest.json"

# What happened to a company in a batch.
BUILT, SKIPPED, TIMED_OUT, STOPPED, FAILED = "built", "skipped", "timed out", "stopped", "failed"

# Errors caused by a bad input file (clean.py raises ValueError with a clear message; a missing
# or unreadable file raises OSError), or by text that can't fit a slide (text_fit.py names the slide
# and the box). Anything else is probably a bug, so its traceback is printed.
INPUT_ERRORS = (ValueError, OSError, TextDoesNotFitError)

# What happened to the AI summary, and how the Result column says it.
AI_OK, AI_SKIPPED, AI_FAILED = "ok", "skipped", "failed"
RESULT_TEXTS = {AI_OK: "OK", AI_SKIPPED: "OK (AI skipped)", AI_FAILED: "OK (AI failed)"}

AI_REUSED = "reused a saved analysis of exactly these numbers (no API call)"   # the manifest's validation

NO_KEY_MESSAGE ="ANTHROPIC_API_KEY isn't set: add it to .env, or run with --skip-ai"


# ---------------------------------------------------------------------------
# Finding the input files
# ---------------------------------------------------------------------------

def find_workbooks(data_dir=DATA_DIR):
    """Every .xlsx in data_dir, sorted by name.

    Skips "~$" files: Excel creates those as lock files while a workbook is open.
    """
    return sorted(path for path in Path(data_dir).glob("*.xlsx") if not path.name.startswith("~$"))


def company_name(workbook_path):
    """data/northwind.xlsx -> 'Northwind' (same rule analyze.py and build_deck.py use)."""
    return Path(workbook_path).stem.title()


# In a company's own thread: its private folder, and the output folder its files will be moved to.
_private = threading.local()


def shown_path(path):
    """A path as printed: relative to the project folder when it's inside it (output/...), else in full.

    A file in a company's private folder is shown where it will end up, in the output folder.
    """
    path = Path(path)
    stage = getattr(_private, "stage", None)
    if stage is not None and path.is_relative_to(stage):
        path = _private.output_dir / path.relative_to(stage)
    return path.relative_to(PROJECT_DIR) if path.is_relative_to(PROJECT_DIR) else path


# ---------------------------------------------------------------------------
# Steps 4 and 5: AI commentary and the deck
# ---------------------------------------------------------------------------

def api_key_problem():
    """None if an API key is set (from the environment or .env), else a message saying what to do."""
    load_dotenv()  # puts ANTHROPIC_API_KEY from .env into the environment
    return None if os.environ.get("ANTHROPIC_API_KEY", "").strip() else NO_KEY_MESSAGE


def ai_cost(run_info):
    """Dollar cost of every attempt in run_info, or None when no tokens were counted (an API error)."""
    run_info = run_info or {}
    model, tokens = run_info.get("model"), (run_info.get("input_tokens"), run_info.get("output_tokens"))
    return round(run_cost(model, *tokens), 4) if model and all(tokens) else None


def ai_step(workbook_path, actuals, next_budget, config, output_dir, client=None, control=None):
    """Ask Claude for the commentary and save output/<company>_analysis.json, whether it passed or not.

    Returns {"analysis_file": the file for the deck or None, "run_info": ..., "validation": what
    happened, "rate_limit_waits": seconds waited before each retry}. Only a validation failure after
    the retry or an API error counts as "AI failed"; any other error is a bug and fails the company.
    client=None uses the real Anthropic client. A rate limit waits and tries again (resilience.py);
    what the calls cost is counted toward the batch's --max-cost ceiling (control).
    """
    control = control or CompanyControl()
    path = analysis_path(workbook_path, output_dir)
    path.unlink(missing_ok=True)  # an old analysis must never look like this run's
    payload = build_payload(company_name(workbook_path), actuals, next_budget, config)
    retrying = RateLimitRetry(client or anthropic.Anthropic(), wait=control.wait)
    try:
        summary, run_info = analyze(payload_to_text(payload), client=retrying)
    except AnalysisError as error:  # both attempts failed validation, and both were paid for
        control.add_cost(ai_cost(error.run_info))
        save_analysis(path, payload, run_info=error.run_info, error=str(error))
        print(f"  ✗ AI commentary: failed validation after the retry ({shown_path(path)}):")
        print("      " + str(error).replace("\n", "\n      "))
        return {"analysis_file": None, "run_info": error.run_info, "validation": f"failed: {error}",
                "rate_limit_waits": retrying.waits}
    except anthropic.AnthropicError as error:  # the API call itself failed (connection, rate limit that never cleared, ...)
        message = f"{type(error).__name__}: {error}"
        save_analysis(path, payload, error=message)
        print(f"  ✗ AI commentary: API error ({shown_path(path)}): {message}")
        return {"analysis_file": None, "run_info": None, "validation": f"failed: {message}",
                "rate_limit_waits": retrying.waits}
    control.add_cost(ai_cost(run_info))
    save_analysis(path, payload, summary, run_info)
    print(f"  ✓ AI commentary: passed validation ({run_info['attempts']} attempt(s), "
          f"{run_info['input_tokens']} in / {run_info['output_tokens']} out tokens, {run_info['seconds']}s): "
          f"{shown_path(path)}")
    return {"analysis_file": path, "run_info": run_info, "validation": "passed", "rate_limit_waits": retrying.waits}


def reusable_analysis(workbook_path, saved_dir, config):
    """The saved analysis in saved_dir if it was made from exactly these numbers and still passes, else None.

    "Exactly these numbers": the facts Claude saw (the payload) must equal today's, not just the
    company and quarter - so an edited workbook is never described by an old analysis.
    """
    path = analysis_path(workbook_path, saved_dir)
    if not path.exists():
        return None
    try:
        saved = json.loads(path.read_text())
    except json.JSONDecodeError:
        return None
    actuals, next_budget = clean_workbook(workbook_path)
    payload = build_payload(company_name(workbook_path), actuals, next_budget, config)
    if not isinstance(saved, dict) or saved.get("payload") != json.loads(json.dumps(payload)):
        return None
    summary, _ = load_analysis(path, payload)  # the same checks the deck makes
    return path if summary else None


def ai_part(workbook_path, actuals, next_budget, config, output_dir, skip_ai, client, reuse_saved, control=None):
    """(what happened to the AI step or None, the analysis file for the deck or None).

    reuse_saved=True (the web page, and main.py --resume) first looks for a saved analysis of exactly
    these numbers: it costs nothing, so it's used even when the AI step is skipped. Otherwise skip, or ask Claude.
    """
    saved = reusable_analysis(workbook_path, output_dir, config) if reuse_saved else None
    if saved:
        details = analysis_details(saved)
        print(f"  ✓ AI commentary: {AI_REUSED} ({shown_path(saved)})")
        return {"analysis_file": saved, "run_info": {"model": details["model"]}, "validation": AI_REUSED}, saved
    if skip_ai:
        print("  - AI commentary: skipped (--skip-ai)")
        return None, None
    ai = ai_step(workbook_path, actuals, next_budget, config, output_dir, client, control)
    return ai, ai["analysis_file"]


def deck_step(workbook_path, config, analysis_file, output_dir, draft=False):
    """Build and save the deck (draft=True: watermarked unless approved). Returns why the AI summary isn't on it, or None."""
    path, why_unavailable = save_deck(workbook_path, config, analysis_file, output_dir=output_dir, draft=draft)
    where = f"on slide {slide_number(commentary_slide)}"   # the AI commentary slide
    if why_unavailable is None:
        print(f"  ✓ Deck: {shown_path(path)} (AI text {where})")
    else:
        print(f"  ✓ Deck: {shown_path(path)} ({PLACEHOLDER_TEXT} {where})")
    if analysis_file is not None and why_unavailable is not None:  # the AI passed, but the deck rejected it
        print(f"      {PLACEHOLDER_TEXT}: {why_unavailable}")
    return why_unavailable


def memo_step(workbook_path, config, analysis_file, output_dir):
    """Build and save the memo (Word and PDF) from the same analysis as the deck. Returns memo.save_memo's result.

    The memo can reject AI text the deck accepted: every number in it must be in the metrics
    workbook, and Claude may quote a raw input (net burn, ending cash) that isn't. The reason is
    printed, as for the deck.
    """
    result = save_memo(workbook_path, config, analysis_file, output_dir=output_dir)
    ai = "AI text" if result["why_unavailable"] is None else MEMO_UNAVAILABLE
    print(f"  ✓ Memo: {shown_path(result['docx'])} + {result['pdf'].name} ({ai})")
    if analysis_file is not None and result["why_unavailable"] is not None:
        print(f"      {MEMO_UNAVAILABLE}: {result['why_unavailable']}")
    return result


def ai_record(ai, analysis_file):
    """What the manifest says about the AI step: model, prompt wording, tokens, cost, result.

    Everything is None when the AI didn't run, rather than a zero that would read like a real number.
    """
    if ai is None:  # --skip-ai
        return {"model": None, "prompt_version": None, "attempts": None, "input_tokens": None,
                "output_tokens": None, "seconds": None, "cost_usd": None, "validation": "skipped",
                "rate_limit_waits_s": None}
    run_info = ai["run_info"] or {}
    details = analysis_details(analysis_file) if analysis_file else None
    return {
        "model": run_info.get("model"),
        "prompt_version": details["prompt_version"] if details else PROMPT_VERSION,
        "attempts": run_info.get("attempts"),
        "input_tokens": run_info.get("input_tokens"), "output_tokens": run_info.get("output_tokens"),
        "seconds": run_info.get("seconds"),
        "cost_usd": ai_cost(run_info),
        "validation": ai["validation"],
        "rate_limit_waits_s": ai.get("rate_limit_waits", []),   # none for a reused analysis: no call was made
    }


def changes_step(workbook_path, actuals, next_budget, config, output_dir):
    """This run's results, and the earlier run they're compared with (diff_runs.py). Prints a one-line summary.

    Returns {"results", "previous_run"} for the manifest. The earlier run is read from the last
    manifest (in a batch, the copy in the company's private folder), before this run replaces it.
    """
    data = collect_deck_data(company_name(workbook_path), Path(workbook_path).name, actuals, next_budget, config)
    results = run_results(data)
    earlier = baseline(read_manifest(manifest_path(workbook_path, output_dir)), results)
    report = None if earlier is None else report_for(earlier, results, move_settings(config))
    print(f"  ✓ {HEADING}: {summary_text(report)}")
    return {"results": results, "previous_run": earlier}


def manifest_step(workbook_path, output_dir, ai, analysis_file, why_unavailable, memo=None, draft=False,
                  changes=None):
    """Write output/<company>_manifest.json: where this deck came from, and who has approved it.

    Any approval already recorded is carried over. Whether it still counts is decided by
    provenance.approval_status, which checks it against today's workbook, thresholds and column
    mapping - so a deck goes back to DRAFT on its own once any of them changes. memo = memo_step's result: the memo's files,
    and whether it carries the AI text (it can differ from the deck). The batch events (skips,
    timeouts, stops, failures) are carried over too, and whether --draft was asked for (--resume checks it).
    changes = changes_step's answer: this run's results and the earlier run they were compared with (Task 10).
    """
    path = manifest_path(workbook_path, output_dir)
    previous = read_manifest(path) or {}
    manifest = build_manifest(
        company_name(workbook_path), workbook_path, CONFIG_PATH,
        deck_path(workbook_path, output_dir).name, ai_record(ai, analysis_file),
        ai_text=why_unavailable is None, approval=previous.get("approval"),
        mapping=mapping_record(workbook_path))
    manifest["deck"]["draft"] = draft
    if previous.get("batch_events"):
        manifest["batch_events"] = previous["batch_events"]
    if memo is not None:
        manifest["memo"] = {"files": [memo["docx"].name, memo["pdf"].name], "ai_text": memo["why_unavailable"] is None}
    if changes is not None:
        manifest.update(changes)
    save_manifest(path, manifest)
    print(f"  ✓ Manifest: {shown_path(path)} (deck status: {manifest['deck']['status']})")
    return manifest


def ai_status(skip_ai, why_unavailable):
    """AI_SKIPPED with --skip-ai; otherwise AI_OK only if the AI text actually made it onto the deck."""
    if skip_ai:
        return AI_SKIPPED
    return AI_OK if why_unavailable is None else AI_FAILED


# ---------------------------------------------------------------------------
# One company
# ---------------------------------------------------------------------------

def blank_quarters(actuals):
    """Quarters with at least one blank input, e.g. ['Q1 2025']."""
    return list(actuals.index[actuals.isna().any(axis=1)])


def company_facts(workbook_path, config):
    """Clean the workbook and work out what the summary table shows: (result, actuals, next_budget)."""
    actuals, next_budget = clean_workbook(workbook_path)
    metrics = compute_metrics(actuals)
    flags = evaluate_flags(metrics, metric_reasons(actuals, metrics), config)  # latest quarter
    result = {
        "company": company_name(workbook_path),
        "quarter": flags[0]["quarter"],
        "flags_total": len(flags),
        "tripped": [flag["flag"] for flag in flags if flag["status"] == TRIP],
        "cannot_evaluate": [flag["flag"] for flag in flags if flag["status"] == CANNOT_EVALUATE],
        "gap_count": len(data_gaps(actuals, metrics, flags)),
        "blank_quarters": blank_quarters(actuals),
    }
    return result, actuals, next_budget


def run_company(workbook_path, config, skip_ai, client=None, output_dir=OUTPUT_DIR, draft=False, reuse_saved=False,
                control=None):
    """Run every step for one workbook and print a line per step.

    Returns a result dict for the summary table. Raises if any step fails.
    reuse_saved=True (the web page, --resume) uses a saved analysis of exactly these numbers instead of asking Claude.
    control (resilience.CompanyControl) is how a batch gives up on a company: it stops before its next step.
    """
    control = control or CompanyControl()
    result, actuals, next_budget = company_facts(workbook_path, config)
    budget_label = next_budget.name if next_budget is not None else "none"
    print(f"  ✓ Cleaned: {len(actuals)} quarters ({actuals.index[0]} to {actuals.index[-1]}), "
          f"budget row: {budget_label}")
    print(f"  ✓ Flags tripped ({result['quarter']}): {flags_text(result)}")
    for name in result["tripped"]:
        print(f"      tripped: {name}")
    print(f"  ✓ Data gaps: {gaps_text(result)}")
    changes = changes_step(workbook_path, actuals, next_budget, config, output_dir)

    control.checkpoint()
    excel_path = save_metrics_workbook(workbook_path, config, output_dir)
    print(f"  ✓ Excel: {shown_path(excel_path)}")
    control.checkpoint()
    ai, analysis_file = ai_part(workbook_path, actuals, next_budget, config, output_dir, skip_ai, client, reuse_saved,
                                control)
    control.checkpoint()
    why_unavailable = deck_step(workbook_path, config, analysis_file, output_dir, draft)
    control.checkpoint()
    memo = memo_step(workbook_path, config, analysis_file, output_dir)
    control.checkpoint()
    manifest = manifest_step(workbook_path, output_dir, ai, analysis_file, why_unavailable, memo, draft, changes)
    result["ai"] = ai_status(ai is None, why_unavailable)   # a reused analysis isn't "skipped"
    result["deck_status"] = manifest["deck"]["status"]
    result["cost_usd"] = manifest["ai"]["cost_usd"]
    note = rate_limit_note((ai or {}).get("rate_limit_waits"))
    result["notes"] = [note] if note else []
    return result


def describe_error(error):
    """Print why a company failed. Unexpected errors also get a traceback, to help fix the bug."""
    print(f"  ✗ FAILED: {type(error).__name__}: {error}")
    if not isinstance(error, INPUT_ERRORS):
        traceback.print_exc()


# ---------------------------------------------------------------------------
# The batch: resume, spend ceiling, timeouts, workers
# ---------------------------------------------------------------------------

def run_batch(workbook_paths, config, skip_ai, client=None, output_dir=OUTPUT_DIR, draft=False, resume=False,
              max_cost=None, timeout=None, workers=1, spend=None):
    """Run every workbook, up to `workers` at a time. A failure, timeout or stop is recorded and the batch moves on.

    Returns one result per workbook, in the order given (whatever order they finished in).
    client and output_dir are for tests: a fake Claude client, and a temporary folder.
    draft=True puts the DRAFT watermark on every deck nobody has approved (--draft).
    resume, max_cost (dollars), timeout (seconds per company), workers: see the module docstring.
    spend (resilience.SpendMeter) adds up the AI cost; main() passes one in to print the total.
    """
    batch = SimpleNamespace(config=config, skip_ai=skip_ai, client=client, output_dir=Path(output_dir), draft=draft,
                            resume=resume, max_cost=max_cost, timeout=timeout, spend=spend or SpendMeter(),
                            finished=queue.Queue())   # each company's thread puts itself here when it's done
    clear_old_stages(batch.output_dir)
    waiting, running, results = list(enumerate(workbook_paths)), {}, {}
    with routed_stdout() as screen:
        batch.screen = screen
        while waiting or running:
            while waiting and len(running) < workers:   # fill every free worker
                index, path = waiting.pop(0)
                start_or_settle(index, path, batch, running, results)
            if running:
                wait_for_a_company(running, results, batch)
    return [results[index] for index in range(len(workbook_paths))]


def start_or_settle(index, path, batch, running, results):
    """Skip the company (--resume, up to date), stop it (--max-cost reached), or start it in its own thread."""
    skipped, why_rebuilt = resume_check(path, batch) if batch.resume else (None, None)
    if skipped:
        results[index] = skipped
    elif batch.max_cost is not None and batch.spend.total >= batch.max_cost:
        results[index] = stopped_result(path, batch)
    else:
        notes = [f"rebuilt (--resume): {why_rebuilt}"] if why_rebuilt else []
        job = SimpleNamespace(index=index, path=Path(path), notes=notes, control=CompanyControl(batch.spend),
                              started=time.monotonic(), stage=None, buffer=None, result=None)
        threading.Thread(target=company_worker, args=(job, batch), daemon=True).start()
        running[index] = job


def resume_check(path, batch):
    """(a SKIPPED result, None) when the saved outputs are up to date, else (None, why the company is rebuilt)."""
    problem = resume_problem(path, batch.output_dir, CONFIG_PATH, want_ai=not batch.skip_ai, draft=batch.draft)
    if problem:
        return None, problem
    try:
        result, _, _ = company_facts(path, batch.config)   # the table still shows the flags and gaps
    except Exception as error:  # noqa: BLE001 - can't read it after all: build it, and let that step report why
        return None, f"could not read it again ({type(error).__name__})"
    print(f"\n=== {Path(path).name} ===\n  ✓ Skipped (--resume): {UP_TO_DATE}")
    record_batch_event(path, batch.output_dir, SKIPPED, f"{UP_TO_DATE} (--resume)")
    result.update(outcome=SKIPPED, error=None, ai=None, cost_usd=None, notes=[])
    return result, None


def stopped_result(path, batch):
    """The result for a company not started because the AI spend reached --max-cost."""
    detail = f"AI spend ${batch.spend.total:.2f} reached the --max-cost ceiling of ${batch.max_cost:.2f}"
    print(f"\n=== {Path(path).name} ===\n  ✗ STOPPED: {detail}")
    record_batch_event(path, batch.output_dir, STOPPED, detail)
    return {"company": company_name(path), "outcome": STOPPED, "error": f"stopped: {detail}", "detail": detail,
            "cost_usd": None, "notes": []}


def company_worker(job, batch):
    """Runs in the company's own thread: build it in a private folder, then hand it back to the batch.

    Its printed lines go to its own buffer, and are printed together when it's done.
    """
    job.buffer = batch.screen.capture()
    print(f"\n=== {job.path.name} ===")
    try:
        job.stage = make_stage(job.path, batch.output_dir)
        _private.stage, _private.output_dir = job.stage, batch.output_dir
        job.result = run_company(job.path, batch.config, batch.skip_ai, batch.client, job.stage, batch.draft,
                                 reuse_saved=batch.resume, control=job.control)
        job.result["error"] = None
    except Cancelled:  # the batch gave up on it (--timeout): nothing to report, it's already recorded
        pass
    except Exception as error:  # noqa: BLE001 - one bad company must not stop the batch
        describe_error(error)
        job.result = {"company": company_name(job.path), "error": f"{type(error).__name__}: {error}"}
    finally:
        if job.control.cancel.is_set() and job.stage:   # given up on: nothing it built may reach output/
            discard_stage(job.stage)
        batch.finished.put(job)


def wait_for_a_company(running, results, batch):
    """Wait until a company finishes or the first one runs out of time, and record what happened."""
    try:
        job = batch.finished.get(timeout=seconds_to_first_deadline(running, batch.timeout))
    except queue.Empty:   # nobody finished in time: give up on whoever is past their timeout
        give_up_on_overdue(running, results, batch)
        return
    if running.get(job.index) is not job:   # it finished after being given up on: throw its files away
        if job.stage:
            discard_stage(job.stage)
        return
    del running[job.index]
    results[job.index] = settle(job, batch)


def seconds_to_first_deadline(running, timeout):
    """How long until the first running company runs out of time (None: no --timeout, wait as long as it takes)."""
    if timeout is None:
        return None
    return max(0, min(job.started + timeout - time.monotonic() for job in running.values()))


def give_up_on_overdue(running, results, batch):
    """Every company past its --timeout: tell it to stop, record it, and free its worker for the next one."""
    for index, job in list(running.items()):
        if time.monotonic() - job.started >= batch.timeout:
            job.control.cancel.set()
            del running[index]
            results[index] = timed_out_result(job, batch)


def timed_out_result(job, batch):
    """Print what the company got through before its timeout, record the timeout, and return its result."""
    printed = job.buffer.getvalue() if job.buffer else f"\n=== {job.path.name} ===\n"
    print(printed, end="")
    print(f"  ✗ TIMED OUT after {batch.timeout:g} s: stopped waiting; earlier outputs left as they were")
    record_batch_event(job.path, batch.output_dir, TIMED_OUT, f"gave up after {batch.timeout:g} s")
    return {"company": company_name(job.path), "outcome": TIMED_OUT, "error": f"timed out after {batch.timeout:g} s",
            "timeout": batch.timeout, "cost_usd": None, "notes": job.notes}


def settle(job, batch):
    """A company's thread is done: print its lines, move its files into output/ if it succeeded, else record the failure.

    A failed company's private folder is thrown away, so its earlier outputs stay exactly as they were.
    """
    print(job.buffer.getvalue(), end="")
    result = job.result
    if result["error"]:
        if job.stage:
            discard_stage(job.stage)
        record_batch_event(job.path, batch.output_dir, FAILED, result["error"])
        result.update(outcome=FAILED, cost_usd=None, notes=job.notes)
    else:
        commit_stage(job.stage, batch.output_dir)
        result.update(outcome=BUILT, notes=job.notes + result["notes"])
    return result


# ---------------------------------------------------------------------------
# Summary table
# ---------------------------------------------------------------------------

def flags_text(result):
    """'6 of 9', or '7 of 9, 1 cannot evaluate' when some flags lack data."""
    text = f"{len(result['tripped'])} of {result['flags_total']}"
    if result["cannot_evaluate"]:
        text += f", {len(result['cannot_evaluate'])} cannot evaluate"
    return text


def gaps_text(result):
    """'none', or '19 metrics/flags (blank: Q1 2025)'."""
    if result["gap_count"] == 0:
        return "none"
    blanks = ", ".join(result["blank_quarters"]) or "none"
    return f"{result['gap_count']} metrics/flags (blank: {blanks})"


def outcome_of(result):
    """BUILT, SKIPPED, TIMED_OUT, STOPPED or FAILED (a result written before Task 7 has no "outcome": work it out)."""
    return result.get("outcome") or (FAILED if result["error"] else BUILT)


def result_text(result):
    """'OK', 'OK (AI skipped)', 'OK (AI failed)', 'SKIPPED ...', 'TIMED OUT ...', 'STOPPED: ...', or 'FAILED: <why>'."""
    outcome = outcome_of(result)
    if outcome == SKIPPED:
        return f"SKIPPED (--resume): {UP_TO_DATE}"
    if outcome == TIMED_OUT:
        return f"TIMED OUT after {result['timeout']:g} s: earlier outputs left as they were"
    if outcome == STOPPED:
        return f"STOPPED: {result['detail']}"
    if outcome == FAILED:
        return f"FAILED: {result['error']}"
    return RESULT_TEXTS[result["ai"]]


def notes_text(result):
    """Anything else worth knowing, e.g. '2 rate-limit retries (waited 15 s)' or why --resume rebuilt it."""
    return "; ".join(result.get("notes", []))


def summary_rows(results):
    """One row of text per company. A company that wasn't built or skipped has no flags or gaps to show."""
    rows = []
    for result in results:
        if result["error"]:
            rows.append([result["company"], "-", "-", result_text(result), notes_text(result)])
        else:
            rows.append([result["company"], flags_text(result), gaps_text(result), result_text(result),
                         notes_text(result)])
    return rows


def outcome_counts_text(results):
    """'1 skipped (--resume), 1 timed out, 1 stopped', or None when every company was simply built or failed."""
    labels = {SKIPPED: "skipped (--resume)", TIMED_OUT: "timed out", STOPPED: "stopped"}
    counts = [f"{sum(outcome_of(result) == outcome for result in results)} {label}" for outcome, label in labels.items()
              if any(outcome_of(result) == outcome for result in results)]
    return ", ".join(counts) or None


def print_summary(results):
    """Print the summary table with columns padded to line up."""
    headers = ["Company", "Flags tripped", "Data gaps", "Result", "Notes"]
    rows = summary_rows(results)
    widths = [max(len(row[i]) for row in [headers] + rows) for i in range(len(headers))]
    print("\n=== Summary ===")
    for row in [headers, ["-" * width for width in widths]] + rows:
        print("  ".join(text.ljust(width) for text, width in zip(row, widths)).rstrip())
    failed = sum(1 for result in results if result["error"])
    print(f"\n{len(results) - failed} of {len(results)} companies succeeded")
    if outcome_counts_text(results):
        print(outcome_counts_text(results))


def spend_text(spend, max_cost):
    """'AI spend this run: $0.27' (and the ceiling, when --max-cost was given)."""
    ceiling = f" (--max-cost ceiling ${max_cost:.2f})" if max_cost is not None else ""
    return f"AI spend this run: ${spend:.2f}{ceiling}"


CSV_HEADERS = ["Company", "Latest quarter", "Flags tripped", "Flags total", "Cannot evaluate", "Data gaps",
               "Blank quarters", "Result", "AI cost (USD)", "Notes"]


def csv_row(result):
    """One company as CSV cells: counts as plain numbers (so a spreadsheet can sort them), blanks if not built."""
    cost = "" if result.get("cost_usd") is None else result["cost_usd"]
    if result["error"]:
        return [result["company"], "", "", "", "", "", "", result_text(result), cost, notes_text(result)]
    return [result["company"], result["quarter"], len(result["tripped"]), result["flags_total"],
            len(result["cannot_evaluate"]), result["gap_count"], ", ".join(result["blank_quarters"]),
            result_text(result), cost, notes_text(result)]


def write_summary_csv(results, path=SUMMARY_CSV_PATH):
    """Save the summary table as a CSV file (overwritten each run). Returns the path."""
    path = Path(path)
    path.parent.mkdir(exist_ok=True)
    with open(path, "w", newline="") as file:  # newline="": the csv module writes its own line endings
        writer = csv.writer(file)
        writer.writerow(CSV_HEADERS)
        writer.writerows(csv_row(result) for result in results)
    return path


def save_batch_manifest(results, options, spend, path=BATCH_MANIFEST_PATH):
    """Save output/batch_manifest.json: the options, the code, the total AI spend, and every company's outcome.

    This is where a company that has never been built (so has no manifest of its own) still has its
    timeout or stop on record. Overwritten each run, like the summary CSV.
    """
    record = {
        "run_at": timestamp(),
        "code": git_commit(),
        "options": options,
        "ai_spend_usd": round(spend, 4),
        "companies": [{"company": result["company"], "outcome": outcome_of(result), "result": result_text(result),
                       "cost_usd": result.get("cost_usd"), "notes": result.get("notes", [])} for result in results],
    }
    return save_manifest(path, record)


def quarter_mismatch_warning(results):
    """A warning line if the successful companies don't all end on the same quarter, else None.

    Not a failure: a company can be a quarter behind. But comparing them side by side needs care.
    """
    companies_by_quarter = {}  # quarter -> company names, in the order first seen
    for result in results:
        if not result["error"]:
            companies_by_quarter.setdefault(result["quarter"], []).append(result["company"])
    if len(companies_by_quarter) <= 1:
        return None
    groups = "; ".join(f"{quarter}: {', '.join(names)}" for quarter, names in companies_by_quarter.items())
    return f"⚠ Companies end on different quarters ({groups}) - compare them with care"


def ai_failed_warning(results):
    """A warning line naming every company whose deck has the AI placeholder because the AI failed, else None."""
    names = [result["company"] for result in results if not result["error"] and result["ai"] == AI_FAILED]
    if not names:
        return None
    companies = "company" if len(names) == 1 else "companies"
    return (f"⚠ AI summary unavailable for {len(names)} {companies} ({', '.join(names)}): their decks show the "
            f"placeholder; the reasons are in output/<company>_analysis.json")


# ---------------------------------------------------------------------------
# Command line
# ---------------------------------------------------------------------------

def more_than_zero(text):
    """argparse type for --max-cost and --timeout: a number above 0."""
    try:
        value = float(text)
    except ValueError:
        raise argparse.ArgumentTypeError(f"{text!r} is not a number") from None
    if value <= 0:
        raise argparse.ArgumentTypeError(f"must be more than 0, not {text}")
    return value


def at_least_one(text):
    """argparse type for --workers: a whole number of 1 or more."""
    try:
        value = int(text)
    except ValueError:
        raise argparse.ArgumentTypeError(f"{text!r} is not a whole number") from None
    if value < 1:
        raise argparse.ArgumentTypeError(f"must be 1 or more, not {text}")
    return value


def parse_args(argv=None):
    """Exactly one of: a workbook path, or --all. Plus the options."""
    parser = argparse.ArgumentParser(description="Build board-pack outputs for one or all KPI workbooks.")
    parser.add_argument("workbook", nargs="?", help="path to one KPI workbook, e.g. data/northwind.xlsx")
    parser.add_argument("--all", action="store_true", help="run every .xlsx workbook in data/")
    parser.add_argument("--skip-ai", action="store_true", help="don't call Claude; decks show the AI placeholder")
    parser.add_argument("--draft", action="store_true", help="stamp DRAFT - NOT REVIEWED on decks nobody has approved")
    parser.add_argument("--resume", action="store_true",
                        help="skip companies whose outputs were built from today's workbook and config.yaml")
    parser.add_argument("--max-cost", type=more_than_zero, metavar="DOLLARS",
                        help="start no more companies once the AI spend reaches this many dollars")
    parser.add_argument("--timeout", type=more_than_zero, metavar="SECONDS",
                        help="give up on a company after this many seconds (default: no limit)")
    parser.add_argument("--workers", type=at_least_one, default=1,
                        help="how many companies to run side by side (default: 1)")
    args = parser.parse_args(argv)
    if bool(args.workbook) == args.all:  # both given, or neither
        parser.error("give one workbook path, or --all (not both)")
    return args


def main(argv=None):
    args = parse_args(argv)
    paths = find_workbooks() if args.all else [Path(args.workbook)]
    if not paths:
        print(f"No .xlsx workbooks found in {DATA_DIR}")
        return 1
    problem = None if args.skip_ai else api_key_problem()
    if problem:  # stop before any company runs, rather than fail the AI step for every one
        print(problem)
        return 1
    options = {"resume": args.resume, "max_cost": args.max_cost, "timeout": args.timeout, "workers": args.workers,
               "skip_ai": args.skip_ai, "draft": args.draft}
    spend = SpendMeter()
    results = run_batch(paths, load_config(), args.skip_ai, draft=args.draft, resume=args.resume,
                        max_cost=args.max_cost, timeout=args.timeout, workers=args.workers, spend=spend)
    print_summary(results)
    if not args.skip_ai or args.max_cost is not None:
        print(spend_text(spend.total, args.max_cost))
    for warning in (quarter_mismatch_warning(results), ai_failed_warning(results)):
        if warning:
            print(warning)
    print(f"Summary saved: {shown_path(write_summary_csv(results))}")
    print(f"Batch manifest saved: {shown_path(save_batch_manifest(results, options, spend.total))}")
    return 1 if any(result["error"] for result in results) else 0   # failed, timed out or stopped


if __name__ == "__main__":
    sys.exit(main())
