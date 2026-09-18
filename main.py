"""Batch runner: turn one KPI workbook, or every workbook in data/, into board-pack outputs.

Steps for each company:
1. Clean the workbook (clean.py)
2. Compute metrics, flags and data gaps (metrics.py)
3. Save output/<company>_metrics.xlsx (excel_output.py)
4. AI commentary (analyze.py): output/<company>_analysis.json, saved whether it passed or failed
5. Deck (build_deck.py): output/<company>_board_pack.pptx
6. Memo (memo.py): output/<company>_board_memo.docx and .pdf, from the same numbers and analysis
7. Manifest (provenance.py): output/<company>_manifest.json

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
(company, flags tripped, data gaps, result), also saved as output/batch_summary.csv, a warning
if companies end on different quarters or any AI summary is unavailable, and exit code 1 if any
company failed. "OK (AI failed)" is not a failed company: its outputs were all built.

Run: python main.py data/northwind.xlsx
     python main.py --all --skip-ai
     python main.py --all --draft
"""

import argparse
import csv
import os
import sys
import traceback
from pathlib import Path

import anthropic
from dotenv import load_dotenv

from analyze import PROMPT_VERSION, AnalysisError, analyze, build_payload, payload_to_text, save_analysis
from build_deck import (PLACEHOLDER_TEXT, analysis_details, analysis_path, commentary_slide, deck_path, save_deck,
                        slide_number)
from clean import clean_workbook
from compare_models import run_cost
from excel_output import save_metrics_workbook
from memo import MEMO_UNAVAILABLE, save_memo
from metrics import (CANNOT_EVALUATE, CONFIG_PATH, TRIP, compute_metrics, data_gaps, evaluate_flags, load_config,
                     metric_reasons)
from provenance import build_manifest, manifest_path, read_manifest, save_manifest
from text_fit import TextDoesNotFitError

PROJECT_DIR = Path(__file__).parent
DATA_DIR = PROJECT_DIR / "data"
OUTPUT_DIR = PROJECT_DIR / "output"
SUMMARY_CSV_PATH = OUTPUT_DIR / "batch_summary.csv"

# Errors caused by a bad input file (clean.py raises ValueError with a clear message; a missing
# or unreadable file raises OSError), or by text that can't fit a slide (text_fit.py names the slide
# and the box). Anything else is probably a bug, so its traceback is printed.
INPUT_ERRORS = (ValueError, OSError, TextDoesNotFitError)

# What happened to the AI summary, and how the Result column says it.
AI_OK, AI_SKIPPED, AI_FAILED = "ok", "skipped", "failed"
RESULT_TEXTS = {AI_OK: "OK", AI_SKIPPED: "OK (AI skipped)", AI_FAILED: "OK (AI failed)"}

NO_KEY_MESSAGE = "ANTHROPIC_API_KEY isn't set: add it to .env, or run with --skip-ai"


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


def shown_path(path):
    """A path as printed: relative to the project folder when it's inside it (output/...), else in full."""
    path = Path(path)
    return path.relative_to(PROJECT_DIR) if path.is_relative_to(PROJECT_DIR) else path


# ---------------------------------------------------------------------------
# Steps 4 and 5: AI commentary and the deck
# ---------------------------------------------------------------------------

def api_key_problem():
    """None if an API key is set (from the environment or .env), else a message saying what to do."""
    load_dotenv()  # puts ANTHROPIC_API_KEY from .env into the environment
    return None if os.environ.get("ANTHROPIC_API_KEY", "").strip() else NO_KEY_MESSAGE


def ai_step(workbook_path, actuals, next_budget, config, output_dir, client=None):
    """Ask Claude for the commentary and save output/<company>_analysis.json, whether it passed or not.

    Returns {"analysis_file": the file for the deck or None, "run_info": ..., "validation": what
    happened}. Only a validation failure after the retry or an API error counts as "AI failed"; any
    other error is a bug and fails the company. client=None uses the real Anthropic client.
    """
    path = analysis_path(workbook_path, output_dir)
    path.unlink(missing_ok=True)  # an old analysis must never look like this run's
    payload = build_payload(company_name(workbook_path), actuals, next_budget, config)
    try:
        summary, run_info = analyze(payload_to_text(payload), client=client)
    except AnalysisError as error:  # both attempts failed validation
        save_analysis(path, payload, run_info=error.run_info, error=str(error))
        print(f"  ✗ AI commentary: failed validation after the retry ({shown_path(path)}):")
        print("      " + str(error).replace("\n", "\n      "))
        return {"analysis_file": None, "run_info": error.run_info, "validation": f"failed: {error}"}
    except anthropic.AnthropicError as error:  # the API call itself failed (connection, rate limit, ...)
        message = f"{type(error).__name__}: {error}"
        save_analysis(path, payload, error=message)
        print(f"  ✗ AI commentary: API error ({shown_path(path)}): {message}")
        return {"analysis_file": None, "run_info": None, "validation": f"failed: {message}"}
    save_analysis(path, payload, summary, run_info)
    print(f"  ✓ AI commentary: passed validation ({run_info['attempts']} attempt(s), "
          f"{run_info['input_tokens']} in / {run_info['output_tokens']} out tokens, {run_info['seconds']}s): "
          f"{shown_path(path)}")
    return {"analysis_file": path, "run_info": run_info, "validation": "passed"}


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
                "output_tokens": None, "seconds": None, "cost_usd": None, "validation": "skipped"}
    run_info = ai["run_info"] or {}
    details = analysis_details(analysis_file) if analysis_file else None
    model, tokens = run_info.get("model"), (run_info.get("input_tokens"), run_info.get("output_tokens"))
    return {
        "model": model,
        "prompt_version": details["prompt_version"] if details else PROMPT_VERSION,
        "attempts": run_info.get("attempts"),
        "input_tokens": tokens[0], "output_tokens": tokens[1], "seconds": run_info.get("seconds"),
        "cost_usd": round(run_cost(model, *tokens), 4) if model and all(tokens) else None,
        "validation": ai["validation"],
    }


def manifest_step(workbook_path, output_dir, ai, analysis_file, why_unavailable, memo=None):
    """Write output/<company>_manifest.json: where this deck came from, and who has approved it.

    Any approval already recorded is carried over. Whether it still counts is decided by
    provenance.approval_status, which checks it against today's workbook and thresholds - so a deck
    goes back to DRAFT on its own once either changes. memo = memo_step's result: the memo's files,
    and whether it carries the AI text (it can differ from the deck).
    """
    path = manifest_path(workbook_path, output_dir)
    previous = read_manifest(path)
    manifest = build_manifest(
        company_name(workbook_path), workbook_path, CONFIG_PATH,
        deck_path(workbook_path, output_dir).name, ai_record(ai, analysis_file),
        ai_text=why_unavailable is None, approval=(previous or {}).get("approval"))
    if memo is not None:
        manifest["memo"] = {"files": [memo["docx"].name, memo["pdf"].name], "ai_text": memo["why_unavailable"] is None}
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


def run_company(workbook_path, config, skip_ai, client=None, output_dir=OUTPUT_DIR, draft=False):
    """Run every step for one workbook and print a line per step.

    Returns a result dict for the summary table. Raises if any step fails.
    """
    actuals, next_budget = clean_workbook(workbook_path)
    budget_label = next_budget.name if next_budget is not None else "none"
    print(f"  ✓ Cleaned: {len(actuals)} quarters ({actuals.index[0]} to {actuals.index[-1]}), "
          f"budget row: {budget_label}")

    metrics = compute_metrics(actuals)
    flags = evaluate_flags(metrics, metric_reasons(actuals, metrics), config)  # latest quarter
    gaps = data_gaps(actuals, metrics, flags)
    result = {
        "company": company_name(workbook_path),
        "quarter": flags[0]["quarter"],
        "flags_total": len(flags),
        "tripped": [flag["flag"] for flag in flags if flag["status"] == TRIP],
        "cannot_evaluate": [flag["flag"] for flag in flags if flag["status"] == CANNOT_EVALUATE],
        "gap_count": len(gaps),
        "blank_quarters": blank_quarters(actuals),
    }
    print(f"  ✓ Flags tripped ({result['quarter']}): {flags_text(result)}")
    for name in result["tripped"]:
        print(f"      tripped: {name}")
    print(f"  ✓ Data gaps: {gaps_text(result)}")

    excel_path = save_metrics_workbook(workbook_path, config, output_dir)
    print(f"  ✓ Excel: {shown_path(excel_path)}")
    if skip_ai:
        ai, analysis_file = None, None
        print("  - AI commentary: skipped (--skip-ai)")
    else:
        ai = ai_step(workbook_path, actuals, next_budget, config, output_dir, client)
        analysis_file = ai["analysis_file"]
    why_unavailable = deck_step(workbook_path, config, analysis_file, output_dir, draft)
    memo = memo_step(workbook_path, config, analysis_file, output_dir)
    manifest = manifest_step(workbook_path, output_dir, ai, analysis_file, why_unavailable, memo)
    result["ai"] = ai_status(skip_ai, why_unavailable)
    result["deck_status"] = manifest["deck"]["status"]
    return result


def describe_error(error):
    """Print why a company failed. Unexpected errors also get a traceback, to help fix the bug."""
    print(f"  ✗ FAILED: {type(error).__name__}: {error}")
    if not isinstance(error, INPUT_ERRORS):
        traceback.print_exc()


def run_batch(workbook_paths, config, skip_ai, client=None, output_dir=OUTPUT_DIR, draft=False):
    """Run every workbook in turn. A failure is recorded and the batch moves on.

    client and output_dir are for tests: a fake Claude client, and a temporary folder.
    draft=True puts the DRAFT watermark on every deck nobody has approved (--draft).
    """
    results = []
    for path in workbook_paths:
        print(f"\n=== {Path(path).name} ===")
        try:
            result = run_company(path, config, skip_ai, client, output_dir, draft)
            result["error"] = None
        except Exception as error:  # noqa: BLE001 - one bad company must not stop the batch
            describe_error(error)
            result = {"company": company_name(path), "error": f"{type(error).__name__}: {error}"}
        results.append(result)
    return results


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


def result_text(result):
    """'OK', 'OK (AI skipped)', 'OK (AI failed)', or 'FAILED: <why>'."""
    if result["error"]:
        return f"FAILED: {result['error']}"
    return RESULT_TEXTS[result["ai"]]


def summary_rows(results):
    """One row of text per company. A failed company has no flags or gaps to show."""
    rows = []
    for result in results:
        if result["error"]:
            rows.append([result["company"], "-", "-", result_text(result)])
        else:
            rows.append([result["company"], flags_text(result), gaps_text(result), result_text(result)])
    return rows


def print_summary(results):
    """Print the summary table with columns padded to line up."""
    headers = ["Company", "Flags tripped", "Data gaps", "Result"]
    rows = summary_rows(results)
    widths = [max(len(row[i]) for row in [headers] + rows) for i in range(len(headers))]
    print("\n=== Summary ===")
    for row in [headers, ["-" * width for width in widths]] + rows:
        print("  ".join(text.ljust(width) for text, width in zip(row, widths)).rstrip())
    failed = sum(1 for result in results if result["error"])
    print(f"\n{len(results) - failed} of {len(results)} companies succeeded")


CSV_HEADERS = ["Company", "Latest quarter", "Flags tripped", "Flags total", "Cannot evaluate", "Data gaps",
               "Blank quarters", "Result"]


def csv_row(result):
    """One company as CSV cells: counts as plain numbers (so a spreadsheet can sort them), blanks if it failed."""
    if result["error"]:
        return [result["company"], "", "", "", "", "", "", result_text(result)]
    return [result["company"], result["quarter"], len(result["tripped"]), result["flags_total"],
            len(result["cannot_evaluate"]), result["gap_count"], ", ".join(result["blank_quarters"]),
            result_text(result)]


def write_summary_csv(results, path=SUMMARY_CSV_PATH):
    """Save the summary table as a CSV file (overwritten each run). Returns the path."""
    path = Path(path)
    path.parent.mkdir(exist_ok=True)
    with open(path, "w", newline="") as file:  # newline="": the csv module writes its own line endings
        writer = csv.writer(file)
        writer.writerow(CSV_HEADERS)
        writer.writerows(csv_row(result) for result in results)
    return path


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

def parse_args(argv=None):
    """Exactly one of: a workbook path, or --all."""
    parser = argparse.ArgumentParser(description="Build board-pack outputs for one or all KPI workbooks.")
    parser.add_argument("workbook", nargs="?", help="path to one KPI workbook, e.g. data/northwind.xlsx")
    parser.add_argument("--all", action="store_true", help="run every .xlsx workbook in data/")
    parser.add_argument("--skip-ai", action="store_true", help="don't call Claude; decks show the AI placeholder")
    parser.add_argument("--draft", action="store_true", help="stamp DRAFT - NOT REVIEWED on decks nobody has approved")
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
    results = run_batch(paths, load_config(), args.skip_ai, draft=args.draft)
    print_summary(results)
    for warning in (quarter_mismatch_warning(results), ai_failed_warning(results)):
        if warning:
            print(warning)
    print(f"Summary saved: {shown_path(write_summary_csv(results))}")
    return 1 if any(result["error"] for result in results) else 0


if __name__ == "__main__":
    sys.exit(main())
