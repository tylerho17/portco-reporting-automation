"""Batch runner: turn one KPI workbook, or every workbook in data/, into board-pack outputs.

Steps for each company:
1. Clean the workbook (clean.py)
2. Compute metrics, flags and data gaps (metrics.py)
3. Save output/<company>_metrics.xlsx (excel_output.py)
4. AI commentary (analyze.py)  - skipped for now, see below
5. Deck (build_deck.py)         - skipped for now, see below

Steps 4 and 5 are skipped with a clear message while build_deck.py doesn't exist:
the commentary has nowhere to go until there is a deck, so no API money is spent.
--skip-ai skips step 4 even once the deck exists.

One company failing never stops the batch. The run ends with a summary table
(company, flags tripped, data gaps, result) and exits with code 1 if any company failed.

Run: python main.py data/northwind.xlsx
     python main.py --all --skip-ai
"""

import argparse
import sys
import traceback
from pathlib import Path

from clean import clean_workbook
from excel_output import save_metrics_workbook
from metrics import MISSING, TRIP, compute_metrics, data_gaps, evaluate_flags, load_config

PROJECT_DIR = Path(__file__).parent
DATA_DIR = PROJECT_DIR / "data"
BUILD_DECK_PATH = PROJECT_DIR / "build_deck.py"

# Errors caused by a bad input file (clean.py raises ValueError with a clear message; a missing
# or unreadable file raises OSError). Anything else is probably a bug, so its traceback is printed.
INPUT_ERRORS = (ValueError, OSError)

NO_DECK_MESSAGE = "skipped (build_deck.py doesn't exist yet - build step 4)"
NOT_WIRED_MESSAGE = "build_deck.py exists but main.py doesn't call it yet - connect it in main.py"


class NotWiredError(Exception):
    """A later build step exists on disk but main.py hasn't been connected to it."""


# ---------------------------------------------------------------------------
# Finding the input files
# ---------------------------------------------------------------------------

def find_workbooks(data_dir=DATA_DIR):
    """Every .xlsx in data_dir, sorted by name.

    Skips "~$" files: Excel creates those as lock files while a workbook is open.
    """
    return sorted(path for path in Path(data_dir).glob("*.xlsx") if not path.name.startswith("~$"))


def company_name(workbook_path):
    """data/northwind.xlsx -> 'Northwind' (same rule analyze.py uses)."""
    return Path(workbook_path).stem.title()


# ---------------------------------------------------------------------------
# Steps 4 and 5: placeholders until build_deck.py exists
# ---------------------------------------------------------------------------

def ai_step(skip_ai):
    """Return a message saying why AI commentary was skipped, or stop if it should have run."""
    if skip_ai:
        return "skipped (--skip-ai)"
    if not BUILD_DECK_PATH.exists():
        return NO_DECK_MESSAGE
    raise NotWiredError(NOT_WIRED_MESSAGE)


def deck_step():
    """Return a message saying why the deck was skipped, or stop if it should have run."""
    if not BUILD_DECK_PATH.exists():
        return NO_DECK_MESSAGE
    raise NotWiredError(NOT_WIRED_MESSAGE)


# ---------------------------------------------------------------------------
# One company
# ---------------------------------------------------------------------------

def blank_quarters(actuals):
    """Quarters with at least one blank input, e.g. ['Q1 2025']."""
    return list(actuals.index[actuals.isna().any(axis=1)])


def run_company(workbook_path, config, skip_ai):
    """Run every step for one workbook and print a line per step.

    Returns a result dict for the summary table. Raises if any step fails.
    """
    actuals, next_budget = clean_workbook(workbook_path)
    budget_label = next_budget.name if next_budget is not None else "none"
    print(f"  ✓ Cleaned: {len(actuals)} quarters ({actuals.index[0]} to {actuals.index[-1]}), "
          f"budget row: {budget_label}")

    metrics = compute_metrics(actuals)
    flags = evaluate_flags(metrics, config)  # latest quarter
    gaps = data_gaps(actuals, metrics, flags)
    result = {
        "company": company_name(workbook_path),
        "quarter": flags[0]["quarter"],
        "flags_total": len(flags),
        "tripped": [flag["flag"] for flag in flags if flag["status"] == TRIP],
        "cannot_evaluate": [flag["flag"] for flag in flags if flag["status"] == MISSING],
        "gap_count": len(gaps),
        "blank_quarters": blank_quarters(actuals),
    }
    print(f"  ✓ Flags tripped ({result['quarter']}): {flags_text(result)}")
    for name in result["tripped"]:
        print(f"      tripped: {name}")
    print(f"  ✓ Data gaps: {gaps_text(result)}")

    excel_path = save_metrics_workbook(workbook_path, config)
    print(f"  ✓ Excel: {excel_path.relative_to(PROJECT_DIR)}")
    print(f"  - AI commentary: {ai_step(skip_ai)}")
    print(f"  - Deck: {deck_step()}")
    return result


def describe_error(error):
    """Print why a company failed. Unexpected errors also get a traceback, to help fix the bug."""
    print(f"  ✗ FAILED: {type(error).__name__}: {error}")
    if not isinstance(error, INPUT_ERRORS + (NotWiredError,)):
        traceback.print_exc()


def run_batch(workbook_paths, config, skip_ai):
    """Run every workbook in turn. A failure is recorded and the batch moves on."""
    results = []
    for path in workbook_paths:
        print(f"\n=== {Path(path).name} ===")
        try:
            result = run_company(path, config, skip_ai)
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
    """'OK', 'OK (AI + deck skipped)' while there is no deck yet, or 'FAILED: <why>'."""
    if result["error"]:
        return f"FAILED: {result['error']}"
    return "OK" if BUILD_DECK_PATH.exists() else "OK (AI + deck skipped)"


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


# ---------------------------------------------------------------------------
# Command line
# ---------------------------------------------------------------------------

def parse_args(argv=None):
    """Exactly one of: a workbook path, or --all."""
    parser = argparse.ArgumentParser(description="Build board-pack outputs for one or all KPI workbooks.")
    parser.add_argument("workbook", nargs="?", help="path to one KPI workbook, e.g. data/northwind.xlsx")
    parser.add_argument("--all", action="store_true", help="run every .xlsx workbook in data/")
    parser.add_argument("--skip-ai", action="store_true", help="don't call Claude for commentary")
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
    results = run_batch(paths, load_config(), args.skip_ai)
    print_summary(results)
    return 1 if any(result["error"] for result in results) else 0


if __name__ == "__main__":
    sys.exit(main())
