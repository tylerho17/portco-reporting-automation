"""Run the evaluation set: clean, metrics, flags and gaps for all 12 companies, each against its answer key.

Run:  python eval/run_eval.py   (make the workbooks first with python eval/make_eval_data.py; they are
saved in eval/data/, so this is only needed after changing make_eval_data.py). No API calls.

For each company, four checks, each giving a list of mismatches (empty = matches the answer key):
- Clean:   clean.py reads back every answer-key value exactly (the blank quarter all empty) and the
           budget-only row; for a workbook that must stop, it stops with the expected message.
- Metrics: the latest quarter's metrics equal the answer key's hand formulas (all 19 on Larkspur's
           numbers, the 8 flag metrics elsewhere); runway at next quarter's
           budgeted burn too; and every value with no number has the reason the rules predict
           (missing input, no prior period, or not meaningful where the answer key lists it).
- Flags:   every flag's status and reason in the latest quarter; for a company that must never trip,
           no flag trips in any quarter.
- Gaps:    data_gaps() lists exactly the metrics and flags the CLAUDE.md rules predict.

Prints a scorecard, one row per company, then every mismatch by name. Exits 1 if any company fails,
so it can sit with the other checks.
"""

import math
import sys
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_DIR))            # the project's own files
sys.path.insert(0, str(Path(__file__).parent))  # make_eval_data.py, when this is imported from elsewhere

import make_eval_data  # noqa: E402  (after the path lines on purpose)
from check_companies import same_number  # noqa: E402
from check_northwind import QOQ_METRICS, YOY_METRICS  # noqa: E402  (typed by hand, not read from metrics.py)
from clean import clean_workbook  # noqa: E402
from metrics import (CANNOT_EVALUATE, FLAG_RULES, MISSING_INPUT, NO_PRIOR_PERIOD, NOT_MEANINGFUL, TRIP,  # noqa: E402
                     compute_metrics, data_gaps, evaluate_flags, load_config, metric_reasons,
                     runway_at_next_budget)

CHECKS = ["Clean", "Metrics", "Flags", "Gaps"]
MAKE_HINT = "run python eval/make_eval_data.py to write it"


# ---------------------------------------------------------------------------
# Clean
# ---------------------------------------------------------------------------

def clean_mismatches(company, actuals, next_budget):
    """Every cleaned value vs the answer key. The blank quarter must be entirely empty (NaN)."""
    if list(actuals.index) != company["quarters"]:
        return [f"quarters: expected {company['quarters']}, got {list(actuals.index)}"]
    found = []
    for column, values in company["true_data"].items():
        for quarter, expected in zip(company["quarters"], values):
            got = actuals.loc[quarter, column]
            if quarter == company["blank_quarter"]:
                if not math.isnan(got):
                    found.append(f"{quarter} {column}: expected blank, got {got}")
            elif got != expected:
                found.append(f"{quarter} {column}: expected {expected}, got {got}")
    got_budget = None if next_budget is None else next_budget.to_dict()
    if got_budget != company["next_quarter_budget"]:
        found.append(f"budget-only row: expected {company['next_quarter_budget']}, got {got_budget}")
    return found


def stop_mismatches(company):
    """A workbook that must stop: clean.py has to raise, and its message must hold every expected phrase."""
    try:
        clean_workbook(company["output_path"])
    except FileNotFoundError:
        return [f"no workbook at {company['output_path']}: {MAKE_HINT}"]
    except ValueError as error:
        return [f"stopped, but the message lacks {phrase!r}: {error}"
                for phrase in company["stop"] if phrase not in str(error)]
    return ["expected clean.py to stop, but the workbook was read"]


# ---------------------------------------------------------------------------
# Metrics
# ---------------------------------------------------------------------------

def lookback(metric):
    """How many quarters back a metric reaches, from the hand-typed lists: QoQ 1, YoY 4, others 0."""
    if metric in QOQ_METRICS:
        return 1
    return 4 if metric in YOY_METRICS else 0


def predicted_reason(company, metric, position):
    """Why the metric at this row should have no number, from the CLAUDE.md rules; None = a number.

    A blank input wins (the blank quarter itself, or the quarter it looks back to); then a quarter the
    workbook doesn't have; then the answer key's list of not-meaningful values.
    """
    quarters, back = company["quarters"], lookback(metric)
    blank = quarters.index(company["blank_quarter"]) if company["blank_quarter"] else None
    if blank is not None and blank in (position, position - back):
        return MISSING_INPUT
    if position - back < 0:
        return NO_PRIOR_PERIOD
    if quarters[position] in company["not_meaningful"].get(metric, []):
        return NOT_MEANINGFUL
    return None


def reason_words(reason):
    """'a number' for no reason, else the reason itself (for mismatch messages)."""
    return reason or "a number"


def reason_mismatches(company, reasons):
    """Every metric in every quarter: the reason it has no number must be the predicted one."""
    found = []
    for metric in reasons.columns:
        for position, quarter in enumerate(reasons.index):
            expected = predicted_reason(company, metric, position)
            got = reasons.loc[quarter, metric]
            if expected != got:
                found.append(f"{metric} in {quarter}: expected {reason_words(expected)}, got {reason_words(got)}")
    return found


def metric_mismatches(company, actuals, next_budget, metrics, reasons):
    """The latest quarter's flag metrics and runway at budget vs the hand formulas, then every reason."""
    latest = company["quarters"][-1]
    found = [f"{metric} in {latest}: expected {expected}, got {metrics.loc[latest, metric]}"
             for metric, expected in company["expected_latest"].items()
             if not same_number(metrics.loc[latest, metric], expected)]
    runway = runway_at_next_budget(actuals, next_budget)
    if not same_number(runway, company["expected_runway_at_budget"]):
        found.append(f"runway at next quarter's budgeted burn: expected {company['expected_runway_at_budget']}, "
                     f"got {runway}")
    return found + reason_mismatches(company, reasons)


# ---------------------------------------------------------------------------
# Flags and gaps
# ---------------------------------------------------------------------------

def flag_words(flag):
    """'trip', 'pass' or 'cannot evaluate: <reason>': the status as the answer key writes it."""
    return f"{CANNOT_EVALUATE}: {flag['reason']}" if flag["status"] == CANNOT_EVALUATE else flag["status"]


def flag_mismatches(company, flags, metrics, reasons, config):
    """Each latest-quarter flag vs the answer key; for a never-trips company, every quarter too."""
    got = {flag["flag"]: flag_words(flag) for flag in flags}
    expected = company["expected_flags"]
    names = list(expected) + [name for name in got if name not in expected]   # answer-key order, extras last
    found = [f"{name}: expected {expected.get(name, 'no such flag')}, got {got.get(name, 'no such flag')}"
             for name in names if expected.get(name) != got.get(name)]
    if company.get("never_trips"):
        for quarter in metrics.index:
            tripped = [f["flag"] for f in evaluate_flags(metrics, reasons, config, quarter) if f["status"] == TRIP]
            if tripped:
                found.append(f"{quarter} tripped {', '.join(tripped)}")
    return found


def predicted_gaps(company, metric_columns):
    """The data gaps the rules predict: every missing-input value, then each flag expected to be
    'cannot evaluate: missing input' (in the latest quarter). 'No prior period' is never a gap."""
    quarters = company["quarters"]
    gaps = {}
    for metric in metric_columns:
        missing = [quarters[p] for p in range(len(quarters)) if predicted_reason(company, metric, p) == MISSING_INPUT]
        if missing:
            gaps[metric] = missing
    for name, status in company["expected_flags"].items():
        if status == f"{CANNOT_EVALUATE}: {MISSING_INPUT}":
            gaps[f"flag: {name}"] = [quarters[-1]]
    return gaps


def gap_mismatches(company, gaps, metric_columns):
    """Each metric or flag whose gap list differs from the prediction (extra, missing or different quarters)."""
    expected = predicted_gaps(company, metric_columns)
    return [f"{name}: expected {expected.get(name, 'no gap')}, got {gaps.get(name, 'no gap')}"
            for name in sorted(set(expected) | set(gaps)) if expected.get(name) != gaps.get(name)]


# ---------------------------------------------------------------------------
# One company, then the scorecard
# ---------------------------------------------------------------------------

def evaluate(company, config):
    """Run one company through clean, metrics, flags and gaps. Returns {check: [mismatches]}.

    A workbook that must stop is only checked under Clean. One that stops unexpectedly fails Clean,
    and the other three can't run.
    """
    if company.get("stop"):
        return {"Clean": stop_mismatches(company)}
    try:
        actuals, next_budget = clean_workbook(company["output_path"])
    except FileNotFoundError:
        return {"Clean": [f"no workbook at {company['output_path']}: {MAKE_HINT}"]}
    except ValueError as error:
        return {"Clean": [f"stopped unexpectedly: {error}"]}
    metrics = compute_metrics(actuals)
    reasons = metric_reasons(actuals, metrics)
    flags = evaluate_flags(metrics, reasons, config)
    return {
        "Clean": clean_mismatches(company, actuals, next_budget),
        "Metrics": metric_mismatches(company, actuals, next_budget, metrics, reasons),
        "Flags": flag_mismatches(company, flags, metrics, reasons, config),
        "Gaps": gap_mismatches(company, data_gaps(actuals, metrics, flags), metrics.columns),
    }


def cell(company, results, check):
    """One scorecard cell: 'ok', 'stopped' (a stop company's Clean), '-' (not run), or 'MISMATCH (n)'."""
    if check not in results:
        return "-"
    if results[check]:
        return f"MISMATCH ({len(results[check])})"
    return "stopped" if company.get("stop") else "ok"


def scorecard_lines(companies, all_results):
    """The table (one row per company), then every mismatch by company and check, then the total."""
    lines = [f"{'Company':<13} {'Clean':<14} {'Metrics':<14} {'Flags':<14} {'Gaps':<14} {'Result':<7} Case"]
    for company, results in zip(companies, all_results):
        passed = not any(results.values())
        cells = " ".join(f"{cell(company, results, check):<14}" for check in CHECKS)
        lines.append(f"{company['name']:<13} {cells} {'PASS' if passed else 'FAIL':<7} {company['story']}")
    mismatches = [f"  {company['name']}, {check}: {problem}"
                  for company, results in zip(companies, all_results)
                  for check in CHECKS for problem in results.get(check, [])]
    if mismatches:
        lines += ["", "Mismatches:", *mismatches]
    stopped = [f"  {company['name']} stopped as expected: {company['stop'][0]}"
               for company, results in zip(companies, all_results) if company.get("stop") and not results["Clean"]]
    if stopped:
        lines += ["", "Workbooks that must stop:", *stopped]
    passed = sum(not any(results.values()) for results in all_results)
    lines += ["", f"{passed} of {len(companies)} companies match their answer keys"]
    return lines


def main(companies=None):
    """Evaluate every company (default: the full set), print the scorecard, return 0 if all match, else 1."""
    companies = make_eval_data.COMPANIES if companies is None else companies
    config = load_config()
    all_results = [evaluate(company, config) for company in companies]
    print(f"Evaluation set: {len(companies)} companies, each against its answer key in eval/make_eval_data.py\n")
    print("\n".join(scorecard_lines(companies, all_results)))
    return 0 if all(not any(results.values()) for results in all_results) else 1


if __name__ == "__main__":
    sys.exit(main())
