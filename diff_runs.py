"""What changed since the last run (Task 10): flags that flipped, metrics that moved, data gaps opened or closed.

Every run saves its results in output/<company>_manifest.json ("results"): the latest quarter,
every metric's value and the words the deck shows for it, every flag's status, and every data gap.
This file compares one run's results with an earlier run's and says what changed:
- Flags that flipped:  a flag whose status changed ("Runway at current burn: Tripped (was Passed)"),
                       including to or from "Cannot evaluate", or from one reason to another
- Metrics that moved:  the latest quarter's value moved more than a set amount. Percentages move
                       by points (default 5 points); everything else ($K, months, multiples) by
                       percent of its old value (default 10%). A value that became or stopped
                       being a number (∞, data missing, n/m) is always listed.
- Data gaps:           gaps new in this run, and gaps the earlier run had that are now filled,
                       quarter by quarter.

Which earlier run: the last one whose results were different. A rebuild from the same numbers
(after approve.py, say) keeps the comparison it already had, instead of comparing the run with
itself and saying "nothing changed". The manifest keeps that run in "previous_run".

The amounts can be set in config.yaml (diff_min_points, diff_min_relative) or on the command line;
without them the defaults below apply.

Rules: no new math beyond the size of each move, and no em dash. Every value shown is the text
the deck shows for it (build_deck.value_text), saved when that run was built.

Run: python diff_runs.py data/northwind.xlsx       (today's workbook vs its last run in output/)
     python diff_runs.py data/northwind.xlsx --min-points 0.02 --min-relative 0.2
"""

import argparse
import json
import math
import sys
from pathlib import Path

from build_deck import OUTPUT_DIR, collect_deck_data, gaps_lines, points_text, value_text
from clean import clean_workbook
from excel_output import status_label
from metrics import DOLLAR_COLUMNS, METRIC_LABELS, MONTH_COLUMNS, load_config
from provenance import manifest_path, read_manifest

PROJECT_DIR = Path(__file__).parent

DEFAULT_MIN_POINTS = 0.05      # a percentage must move more than 5 points (NRR 102% to 96.9%)
DEFAULT_MIN_RELATIVE = 0.10    # anything else must move more than 10% of its old value
SETTING_KEYS = {"min_points": "diff_min_points", "min_relative": "diff_min_relative"}   # optional in config.yaml
ROUND_TO = 6                   # decimals kept: float noise (0.30000000000000004) is never a change

# Metrics that aren't percentages, so they move by percent of their old value, not by points.
NOT_PERCENT = DOLLAR_COLUMNS | MONTH_COLUMNS | {"burn_multiple"}

HEADING = "What changed since the last run"
NO_EARLIER_RUN = "Nothing to compare with yet: this is the first run on record."
NOT_CHECKED = "Not checked"            # a flag the other run didn't evaluate (the combo rule switched off)
FROM_ZERO = "from zero"                # no percent of zero exists: any move away from it counts
RESULT_KEYS = {"quarter", "metrics", "flags", "gaps"}


# ---------------------------------------------------------------------------
# How far a metric must move to be listed
# ---------------------------------------------------------------------------

def setting_value(name, value):
    """The setting if it's a number of 0 or more, else stop naming its config.yaml key."""
    if not isinstance(value, (int, float)) or isinstance(value, bool) or value < 0:
        raise ValueError(f"config.yaml: {SETTING_KEYS[name]} must be a number of 0 or more (got {value!r})")
    return value


def move_settings(config=None, min_points=None, min_relative=None):
    """{"min_points", "min_relative"}: an argument (the command line) beats config.yaml, which beats the default."""
    config = config or {}
    chosen = {"min_points": min_points, "min_relative": min_relative}
    defaults = {"min_points": DEFAULT_MIN_POINTS, "min_relative": DEFAULT_MIN_RELATIVE}
    for name, key in SETTING_KEYS.items():
        if chosen[name] is None:
            chosen[name] = config.get(key, defaults[name])
    return {name: setting_value(name, value) for name, value in chosen.items()}


# ---------------------------------------------------------------------------
# What a run records
# ---------------------------------------------------------------------------

def number_or_none(value):
    """A finite value, rounded (JSON has no NaN or infinity); None when there's no number to compare."""
    return round(float(value), ROUND_TO) if math.isfinite(value) else None


def run_results(data):
    """What the manifest saves about a run (data = build_deck.collect_deck_data's dict).

    The latest quarter; each metric's value and the words the deck shows for it; each flag's
    status as the Excel Flags sheet words it; and every data gap ({metric or "flag: <name>": [quarters]}).
    """
    latest = data["latest"]
    results = {
        "quarter": latest,
        "metrics": {metric: {"value": number_or_none(data["metrics"].loc[latest, metric]),
                             "shown": value_text(data, metric, latest)}
                    for metric in data["metrics"].columns},
        "flags": {flag["flag"]: status_label(flag) for flag in data["flags"]},
        "gaps": {name: list(quarters) for name, quarters in data["gaps"].items()},
    }
    return json.loads(json.dumps(results))   # exactly what reading it back from the manifest gives


def looks_like_results(results):
    """True if it has a run's four parts (a hand-edited manifest might not)."""
    return isinstance(results, dict) and RESULT_KEYS <= set(results)


def baseline(manifest, results):
    """The earlier run to compare these results with: {"run_at", "results"}, or None if there isn't one.

    The manifest's own run if its results differ. If they're the same (a rebuild from the same
    numbers), the run that manifest was already compared with, so a rebuild keeps its comparison.
    """
    saved = (manifest or {}).get("results")
    if not looks_like_results(saved):
        return None   # no manifest, or one written before Task 10
    if saved != results:
        return {"run_at": manifest.get("run_at"), "results": saved}
    return manifest.get("previous_run")


# ---------------------------------------------------------------------------
# Comparing two runs
# ---------------------------------------------------------------------------

def flag_flips(before, after):
    """[(flag, status before, status after)] for every flag whose status changed, in this run's order."""
    names = list(after) + [name for name in before if name not in after]
    return [(name, before.get(name, NOT_CHECKED), after.get(name, NOT_CHECKED)) for name in names
            if before.get(name) != after.get(name)]


def move_text(metric, before, after, settings):
    """'down 5.1 pts' or 'up 10.1%' if the move is more than the setting, else None."""
    if after == before:
        return None   # zero to zero is no move, whatever the setting
    direction = "up" if after > before else "down"
    if metric not in NOT_PERCENT:
        points = round(abs(after - before), ROUND_TO)
        return f"{direction} {points_text(points)}" if points > settings["min_points"] else None
    if before == 0:
        return f"{direction} {FROM_ZERO}"
    relative = round(abs(after - before) / abs(before), ROUND_TO)
    return f"{direction} {relative:.1%}" if relative > settings["min_relative"] else None


def metric_moves(before, after, settings):
    """[(label, shown before, shown after, size of the move or None)] for every metric that moved enough.

    Two numbers: listed if the move is more than the setting. Otherwise (∞, data missing, n/m on
    either side): listed if the words changed, with no size, since there's no number to subtract.
    """
    moved = []
    for metric, now in after.items():
        then = before.get(metric)
        if then is None:
            continue   # a metric the earlier run didn't have
        if then["value"] is not None and now["value"] is not None:
            size = move_text(metric, then["value"], now["value"], settings)
            if size:
                moved.append((METRIC_LABELS[metric], then["shown"], now["shown"], size))
        elif then["shown"] != now["shown"]:
            moved.append((METRIC_LABELS[metric], then["shown"], now["shown"], None))
    return moved


def gaps_only_in(gaps, other):
    """The gaps in `gaps` that `other` doesn't have, quarter by quarter: {name: [quarters]}."""
    left = {}
    for name, quarters in gaps.items():
        missing = [quarter for quarter in quarters if quarter not in other.get(name, [])]
        if missing:
            left[name] = missing
    return left


def compare(previous, current, settings):
    """Everything that changed between two runs' results: flags, moved metrics, new and resolved gaps."""
    return {
        "flags": flag_flips(previous["flags"], current["flags"]),
        "moved": metric_moves(previous["metrics"], current["metrics"], settings),
        "new_gaps": gaps_only_in(current["gaps"], previous["gaps"]),
        "resolved_gaps": gaps_only_in(previous["gaps"], current["gaps"]),
    }


def report_for(earlier, current, settings):
    """The comparison with an earlier run (baseline's answer), with when it ran and both latest quarters."""
    return {"since": earlier.get("run_at"), "quarter_before": earlier["results"]["quarter"],
            "quarter_now": current["quarter"], **compare(earlier["results"], current, settings)}


def changes_since_last_run(data, manifest, settings):
    """What changed between the last different run in `manifest` and today's data, or None if there's none."""
    current = run_results(data)
    earlier = baseline(manifest, current)
    return None if earlier is None else report_for(earlier, current, settings)


# ---------------------------------------------------------------------------
# The words (the web page, the memo and the command line all use these)
# ---------------------------------------------------------------------------

def when_text(run_at):
    """'2026-06-18T09:05:41' -> '2026-06-18 09:05' (portfolio.last_run_text's form)."""
    return (run_at or "an unknown time").replace("T", " ")[:16]


def compared_with_text(report):
    """Which run this is compared with, and both latest quarters."""
    when = when_text(report["since"])
    if report["quarter_before"] == report["quarter_now"]:
        return f"Compared with the run of {when}, on the same latest quarter ({report['quarter_now']})."
    return (f"Compared with the run of {when}, whose latest quarter was {report['quarter_before']} "
            f"(now {report['quarter_now']}).")


def threshold_words(settings):
    """'5.0 pts (percentages) or 10.0% (other metrics)'."""
    return f"{points_text(settings['min_points'])} (percentages) or {settings['min_relative']:.1%} (other metrics)"


def flip_line(name, before, after):
    """'Runway at current burn: Tripped (was Passed)': today's status first.

    Not "Passed to Tripped": a status with a reason has a colon of its own, and "Cannot evaluate:
    missing input to Tripped" reads as one muddle.
    """
    return f"{name}: {after} (was {before})"


def moved_line(label, before, after, size):
    """'NRR (annualized): 102.0% to 90.0% (down 12.0 pts)'."""
    return f"{label}: {before} to {after}" + (f" ({size})" if size else "")


def change_sections(report, settings):
    """[(title, lines)]: flags that flipped, metrics that moved, new gaps, resolved gaps; empty ones left out.

    When nothing changed, one section says so and what was checked.
    """
    sections = [
        ("Flags that flipped", [flip_line(*flip) for flip in report["flags"]]),
        (f"Metrics that moved more than {threshold_words(settings)}",
         [moved_line(*move) for move in report["moved"]]),
        ("New data gaps", gaps_lines(report["new_gaps"]) if report["new_gaps"] else []),
        ("Resolved data gaps", gaps_lines(report["resolved_gaps"]) if report["resolved_gaps"] else []),
    ]
    sections = [(title, lines) for title, lines in sections if lines]
    if not sections:
        return [("Nothing changed", [f"No flag flipped, no metric moved more than {threshold_words(settings)}, "
                                     f"and no data gap opened or closed."])]
    return sections


def count_text(count, one, many):
    return f"{count} {one if count == 1 else many}"


def summary_text(report):
    """One line for the command line: '1 flag flipped, 0 metrics moved, 0 new data gaps, 1 resolved (since ...)'."""
    if report is None:
        return NO_EARLIER_RUN
    new = sum(len(quarters) for quarters in report["new_gaps"].values())
    resolved = sum(len(quarters) for quarters in report["resolved_gaps"].values())
    return (f"{count_text(len(report['flags']), 'flag flipped', 'flags flipped')}, "
            f"{count_text(len(report['moved']), 'metric moved', 'metrics moved')}, "
            f"{count_text(new, 'new data gap', 'new data gaps')}, {resolved} resolved "
            f"(since the run of {when_text(report['since'])})")


# ---------------------------------------------------------------------------
# Command line
# ---------------------------------------------------------------------------

def workbook_data(workbook_path, config):
    """Clean the workbook and compute what the deck shows (build_deck.collect_deck_data)."""
    actuals, next_budget = clean_workbook(workbook_path)
    return collect_deck_data(Path(workbook_path).stem.title(), Path(workbook_path).name, actuals, next_budget, config)


def printed_lines(company, report, settings):
    """What the command line prints: the heading, which run, then each section and its lines."""
    lines = [f"{company}: {HEADING}"]
    if report is None:
        return lines + [NO_EARLIER_RUN]
    lines.append(compared_with_text(report))
    for title, items in change_sections(report, settings):
        lines += [f"{title}:"] + [f"  - {item}" for item in items]
    return lines


def main(argv=None):
    parser = argparse.ArgumentParser(description="What changed between a company's workbook today and its last run.")
    parser.add_argument("workbook", help="path to a KPI workbook, e.g. data/northwind.xlsx")
    parser.add_argument("--min-points", type=float, help="percentages: list a move of more than this (0.05 = 5 points)")
    parser.add_argument("--min-relative", type=float,
                        help="other metrics: list a move of more than this share of the old value (0.1 = 10%%)")
    parser.add_argument("--output-dir", default=OUTPUT_DIR, help="where the manifest is (default: output/)")
    args = parser.parse_args(argv)
    config = load_config()
    settings = move_settings(config, args.min_points, args.min_relative)
    data = workbook_data(args.workbook, config)
    manifest = read_manifest(manifest_path(args.workbook, args.output_dir))
    report = changes_since_last_run(data, manifest, settings)
    print("\n".join(printed_lines(data["company"], report, settings)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
