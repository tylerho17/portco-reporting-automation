"""Build step 3b: compare claude-sonnet-5 and claude-haiku-4-5 on the same job, scored blind.

Two commands:
  python compare_models.py run               # 6 live runs (costs money), saves shuffled answers A-F
  python compare_models.py score A=4 B=3 ... # free: reveals the key, writes README table + recommendation

Rules (agreed before running):
- Same Northwind payload and v3 prompt for every run. Models alternate S, H, S, H, S, H.
- A run that fails (after analyze.py's built-in single retry) is never rerun. It's logged
  with its problems and counts in the pass rate and costs. A failure is data.
- Answers are shuffled and saved without model names or stats, so scoring is blind.
- Recommendation rule: Haiku becomes the batch default only if it averages >= 4.0
  with a 100% pass rate. Otherwise Sonnet stays the default.
"""

import argparse
import json
import random
import sys
import time
from datetime import date
from pathlib import Path

import anthropic
from dotenv import load_dotenv

from analyze import AnalysisError, BoardSummary, analyze, build_payload, payload_to_text, print_commentary
from clean import clean_workbook
from metrics import load_config

# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------

PROJECT_DIR = Path(__file__).parent
WORKBOOK = PROJECT_DIR / "data" / "northwind.xlsx"
COMPANY = "Northwind"
COMPARE_DIR = PROJECT_DIR / "output" / "compare"
KEY_PATH = COMPARE_DIR / "key_DO_NOT_OPEN.json"
README_PATH = PROJECT_DIR / "README.md"
LEARNINGS_PATH = PROJECT_DIR / "LEARNINGS.md"

SONNET = "claude-sonnet-5"
HAIKU = "claude-haiku-4-5"
MODELS = [SONNET, HAIKU]
RUNS_PER_MODEL = 3
COMPANIES_PER_QUARTER = 275

# $ per million tokens: (input, output). Source: Anthropic API pricing table as of PRICES_AS_OF.
PRICES = {SONNET: (2.00, 10.00), HAIKU: (1.00, 5.00)}
PRICES_AS_OF = "2026-06-24"

# Recommendation rule thresholds.
HAIKU_MIN_SCORE = 4.0
HAIKU_MIN_PASS_RATE = 1.0

RUBRIC = [
    (5, "board-ready, no edits"),
    (4, "one or two word-level fixes"),
    (3, "framing or emphasis wrong somewhere"),
    (2, "misleading or missing a key point"),
    (1, "wrong or unusable"),
]


# ---------------------------------------------------------------------------
# Running the models
# ---------------------------------------------------------------------------

def run_cost(model, input_tokens, output_tokens):
    """Dollar cost of a run: tokens x price per million."""
    input_price, output_price = PRICES[model]
    return input_tokens * input_price / 1_000_000 + output_tokens * output_price / 1_000_000


def make_record(model, run_number, run_info, problems, error_type, summary):
    """One run's result in a fixed shape, whether it passed or failed."""
    return {
        "model": model,
        "run": run_number,
        "passed": error_type is None,
        "error_type": error_type,          # None, "validation" or "api"
        "attempts": run_info.get("attempts"),
        "input_tokens": run_info.get("input_tokens", 0),
        "output_tokens": run_info.get("output_tokens", 0),
        "seconds": run_info.get("seconds"),
        "cost": run_cost(model, run_info.get("input_tokens", 0), run_info.get("output_tokens", 0)),
        "problems": problems,
        "summary": summary,
    }


def run_once(client, model, run_number, payload_text):
    """Run analyze() once. Never reruns on failure - the failure itself is recorded."""
    start = time.perf_counter()
    try:
        summary, run_info = analyze(payload_text, model, client)
        return make_record(model, run_number, run_info, [], None, summary.model_dump())
    except AnalysisError as error:  # failed validation on both attempts
        run_info = error.run_info
        problems = run_info["attempt_log"][-1]["problems"]
        return make_record(model, run_number, run_info, problems, "validation", None)
    except anthropic.APIError as error:  # network, auth, overload... still a failed run
        run_info = {"attempts": None, "seconds": round(time.perf_counter() - start, 2)}
        return make_record(model, run_number, run_info, [f"API error: {error}"], "api", None)


def assign_letters(records):
    """Shuffle the runs and label them A, B, C... so the order says nothing about the model."""
    shuffled = list(records)
    random.SystemRandom().shuffle(shuffled)
    return dict(zip("ABCDEFGHIJ", shuffled))


def save_blind_files(lettered):
    """answer_X.json = commentary only. The key file holds models and stats."""
    COMPARE_DIR.mkdir(parents=True, exist_ok=True)
    for letter, record in lettered.items():
        if record["passed"]:
            answer = {"letter": letter, "status": "passed", "summary": record["summary"]}
        else:
            status = "FAILED validation" if record["error_type"] == "validation" else "FAILED (API error)"
            answer = {"letter": letter, "status": status, "problems": record["problems"]}
        (COMPARE_DIR / f"answer_{letter}.json").write_text(json.dumps(answer, indent=2, ensure_ascii=False))

    key = {"created": date.today().isoformat(), "company": COMPANY, "prices_as_of": PRICES_AS_OF,
           "runs": {letter: {k: v for k, v in r.items() if k != "summary"} for letter, r in lettered.items()},
           "scores": None}
    KEY_PATH.write_text(json.dumps(key, indent=2))


def print_blind_answers(lettered):
    """Print every answer in letter order, with no model names or stats."""
    for letter, record in lettered.items():
        print(f"\n{'=' * 30} ANSWER {letter} {'=' * 30}")
        if record["passed"]:
            print_commentary(BoardSummary(**record["summary"]))
        else:
            print("FAILED - not scored. Problems:")
            for problem in record["problems"]:
                print(f"- {problem}")

    print("\nRubric: " + "; ".join(f"{score} = {text}" for score, text in RUBRIC))
    passed = [letter for letter, r in lettered.items() if r["passed"]]
    print("Score every passed answer, then run:\n  python compare_models.py score "
          + " ".join(f"{letter}=?" for letter in passed))


def command_run(force):
    """Run each model RUNS_PER_MODEL times on the same payload and save the blind set."""
    if KEY_PATH.exists() and not force:
        sys.exit(f"{KEY_PATH.relative_to(PROJECT_DIR)} already exists - a blind set was already made. "
                 "Use --force to replace it (this spends money again).")

    load_dotenv()
    actuals, next_budget = clean_workbook(WORKBOOK)
    payload_text = payload_to_text(build_payload(COMPANY, actuals, next_budget, load_config()))
    client = anthropic.Anthropic()

    schedule = [model for _ in range(RUNS_PER_MODEL) for model in MODELS]  # S, H, S, H, S, H
    run_numbers = {model: 0 for model in MODELS}
    records = []
    for position, model in enumerate(schedule, start=1):
        run_numbers[model] += 1
        records.append(run_once(client, model, run_numbers[model], payload_text))
        print(f"Run {position} of {len(schedule)} done")  # no model name: keeps scoring blind

    lettered = assign_letters(records)
    save_blind_files(lettered)
    print_blind_answers(lettered)


# ---------------------------------------------------------------------------
# Scoring and the recommendation
# ---------------------------------------------------------------------------

def parse_scores(score_args, runs):
    """Turn ["A=4", "B=3"] into {"A": 4, "B": 3}. Every passed answer needs a 1-5 score; failed ones get none."""
    scores = {}
    for arg in score_args:
        letter, _, value = arg.partition("=")
        letter = letter.strip().upper()
        if letter not in runs:
            raise ValueError(f"{arg!r}: there is no answer {letter}")
        if not runs[letter]["passed"]:
            raise ValueError(f"{letter} failed validation - it isn't scored (it already counts in the pass rate)")
        if not value.strip().isdigit() or not 1 <= int(value) <= 5:
            raise ValueError(f"{arg!r}: score must be a whole number from 1 to 5")
        scores[letter] = int(value)

    missing = [letter for letter, run in runs.items() if run["passed"] and letter not in scores]
    if missing:
        raise ValueError(f"missing scores for: {', '.join(missing)}")
    return scores


def average(values):
    """Mean of the values that exist; None if there are none."""
    values = [v for v in values if v is not None]
    return sum(values) / len(values) if values else None


def model_stats(model, runs, scores):
    """Pass rate and averages for one model. Cost/time/attempts use all runs; score uses passed runs."""
    mine = {letter: run for letter, run in runs.items() if run["model"] == model}
    passed = [letter for letter, run in mine.items() if run["passed"]]
    avg_cost = average([run["cost"] for run in mine.values()])
    return {
        "model": model,
        "runs": len(mine),
        "passed": len(passed),
        "pass_rate": len(passed) / len(mine),
        "avg_attempts": average([run["attempts"] for run in mine.values()]),
        "avg_cost": avg_cost,
        "avg_seconds": average([run["seconds"] for run in mine.values()]),
        "avg_score": average([scores[letter] for letter in passed]),
        "cost_per_quarter": avg_cost * COMPANIES_PER_QUARTER,
    }


def recommend(haiku):
    """The agreed rule: Haiku only if avg score >= 4.0 AND pass rate is 100%. Returns the default model."""
    score = haiku["avg_score"]
    score_ok = score is not None and round(score, 6) >= HAIKU_MIN_SCORE  # round: avoid float noise
    pass_ok = haiku["pass_rate"] >= HAIKU_MIN_PASS_RATE
    return HAIKU if score_ok and pass_ok else SONNET


def fmt(value, pattern, missing="n/a"):
    """Format a number, or show 'n/a' if it doesn't exist."""
    return missing if value is None else format(value, pattern)


def recommendation_text(sonnet, haiku):
    """The one-line recommendation (plus a cost-range note if Haiku wins)."""
    if recommend(haiku) == HAIKU:
        return (f"Use {HAIKU} as the batch default (${haiku['cost_per_quarter']:,.2f} per quarter for "
                f"{COMPANIES_PER_QUARTER} companies) and {SONNET} for flagged or board-critical companies. "
                f"Actual cost lands between ${haiku['cost_per_quarter']:,.2f} (all Haiku) and "
                f"${sonnet['cost_per_quarter']:,.2f} (all Sonnet), depending on how many companies are escalated.")
    reasons = []
    if haiku["avg_score"] is None or round(haiku["avg_score"], 6) < HAIKU_MIN_SCORE:
        reasons.append(f"Haiku averaged {fmt(haiku['avg_score'], '.1f')} (needs ≥ {HAIKU_MIN_SCORE:.1f})")
    if haiku["pass_rate"] < HAIKU_MIN_PASS_RATE:
        reasons.append(f"Haiku passed {haiku['passed']} of {haiku['runs']} runs (needs all)")
    return (f"Keep {SONNET} as the default (${sonnet['cost_per_quarter']:,.2f} per quarter for "
            f"{COMPANIES_PER_QUARTER} companies): {'; '.join(reasons)}.")


def readme_section(stats, runs, scores):
    """The Markdown for the README's model comparison section."""
    rows = "\n".join(
        f"| {s['model']} | {s['passed']}/{s['runs']} ({s['pass_rate']:.0%}) | {fmt(s['avg_attempts'], '.1f')} "
        f"| ${s['avg_cost']:.4f} | {fmt(s['avg_seconds'], '.1f')} | {fmt(s['avg_score'], '.1f')} "
        f"| ${s['cost_per_quarter']:,.2f} |"
        for s in stats.values())
    rubric = "\n".join(f"| {score} | {text[0].upper() + text[1:]} |" for score, text in RUBRIC)
    by_answer = ", ".join(
        f"{letter} {run['model']} = {scores[letter] if run['passed'] else 'failed'}"
        for letter, run in sorted(runs.items()))
    input_s, output_s = PRICES[SONNET]
    input_h, output_h = PRICES[HAIKU]

    return f"""## Model comparison (build step 3b)

**Setup:** {COMPANY} (fictional data), analyze.py v3 prompt, identical input for every run, {RUNS_PER_MODEL} runs per model, alternating order. Answers were shuffled and scored blind before the model key was opened. Each model ran at its default settings (Sonnet 5 uses adaptive thinking by default; Haiku 4.5 does not), which is how they would be deployed. Runs that failed validation after analyze.py's built-in retry were not rerun: they count in the pass rate, cost and time, but aren't scored.

**Scoring rubric**

| Score | Meaning |
|---|---|
{rubric}

**Results**

| Model | Pass rate | Avg attempts | Avg cost / run | Avg seconds | Avg score | Cost for {COMPANIES_PER_QUARTER} companies / quarter |
|---|---|---|---|---|---|---|
{rows}

Blind scores by answer: {by_answer}.

**Recommendation rule (set before running):** if Haiku averages ≥ {HAIKU_MIN_SCORE:.1f} with a 100% pass rate, Haiku becomes the batch default with Sonnet for flagged or board-critical companies. Otherwise Sonnet stays the default.

**Recommendation:** {recommendation_text(stats[SONNET], stats[HAIKU])}

**Caveats:** {RUNS_PER_MODEL} runs per model on one company, so averages are indicative, not conclusive. Prices as of {PRICES_AS_OF} (per million input/output tokens: Sonnet 5 ${input_s:.0f}/${output_s:.0f}, Haiku 4.5 ${input_h:.0f}/${output_h:.0f}). Projected cost = average cost per run × {COMPANIES_PER_QUARTER}, assuming other companies' data is similar in size to {COMPANY}'s."""


README_START = "<!-- model-comparison:start -->"
README_END = "<!-- model-comparison:end -->"
README_STUB = ("# Board Pack Generator\n\nMessy portfolio-company KPI workbook (.xlsx) → board update deck (.pptx) "
               "+ AI summary. Full README comes in build step 6.\n")


def write_readme_section(section):
    """Put the section between marker comments: replace it if present, append it if not."""
    readme = README_PATH.read_text() if README_PATH.exists() else README_STUB
    block = f"{README_START}\n{section}\n{README_END}"
    if README_START in readme:
        before = readme.split(README_START)[0]
        after = readme.split(README_END, 1)[1]
        readme = before + block + after
    else:
        readme = readme.rstrip() + "\n\n" + block + "\n"
    README_PATH.write_text(readme)


def append_learning(stats):
    """Add a one-entry summary of the result to LEARNINGS.md."""
    sonnet, haiku = stats[SONNET], stats[HAIKU]
    entry = (f"\n## Model comparison (step 3b)\n\n"
             f"- **{date.today().isoformat()}, Sonnet 5 vs Haiku 4.5 on {COMPANY} (v3 prompt, {RUNS_PER_MODEL} runs each, blind)**\n"
             f"  - **Sonnet:** passed {sonnet['passed']}/{sonnet['runs']}, avg score {fmt(sonnet['avg_score'], '.1f')}, "
             f"${sonnet['avg_cost']:.4f}/run, {fmt(sonnet['avg_seconds'], '.1f')}s\n"
             f"  - **Haiku:** passed {haiku['passed']}/{haiku['runs']}, avg score {fmt(haiku['avg_score'], '.1f')}, "
             f"${haiku['avg_cost']:.4f}/run, {fmt(haiku['avg_seconds'], '.1f')}s\n"
             f"  - **Result:** {recommendation_text(sonnet, haiku)}\n")
    with open(LEARNINGS_PATH, "a") as file:
        file.write(entry)


def command_score(score_args):
    """Check scores, reveal the key, compute per-model stats, write README and LEARNINGS."""
    if not KEY_PATH.exists():
        sys.exit("No blind set found - run `python compare_models.py run` first.")
    key = json.loads(KEY_PATH.read_text())
    runs = key["runs"]
    try:
        scores = parse_scores(score_args, runs)
    except ValueError as error:
        sys.exit(f"Score error: {error}")
    already_scored = key["scores"] is not None

    print("Key revealed:")
    for letter, run in sorted(runs.items()):
        result = f"score {scores[letter]}" if run["passed"] else f"FAILED ({run['error_type']})"
        print(f"  {letter} = {run['model']} (run {run['run']}): {result}")

    stats = {model: model_stats(model, runs, scores) for model in MODELS}
    write_readme_section(readme_section(stats, runs, scores))
    key["scores"] = scores
    KEY_PATH.write_text(json.dumps(key, indent=2))
    if already_scored:
        print("Scores were already recorded before - LEARNINGS.md not appended again.")
    else:
        append_learning(stats)

    print(f"\nRecommendation: {recommendation_text(stats[SONNET], stats[HAIKU])}")
    print("README.md model comparison section written.")


def main():
    parser = argparse.ArgumentParser(description="Blind comparison of Claude models for analyze.py.")
    commands = parser.add_subparsers(dest="command", required=True)
    run_parser = commands.add_parser("run", help="run both models (costs money) and save the blind set")
    run_parser.add_argument("--force", action="store_true", help="replace an existing blind set")
    score_parser = commands.add_parser("score", help="enter scores, e.g. A=4 B=3, and reveal the key")
    score_parser.add_argument("scores", nargs="+")
    args = parser.parse_args()

    if args.command == "run":
        command_run(args.force)
    else:
        command_score(args.scores)


if __name__ == "__main__":
    main()
