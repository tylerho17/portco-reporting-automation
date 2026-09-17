"""Ask Claude to write the board commentary from the computed metrics.

Flow: clean workbook -> metrics/flags/gaps (Python) -> payload (formatted facts)
      -> Claude -> validate -> retry once if needed -> save JSON.

Guardrails (CLAUDE.md):
- Python computes every number. Claude only interprets.
- Every number in Claude's answer must appear in the payload (the number check).
- The answer must match the BoardSummary shape, with exactly 3 wins/risks/questions.
- One retry with feedback; if that fails too, stop with an error. Nothing
  unvalidated is ever returned.

Run: python analyze.py data/northwind.xlsx [--model claude-haiku-4-5]
"""

import argparse
import json
import re
import sys
import time
from pathlib import Path

import anthropic
import pandas as pd
from dotenv import load_dotenv
from pydantic import BaseModel, ValidationError

from clean import clean_workbook
from metrics import (CANNOT_EVALUATE, INPUT_LABELS, METRIC_LABELS, MISSING_INPUT, PASS, REASON_DISPLAY, TRIP,
                     compute_metrics, data_gaps, display_value, evaluate_flags, flag_status_text, format_value,
                     load_config, metric_reasons, runway_at_next_budget, runway_context_label)

DEFAULT_MODEL = "claude-sonnet-5"
MAX_ATTEMPTS = 2           # first try + one retry
MAX_TOKENS = 16000         # room for thinking + the answer
MAX_HEADLINE_WORDS = 30    # prompt asks for 25; small buffer before we fail it
MAX_DETAIL_WORDS = 45      # prompt asks for ~40; small buffer before we fail it
MAX_DETAIL_SENTENCES = 2
OUTPUT_DIR = Path(__file__).parent / "output"


# ---------------------------------------------------------------------------
# 1. The shape Claude's answer must have
# ---------------------------------------------------------------------------

class Point(BaseModel):
    title: str   # short phrase for the slide
    detail: str  # 1-2 sentences with the evidence


class BoardSummary(BaseModel):
    headline: str
    wins: list[Point]
    risks: list[Point]
    questions: list[str]


# ---------------------------------------------------------------------------
# 2. The payload: the facts Claude sees, already formatted as text
# ---------------------------------------------------------------------------

# Labels (INPUT_LABELS, METRIC_LABELS) come from metrics.py, so every output uses the same words.
STATUS_LABELS = {TRIP: "TRIPPED", PASS: "passed"}  # a flag that can't be evaluated shows its reason instead


def input_trend(actuals, column):
    """One raw input across all quarters, e.g. {"Q3 2024": "2,400", ...}. A blank input is data missing."""
    return {q: REASON_DISPLAY[MISSING_INPUT] if pd.isna(v) else format_value(column, v)
            for q, v in actuals[column].items()}


def metric_trend(actuals, metrics, reasons, column):
    """One metric across all quarters: its value, or why there isn't one (data missing / no prior period / n/m)."""
    return {quarter: display_value(actuals, metrics, reasons, column, quarter) for quarter in metrics.index}


def flag_status(flag):
    """'TRIPPED', 'passed', or 'cannot evaluate — <reason>'."""
    return flag_status_text(flag) if flag["status"] == CANNOT_EVALUATE else STATUS_LABELS[flag["status"]]


def describe_flag(flag, config, actuals, metrics, reasons):
    """One flag as text: name, status, value, threshold."""
    if flag["metric"] is None:  # the combo rule has no single value
        size = config["combo_lookback_quarters"]
        return {"flag": flag["flag"], "status": flag_status(flag),
                "value": "see NRR and Pipeline trends",
                "rule": f"NRR down and pipeline up at every step over the last {size} quarters"}
    return {"flag": flag["flag"], "status": flag_status(flag),
            "value": display_value(actuals, metrics, reasons, flag["metric"], flag["quarter"]),
            "threshold": format_value(flag["metric"], flag["threshold"])}


def build_payload(company, actuals, next_budget, config):
    """Run the step 2 math and package every fact Claude may use, formatted as text."""
    metrics = compute_metrics(actuals)
    reasons = metric_reasons(actuals, metrics)
    flags = evaluate_flags(metrics, reasons, config)
    gaps = data_gaps(actuals, metrics, flags)
    runway = runway_at_next_budget(actuals, next_budget)

    trends = {label: input_trend(actuals, column) for column, label in INPUT_LABELS.items()}
    trends.update({label: metric_trend(actuals, metrics, reasons, column) for column, label in METRIC_LABELS.items()})

    return {
        "company": company,
        "latest_quarter": metrics.index[-1],
        "units": "Money figures are $K (thousands of US dollars). NRR and GRR are annualized from one quarter.",
        # Counts are computed here so Claude quotes them instead of counting.
        "flags_tripped": sum(f["status"] == TRIP for f in flags),
        "flags_passed": sum(f["status"] == PASS for f in flags),
        "flags_cannot_evaluate": sum(f["status"] == CANNOT_EVALUATE for f in flags),
        "flags_total": len(flags),
        "flags_latest_quarter": [describe_flag(f, config, actuals, metrics, reasons) for f in flags],
        # latest cash at next quarter's budgeted burn
        "runway_if_burn_returns_to_plan": (runway_context_label(runway, next_budget is not None)
                                           or format_value("runway_months", runway)),
        "trends_by_quarter": trends,
        "data_gaps": {METRIC_LABELS.get(name, name): quarters for name, quarters in gaps.items()},
    }


def payload_to_text(payload):
    """The exact JSON text sent to Claude (and used by the number check)."""
    return json.dumps(payload, indent=2, ensure_ascii=False)


# ---------------------------------------------------------------------------
# 3. Instructions for Claude
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = """You are a private equity portfolio analyst writing the commentary for a quarterly board update. Your readers sit on the investment committee: they want the few things that matter, stated plainly, with the evidence.

You will receive a portfolio company's KPI data as JSON. Every number in it was computed and checked in Python. Your job is interpretation, not arithmetic.

Rules for numbers - an automated check enforces these, and any violation fails your answer:
- Use only numbers that appear in the data, written exactly as shown. Do not round or reformat them.
- Never calculate a new number: no differences, sums, averages, ratios or percentage-point changes. To describe a change, quote the start value and the end value ("went from A to B").
- Money figures are in $K. Quote them as shown followed by "K" (for example "$" + value + "K"); do not convert to millions.
- Keep minus signs exactly as shown: a negative value is written with its minus sign, never as a positive number with "below" or "short".
- Where the data says "data missing", say the data is missing. Never estimate it. "n/a (no prior period)" means there is no earlier quarter to compare with, and "n/m" means the figure isn't meaningful - neither is missing data.

What to write:
- headline: one sentence, at most 25 words, with the single most important takeaway for the board.
- wins: exactly 3 genuine strengths supported by the data. Lead with the strongest growth or scale metric (such as ARR growth YoY) before any threshold passes. A metric that is merely "within threshold" counts as a win only if nothing stronger exists. If strengths are thin, don't overstate them.
- risks: exactly 3, starting with the most serious tripped flags. Explain why each matters to an investor, not just that a threshold was crossed. Where several flags point to one underlying problem, combine them into one risk.
- questions: exactly 3 specific questions for management that the data raises but cannot answer, such as what is driving a trend. No yes/no questions.
Each title is a short phrase. Each detail is at most 2 sentences and about 40 words - it goes on a slide, so pick the evidence that matters most rather than listing everything. Answers with longer details are rejected.

Accuracy of framing:
- When a flag passed, say so explicitly, quoting its value next to its threshold (for example: "passed, but close to its threshold (value vs threshold) - watch"). Never place a passing metric where it reads as a breach.
- Describe a trend from its peak, or from the start of the flag's lookback window, not from the first quarter in the data.
- Never call a quarter a "significant" (or large, major, sharp) miss or beat against budget unless you quote its value and that value is more than 10% away from budget - above +10.0% or below -10.0%. Smaller variances are described plainly, with their value.
- "runway_if_burn_returns_to_plan" is runway if burn returns to plan. Describe it only in those words. It is not a projection, forecast or improvement."""


# ---------------------------------------------------------------------------
# 4-5. Validation
# ---------------------------------------------------------------------------

# A number: an optional minus sign, digits with optional thousands commas, optional decimals.
# The minus can be "-", "−" (true minus) or "–" (en dash), optionally before a "$" ("-$240K").
# It only counts when it isn't right after a letter or digit, so "2025–2026" and "Q4 2025-Q2 2026"
# stay ranges, not negative numbers.
NUMBER_PATTERN = re.compile(r"(?:(?<![0-9A-Za-z])([-−–])\$?)?(\d[\d,]*(?:\.\d+)?)")


def numbers_in(text):
    """Every number in a piece of text, as floats, with its sign. Ignores $, %, x, K and commas."""
    numbers = set()
    for sign, digits in NUMBER_PATTERN.findall(text):
        value = float(digits.replace(",", ""))
        numbers.add(-value if sign else value)
    return numbers


def count_sentences(text):
    """Sentences = pieces ending in . ! or ? followed by a space or the end.

    The decimal point in "97.1%" isn't followed by a space, so it doesn't count.
    Known limit: an abbreviation like "vs. " would count as a sentence end.
    """
    return len([s for s in re.split(r"(?<=[.!?])\s+", text.strip()) if s])


def summary_texts(summary):
    """All the text Claude wrote, as one list."""
    texts = [summary.headline] + summary.questions
    for point in summary.wins + summary.risks:
        texts += [point.title, point.detail]
    return texts


def find_ungrounded_numbers(summary, payload_text):
    """Numbers Claude wrote that don't appear anywhere in the payload (sorted)."""
    allowed = numbers_in(payload_text)
    used = numbers_in(" ".join(summary_texts(summary)))
    return sorted(used - allowed)


def validate_summary(summary, payload_text):
    """Return a list of problems with Claude's answer. Empty list = pass."""
    problems = []
    for field in ("wins", "risks", "questions"):
        count = len(getattr(summary, field))
        if count != 3:
            problems.append(f"{field} must have exactly 3 items, got {count}")
    if any(not text.strip() for text in summary_texts(summary)):
        problems.append("some text fields are empty")
    headline_words = len(summary.headline.split())
    if headline_words > MAX_HEADLINE_WORDS:
        problems.append(f"headline has {headline_words} words, max is {MAX_HEADLINE_WORDS}")
    for section in ("wins", "risks"):
        for number, point in enumerate(getattr(summary, section), start=1):
            words, sentences = len(point.detail.split()), count_sentences(point.detail)
            if words > MAX_DETAIL_WORDS or sentences > MAX_DETAIL_SENTENCES:
                problems.append(f"{section} #{number} detail has {words} words and {sentences} sentences; "
                                f"max is {MAX_DETAIL_WORDS} words and {MAX_DETAIL_SENTENCES} sentences")
    ungrounded = find_ungrounded_numbers(summary, payload_text)
    if ungrounded:
        problems.append("these numbers are not in the data (rounded or calculated?): "
                        + ", ".join(f"{n:g}" for n in ungrounded))
    return problems


# ---------------------------------------------------------------------------
# 6-7. Calling Claude, with one retry
# ---------------------------------------------------------------------------

class AnalysisError(Exception):
    """Both attempts failed validation. Carries run_info so failures can be logged (step 3b)."""

    def __init__(self, problems, run_info):
        super().__init__("Claude's answer failed validation twice:\n- " + "\n- ".join(problems))
        self.run_info = run_info


def build_messages(payload_text, previous=None):
    """The conversation to send. On a retry, include the last answer and what was wrong with it."""
    messages = [{"role": "user", "content": f"KPI data:\n\n{payload_text}"}]
    if previous is not None:
        if previous["answer_text"]:
            messages.append({"role": "assistant", "content": previous["answer_text"]})
        messages.append({"role": "user", "content": (
            "That answer failed validation:\n- " + "\n- ".join(previous["problems"])
            + "\nReturn a corrected answer that fixes every problem.")})
    return messages


def call_claude(client, model, payload_text, previous=None):
    """One API call. Returns the summary (or None), its problems, and a log entry."""
    start = time.perf_counter()
    try:
        response = client.messages.parse(
            model=model,
            max_tokens=MAX_TOKENS,
            system=SYSTEM_PROMPT,
            messages=build_messages(payload_text, previous),
            output_format=BoardSummary,  # structured outputs: forces this JSON shape
        )
    except ValidationError as error:  # answer didn't match the shape (e.g. cut off)
        seconds = time.perf_counter() - start
        problems = [f"answer did not match the required JSON shape: {error.error_count()} errors"]
        return {"summary": None, "problems": problems, "answer_text": "",
                "log": {"seconds": round(seconds, 2), "input_tokens": None, "output_tokens": None,
                        "stop_reason": None, "problems": problems}}
    seconds = time.perf_counter() - start

    answer_text = "".join(block.text for block in response.content if block.type == "text")
    summary = response.parsed_output
    if response.stop_reason in ("refusal", "max_tokens"):
        problems = [f"Claude stopped early (stop_reason: {response.stop_reason})"]
        summary = None
    elif summary is None:
        problems = ["no structured answer was returned"]
    else:
        problems = validate_summary(summary, payload_text)

    log = {"seconds": round(seconds, 2),
           "input_tokens": response.usage.input_tokens,
           "output_tokens": response.usage.output_tokens,
           "stop_reason": response.stop_reason,
           "problems": problems}
    return {"summary": summary, "problems": problems, "answer_text": answer_text, "log": log}


def make_run_info(model, attempts, passed):
    """Totals across attempts, for printing now and the model comparison in step 3b."""
    return {
        "model": model,
        "passed": passed,
        "attempts": len(attempts),
        "input_tokens": sum(a["input_tokens"] or 0 for a in attempts),
        "output_tokens": sum(a["output_tokens"] or 0 for a in attempts),
        "seconds": round(sum(a["seconds"] for a in attempts), 2),
        "attempt_log": attempts,
    }


def analyze(payload_text, model=DEFAULT_MODEL, client=None):
    """Get a validated BoardSummary from Claude. Returns (summary, run_info) or raises AnalysisError."""
    client = client or anthropic.Anthropic()  # reads ANTHROPIC_API_KEY from the environment
    attempts = []
    previous = None
    for _ in range(MAX_ATTEMPTS):
        result = call_claude(client, model, payload_text, previous)
        attempts.append(result["log"])
        if not result["problems"]:
            return result["summary"], make_run_info(model, attempts, passed=True)
        previous = result  # feed the problems into the retry
    raise AnalysisError(previous["problems"], make_run_info(model, attempts, passed=False))


# ---------------------------------------------------------------------------
# 8. Command line
# ---------------------------------------------------------------------------

def save_analysis(path, payload, summary=None, run_info=None, error=None):
    """Save what Claude saw next to what it wrote, so every claim can be traced. Returns the path.

    summary=None means there is no validated answer (error says why), and build_deck.py then
    shows "AI summary unavailable". analyze.py and main.py both save through here.
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    record = {"summary": summary.model_dump() if summary else None, "run_info": run_info,
              "error": error, "payload": payload}
    path.write_text(json.dumps(record, indent=2, ensure_ascii=False))
    return path


def print_commentary(summary):
    """Show the commentary only (no model or stats - used for blind scoring too)."""
    print(f"\nHEADLINE: {summary.headline}")
    for section in ("wins", "risks"):
        print(f"\n{section.upper()}")
        for point in getattr(summary, section):
            print(f"- {point.title}: {point.detail}")
    print("\nQUESTIONS FOR MANAGEMENT")
    for question in summary.questions:
        print(f"- {question}")


def print_summary(summary, run_info):
    """Show the commentary and the run stats in the terminal."""
    print_commentary(summary)
    print(f"\nRun: {run_info['model']}, {run_info['attempts']} attempt(s), "
          f"{run_info['input_tokens']} in / {run_info['output_tokens']} out tokens, {run_info['seconds']}s")


def main():
    parser = argparse.ArgumentParser(description="Write board commentary with Claude.")
    parser.add_argument("workbook", help="path to a KPI workbook, e.g. data/northwind.xlsx")
    parser.add_argument("--model", default=DEFAULT_MODEL, help=f"Claude model (default {DEFAULT_MODEL})")
    args = parser.parse_args()

    load_dotenv()  # puts ANTHROPIC_API_KEY from .env into the environment
    company_key = Path(args.workbook).stem              # "northwind"
    actuals, next_budget = clean_workbook(args.workbook)
    payload = build_payload(company_key.title(), actuals, next_budget, load_config())
    payload_text = payload_to_text(payload)

    output_path = OUTPUT_DIR / f"{company_key}_analysis.json"
    try:
        summary, run_info = analyze(payload_text, args.model)
    except AnalysisError as error:
        save_analysis(output_path, payload, run_info=error.run_info, error=str(error))
        print(error)
        sys.exit(1)

    save_analysis(output_path, payload, summary, run_info)
    print_summary(summary, run_info)
    print(f"Saved {output_path.relative_to(Path(__file__).parent)}")


if __name__ == "__main__":
    main()
