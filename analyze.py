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
from metrics import (MISSING, PASS, TRIP, compute_metrics, data_gaps, evaluate_flags,
                     format_value, load_config, runway_at_next_budget)

DEFAULT_MODEL = "claude-sonnet-5"
MAX_ATTEMPTS = 2           # first try + one retry
MAX_TOKENS = 16000         # room for thinking + the answer
MAX_HEADLINE_WORDS = 30    # prompt asks for 25; small buffer before we fail it
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

# Column -> label shown to Claude. Raw inputs first, then computed metrics.
INPUT_LABELS = {
    "net_burn": "Net burn ($K)",
    "ending_cash": "Ending cash ($K)",
}
METRIC_LABELS = {
    "ending_arr": "Ending ARR ($K)",
    "net_new_arr": "Net new ARR ($K)",
    "nrr": "NRR (annualized)",
    "grr": "GRR (annualized)",
    "gross_margin": "Gross margin",
    "arr_qoq": "ARR growth QoQ",
    "arr_yoy": "ARR growth YoY",
    "revenue_qoq": "Revenue growth QoQ",
    "revenue_yoy": "Revenue growth YoY",
    "pipeline": "Pipeline ($K)",
    "pipeline_qoq": "Pipeline growth QoQ",
    "fcf_margin": "FCF margin",
    "rule_of_40": "Rule of 40",
    "burn_multiple": "Burn multiple",
    "burn_vs_budget": "Net burn vs budget",
    "net_new_arr_vs_budget": "Net new ARR vs budget",
    "arr_vs_budget": "Ending ARR vs budget",
    "cac_payback_months": "CAC payback",
    "runway_months": "Runway at current burn",
}
STATUS_LABELS = {TRIP: "TRIPPED", PASS: "passed", MISSING: MISSING}


def input_trend(actuals, column):
    """One raw input across all quarters, e.g. {"Q3 2024": "2,400", ...}."""
    return {q: "data missing" if pd.isna(v) else format_value(column, v)
            for q, v in actuals[column].items()}


def metric_trend(metrics, gaps, column):
    """One metric across all quarters. Distinguishes a data gap from 'no prior period'."""
    trend = {}
    for quarter, value in metrics[column].items():
        if quarter in gaps.get(column, []):
            trend[quarter] = "data missing"
        elif pd.isna(value):
            trend[quarter] = "n/a (no prior period)"  # e.g. YoY in the first 4 quarters
        else:
            trend[quarter] = format_value(column, value)
    return trend


def describe_flag(flag, config):
    """One flag as text: name, status, value, threshold."""
    if flag["metric"] is None:  # the combo rule has no single value
        size = config["combo_lookback_quarters"]
        return {"flag": flag["flag"], "status": STATUS_LABELS[flag["status"]],
                "value": "see NRR and Pipeline trends",
                "rule": f"NRR down and pipeline up at every step over the last {size} quarters"}
    return {"flag": flag["flag"], "status": STATUS_LABELS[flag["status"]],
            "value": format_value(flag["metric"], flag["value"]),
            "threshold": format_value(flag["metric"], flag["threshold"])}


def build_payload(company, actuals, next_budget, config):
    """Run the step 2 math and package every fact Claude may use, formatted as text."""
    metrics = compute_metrics(actuals)
    flags = evaluate_flags(metrics, config)
    gaps = data_gaps(actuals, metrics, flags)

    trends = {label: input_trend(actuals, column) for column, label in INPUT_LABELS.items()}
    trends.update({label: metric_trend(metrics, gaps, column) for column, label in METRIC_LABELS.items()})

    return {
        "company": company,
        "latest_quarter": metrics.index[-1],
        "units": "Money figures are $K (thousands of US dollars). NRR and GRR are annualized from one quarter.",
        # Counts are computed here so Claude quotes them instead of counting.
        "flags_tripped": sum(f["status"] == TRIP for f in flags),
        "flags_passed": sum(f["status"] == PASS for f in flags),
        "flags_cannot_evaluate": sum(f["status"] == MISSING for f in flags),
        "flags_total": len(flags),
        "flags_latest_quarter": [describe_flag(f, config) for f in flags],
        "runway_at_next_quarter_budgeted_burn": format_value(
            "runway_months", runway_at_next_budget(actuals, next_budget)),
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
- Where the data says "data missing", say the data is missing. Never estimate it.

What to write:
- headline: one sentence, at most 25 words, with the single most important takeaway for the board.
- wins: exactly 3 genuine strengths supported by the data. If strengths are thin, pick the most meaningful and don't overstate them.
- risks: exactly 3, starting with the most serious tripped flags. Explain why each matters to an investor, not just that a threshold was crossed. Where several flags point to one underlying problem, combine them into one risk.
- questions: exactly 3 specific questions for management that the data raises but cannot answer, such as what is driving a trend. No yes/no questions.
Each title is a short phrase. Each detail is 1-2 sentences."""


# ---------------------------------------------------------------------------
# 4-5. Validation
# ---------------------------------------------------------------------------

# A number: digits, optional thousands commas, optional decimals. "$27,470K" -> "27,470".
NUMBER_PATTERN = re.compile(r"\d[\d,]*(?:\.\d+)?")


def numbers_in(text):
    """Every number in a piece of text, as floats. Ignores $, %, x, K, signs and commas."""
    return {float(match.replace(",", "")) for match in NUMBER_PATTERN.findall(text)}


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

def print_summary(summary, run_info):
    """Show the commentary and the run stats in the terminal."""
    print(f"\nHEADLINE: {summary.headline}")
    for section in ("wins", "risks"):
        print(f"\n{section.upper()}")
        for point in getattr(summary, section):
            print(f"- {point.title}: {point.detail}")
    print("\nQUESTIONS FOR MANAGEMENT")
    for question in summary.questions:
        print(f"- {question}")
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

    OUTPUT_DIR.mkdir(exist_ok=True)
    output_path = OUTPUT_DIR / f"{company_key}_analysis.json"
    try:
        summary, run_info = analyze(payload_text, args.model)
    except AnalysisError as error:
        output_path.write_text(json.dumps({"summary": None, "run_info": error.run_info,
                                           "payload": payload}, indent=2, ensure_ascii=False))
        print(error)
        sys.exit(1)

    # Save what Claude saw next to what it wrote, so every claim can be traced.
    output_path.write_text(json.dumps({"summary": summary.model_dump(), "run_info": run_info,
                                       "payload": payload}, indent=2, ensure_ascii=False))
    print_summary(summary, run_info)
    print(f"Saved {output_path.relative_to(Path(__file__).parent)}")


if __name__ == "__main__":
    main()
