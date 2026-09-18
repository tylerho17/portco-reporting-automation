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
from text_fit import preview  # the start of a text, for error messages

DEFAULT_MODEL = "claude-sonnet-5"
# The wording of SYSTEM_PROMPT below. Bump this whenever that text changes, so a saved answer
# says which instructions produced it. A test pins each version's checksum, so the bump isn't
# something anyone has to remember (tests/test_analyze.py). v1-v3 are logged in LEARNINGS.md.
PROMPT_VERSION = "v4"
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
- "runway_if_burn_returns_to_plan" is runway if burn returns to plan. Describe it only in those words. It is not a projection, forecast or improvement.
- Only call a trend rising, falling, improving or deteriorating if every step in the window moves that way. Check the quarter-by-quarter values before you write it. If the figure moves up and down, say so ("moved between A and B") and quote the start and end values. Never call a trend persistent, consistent or steady unless every step moves the same way.
- A flag that passed is a win only if its value is good in itself, not merely inside its threshold. If it passes for a bad reason - for example "NRR falling while pipeline rising" passing because pipeline is falling as well - say so and name the figures behind it, rather than presenting it as good news."""


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


# ---------------------------------------------------------------------------
# Does the text fit the slides? (review finding 1)
# ---------------------------------------------------------------------------

def slide_fit_problems(summary):
    """Problems if the text is too long for slides 1 and 5, even at the 12 pt floor.

    build_deck is imported inside the function, not at the top of the file, because build_deck
    imports this module: importing both ways at load time would fail.
    """
    from build_deck import ai_text_problems
    return ai_text_problems(summary)


# ---------------------------------------------------------------------------
# Does the text agree with the direction of the data? (Task 4)
# ---------------------------------------------------------------------------

# Words saying which way a figure moved. "Improved" and "deteriorated" are deliberately absent: an
# improvement is a smaller burn multiple but a bigger NRR, so they have no direction of their own.
DIRECTION_WORDS = {
    "rising": {"rose", "rise", "rises", "rising", "grew", "grow", "grows", "growing", "increased",
               "increase", "increases", "increasing", "climbed", "climbs", "climbing", "up"},
    "falling": {"fell", "fall", "falls", "falling", "declined", "decline", "declines", "declining",
                "dropped", "drops", "dropping", "decreased", "decrease", "decreases", "decreasing",
                "shrank", "shrunk", "shrinking", "slipped", "slips", "slipping", "eased", "eases",
                "easing", "down"},
}
# Words claiming a whole run of quarters moved one way, not just the two ends.
PERSISTENCE_WORDS = {"persistent", "persistently", "consistent", "consistently", "steady", "steadily",
                     "continued", "continuous", "continuously", "uninterrupted", "straight", "every", "each"}

# A number carrying a unit: 97.1%, 2.35x, 11.0 mo, $27,470K. Without a unit it could be a year or a
# quarter number, so "Q3 2024" is never read as a value to compare.
UNIT_NUMBER_PATTERN = re.compile(r"(?:(?<![0-9A-Za-z])([-−–]))?(\$)?(\d[\d,]*(?:\.\d+)?)\s?(%|x|mo|K)?(?![A-Za-z])")
QUARTER_PATTERN = re.compile(r"Q[1-4] \d{4}")
TOLERANCE = 1e-9  # two formatted values count as the same number within this


def words_in(text):
    """Every word in a text, lower-cased, as a set."""
    return {word.lower() for word in re.findall(r"[A-Za-z]+", text)}


def unit_numbers(text):
    """Every number in `text` that carries a unit, as dicts of value, unit and where it sits."""
    found = []
    for match in UNIT_NUMBER_PATTERN.finditer(text):
        sign, dollar, digits, unit = match.groups()
        if not dollar and not unit:
            continue  # no unit: it could be a year, a quarter number or a count
        value = float(digits.replace(",", ""))
        found.append({"value": -value if sign else value, "unit": (dollar or "") + (unit or ""),
                      "start": match.start(), "end": match.end()})
    return found


def from_to_pairs(sentence):
    """Pairs written as "from A ... to B", or "to B from A", each as (start number, end number).

    Only neighbouring numbers are paired, and both must carry the same unit, so a second claim in
    the same sentence can't be read as the other half of the first one. "to B from A" is read only
    when "from" comes straight after B: with other words in between, which number belongs to which
    claim would be a guess.
    """
    numbers = unit_numbers(sentence)
    pairs = []
    for first, second in zip(numbers, numbers[1:]):
        if first["unit"] != second["unit"]:
            continue
        before = re.findall(r"[A-Za-z]+", sentence[:first["start"]])
        between = re.findall(r"[A-Za-z]+", sentence[first["end"]:second["start"]])
        if before[-1:] == ["from"] and "to" in between:
            pairs.append((first, second))
        elif before[-1:] == ["to"] and between[:1] == ["from"]:
            pairs.append((second, first))
    return pairs


def direction_of(text):
    """"rising", "falling", or None - which also covers text using both words, which is ambiguous."""
    words = words_in(text)
    senses = [sense for sense, sense_words in DIRECTION_WORDS.items() if words & sense_words]
    return senses[0] if len(senses) == 1 else None


def direction_of_pair(sentence, pair, numbers):
    """The direction word governing a pair, or None.

    Only the words introducing the pair count: from the number before it up to the pair itself. A
    sentence usually makes more than one claim, and each claim's word belongs to its own numbers:
    - "NRR went from 108.0% to 97.1% while pipeline rose" - "rose" is about pipeline, after the pair
    - "Ending ARR rose from $9,000K to $18,450K, with ARR growth YoY moderating from 53.9% to 47.5%"
      - "rose" is about ARR, in the clause before the pair
    Reading either as the pair's direction would fail a true claim, so both are left alone.
    """
    first_in_sentence = min(pair, key=lambda number: number["start"])
    earlier_ends = [number["end"] for number in numbers if number["end"] <= first_in_sentence["start"]]
    return direction_of(sentence[max(earlier_ends, default=0):first_in_sentence["start"]])


def quarter_after(text, number, numbers):
    """The quarter named just after a number ("110.0% (Q3 2024)"), before the next number, or None."""
    later = [other["start"] for other in numbers if other["start"] > number["start"]]
    found = QUARTER_PATTERN.search(text, number["end"], min(later, default=len(text)))
    return found.group() if found else None


def series_value(text):
    """One value from a trend as a number, or None when it isn't a single number ("data missing")."""
    if not isinstance(text, str):
        return None
    found = NUMBER_PATTERN.findall(text)
    if len(found) != 1:
        return None
    sign, digits = found[0]
    value = float(digits.replace(",", ""))
    return -value if sign else value


def is_value(text, value):
    """True if a trend's formatted value is this number ("$9,000K" in the text is 9,000 in the data)."""
    found = series_value(text)
    return found is not None and abs(found - value) < TOLERANCE


def matching_series(trends, start, end, start_quarter, end_quarter):
    """The one trend holding both of the claim's values at both of its quarters, or None.

    None when no trend matches, or when several do: the check stays quiet rather than guess.
    """
    matches = [label for label, series in trends.items()
               if isinstance(series, dict)
               and is_value(series.get(start_quarter), start["value"])
               and is_value(series.get(end_quarter), end["value"])]
    return matches[0] if len(matches) == 1 else None


def steps_against(series, start_quarter, end_quarter, direction):
    """(steps moving the other way, steps in total) between two quarters, or None if a value is missing."""
    quarters = list(series)
    window = quarters[quarters.index(start_quarter):quarters.index(end_quarter) + 1]
    values = [series_value(series[quarter]) for quarter in window]
    if any(value is None for value in values) or len(values) < 2:
        return None  # a quarter with no number in between: nothing can be proven either way
    steps = [later - earlier for earlier, later in zip(values, values[1:])]
    wrong_way = [step for step in steps if (step > 0 if direction == "falling" else step < 0)]
    return len(wrong_way), len(steps)


def contradiction_problem(sentence, direction, start, end):
    """Layer 1: the direction word and its own two numbers disagree."""
    goes_up = end["value"] > start["value"]
    if (direction == "rising") == goes_up or end["value"] == start["value"]:
        return None
    return (f"'{preview(sentence)}': says the figure is {direction}, but its numbers go from "
            f"{start['value']:g} to {end['value']:g}")


def persistence_problem(sentence, direction, start, end, trends):
    """Layer 2: a trend called persistent that moves the other way somewhere in between."""
    numbers = unit_numbers(sentence)
    start_quarter = quarter_after(sentence, start, numbers)
    end_quarter = quarter_after(sentence, end, numbers)
    if not start_quarter or not end_quarter:
        return None  # no quarters named, so there is no series to look up
    label = matching_series(trends, start, end, start_quarter, end_quarter)
    if label is None:
        return None
    counted = steps_against(trends[label], start_quarter, end_quarter, direction)
    if counted is None or counted[0] == 0:
        return None
    wrong_way, total = counted
    return (f"'{preview(sentence)}': calls {label} persistently {direction} from {start_quarter} to "
            f"{end_quarter}, but it moves the other way at {wrong_way} of {total} steps")


def claim_problems(text, trends):
    """Problems with the direction claims in one piece of Claude's text. Empty list = nothing to say."""
    problems = []
    for sentence in re.split(r"(?<=[.!?])\s+", text.strip()):
        persistent = bool(words_in(sentence) & PERSISTENCE_WORDS)  # can be said after the numbers
        numbers = unit_numbers(sentence)
        for start, end in from_to_pairs(sentence):
            direction = direction_of_pair(sentence, (start, end), numbers)
            if direction is None:
                continue
            problem = contradiction_problem(sentence, direction, start, end)
            if problem is None and persistent:
                problem = persistence_problem(sentence, direction, start, end, trends)
            if problem:
                problems.append(problem)
    return problems


def direction_problems(summary, payload_text):
    """Direction problems anywhere in Claude's answer, checked against the payload's own trends."""
    try:
        trends = json.loads(payload_text).get("trends_by_quarter") or {}
    except (json.JSONDecodeError, AttributeError):
        trends = {}  # not our payload shape: the step-by-step check has nothing to read
    return [problem for text in summary_texts(summary) for problem in claim_problems(text, trends)]


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
    problems += direction_problems(summary, payload_text)
    problems += slide_fit_problems(summary)
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
              "prompt_version": PROMPT_VERSION,  # which wording wrote this, for the run manifest
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
