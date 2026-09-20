"""Score a generated management-question set against the target set. No API calls, no model.

Run:  python eval/score_questions.py output/northwind_analysis.json
      python eval/score_questions.py my_questions.md --target tests/fixtures/target_questions_northwind.md
      python eval/score_questions.py --rubric

The target set is the gold standard the prompt is tuned against: the worked ten-question set from
section 6 of docs/RESEARCH_BOARD_PACKS.md, saved as tests/fixtures/target_questions_northwind.md
with each question's provenance marker kept. This file never writes it and never judges whether a
question is a good one: it measures the five things a string or a structure can settle, and leaves
the judgement to the 1 to 5 rubric below (--rubric), which a person applies by reading.

The five checks, one line each:
1. Decomposition, not explanation: does the question send management to a cut of the data
   ("which cohorts", "how much of") rather than to a story ("why", "what is driving")?
2. Names a metric and a value: does it quote one of metrics.py's metrics and a figure with a unit?
3. Grouped under a theme: does it sit under a "## Theme:" heading, as the target's questions do?
4. Duplicates: does another question in the same set ask the same thing in other words?
5. Theme coverage: which of the target's themes does the set reach, and which does it miss?

Input formats:
- .md   a numbered list, optionally under "## Theme: name" headings, optionally with a
        "[verbatim]", "[verbatim with values]" or "[constructed]" provenance marker per question.
        The target must have a theme and a marker on every question; a generated set need not.
- .json an analysis file written by analyze.py: summary.questions, each one a theme and a question
        since prompt v5 (a plain string in the v4 files, which had no themes).

Exit codes: 0 scored, 1 below --fail-under, 2 a file could not be read.
"""

import argparse
import itertools
import json
import re
import sys
from dataclasses import dataclass
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_DIR))            # the project's own files

from metrics import INPUT_LABELS, METRIC_LABELS  # noqa: E402  (after the path line on purpose)

# Where the gold standard lives. Kept as a constant so the error message can name it.
TARGET_PATH = PROJECT_DIR / "tests" / "fixtures" / "target_questions_northwind.md"

# The three provenance markers section 6 uses, in the order the fixture explains them.
PROVENANCE_MARKERS = ("verbatim", "verbatim with values", "constructed")

# The three answers check 1 can give.
DECOMPOSITION = "decomposition"
EXPLANATION = "explanation"
NEITHER = "neither"


@dataclass(frozen=True)
class Question:
    """One question: its text, the theme it sits under (None if ungrouped), its marker and its number."""
    text: str
    theme: str | None
    provenance: str | None
    number: int


@dataclass(frozen=True)
class QuestionSet:
    """A parsed file: where it came from, its questions, and its themes in the order they appear."""
    source: str
    questions: list
    themes: list


# ---------------------------------------------------------------------------
# Reading a question file
# ---------------------------------------------------------------------------

# "1. [verbatim with values] text" or "1. text": the number, an optional marker, then the question.
QUESTION_LINE = re.compile(r"^\s*\d+[.)]\s*(?:\[([^\]]*)\]\s*)?(.+?)\s*$")
THEME_LINE = re.compile(r"^\s*##\s*Theme:\s*(.+?)\s*$")


def parse_markdown(text, source, strict):
    """Read a numbered question list. strict=True is the target: every question needs a theme and a marker.

    Anything it cannot read with certainty stops the run and names the line, the same rule clean.py
    follows for a workbook: a guessed theme or a guessed marker would quietly change the score.
    """
    questions, themes, theme = [], [], None
    for line in text.splitlines():
        heading = THEME_LINE.match(line)
        if heading:
            theme = heading.group(1)
            if theme not in themes:
                themes.append(theme)
            continue
        found = QUESTION_LINE.match(line)
        if not found:
            continue
        marker, body = found.group(1), found.group(2)
        if marker is not None and marker not in PROVENANCE_MARKERS:
            allowed = ", ".join(PROVENANCE_MARKERS)
            raise ValueError(f"{source}: unknown provenance marker [{marker}] on {body!r}. Allowed: {allowed}")
        if strict and marker is None:
            raise ValueError(f"{source}: no provenance marker on {body!r}. "
                             f"Every target question carries one of: {', '.join(PROVENANCE_MARKERS)}")
        if strict and theme is None:
            raise ValueError(f"{source}: {body!r} sits under no theme. "
                             "Every target question belongs to a '## Theme: name' heading")
        questions.append(Question(body, theme, marker, len(questions) + 1))
    if not questions:
        raise ValueError(f"{source}: no numbered questions found (a question line looks like '1. text')")
    return QuestionSet(source, questions, themes)


def question_and_theme(found, source):
    """One saved question as (text, theme). v5 saves {"theme": ..., "question": ...}; v4 saved a string."""
    if isinstance(found, str):
        return found, None
    if isinstance(found, dict) and isinstance(found.get("question"), str):
        theme = found.get("theme")
        return found["question"], theme if isinstance(theme, str) else None
    raise ValueError(f"{source}: a question in summary.questions is neither a string nor "
                     f"a theme and question: {found!r}")


def parse_analysis_json(text, source):
    """Read analyze.py's summary.questions: since v5 each one carries the theme it sits under."""
    data = json.loads(text)
    found = data.get("summary", {}).get("questions")
    if not isinstance(found, list) or not found:
        raise ValueError(f"{source}: no summary.questions list in the analysis JSON")
    pairs = [question_and_theme(item, source) for item in found]
    themes = list(dict.fromkeys(theme for _, theme in pairs if theme is not None))
    return QuestionSet(source, [Question(text, theme, None, number)
                                for number, (text, theme) in enumerate(pairs, start=1)], themes)


def read_text(path):
    """The file's text, or a stop that names the path (a missing target is the likeliest mistake)."""
    path = Path(path)
    if not path.exists():
        raise ValueError(f"no question file at {path}")
    return path.read_text()


def load_target(path=TARGET_PATH):
    """The gold standard: markdown only, every question themed and marked."""
    path = Path(path)
    if path == TARGET_PATH and not path.exists():
        # The likeliest reason the default is absent: section 6 has not been copied across yet.
        raise ValueError(f"no question file at {path}. It is the worked ten-question set from "
                         "section 6 of docs/RESEARCH_BOARD_PACKS.md, saved with each question's "
                         "provenance marker. Copy it there before scoring against it")
    return parse_markdown(read_text(path), str(path), strict=True)


def load_generated(path):
    """A set to score: an analysis JSON from analyze.py, or the same markdown format, themes optional."""
    path = Path(path)
    text = read_text(path)
    if path.suffix == ".json":
        return parse_analysis_json(text, str(path))
    return parse_markdown(text, str(path), strict=False)


# ---------------------------------------------------------------------------
# Check 1: a decomposition, not an explanation
# ---------------------------------------------------------------------------

# Phrases that ask for the parts of a number outright.
DECOMPOSITION_PHRASES = ("break down", "breakdown", "broken down", "split", "decompose", "disaggregate",
                         "how much of", "how many of", "what share", "what portion", "what proportion",
                         "what percentage of", "which of the")

# The cuts a board asks for. Plurals are allowed by the "s?" in the pattern below.
DIMENSIONS = ("segment", "cohort", "account", "customer", "logo", "product", "region", "geography",
              "channel", "vertical", "tier", "deal", "cost line", "line item", "category", "bucket",
              "business unit")

# The words that turn a dimension into a cut. Without one, "our largest customers" is just a noun.
SPLIT_LEADS = ("by", "which", "across", "per", "within", "among", "between", "concentrated in",
               "attributable to", "made up of", "composed of", "sits in", "sit in")

# A lead, then at most two filler words, then a dimension: "which cost lines", "by customer segment".
# The two-word gap is what keeps "driven by the slowdown you flagged" from counting as a cut.
SPLIT_PATTERN = re.compile(
    r"\b(?:" + "|".join(SPLIT_LEADS) + r")\s+(?:\w+\s+){0,2}(?:" + "|".join(DIMENSIONS) + r")s?\b")

# Phrases that ask for a story instead. A question with both is still a decomposition.
EXPLANATION_PHRASES = ("why", "what is driving", "what's driving", "what drove", "what caused",
                       "what is causing", "what led to", "what explains", "what accounts for", "explain")


def demand_label(text):
    """DECOMPOSITION, EXPLANATION or NEITHER for one question.

    Known limit: the check reads words, not meaning. A long compound question that happens to hold
    a lead and a dimension far apart can read as a decomposition when it is not, so the 1 to 5
    rubric asks a person to confirm the label on anything that matters.
    """
    lower = text.lower()
    if any(phrase in lower for phrase in DECOMPOSITION_PHRASES) or SPLIT_PATTERN.search(lower):
        return DECOMPOSITION
    if any(phrase in lower for phrase in EXPLANATION_PHRASES):
        return EXPLANATION
    return NEITHER


# ---------------------------------------------------------------------------
# Check 2: names a metric and a value
# ---------------------------------------------------------------------------

def alias_forms(label):
    """The words a question may use for one metric, worked out from its label in metrics.py.

    These are not display names (metrics.py owns those, and the project keeps one label set): they
    are only the spellings this file will recognise. "Runway at current burn" gives "runway at
    current burn" and "runway"; "ARR growth QoQ" gives "arr growth qoq" and "arr growth".
    """
    plain = re.sub(r"\s*\([^)]*\)", "", label).strip().lower()   # drop "(annualized)", "($K)"
    forms = {plain}
    for separator in (" at ", " vs "):
        if separator in plain:
            forms.add(plain.split(separator)[0])
    for suffix in (" qoq", " yoy"):
        if plain.endswith(suffix):
            forms.add(plain[: -len(suffix)])
    return forms


def build_alias_table():
    """alias -> the labels it can mean. An alias like "net burn" means more than one, which is fine."""
    table = {}
    for label in list(METRIC_LABELS.values()) + list(INPUT_LABELS.values()):
        for alias in alias_forms(label):
            table.setdefault(alias, set()).add(label)
    # Inputs CLAUDE.md lists that have no display label of their own, so alias_forms cannot reach them.
    for alias, shown in (("new arr", "new_arr"), ("expansion", "expansion_arr"), ("contraction", "contraction_arr"),
                         ("churn", "churned_arr"), ("churned arr", "churned_arr"), ("revenue", "revenue"),
                         ("gross profit", "gross_profit"), ("headcount", "headcount"),
                         ("new customers", "new_customers"), ("sales and marketing spend", "sm_spend"),
                         ("s&m spend", "sm_spend")):
        table.setdefault(alias, set()).add(shown)
    return table


ALIASES = build_alias_table()

# Longest alias first, so "net burn vs budget" is found before the "net burn" inside it. The
# trailing "s?" lets a plural count ("pipelines"), and the word boundaries keep an alias from
# matching inside a longer word ("new arr" must not be found in "new arrangement").
ALIAS_PATTERN = re.compile(
    r"\b(?:" + "|".join(re.escape(a) for a in sorted(ALIASES, key=len, reverse=True)) + r")s?\b")

# A value is a figure with a unit, or a decimal. A quarter label ("Q2 2026") and a plain count
# ("the last 3 quarters") are not values, so they must not let a vague question pass this check.
VALUE_PATTERN = re.compile(r"""
      -?\$\d[\d,]*(?:\.\d+)?[KkMm]?                         # $3,900K
    | -?\d[\d,]*(?:\.\d+)?\s?%                              # 97.1%
    | -?\d[\d,]*(?:\.\d+)?\s?x\b                            # 2.35x
    | -?\d[\d,]*(?:\.\d+)?\s?(?:mo\b|months?\b|weeks?\b)    # 11.0 mo
    | -?\d[\d,]*\.\d+                                       # a bare decimal
""", re.VERBOSE)


def metric_names_in(text):
    """Every metric label (or input column) the question names, sorted, without repeats."""
    found = set()
    for alias in ALIAS_PATTERN.findall(text.lower()):
        # The match may carry the plural "s" the pattern allows, so try the singular too.
        found |= ALIASES.get(alias) or ALIASES.get(alias.rstrip("s"), set())
    return sorted(found)


def values_in(text):
    """Every figure with a unit or a decimal point, in the order written."""
    return [match.group(0).strip() for match in VALUE_PATTERN.finditer(text)]


def names_metric_and_value(text):
    """True when the question quotes both a metric and a figure, so management knows what is asked."""
    return bool(metric_names_in(text)) and bool(values_in(text))


# ---------------------------------------------------------------------------
# Check 4: duplicates
# ---------------------------------------------------------------------------

# Words carried by almost every question, so they say nothing about what it asks.
STOPWORDS = {"a", "an", "and", "are", "as", "at", "be", "but", "by", "can", "do", "does", "for", "from",
             "has", "have", "how", "in", "is", "it", "its", "of", "on", "or", "our", "over", "that",
             "the", "their", "there", "this", "to", "was", "we", "were", "what", "when", "where",
             "which", "who", "why", "will", "with", "you", "your", "did", "each", "into", "any"}


def content_words(text):
    """The words that carry meaning: lowercase, no punctuation, no stopwords."""
    words = re.findall(r"[\w$%.]+", text.lower())
    return {word.strip(".") for word in words if word.strip(".") and word.strip(".") not in STOPWORDS}


def overlap(first, second):
    """How much two questions share, 0 to 1 (shared words divided by all their words)."""
    left, right = content_words(first), content_words(second)
    if not left or not right:
        return 0.0
    return len(left & right) / len(left | right)


# A pair counts as a duplicate when it is about the same metric and half its words agree, or when
# the words agree so far that the metric no longer matters.
SAME_METRIC_OVERLAP = 0.3
ANY_OVERLAP = 0.6


def duplicate_pairs(questions):
    """[(question, question, overlap)] for every pair that asks the same thing in other words."""
    pairs = []
    for first, second in itertools.combinations(questions, 2):
        shared_metrics = set(metric_names_in(first.text)) & set(metric_names_in(second.text))
        score = overlap(first.text, second.text)
        if score >= ANY_OVERLAP or (shared_metrics and score >= SAME_METRIC_OVERLAP):
            pairs.append((first, second, round(score, 2)))
    return pairs


# ---------------------------------------------------------------------------
# Check 5: theme coverage
# ---------------------------------------------------------------------------

# A generated question reaches a target question when they name a metric in common, or when this
# much of their wording agrees. Lower than the duplicate bar on purpose: this asks "same subject",
# not "same question".
COVERS_OVERLAP = 0.3


def covers(question, target_question):
    """True when the generated question is about the same thing as this target question."""
    shared = set(metric_names_in(question.text)) & set(metric_names_in(target_question.text))
    return bool(shared) or overlap(question.text, target_question.text) >= COVERS_OVERLAP


def theme_coverage(target, generated):
    """(themes reached, themes missed), each in the target's own order."""
    covered = [theme for theme in target.themes
               if any(covers(question, wanted)
                      for wanted in target.questions if wanted.theme == theme
                      for question in generated.questions)]
    missed = [theme for theme in target.themes if theme not in covered]
    return covered, missed


# ---------------------------------------------------------------------------
# The scorecard
# ---------------------------------------------------------------------------

def score_set(target, generated):
    """Every check for one generated set, as a dictionary the printout and --json both use."""
    questions = generated.questions
    duplicates = duplicate_pairs(questions)
    in_a_duplicate = {q.number for pair in duplicates for q in pair[:2]}
    covered, missed = theme_coverage(target, generated)

    rows = [{"number": q.number,
             "text": q.text,
             "theme": q.theme,
             "provenance": q.provenance,
             "demand": demand_label(q.text),
             "metrics": metric_names_in(q.text),
             "values": values_in(q.text),
             "metric_and_value": names_metric_and_value(q.text),
             "duplicate_of": sorted({other.number for pair in duplicates for other in pair[:2]
                                     if q.number in (pair[0].number, pair[1].number)} - {q.number})}
            for q in questions]

    total = len(questions)
    counts = {"questions": total,
              "decomposition": sum(row["demand"] == DECOMPOSITION for row in rows),
              "explanation": sum(row["demand"] == EXPLANATION for row in rows),
              "metric_and_value": sum(row["metric_and_value"] for row in rows),
              "themed": sum(row["theme"] is not None for row in rows),
              "duplicate_pairs": len(duplicates)}
    rates = {"decomposition": counts["decomposition"] / total,
             "metric_and_value": counts["metric_and_value"] / total,
             "themed": counts["themed"] / total,
             "unique": (total - len(in_a_duplicate)) / total,
             "theme_coverage": len(covered) / len(target.themes) if target.themes else 0.0}
    return {"target": target.source,
            "generated": generated.source,
            "questions": rows,
            "counts": counts,
            "rates": rates,
            "theme_coverage": {"covered": covered, "missed": missed},
            "duplicates": [(a.number, b.number, score) for a, b, score in duplicates],
            # An unweighted mean of the five checks. A convenience number for watching a prompt
            # change, not a benchmark: no source says these five matter equally.
            "overall": sum(rates.values()) / len(rates)}


def percent(rate):
    """A rate as a percentage string, e.g. 0.5 -> '50%'."""
    return f"{rate * 100:.0f}%"


def format_report(report):
    """The scorecard as text: every question with its verdict, then the totals and the themes."""
    lines = [f"Target:    {report['target']}",
             f"Generated: {report['generated']}",
             ""]
    for row in report["questions"]:
        theme = row["theme"] or "no theme"
        lines.append(f"{row['number']}. [{theme}] {row['text']}")
        metrics = ", ".join(row["metrics"]) or "none"
        values = ", ".join(row["values"]) or "none"
        lines.append(f"   asks for: {row['demand']} | metric: {metrics} | value: {values}")
        if row["duplicate_of"]:
            lines.append(f"   duplicates: {', '.join(str(n) for n in row['duplicate_of'])}")
    counts, rates = report["counts"], report["rates"]
    total = counts["questions"]
    lines += ["",
              f"Decomposition, not explanation: {counts['decomposition']} of {total} "
              f"({percent(rates['decomposition'])})",
              f"Names a metric and a value:     {counts['metric_and_value']} of {total} "
              f"({percent(rates['metric_and_value'])})",
              f"Grouped under a theme:          {counts['themed']} of {total} ({percent(rates['themed'])})",
              f"Not a duplicate:                {percent(rates['unique'])} "
              f"({counts['duplicate_pairs']} duplicate pairs)",
              f"Themes covered:                 {len(report['theme_coverage']['covered'])} of "
              f"{len(report['theme_coverage']['covered']) + len(report['theme_coverage']['missed'])} "
              f"({percent(rates['theme_coverage'])})"]
    covered = ", ".join(report["theme_coverage"]["covered"]) or "none"
    missed = ", ".join(report["theme_coverage"]["missed"]) or "none"
    lines += [f"  covered: {covered}",
              f"  missed:  {missed}",
              "",
              f"Overall (unweighted mean of the five): {percent(report['overall'])}",
              "",
              "These are string and structure checks only. Read the set against the 1 to 5 rubric",
              "(python eval/score_questions.py --rubric) before trusting the number."]
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# The rubric a person applies
# ---------------------------------------------------------------------------

RUBRIC = """Question set rubric, 1 to 5. One score for the set, given by reading it.

What the machine checks cannot see: whether a question is worth a board's time, whether the company
could actually answer it, and whether the answer would change a decision. That is what this rubric
is for. Score the set as a whole, then write one line saying what cost it the missing points.

5. Every question sends management to a specific cut of the data that the pack does not already
   show, names the metric and the figure that prompted it, and would change what the board decides.
   The set spans the themes the quarter raises, with no two questions on the same ground.
4. As above, but one question is softer than the rest: it asks for a cause where a cut was
   available, or it repeats a number the pack already explains.
3. The questions are on the right subjects and quote real figures, but most ask for a story rather
   than a decomposition. Management can answer them in a sentence with no work, which means the
   answers will not be evidence.
2. The questions restate the pack. They name metrics without figures, or figures without a metric,
   or they ask something the slide already answers. One theme the quarter clearly raises is missing.
1. The questions could be asked of any company in any quarter. Nothing ties them to this data, or a
   question invents a number, assumes a cause the data does not show, or can be answered yes or no.

Scoring notes:
- A question that can be answered yes or no caps the set at 2. The board learns nothing from "yes".
- A question that asks management to explain a figure the pack has already explained is a repeat,
  even when the wording differs. Count it as a duplicate.
- Where the data is missing, the right question asks when the number will exist and who owns it.
  Treat that as a decomposition-grade question, not a soft one.
- Note the provenance mix. A set that reaches the target's themes only through questions the target
  marked "constructed" is weaker evidence than one that lands on the sourced questions."""


# ---------------------------------------------------------------------------
# Command line
# ---------------------------------------------------------------------------

def build_parser():
    """The command line: which set to score, which target to score it against, and the floor."""
    parser = argparse.ArgumentParser(description="Score a generated question set against the target set.")
    parser.add_argument("generated", nargs="?",
                        help="the set to score: an analysis JSON from analyze.py, or a markdown list")
    parser.add_argument("--target", default=str(TARGET_PATH), help=f"the gold standard (default: {TARGET_PATH})")
    parser.add_argument("--fail-under", type=float, default=None,
                        help="exit 1 if the overall percentage is below this")
    parser.add_argument("--json", action="store_true", help="print the scorecard as JSON")
    parser.add_argument("--rubric", action="store_true", help="print the 1 to 5 rubric and stop")
    return parser


def main(argv=None):
    """Score one set. 0 scored, 1 below the floor, 2 a file could not be read."""
    args = build_parser().parse_args(argv)
    if args.rubric:
        print(RUBRIC)
        return 0
    if not args.generated:
        print("give a question set to score, or --rubric. See --help.")
        return 2
    try:
        target = load_target(args.target)
        generated = load_generated(args.generated)
    except ValueError as stopped:
        print(f"Stopped: {stopped}")
        return 2
    report = score_set(target, generated)
    print(json.dumps(report, indent=2) if args.json else format_report(report))
    if args.fail_under is not None and report["overall"] * 100 < args.fail_under:
        print(f"\nOverall {percent(report['overall'])} is below the floor of {args.fail_under:.0f}%.")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
