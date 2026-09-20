"""Tests for eval/score_questions.py: scoring a generated question set against the target set.

No API calls and no model: every check is a string or structure check, so a test can state the
expected answer by hand. The target set lives in tests/fixtures/target_questions_northwind.md;
these tests build their own small targets in a temporary folder so they never depend on it.

The tests are in the order a reader meets the code: the file format first, then the five checks
(decomposition, metric and value, theme, duplicate, theme coverage), then the whole scorecard
and the command line.
"""

import json
import re

import pytest

import score_questions as sq

# A small target set: two themes, three questions, one of each provenance marker.
TARGET_MD = """# Target question set: Example

Source: docs/RESEARCH_BOARD_PACKS.md, section 6.

## Theme: Cash and runway

1. [verbatim] Which cost lines carry the increase in net burn?
2. [verbatim with values] What share of the 11.0 mo runway depends on the budgeted burn?

## Theme: Retention

3. [constructed] Which customer cohorts drove NRR from 108.0% to 97.1%?
"""


def write(path, text):
    """Write text to a path and give the path back, so a test can make a file in one line."""
    path.write_text(text)
    return path


@pytest.fixture
def target(tmp_path):
    """The small target set above, parsed."""
    return sq.load_target(write(tmp_path / "target.md", TARGET_MD))


# ---------------------------------------------------------------------------
# The file format
# ---------------------------------------------------------------------------

def test_the_target_parses_into_themes_questions_and_markers(target):
    assert [q.theme for q in target.questions] == ["Cash and runway", "Cash and runway", "Retention"]
    assert [q.provenance for q in target.questions] == ["verbatim", "verbatim with values", "constructed"]
    assert target.themes == ["Cash and runway", "Retention"]
    assert target.questions[0].text.startswith("Which cost lines")


def test_an_unknown_provenance_marker_stops_and_names_the_line(tmp_path):
    broken = TARGET_MD.replace("[constructed]", "[made up]")
    with pytest.raises(ValueError) as stopped:
        sq.load_target(write(tmp_path / "target.md", broken))
    assert "made up" in str(stopped.value)
    assert "verbatim with values" in str(stopped.value)   # the message lists what is allowed


def test_a_target_question_with_no_marker_stops(tmp_path):
    broken = TARGET_MD.replace("1. [verbatim] Which", "1. Which")
    with pytest.raises(ValueError) as stopped:
        sq.load_target(write(tmp_path / "target.md", broken))
    assert "provenance" in str(stopped.value).lower()


def test_a_target_question_outside_a_theme_stops(tmp_path):
    broken = TARGET_MD.replace("## Theme: Cash and runway\n\n", "")
    with pytest.raises(ValueError) as stopped:
        sq.load_target(write(tmp_path / "target.md", broken))
    assert "theme" in str(stopped.value).lower()


def test_a_missing_target_file_stops_and_names_it(tmp_path):
    with pytest.raises(ValueError) as stopped:
        sq.load_target(tmp_path / "not_here.md")
    # The message has to say the file is absent, not just fail somewhere downstream on empty text:
    # a scorecard built from a file that isn't there would look like a real result.
    assert "no question file at" in str(stopped.value)
    assert "not_here.md" in str(stopped.value)


def test_the_missing_gold_standard_points_at_section_6(tmp_path, monkeypatch):
    # The default target is missing only if someone deletes it, so this test points the default at
    # an empty spot. The stop has to say where the file comes from, not just that it is absent.
    # (It used to return early when the real file existed, which made it pass without checking.)
    absent = tmp_path / "target_questions_northwind.md"
    monkeypatch.setattr(sq, "TARGET_PATH", absent)
    with pytest.raises(ValueError) as stopped:
        sq.load_target(absent)
    assert "section 6" in str(stopped.value)
    assert "RESEARCH_BOARD_PACKS.md" in str(stopped.value)


def test_a_missing_generated_file_stops_too(tmp_path):
    with pytest.raises(ValueError) as stopped:
        sq.load_generated(tmp_path / "also_gone.json")
    assert "no question file at" in str(stopped.value)


def test_a_generated_set_may_be_a_flat_markdown_list_with_no_theme(tmp_path):
    flat = "1. Why did NRR fall?\n2. Which segments drove the fall?\n"
    generated = sq.load_generated(write(tmp_path / "gen.md", flat))
    assert [q.theme for q in generated.questions] == [None, None]
    assert [q.provenance for q in generated.questions] == [None, None]


def test_a_generated_set_may_be_an_analysis_json_from_before_v5(tmp_path):
    # v4 saved summary.questions as plain strings, with no themes. Those files still score.
    path = tmp_path / "example_analysis.json"
    path.write_text(json.dumps({"summary": {"questions": ["Why did NRR fall?", "Which cohorts churned?"]}}))
    generated = sq.load_generated(path)
    assert [q.text for q in generated.questions] == ["Why did NRR fall?", "Which cohorts churned?"]
    assert all(q.theme is None for q in generated.questions)


def test_an_analysis_json_carries_the_theme_each_question_sits_under(tmp_path):
    # Since v5 analyze.py groups its questions, so the "grouped under a theme" check can see them.
    path = tmp_path / "example_analysis.json"
    path.write_text(json.dumps({"summary": {"questions": [
        {"theme": "Retention decomposition", "question": "Which cohorts churned?"},
        {"theme": "Liquidity and runway", "question": "How far is runway from 12.0 mo?"},
        {"theme": "Retention decomposition", "question": "How much of NRR at 97.1% is expansion?"}]}}))
    generated = sq.load_generated(path)
    assert [q.theme for q in generated.questions] == [
        "Retention decomposition", "Liquidity and runway", "Retention decomposition"]
    assert generated.themes == ["Retention decomposition", "Liquidity and runway"]   # in the order first seen


def test_a_question_that_is_neither_a_string_nor_a_theme_and_question_stops(tmp_path):
    path = tmp_path / "example_analysis.json"
    path.write_text(json.dumps({"summary": {"questions": [{"text": "Which cohorts churned?"}]}}))
    with pytest.raises(ValueError) as stopped:
        sq.load_generated(path)
    assert "neither a string nor" in str(stopped.value)


def test_an_analysis_json_with_no_questions_stops(tmp_path):
    path = tmp_path / "example_analysis.json"
    path.write_text(json.dumps({"summary": {"headline": "no questions here"}}))
    with pytest.raises(ValueError) as stopped:
        sq.load_generated(path)
    assert "questions" in str(stopped.value)


# ---------------------------------------------------------------------------
# The real gold standard: tests/fixtures/target_questions_northwind.md
#
# Every expected value below was read off the table in section 6 of docs/RESEARCH_BOARD_PACKS.md
# by hand: the theme of each row, and its provenance mark ([V] verbatim, [V+] verbatim with the
# company's values, [C] constructed). The fixture is the standard the prompt is tuned against, so
# a stray edit to it must fail a test rather than quietly move the goalposts.
# ---------------------------------------------------------------------------

SECTION_6_THEMES = ["Retention decomposition", "Burn variance and efficiency", "Liquidity and downside",
                    "Pipeline versus bookings", "Definitions and assumptions"]

# Rows 1 to 10 of section 6's table. A question here takes exactly one marker, and two rows mix:
# row 3 is [V] with a [C] scoping clause, so it keeps the table's own [V]; row 10 is a [C] clause
# plus a [V] clause with no lead named, so it takes the weaker marker, constructed. The source line
# under each of the two says which part is which (a test below checks it does).
SECTION_6_MARKERS = ["verbatim with values", "constructed", "verbatim", "constructed", "verbatim with values",
                     "verbatim with values", "verbatim", "verbatim with values", "verbatim", "constructed"]


@pytest.fixture(scope="module")
def gold():
    """The real target set, parsed with the strict rules the scorer applies to it."""
    return sq.load_target(sq.TARGET_PATH)


def test_the_gold_standard_has_the_ten_questions_and_five_themes_of_section_6(gold):
    assert len(gold.questions) == 10
    assert gold.themes == SECTION_6_THEMES
    assert [sum(q.theme == theme for q in gold.questions) for theme in SECTION_6_THEMES] == [3, 2, 2, 2, 1]


def test_the_gold_standard_keeps_every_provenance_marker_of_section_6(gold):
    assert [q.provenance for q in gold.questions] == SECTION_6_MARKERS


@pytest.mark.parametrize("number, opening", [
    (1, "NRR fell from 108% to 97% over three quarters. Rebuild the monthly waterfall for the last eight quarters"),
    (4, "Burn is 20% over budget. Split the burn multiple"),
    (6, "At 11 months of runway, how much cushion remains after downside sensitivity"),
    (8, "Pipeline is up while bookings are flat. What is conversion at each stage entered"),
    (10, "Is the 97% computed on exactly the same basis as the 108%"),
])
def test_the_gold_standard_questions_open_with_section_6s_own_words(gold, number, opening):
    assert gold.questions[number - 1].text.startswith(opening)


def test_the_gold_standard_keeps_a_short_verbatim_question_word_for_word(gold):
    assert gold.questions[6].text == "What is the worst outcome here, and how likely is that outcome?"


EM_DASH = chr(0x2014)    # section 6 uses them; written by number so this file holds none


def section_6_words(text):
    """The words of one section 6 table cell or fixture question: no quotes, no em dashes, lowercase."""
    plain = text.replace(chr(0x201C), "").replace(chr(0x201D), "").replace('"', "").replace(EM_DASH, " ")
    return re.findall(r"[\w$%.'–-]+", plain.lower())


def test_the_gold_standard_is_word_for_word_section_6_apart_from_one_stated_edit(gold):
    # Reads the question column of the table in docs/RESEARCH_BOARD_PACKS.md and compares it with
    # the fixture. The only difference allowed is the one the fixture's header owns up to: row 3's
    # trailing "applied to" note written as "Apply this to". Anything else means the fixture drifted
    # from its source, or the source changed and the fixture was not updated.
    doc = (sq.PROJECT_DIR / "docs" / "RESEARCH_BOARD_PACKS.md").read_text()
    rows = {int(m.group(1)): m.group(2) for m in
            re.finditer(r"^\| (\d+) \| [^|]+\| (.+?) \| .+ \|$", doc, flags=re.MULTILINE) if int(m.group(1)) <= 10}
    assert sorted(rows) == list(range(1, 11))
    rows[3] = rows[3].replace(EM_DASH + " applied to", "Apply this to")    # the one stated edit
    for question in gold.questions:
        assert section_6_words(question.text) == section_6_words(rows[question.number]), \
            f"question {question.number} differs from section 6"


def test_every_gold_standard_question_is_followed_by_the_source_it_was_taken_from():
    # The marker says how a question relates to its source; the source line says which source.
    # The scorer ignores these lines, so only a test can notice one going missing.
    lines = sq.TARGET_PATH.read_text().splitlines()
    for index, line in enumerate(lines):
        if sq.QUESTION_LINE.match(line):
            following = next(later for later in lines[index + 1:] if later.strip())
            assert following.strip().startswith("Source:"), f"no source line under: {line[:60]}"


def test_the_two_mixed_questions_say_which_part_is_which():
    # Rows 3 and 10 of section 6 mix a quoted part and a constructed part. One marker cannot say
    # that, so the source line under each must, or the fixture claims more than its source does.
    text = sq.TARGET_PATH.read_text()
    third = text.split("3. [verbatim]")[1].split("4. [constructed]")[0]
    assert "scoping clause" in third and "constructed" in third
    tenth = text.split("10. [constructed]")[1]
    assert "second clause" in tenth and "verbatim" in tenth


def test_the_gold_standard_says_where_it_came_from_and_what_it_leaves_out():
    text = sq.TARGET_PATH.read_text()
    assert "docs/RESEARCH_BOARD_PACKS.md" in text
    assert "section 6" in text
    assert "reserve" in text.lower()      # the two reserve questions are named as left out, not silently dropped


def test_the_gold_standard_reaches_every_one_of_its_own_themes(gold):
    # Score the target's questions, flattened to a plain list, against the target: a set that is the
    # standard itself must miss no theme. Its other rates are deliberately not asserted here:
    # section 6's own rule says every question names a metric and a value, and some of its ten do
    # not, so demanding 100% would demand something the source does not meet.
    flat = "\n".join(f"{q.number}. {q.text}" for q in gold.questions)
    report = sq.score_set(gold, sq.parse_markdown(flat, "itself", strict=False))
    assert report["theme_coverage"]["missed"] == []
    assert report["counts"]["questions"] == 10


# ---------------------------------------------------------------------------
# Check 1: a decomposition, not an explanation
# ---------------------------------------------------------------------------

DECOMPOSITIONS = [
    "Which customer cohorts drove NRR from 108.0% to 97.1%?",
    "How much of the 20.0% burn overrun sits in sales and marketing?",
    "Can you break down net new ARR by segment?",
    "What share of churned ARR is concentrated in the ten largest accounts?",
    "Which cost lines carry the increase in net burn?",
    # Live run 2: the plural of "category" is not "categorys", and three plain cuts were labelled "neither".
    "Which cost categories make up the net burn vs budget variance of 20.0%?",
    "Which pipeline segments make up the growth to $12,500K?",
]

EXPLANATIONS = [
    "What is driving the decline in NRR from 108.0% to 97.1%?",
    "Why did runway fall to 11.0 mo?",
    "What caused the burn multiple to reach 2.35x?",
    "What accounts for the Rule of 40 at -12.2%?",
]


@pytest.mark.parametrize("text", DECOMPOSITIONS)
def test_a_question_that_asks_for_the_parts_is_a_decomposition(text):
    assert sq.demand_label(text) == sq.DECOMPOSITION


@pytest.mark.parametrize("text", EXPLANATIONS)
def test_a_question_that_asks_for_a_cause_is_an_explanation(text):
    assert sq.demand_label(text) == sq.EXPLANATION


def test_a_question_that_asks_for_neither_is_labelled_neither():
    assert sq.demand_label("What cost actions are planned for next quarter?") == sq.NEITHER


def test_a_decomposition_wins_when_a_question_asks_for_both():
    # "Why ... and which segments" still sends management to the data, so it counts as the better kind.
    assert sq.demand_label("Why did NRR fall, and which segments drove it?") == sq.DECOMPOSITION


def test_driven_by_a_team_is_not_a_decomposition():
    # "by" only counts when a dimension to cut by follows it closely.
    assert sq.demand_label("Was the miss driven by the slowdown you flagged last quarter?") == sq.NEITHER


def test_a_dimension_far_from_the_split_word_is_not_a_decomposition():
    # "by" sits eight words from "product" here and does not cut anything by it. Widening the gap
    # the pattern allows would read this as a decomposition, which is why the gap is two words.
    question = "Is the burn increase driven by the hiring plan we approved for the new product line?"
    assert sq.demand_label(question) == sq.NEITHER


# ---------------------------------------------------------------------------
# Check 2: names a metric and a value
# ---------------------------------------------------------------------------

def test_metric_names_come_from_the_metrics_py_labels():
    assert "NRR (annualized)" in sq.metric_names_in("Why did NRR fall to 97.1%?")
    assert "Runway at current burn" in sq.metric_names_in("Runway is 11.0 mo at current burn.")
    assert "Burn multiple" in sq.metric_names_in("The burn multiple reached 2.35x.")
    assert sq.metric_names_in("What are the plans for next quarter?") == []


def test_a_value_needs_a_unit_or_a_decimal_point():
    assert sq.values_in("NRR fell to 97.1% and runway to 11.0 mo") == ["97.1%", "11.0 mo"]
    assert sq.values_in("net burn of $3,900K against a 2.35x burn multiple") == ["$3,900K", "2.35x"]
    # A quarter label and a plain count are not metric values, so they must not pass the check.
    assert sq.values_in("Over the last 3 quarters, and in Q2 2026") == []


def test_a_metric_alias_is_not_found_inside_a_longer_word():
    # "new arr" sits inside "new arrangement". Without the word boundary this question would
    # count as naming a metric, and a vague question would pass the check.
    assert sq.metric_names_in("Did the new arrangement with the reseller close?") == []


def test_a_plural_names_the_same_metric():
    assert "Pipeline ($K)" in sq.metric_names_in("How do the two pipelines compare?")


def test_naming_a_metric_and_a_value_needs_both():
    assert sq.names_metric_and_value("Why did NRR fall to 97.1%?")
    assert not sq.names_metric_and_value("Why did NRR fall?")              # metric, no value
    assert not sq.names_metric_and_value("Why did it fall to 97.1%?")      # value, no metric


# ---------------------------------------------------------------------------
# Check 3: grouped under a theme
# ---------------------------------------------------------------------------

def test_every_target_question_is_themed_and_a_flat_list_is_not(target, tmp_path):
    assert all(q.theme for q in target.questions)
    flat = sq.load_generated(write(tmp_path / "gen.md", "1. Why did NRR fall?\n"))
    assert not any(q.theme for q in flat.questions)


# ---------------------------------------------------------------------------
# Check 4: duplicates
# ---------------------------------------------------------------------------

def test_two_questions_about_the_same_metric_in_the_same_words_are_duplicates():
    pairs = sq.duplicate_pairs([
        sq.Question("What is driving the decline in NRR from 108.0% to 97.1%?", None, None, 1),
        sq.Question("What is driving NRR down from 108.0% to 97.1%?", None, None, 2),
    ])
    assert [(a.number, b.number) for a, b, _ in pairs] == [(1, 2)]


def test_two_questions_about_one_metric_are_duplicates_even_in_different_words():
    # Only a third of the wording agrees, so the words alone would not settle it. Both ask about
    # the same fall in the same metric, so asking both wastes a slot: the metric rule catches it.
    pairs = sq.duplicate_pairs([
        sq.Question("What is driving the decline in NRR from 108.0% to 97.1%?", None, None, 1),
        sq.Question("Which accounts explain the NRR decline to 97.1%?", None, None, 2),
    ])
    assert [(a.number, b.number) for a, b, _ in pairs] == [(1, 2)]
    assert pairs[0][2] < sq.ANY_OVERLAP    # the wording rule on its own would have missed it


def test_two_questions_about_different_things_are_not_duplicates():
    assert sq.duplicate_pairs([
        sq.Question("Which cost lines carry the increase in net burn?", None, None, 1),
        sq.Question("Which customer cohorts drove NRR from 108.0% to 97.1%?", None, None, 2),
    ]) == []


# ---------------------------------------------------------------------------
# Check 5: theme coverage
# ---------------------------------------------------------------------------

def test_a_set_that_misses_a_theme_has_it_named(target, tmp_path):
    generated = sq.load_generated(write(tmp_path / "gen.md", "1. Which cost lines carry the rise in net burn?\n"))
    covered, missed = sq.theme_coverage(target, generated)
    assert covered == ["Cash and runway"]
    assert missed == ["Retention"]


@pytest.mark.xfail(strict=True, reason="known limit: a shared metric name counts as covering a theme, and "
                                       "the real target's themes share metrics (see covers())")
def test_a_theme_is_missed_when_only_a_neighbouring_theme_shares_its_metric(gold):
    # The real target: "Definitions and assumptions" is target question 10, which names GRR. A set
    # with one retention question that names GRR reaches "Retention decomposition" and nothing
    # about definitions, so the right answer is that "Definitions and assumptions" is missed.
    # Today the shared GRR counts as covering it. strict=True: when the check is made stricter,
    # this starts passing, the suite says so, and the marker comes off.
    generated = sq.QuestionSet("one", [sq.Question("Which accounts explain GRR at 88.1%?", None, None, 1)], [])
    _, missed = sq.theme_coverage(gold, generated)
    assert "Definitions and assumptions" in missed


def test_a_set_that_covers_every_theme_misses_none(target, tmp_path):
    both = "1. Which cost lines carry the rise in net burn?\n2. Which cohorts drove NRR to 97.1%?\n"
    generated = sq.load_generated(write(tmp_path / "gen.md", both))
    covered, missed = sq.theme_coverage(target, generated)
    assert covered == ["Cash and runway", "Retention"]
    assert missed == []


# ---------------------------------------------------------------------------
# The whole scorecard
# ---------------------------------------------------------------------------

def test_the_scorecard_counts_every_check(target, tmp_path):
    generated = sq.load_generated(write(tmp_path / "gen.md", (
        "1. Which customer cohorts drove NRR from 108.0% to 97.1%?\n"
        "2. Why did runway fall to 11.0 mo?\n"
    )))
    report = sq.score_set(target, generated)
    assert report["counts"]["questions"] == 2
    assert report["counts"]["decomposition"] == 1        # only the first asks for the parts
    assert report["counts"]["metric_and_value"] == 2     # both name a metric and a value
    assert report["counts"]["themed"] == 0               # a flat list has no themes
    assert report["counts"]["duplicate_pairs"] == 0
    assert report["theme_coverage"]["missed"] == []


def test_the_scorecard_shows_what_the_target_itself_scores(target, tmp_path):
    # In the small target: all three questions ask for the parts (1.0), and two of three quote a
    # metric and a value (question 1 names net burn but no figure). A set cannot be asked to beat
    # the standard on a check the standard does not pass, so the printout shows this next to it.
    generated = sq.load_generated(write(tmp_path / "gen.md", "1. Which cohorts drove NRR to 97.1%?\n"))
    report = sq.score_set(target, generated)
    assert report["target_reference"]["decomposition"] == 1.0
    assert report["target_reference"]["metric_and_value"] == pytest.approx(2 / 3)
    text = sq.format_report(report)
    assert "the target itself: 100%" in text
    assert "the target itself: 67%" in text


def test_the_report_prints_every_question_with_its_verdict(target, tmp_path):
    generated = sq.load_generated(write(tmp_path / "gen.md", "1. Why did runway fall to 11.0 mo?\n"))
    text = sq.format_report(sq.score_set(target, generated))
    assert "Why did runway fall" in text
    assert sq.EXPLANATION in text
    assert "Retention" in text            # the missed theme is named
    assert "no theme" in text


def test_the_overall_number_is_the_mean_of_the_five_checks(target, tmp_path):
    generated = sq.load_generated(write(tmp_path / "gen.md", "1. Which cohorts drove NRR to 97.1%?\n"))
    report = sq.score_set(target, generated)
    rates = report["rates"]
    assert rates["decomposition"] == 1.0
    assert rates["metric_and_value"] == 1.0
    assert rates["themed"] == 0.0
    assert rates["unique"] == 1.0                         # nothing to duplicate
    assert rates["theme_coverage"] == 0.5                 # one of the two themes
    assert report["overall"] == pytest.approx(sum(rates.values()) / len(rates))


# ---------------------------------------------------------------------------
# The rubric and the command line
# ---------------------------------------------------------------------------

def test_the_rubric_has_a_line_for_every_score_from_one_to_five():
    for score in range(1, 6):
        assert f"{score}." in sq.RUBRIC, f"the rubric has no {score}"
    assert "cannot see" in sq.RUBRIC.lower()   # it says what the machine checks miss


def test_the_command_line_scores_a_set_and_exits_zero(target, tmp_path, capsys):
    target_path = write(tmp_path / "target.md", TARGET_MD)
    gen_path = write(tmp_path / "gen.md", "1. Which cohorts drove NRR to 97.1%?\n")
    code = sq.main([str(gen_path), "--target", str(target_path)])
    assert code == 0
    assert "Which cohorts" in capsys.readouterr().out


def test_the_command_line_fails_under_a_floor(tmp_path, capsys):
    target_path = write(tmp_path / "target.md", TARGET_MD)
    gen_path = write(tmp_path / "gen.md", "1. What are the plans for next quarter?\n")
    code = sq.main([str(gen_path), "--target", str(target_path), "--fail-under", "90"])
    assert code == 1
    assert "below" in capsys.readouterr().out


def test_the_command_line_says_which_file_is_missing(tmp_path, capsys):
    gen_path = write(tmp_path / "gen.md", "1. Why did NRR fall?\n")
    code = sq.main([str(gen_path), "--target", str(tmp_path / "gone.md")])
    assert code == 2
    assert "gone.md" in capsys.readouterr().out


def test_the_command_line_prints_the_rubric_on_its_own(capsys):
    assert sq.main(["--rubric"]) == 0
    assert "1." in capsys.readouterr().out


# ---------------------------------------------------------------------------
# A run that failed validation has no summary (Night two, Task 3, live run 1)
# ---------------------------------------------------------------------------

REJECTED_QUESTIONS = [{"theme": "Retention decomposition", "question": "Which cohorts drove NRR to 97.1%?"},
                      {"theme": "Liquidity and runway", "question": "How far is runway from 12.0 mo?"}]


def failed_analysis(tmp_path, attempts):
    """What analyze.py saves when both answers fail validation: no summary, an error, and the attempts."""
    path = tmp_path / "failed_analysis.json"
    path.write_text(json.dumps({"summary": None, "error": "Claude's answer failed validation twice.\n- too long",
                                "run_info": {"passed": False, "attempt_log": attempts}}))
    return path


def test_an_analysis_that_failed_validation_stops_in_plain_words_not_a_traceback(tmp_path):
    # summary is null, not missing: data.get("summary", {}) returned None and .get on it crashed.
    path = failed_analysis(tmp_path, [{"problems": ["too long"], "answer": {"questions": REJECTED_QUESTIONS}}])
    with pytest.raises(ValueError) as stopped:
        sq.load_generated(path)
    assert "no validated answer" in str(stopped.value) and "--rejected" in str(stopped.value)


def test_the_last_rejected_answer_can_be_scored_on_request(tmp_path):
    first = {"problems": ["a"], "answer": {"questions": [{"theme": "Liquidity and runway", "question": "Old?"}]}}
    last = {"problems": ["b"], "answer": {"questions": REJECTED_QUESTIONS}}
    generated = sq.load_generated(failed_analysis(tmp_path, [first, last]), rejected=True)
    assert [q.text for q in generated.questions] == [item["question"] for item in REJECTED_QUESTIONS]
    assert generated.themes == ["Retention decomposition", "Liquidity and runway"]


def test_asking_for_the_rejected_answer_when_none_was_kept_says_so(tmp_path):
    # Analyses saved before the rejected answer was kept have nothing to score.
    path = failed_analysis(tmp_path, [{"problems": ["too long"]}, {"problems": ["too long"]}])
    with pytest.raises(ValueError) as stopped:
        sq.load_generated(path, rejected=True)
    assert "no rejected answer was kept" in str(stopped.value)


def test_a_validated_analysis_is_scored_as_it_always_was_even_with_rejected(tmp_path):
    path = tmp_path / "ok_analysis.json"
    path.write_text(json.dumps({"summary": {"questions": REJECTED_QUESTIONS}, "run_info": {"attempt_log": [
        {"problems": ["x"], "answer": {"questions": [{"theme": "Liquidity and runway", "question": "Old?"}]}}]}}))
    assert [q.text for q in sq.load_generated(path, rejected=True).questions] == [i["question"] for i in REJECTED_QUESTIONS]


def test_the_command_line_stops_with_2_on_a_failed_analysis_and_names_the_way_out(tmp_path, capsys):
    target_path = write(tmp_path / "target.md", TARGET_MD)
    path = failed_analysis(tmp_path, [{"problems": ["x"], "answer": {"questions": REJECTED_QUESTIONS}}])
    assert sq.main([str(path), "--target", str(target_path)]) == 2
    assert "--rejected" in capsys.readouterr().out


def test_the_command_line_says_when_it_is_scoring_a_rejected_answer(tmp_path, capsys):
    target_path = write(tmp_path / "target.md", TARGET_MD)
    path = failed_analysis(tmp_path, [{"problems": ["x"], "answer": {"questions": REJECTED_QUESTIONS}}])
    assert sq.main([str(path), "--target", str(target_path), "--rejected"]) == 0
    assert "REJECTED" in capsys.readouterr().out
