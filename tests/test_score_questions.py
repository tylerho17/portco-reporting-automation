"""Tests for eval/score_questions.py: scoring a generated question set against the target set.

No API calls and no model: every check is a string or structure check, so a test can state the
expected answer by hand. The target set lives in tests/fixtures/target_questions_northwind.md;
these tests build their own small targets in a temporary folder so they never depend on it.

The tests are in the order a reader meets the code: the file format first, then the five checks
(decomposition, metric and value, theme, duplicate, theme coverage), then the whole scorecard
and the command line.
"""

import json

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


def test_the_missing_gold_standard_points_at_section_6():
    # Until section 6 is copied across, the default target does not exist. The stop has to say
    # where it comes from, not just that a file is absent.
    if sq.TARGET_PATH.exists():
        return
    with pytest.raises(ValueError) as stopped:
        sq.load_target()
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
# Check 1: a decomposition, not an explanation
# ---------------------------------------------------------------------------

DECOMPOSITIONS = [
    "Which customer cohorts drove NRR from 108.0% to 97.1%?",
    "How much of the 20.0% burn overrun sits in sales and marketing?",
    "Can you break down net new ARR by segment?",
    "What share of churned ARR is concentrated in the ten largest accounts?",
    "Which cost lines carry the increase in net burn?",
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
