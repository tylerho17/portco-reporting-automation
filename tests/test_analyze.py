"""Unit tests for analyze.py that need no API call: the number check and the payload.

Expected values are worked out by hand (shown in comments), not copied from the code.
Run from the project folder:  pytest
"""

import hashlib
import json
from pathlib import Path

import pytest

from analyze import (MAX_DETAIL_WORDS, MAX_QUESTIONS, MIN_QUESTIONS, PROMPT_VERSION, SYSTEM_PROMPT, THEMES,
                     BoardSummary, Point, Question, build_payload, claim_problems, find_ungrounded_numbers,
                     grouped_questions, numbers_in, save_analysis, slide_fit_problems, validate_summary)
from build_deck import load_analysis
from clean import clean_workbook
from make_data import OUTPUT_PATH as NORTHWIND
from metrics import load_config

# The approved Northwind answer the golden deck, memo and workbook are rebuilt from. Used here as a
# valid v5 answer for this company's numbers, rather than typing a ninth copy of one.
GOLDEN_ANALYSIS = Path(__file__).parent / "golden" / "analysis" / "northwind_analysis.json"


def golden_summary():
    return BoardSummary.model_validate(json.loads(GOLDEN_ANALYSIS.read_text())["summary"])


# ---------------------------------------------------------------------------
# The number check reads minus signs (decision O)
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("text, expected", [
    ("-19.0%", {-19.0}),                      # hyphen as a minus sign
    ("−12.2%", {-12.2}),                      # true minus sign (U+2212)
    ("–3.8%", {-3.8}),                        # en dash used as a minus sign
    ("(-19.0%)", {-19.0}),                    # after a bracket
    ("-$240K", {-240.0}),                     # minus before a dollar sign
    ("fell to -12.2% from -3.8%", {-12.2, -3.8}),
    ("Q4 2025-Q2 2026", {4.0, 2025.0, 2.0, 2026.0}),  # a range between quarters, not negative 2
    ("2025–2026", {2025.0, 2026.0}),          # en dash between numbers is a range
    ("74–75%", {74.0, 75.0}),
    ("1,660 - 2,000", {1660.0, 2000.0}),      # a spaced dash is not attached to the number
    ("$27,470K and 97.1%", {27470.0, 97.1}),  # plain numbers still work
])
def test_numbers_in_reads_signs(text, expected):
    assert numbers_in(text) == expected


def summary_saying(text):
    """A BoardSummary whose only numbers are in `text` (the headline) and in the question below.

    The question quotes -12.2%, which PAYLOAD_TEXT holds, so it adds nothing to the ungrounded list.
    """
    point = Point(title="Title", detail="Detail.")
    question = Question(theme="Definitions and assumptions",
                        question="Which definition gives Rule of 40 at -12.2%?")
    return BoardSummary(headline=text, diagnosis="Retention is the question.", risks=[point] * 3,
                        questions=[question] * MIN_QUESTIONS)


PAYLOAD_TEXT = '{"Net new ARR vs budget": "-19.0%", "Rule of 40": "-12.2%"}'


@pytest.mark.parametrize("headline, ungrounded", [
    ("Net new ARR vs budget was -19.0%.", []),
    ("Net new ARR vs budget was −19.0%.", []),        # typographic minus reads the same
    ("Net new ARR missed budget by 19.0%.", [19.0]),  # the sign was dropped: old code let this through
    ("Rule of 40 is 12.2%.", [12.2]),
])
def test_positive_number_is_not_grounded_by_a_negative_one(headline, ungrounded):
    assert find_ungrounded_numbers(summary_saying(headline), PAYLOAD_TEXT) == ungrounded


# ---------------------------------------------------------------------------
# The payload uses the shared labels and reasons (decisions A, C, I, J)
# ---------------------------------------------------------------------------

@pytest.fixture(scope="module")
def northwind():
    return clean_workbook(NORTHWIND)


def test_payload_says_data_missing_or_no_prior_period(northwind):
    actuals, next_budget = northwind
    trends = build_payload("Northwind", actuals, next_budget, load_config())["trends_by_quarter"]
    assert trends["ARR growth YoY"]["Q3 2024"] == "n/a (no prior period)"  # nothing a year earlier
    assert trends["ARR growth YoY"]["Q1 2025"] == "data missing"           # the blank quarter
    assert trends["ARR growth YoY"]["Q1 2026"] == "data missing"           # a year after the blank
    assert trends["Net burn ($K)"]["Q1 2025"] == "data missing"


def test_payload_flag_names_are_the_metric_labels(northwind):
    actuals, next_budget = northwind
    flags = build_payload("Northwind", actuals, next_budget, load_config())["flags_latest_quarter"]
    assert [flag["flag"] for flag in flags] == [
        "NRR (annualized)", "GRR (annualized)", "Burn multiple", "Net burn vs budget",
        "Runway at current burn", "CAC payback", "Net new ARR vs budget", "Rule of 40",
        "NRR falling while pipeline rising"]


def test_payload_not_meaningful_flag_shows_the_k_figures(northwind):
    # Decision C: Northwind Q2 2026 net burn is 3,900; with a budgeted burn of 0 the % means nothing.
    actuals, next_budget = northwind
    actuals = actuals.copy()
    actuals.loc["Q2 2026", "budget_net_burn"] = 0
    payload = build_payload("Northwind", actuals, next_budget, load_config())
    flag = next(f for f in payload["flags_latest_quarter"] if f["flag"] == "Net burn vs budget")
    assert flag["status"] == "cannot evaluate: not meaningful"
    assert flag["value"] == "n/m: net burn 3,900 vs budget 0 ($K)"
    assert payload["flags_tripped"] == 5 and payload["flags_cannot_evaluate"] == 1
    # Not meaningful is not missing data. (Q1 2025 is still a gap: Northwind's blank quarter.)
    assert payload["data_gaps"]["Net burn vs budget"] == ["Q1 2025"]
    assert "flag: Net burn vs budget" not in payload["data_gaps"]


def test_save_analysis_passed_answer_loads_on_the_deck(northwind, tmp_path):
    # analyze.py and main.py save through save_analysis, so build_deck.py can read what either wrote.
    actuals, next_budget = northwind
    payload = build_payload("Northwind", actuals, next_budget, load_config())
    answer = golden_summary()
    path = save_analysis(tmp_path / "sub" / "northwind_analysis.json", payload, answer, {"passed": True})
    saved = json.loads(path.read_text())
    assert saved == {"summary": answer.model_dump(), "run_info": {"passed": True},
                     "prompt_version": PROMPT_VERSION, "error": None, "payload": payload}
    assert load_analysis(path, payload) == (answer, None)


def test_the_analysis_records_which_prompt_wrote_it(northwind, tmp_path):
    # The manifest and the audit trail need to know which wording produced this text.
    actuals, next_budget = northwind
    payload = build_payload("Northwind", actuals, next_budget, load_config())
    path = save_analysis(tmp_path / "northwind_analysis.json", payload, error="AnthropicError: outage")
    assert json.loads(path.read_text())["prompt_version"] == PROMPT_VERSION  # recorded even on a failure


# The checksum of each prompt wording. A change to SYSTEM_PROMPT without a new PROMPT_VERSION fails
# here: a version number nobody remembers to bump is worse than no version number at all.
PROMPT_CHECKSUMS = {"v4": "1500e06a2299c7717b87f248a60a47404b32ebcefcb70eac4e1c83f22a3fa6d2",
                    "v5": "9c15b357297c1e58ed1a6a558f94a96fdb5b525563599a3ee6e5fad4eb790ed5"}


def test_the_prompt_version_is_bumped_whenever_the_prompt_changes():
    digest = hashlib.sha256(SYSTEM_PROMPT.encode()).hexdigest()
    assert PROMPT_VERSION in PROMPT_CHECKSUMS, (
        f"PROMPT_VERSION is {PROMPT_VERSION!r}: add its checksum {digest!r} to PROMPT_CHECKSUMS")
    assert digest == PROMPT_CHECKSUMS[PROMPT_VERSION], (
        f"SYSTEM_PROMPT changed but PROMPT_VERSION is still {PROMPT_VERSION!r} - bump it and add "
        f"the new checksum {digest!r}")


def test_save_analysis_failure_keeps_the_reason_and_gives_the_placeholder(northwind, tmp_path):
    actuals, next_budget = northwind
    payload = build_payload("Northwind", actuals, next_budget, load_config())
    path = save_analysis(tmp_path / "northwind_analysis.json", payload, error="AnthropicError: outage")
    saved = json.loads(path.read_text())
    assert saved["summary"] is None and saved["run_info"] is None and saved["error"] == "AnthropicError: outage"
    summary, why = load_analysis(path, payload)
    assert summary is None and "failed validation when it was made" in why


# ---------------------------------------------------------------------------
# Text that doesn't fit the slides is a validation problem (review finding 1b)
# ---------------------------------------------------------------------------

SHORT_QUESTION = Question(theme="Definitions and assumptions",
                          question="Which definition gives Rule of 40 at -12.2%?")


# 69 words, no numbers in it, so it passes the word check and adds nothing to the number check.
PLAIN_DIAGNOSIS = ("Retention is the question this quarter. " + "The installed base shrank while new sales held "
                   "up, and spend stayed where the plan put it. ") * 3


def summary_with(headline="Retention is the main question for the board.", detail="Customers stayed.",
                 questions=None, diagnosis=PLAIN_DIAGNOSIS):
    """An answer with no numbers of its own, so only the length and fit checks can fail it."""
    point = Point(title="Steady base", detail=detail)
    return BoardSummary(headline=headline, diagnosis=diagnosis, risks=[point] * 3,
                        questions=questions or [SHORT_QUESTION] * MIN_QUESTIONS)


LONGEST_ALLOWED_DETAIL = "customers " * MAX_DETAIL_WORDS  # the longest detail the word check lets through
# No data gaps, so the risks have slide 3's right-hand column to themselves, and no NRR trend, so
# the divergence test doesn't apply. It grounds the numbers the fixture's question quotes.
NO_GAPS_PAYLOAD = json.dumps({"Rule of 40": "-12.2%"})


def test_a_short_answer_fits_the_slides():
    assert slide_fit_problems(summary_with(), NO_GAPS_PAYLOAD) == []


def test_details_at_the_word_limit_do_not_fit_slide_3():
    # The word limit was set in step 3, before the slides existed: 45 words per detail is more than
    # the risks hold beside the data gaps on slide 3 (Risks and flags) at the 12 pt floor.
    problems = slide_fit_problems(summary_with(detail=LONGEST_ALLOWED_DETAIL.strip()), NO_GAPS_PAYLOAD)
    assert [p for p in problems if "slide 3 (Risks)" in p]


def test_a_companys_data_gaps_count_against_the_risks_on_slide_3():
    # The risks sit under the data gaps in one column, so what the gaps take is not the risks' to use.
    # With every metric missing a quarter, the same answer that fits an empty column no longer does.
    gaps = {"data_gaps": {f"Metric number {number} ($K)": ["Q1 2025"] for number in range(30)}}
    detail = "Net burn ran over budget while retention slipped and the pipeline kept growing. " * 2
    assert slide_fit_problems(summary_with(detail=detail), NO_GAPS_PAYLOAD) == []
    problems = slide_fit_problems(summary_with(detail=detail), json.dumps(gaps))
    assert any("slide 3 (Risks)" in problem for problem in problems)


def test_a_very_long_headline_does_not_fit_slide_4():
    # A real headline is 25 words at most, so the headline box holds two lines with room to shrink:
    # it takes about 60 words to overflow it. The check is here so a runaway headline can't reach a
    # slide either.
    problems = slide_fit_problems(summary_with(headline="retention " * 60), NO_GAPS_PAYLOAD)
    assert any("slide 4 (Headline)" in problem for problem in problems)


def test_a_very_long_diagnosis_does_not_fit_slide_4():
    # The word check allows 90 words; this is 200, which the diagnosis box cannot hold at 12 pt.
    problems = slide_fit_problems(summary_with(diagnosis="retention " * 200), NO_GAPS_PAYLOAD)
    assert any("slide 4 (Diagnosis)" in problem for problem in problems)


def test_long_questions_do_not_fit_slide_4():
    # The questions have slide 4's two columns under the diagnosis. A real question is one sentence;
    # this one is about 200 words, far past anything Claude writes.
    long_question = Question(theme="Liquidity and runway",
                             question="Which cost lines carry the runway at 11.0 mo, and who owns each? " * 15)
    problems = slide_fit_problems(summary_with(questions=[long_question] * MIN_QUESTIONS), NO_GAPS_PAYLOAD)
    assert any("slide 4 (Questions for management" in problem for problem in problems)


def test_text_too_long_for_a_slide_is_a_validation_problem_so_the_retry_handles_it():
    # Before this, an answer this long passed validation and then stopped the deck build, so the
    # company failed with no deck at all (review finding 1).
    problems = validate_summary(summary_with(detail=LONGEST_ALLOWED_DETAIL.strip()), NO_GAPS_PAYLOAD)
    assert problems and all("does not fit" in problem for problem in problems)


# ---------------------------------------------------------------------------
# The v5 shape: a diagnosis, 3 risks, and 8 to 10 questions that do a job
# ---------------------------------------------------------------------------

# A payload holding one metric trend, so the questions below can quote grounded values. NRR moves
# from 108.0% to 97.1%, which is what makes the gross versus net divergence test compulsory.
MOVING_NRR = json.dumps({"trends_by_quarter": {"NRR (annualized)": {"Q1 2026": "108.0%", "Q2 2026": "97.1%"},
                                               "GRR (annualized)": {"Q1 2026": "92.0%", "Q2 2026": "88.1%"}}})
FLAT_NRR = json.dumps({"trends_by_quarter": {"NRR (annualized)": {"Q1 2026": "97.1%", "Q2 2026": "97.1%"},
                                             "GRR (annualized)": {"Q1 2026": "88.1%", "Q2 2026": "88.1%"}}})

DIVERGENCE = Question(theme="Retention decomposition",
                      question="How much of the NRR (annualized) fall to 97.1% sits in GRR (annualized) at 88.1%?")
DEFINITION = Question(theme="Definitions and assumptions",
                      question="Which definition gives GRR (annualized) at 88.1%?")


def questions(count=MIN_QUESTIONS, first=DIVERGENCE):
    """A set that passes every question rule: the divergence test, then padding, then a definition."""
    padding = Question(theme="Liquidity and runway",
                       question="Which cost lines hold NRR (annualized) at 97.1% this quarter?")
    return [first] + [padding] * (count - 2) + [DEFINITION]


@pytest.mark.parametrize("words, caught", [(59, True), (60, False), (90, False), (91, True)])
def test_the_diagnosis_must_be_60_to_90_words(words, caught):
    summary = summary_with(diagnosis="retention " * words)
    problems = validate_summary(summary, MOVING_NRR)
    assert any("diagnosis has" in problem for problem in problems) == caught


@pytest.mark.parametrize("count, caught", [(MIN_QUESTIONS - 1, True), (MIN_QUESTIONS, False),
                                           (MAX_QUESTIONS, False), (MAX_QUESTIONS + 1, True)])
def test_there_must_be_8_to_10_questions(count, caught):
    summary = summary_with(questions=questions(count))
    problems = validate_summary(summary, MOVING_NRR)
    assert any("questions must have" in problem for problem in problems) == caught


def test_a_theme_that_is_not_one_of_the_five_is_caught():
    off_theme = Question(theme="Retention", question="Which accounts hold GRR (annualized) at 88.1%?")
    problems = validate_summary(summary_with(questions=questions(first=off_theme)), FLAT_NRR)
    assert any("has the theme 'Retention'" in problem for problem in problems)
    assert all(theme in SYSTEM_PROMPT for theme in THEMES), "the prompt must spell out the same five themes"


@pytest.mark.parametrize("text, caught", [
    ("Which cost lines carry net burn vs budget at 20.0%?", False),      # a metric and a value
    ("Which cost lines carry the net burn variance?", True),             # a metric, no value
    ("Which parts of the plan carry the 20.0% increase?", True),         # a value, no metric
    ("Which accounts churned in Q2 2026?", True),                        # a quarter is not a value
])
def test_every_question_names_a_metric_and_its_value(text, caught):
    question = Question(theme="Burn and budget variance", question=text)
    problems = validate_summary(summary_with(questions=questions(first=question)), FLAT_NRR)
    assert any("name a metric and quote its value" in problem for problem in problems) == caught


@pytest.mark.parametrize("text, caught", [
    ("Which cost lines carry net burn vs budget at 20.0%?", False),                   # decomposition
    ("How much of net burn vs budget at 20.0% is headcount?", False),                 # decomposition
    ("How does net burn vs budget at 20.0% reconcile to the 2.35x burn multiple?", False),  # reconciliation
    ("How far must net burn vs budget fall from 20.0% before the flag clears?", False),     # distance
    ("What is driving net burn vs budget to 20.0%?", True),                           # an explanation
    ("Why is net burn vs budget at 20.0%?", True),
])
def test_a_question_must_ask_for_a_decomposition_a_reconciliation_or_a_distance(text, caught):
    question = Question(theme="Burn and budget variance", question=text)
    problems = validate_summary(summary_with(questions=questions(first=question)), FLAT_NRR)
    assert any("asks for an explanation" in problem for problem in problems) == caught


def test_a_question_asking_for_both_a_cause_and_a_cut_is_allowed():
    # "Why did it fall, and which cohorts drove it?" still sends management to the data, so it
    # counts, the same way eval/score_questions.py scores it.
    both = Question(theme="Burn and budget variance",
                    question="Why is net burn vs budget at 20.0%, and which cost lines carry it?")
    problems = validate_summary(summary_with(questions=questions(first=both)), FLAT_NRR)
    assert not any("asks for an explanation" in problem for problem in problems)


def test_one_question_must_interrogate_a_definition_a_restatement_or_a_missing_input():
    without = questions()[:-1] + [Question(theme="Pipeline and sales efficiency",
                                           question="Which segments hold NRR (annualized) at 97.1%?")]
    problems = validate_summary(summary_with(questions=without), FLAT_NRR)
    assert any("interrogates a definition" in problem for problem in problems)
    assert not any("interrogates a definition" in problem
                   for problem in validate_summary(summary_with(questions=questions()), FLAT_NRR))


def test_when_nrr_has_moved_the_first_retention_question_is_the_gross_versus_net_test():
    net_only = Question(theme="Retention decomposition",
                        question="Which accounts make up the NRR (annualized) fall to 97.1%?")
    problems = validate_summary(summary_with(questions=questions(first=net_only)), MOVING_NRR)
    assert any("divergence test" in problem for problem in problems)
    assert validate_summary(summary_with(questions=questions()), MOVING_NRR) == []


def test_a_flat_nrr_does_not_demand_the_divergence_test():
    # Nothing moved, so there is no move to decompose: the set may leave retention out entirely.
    net_only = Question(theme="Retention decomposition",
                        question="Which accounts make up the NRR (annualized) at 97.1%?")
    assert validate_summary(summary_with(questions=questions(first=net_only)), FLAT_NRR) == []


def test_a_moved_nrr_with_no_retention_question_at_all_is_caught():
    no_retention = [Question(theme="Liquidity and runway",
                             question="Which cost lines hold NRR (annualized) at 97.1%?")] * 7 + [DEFINITION]
    problems = validate_summary(summary_with(questions=no_retention), MOVING_NRR)
    assert any("Retention decomposition" in problem for problem in problems)


def test_questions_are_grouped_in_theme_order_for_every_output():
    summary = golden_summary()
    groups = grouped_questions(summary.questions)
    assert [theme for theme, _ in groups] == [theme for theme in THEMES if any(
        item.theme == theme for item in summary.questions)]
    assert sum(len(items) for _, items in groups) == len(summary.questions)


# ---------------------------------------------------------------------------
# Direction claims (Task 4). The claims below are Claude's real words from the live run.
# ---------------------------------------------------------------------------

# Alderpeak's real trends, from output/alderpeak_analysis.json's payload.
TRENDS = {
    "NRR (annualized)": {"Q3 2024": "110.0%", "Q4 2024": "110.2%", "Q1 2025": "109.5%", "Q2 2025": "110.0%",
                         "Q3 2025": "109.6%", "Q4 2025": "109.8%", "Q1 2026": "109.7%", "Q2 2026": "109.8%"},
    "ARR growth QoQ": {"Q3 2024": "n/a (no prior period)", "Q4 2024": "12.1%", "Q1 2025": "11.5%",
                       "Q2 2025": "11.2%", "Q3 2025": "10.7%", "Q4 2025": "10.4%", "Q1 2026": "10.0%",
                       "Q2 2026": "9.7%"},
    "ARR growth YoY": {"Q3 2024": "n/a (no prior period)", "Q4 2024": "n/a (no prior period)",
                       "Q1 2025": "n/a (no prior period)", "Q2 2025": "n/a (no prior period)",
                       "Q3 2025": "53.9%", "Q4 2025": "51.5%", "Q1 2026": "49.5%", "Q2 2026": "47.5%"},
    "Ending ARR ($K)": {"Q3 2024": "9,000", "Q4 2024": "10,090", "Q1 2025": "11,250", "Q2 2025": "12,510",
                        "Q3 2025": "13,850", "Q4 2025": "15,290", "Q1 2026": "16,820", "Q2 2026": "18,450"},
}

ALDERPEAK_WIN_1 = ("Ending ARR grew from $9,000K (Q3 2024) to $18,450K (Q2 2026), with ARR growth YoY at 47.5% "
                   "this quarter, well above prior-period levels.")
ALDERPEAK_WIN_3 = ("Burn multiple improved from 0.70x (Q3 2024) to 0.12x (Q2 2026), passing the 2.00x threshold "
                   "by a wide margin and signaling efficient capital use.")
ALDERPEAK_RISK_1 = ("ARR growth YoY fell from 53.9% (Q3 2025) to 47.5% (Q2 2026), and ARR growth QoQ eased from "
                    "12.1% (Q4 2024) to 9.7% (Q2 2026), a consistent slowing trend investors should track.")
ALDERPEAK_RISK_3 = ("NRR (annualized) has drifted down from 110.0% (Q3 2024) to 109.8% (Q2 2026); still well above "
                    "the 100.0% threshold but a persistent gentle decline worth monitoring.")
FERNHOLLOW_WIN_3 = ("The 'NRR falling while pipeline rising' flag passed, meaning this specific compounding risk "
                    "pattern has not been triggered this quarter.")
NORTHWIND_RISK_3 = ("NRR tripped, falling to 97.1% from 108.9% in Q3 2025, while pipeline rose to $12,500K from "
                    "$9,200K over the same period - the flagged divergence suggests new bookings may be masking "
                    "churn.")


# --- Layer 1: the direction word and its own two numbers disagree ---

@pytest.mark.parametrize("claim", [
    "NRR fell from 97.1% (Q3 2025) to 108.9% (Q2 2026).",   # "fell", but the numbers rise
    "Runway at current burn rose to 6.0 mo from 35.8 mo.",  # "rose" with the values the other way round
    "Net burn vs budget declined from 15.0% to 20.0%.",
])
def test_a_direction_word_its_numbers_contradict_is_flagged(claim):
    assert claim_problems(claim, TRENDS)


@pytest.mark.parametrize("claim", [
    ALDERPEAK_RISK_1,   # "fell ... eased", and both pairs really do fall
    ALDERPEAK_WIN_1,    # "grew from $9,000K to $18,450K": true
    "GRR fell from 88.7% to 74.7%, pointing to accelerating churn.",
])
def test_a_claim_whose_numbers_match_its_direction_word_passes(claim):
    assert claim_problems(claim, TRENDS) == []


def test_a_direction_word_after_the_numbers_belongs_to_something_else():
    # check_northwind.py's hand-written summary: "rose" is about pipeline, not about the NRR pair.
    # Reading it as the pair's direction word failed this true claim until the check was narrowed to
    # the words leading up to the numbers.
    assert claim_problems("NRR went from 108.0% to 97.1% while pipeline rose.", TRENDS) == []


def test_a_direction_word_from_an_earlier_claim_does_not_carry_over():
    # The live run's first Alderpeak answer. "Rose" is about Ending ARR; the growth figures that
    # follow are a separate claim. Counting "rose" as their direction cost a retry until the check
    # was narrowed to the words between the previous number and the pair.
    claim = "Ending ARR rose from $9,000K to $18,450K, with ARR growth YoY moderating from 53.9% to 47.5%."
    assert claim_problems(claim, TRENDS) == []


def test_a_sentence_making_two_claims_at_once_is_read_claim_by_claim():
    # Northwind's real risk 3 is correct: NRR is falling and pipeline is rising, in one sentence.
    # Each pair is judged by the words in front of it, so neither claim borrows the other's word.
    assert claim_problems(NORTHWIND_RISK_3, TRENDS) == []


def test_improved_and_deteriorated_are_left_to_the_prompt():
    # "Improved" means a smaller number for burn multiple and CAC payback, and a bigger one for NRR,
    # so the words have no direction of their own to check.
    assert claim_problems(ALDERPEAK_WIN_3, TRENDS) == []


def test_numbers_without_units_are_not_compared():
    # "Q3 2024" and "2,500" would otherwise be read as values to compare.
    assert claim_problems("Pipeline fell from Q3 2024 to Q2 2026.", TRENDS) == []
    assert claim_problems("Net new ARR grew from 12.1% to 9,700.", TRENDS) == []


# --- Layer 2: a trend claimed to be persistent that moves both ways in the data ---

def test_a_persistent_decline_that_moves_both_ways_is_flagged():
    # Alderpeak risk 3 from the live run. True at the two ends, false as a trend: NRR rises at
    # 4 of the 7 steps in between (110.0 -> 110.2, 109.5 -> 110.0, 109.6 -> 109.8, 109.7 -> 109.8).
    problems = claim_problems(ALDERPEAK_RISK_3, TRENDS)
    assert len(problems) == 1
    assert "NRR (annualized)" in problems[0] and "persistent" in problems[0]


def test_a_persistent_claim_that_is_true_passes():
    # Alderpeak risk 1 calls ARR growth QoQ "a consistent slowing trend". It really does fall at
    # every step (12.1, 11.5, 11.2, 10.7, 10.4, 10.0, 9.7), so it must not be flagged.
    assert claim_problems(ALDERPEAK_RISK_1, TRENDS) == []


def test_a_trend_claim_without_persistence_words_is_not_measured_step_by_step():
    # "Fell from A to B" is a true statement about the two ends, whatever happened in between.
    assert claim_problems("NRR (annualized) fell from 110.0% (Q3 2024) to 109.8% (Q2 2026).", TRENDS) == []


def test_a_persistent_claim_whose_series_cannot_be_identified_is_skipped():
    # No quarters named, so there is nothing to look up: the check stays quiet rather than guess.
    assert claim_problems("NRR shows a persistent decline from 110.0% to 109.8%.", TRENDS) == []
    # And with no trends at all (a payload that isn't ours), layer 2 has nothing to read.
    assert claim_problems(ALDERPEAK_RISK_3, {}) == []


# --- What no code check can catch: the prompt rules own these ---

@pytest.mark.parametrize("claim", [
    ALDERPEAK_WIN_1,    # "well above prior-period levels" is false (YoY fell every quarter), but
                        # there is no second number to compare it with.
    FERNHOLLOW_WIN_3,   # the combo flag passing is not good news (pipeline is falling too), but the
                        # claim has no numbers and no direction word at all.
])
def test_claims_only_the_prompt_rules_can_catch(claim):
    assert claim_problems(claim, TRENDS) == []


def test_a_direction_problem_fails_validation():
    payload_text = json.dumps({"trends_by_quarter": TRENDS})
    summary = summary_with(detail="ARR growth YoY fell from 47.5% (Q2 2026) to 53.9% (Q3 2025).")
    assert any("fell" in problem for problem in validate_summary(summary, payload_text))
