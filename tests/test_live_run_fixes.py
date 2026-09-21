"""What the first live v5 run showed, one test per finding. No API call: a recording fake client.

The run failed all three companies twice. The slides hold about 70 words of diagnosis and, for a
company with a blank quarter, about 20 words per risk detail, while the prompt asked for 65 to 85
and about 40, and the retry message only said "shorten". A "who owns the missing input" question,
which the prompt asks for, was rejected for having no value, and "what line items make up X" was
rejected as an explanation. Each test below is one of those, worked from the run's own error text.

Run from the project folder:  pytest
"""

import json
import re
from types import SimpleNamespace

import pytest

from analyze import (MIN_DIAGNOSIS_WORDS, SYSTEM_PROMPT, THEMES, AnalysisError, Question, analyze, slide_fit_problems,
                     slide_space_text, validate_summary)
from build_deck import slide_word_limits
from test_analyze import FLAT_NRR, NO_GAPS_PAYLOAD, golden_summary, questions, summary_with

# Twenty metrics missing one quarter: what a company with a blank quarter has on slide 3.
TWENTY_GAPS = json.dumps({"data_gaps": {f"Metric number {number} ($K)": ["Q1 2025"] for number in range(20)}})

DASHES = (chr(0x2014), chr(0x2013))   # an em dash and an en dash, by code point: none may reach Claude


def words_of(text, count):
    """`count` words of realistic text: the approved Northwind answer's own words, repeated as needed."""
    return " ".join((text.split() * 20)[:count])


def gold_prose(count):
    return words_of(golden_summary().diagnosis, count)


def gold_question(count, theme="Liquidity and runway"):
    return Question(theme=theme, question=words_of(golden_summary().questions[0].question, count))


def theme_spread(count, words):
    """`count` questions of `words` words each, spread over the five themes as a real set is."""
    return [gold_question(words, THEMES[number % len(THEMES)]) for number in range(count)]


def problems_about(name, summary, payload):
    return [problem for problem in slide_fit_problems(summary, payload) if name in problem]


# ---------------------------------------------------------------------------
# The room on the slides is measured, told to Claude, and named when an answer is too long
# ---------------------------------------------------------------------------

def test_the_slide_limits_are_measured_for_this_company():
    empty, gapped = slide_word_limits(NO_GAPS_PAYLOAD), slide_word_limits(TWENTY_GAPS)
    assert gapped["risk_detail"] < empty["risk_detail"]     # the data gaps take room the risks would use
    assert gapped["diagnosis"] == empty["diagnosis"]        # slide 4 has no data gaps
    assert MIN_DIAGNOSIS_WORDS < empty["diagnosis"] < 90    # the range the prompt asks for has to exist
    rooms = empty["questions"]
    # More questions never means more room for each, and the tenth costs a lot: a line of the two columns.
    assert rooms[8] >= rooms[9] > rooms[10] > 0


def test_text_at_the_measured_limit_fits_and_text_well_over_it_does_not():
    limits = slide_word_limits(TWENTY_GAPS)
    diagnosis, detail = limits["diagnosis"], limits["risk_detail"]
    assert problems_about("(Diagnosis)", summary_with(diagnosis=gold_prose(diagnosis)), TWENTY_GAPS) == []
    assert problems_about("(Diagnosis)", summary_with(diagnosis=gold_prose(diagnosis + 15)), TWENTY_GAPS)
    assert problems_about("(Risks)", summary_with(detail=gold_prose(detail)), TWENTY_GAPS) == []
    assert problems_about("(Risks)", summary_with(detail=gold_prose(detail + 15)), TWENTY_GAPS)
    for count in (8, 9, 10):
        room = limits["questions"][count]
        assert problems_about("(Questions", summary_with(questions=theme_spread(count, room)), TWENTY_GAPS) == []
        assert problems_about("(Questions", summary_with(questions=theme_spread(count, room + 10)), TWENTY_GAPS)


def test_the_diagnosis_fit_message_says_how_many_words_fit():
    long = gold_prose(85)
    message = problems_about("(Diagnosis)", summary_with(diagnosis=long), NO_GAPS_PAYLOAD)[0]
    found = re.search(r"to (\d+) words or fewer \(it has (\d+)\)", message)
    assert found and int(found[2]) == 85, message
    # Cutting the answer to the number in the message is enough: the message can be acted on.
    cut = " ".join(long.split()[:int(found[1])])
    assert problems_about("(Diagnosis)", summary_with(diagnosis=cut), NO_GAPS_PAYLOAD) == []


def test_the_risk_fit_message_says_how_many_words_each_detail_may_have():
    message = problems_about("(Risks)", summary_with(detail=gold_prose(40)), TWENTY_GAPS)[0]
    found = re.search(r"to (\d+) words or fewer \(yours have 40, 40 and 40\)", message)
    assert found, message
    assert problems_about("(Risks)", summary_with(detail=gold_prose(int(found[1]))), TWENTY_GAPS) == []


def test_the_question_fit_message_gives_the_room_for_that_many_questions():
    message = problems_about("(Questions", summary_with(questions=theme_spread(10, 20)), NO_GAPS_PAYLOAD)[0]
    found = re.search(r"with 10 questions .* (\d+) words each", message)
    assert found, message
    fitted_set = summary_with(questions=theme_spread(10, int(found[1])))
    assert problems_about("(Questions", fitted_set, NO_GAPS_PAYLOAD) == []


class RecordingClient:
    """Answers with `answers` in turn (the last repeats), and keeps every request it was sent."""

    def __init__(self, *answers):
        self.answers, self.requests, self.messages = list(answers), [], self   # analyze.py calls .messages.parse

    def parse(self, **kwargs):
        self.requests.append(kwargs)
        summary = self.answers[min(len(self.requests), len(self.answers)) - 1]
        return SimpleNamespace(content=[], parsed_output=summary, stop_reason="end_turn",
                               usage=SimpleNamespace(input_tokens=10, output_tokens=5))


def test_claude_is_told_this_companys_slide_space_in_the_prompt():
    client = RecordingClient(summary_with(diagnosis="retention " * 200))
    with pytest.raises(AnalysisError):
        analyze(TWENTY_GAPS, client=client)
    limits = slide_word_limits(TWENTY_GAPS)
    assert len(client.requests) == 2
    for request in client.requests:
        assert request["system"].startswith(SYSTEM_PROMPT)
        told = request["system"][len(SYSTEM_PROMPT):]
        assert told == slide_space_text(TWENTY_GAPS)
        assert f"{MIN_DIAGNOSIS_WORDS} to {limits['diagnosis']} words" in told
        assert f"at most {limits['risk_detail']} words" in told
        assert all(str(limits["questions"][count]) in told for count in (8, 9, 10))
        assert not any(dash in told for dash in DASHES)


def test_the_first_message_is_still_the_data_alone():
    # Other code, and the batch tests' fake client, read the company back out of that message's JSON.
    client = RecordingClient(summary_with(diagnosis="retention " * 200))
    with pytest.raises(AnalysisError):
        analyze(TWENTY_GAPS, client=client)
    assert client.requests[0]["messages"][0]["content"] == f"KPI data:\n\n{TWENTY_GAPS}"


def test_the_prompt_says_what_the_first_run_showed_it_did_not():
    assert "missing input has no value to quote" in SYSTEM_PROMPT
    assert "Slide space" in SYSTEM_PROMPT     # where the limits are, and that they are hard


# ---------------------------------------------------------------------------
# A rejected answer is kept, so a failed run can still be read
# ---------------------------------------------------------------------------

def test_a_rejected_answer_is_kept_in_the_run_record_so_it_can_be_read():
    # The first live run threw every answer away: there was no way to read what Claude wrote.
    first, second = summary_with(diagnosis="retention " * 200), summary_with(diagnosis="retention " * 150)
    with pytest.raises(AnalysisError) as raised:
        analyze(NO_GAPS_PAYLOAD, client=RecordingClient(first, second))
    log = raised.value.run_info["attempt_log"]
    assert [entry["answer"] for entry in log] == [first.model_dump(), second.model_dump()]


def test_an_accepted_answer_is_not_copied_into_the_attempt_log():
    _, run_info = analyze(NO_GAPS_PAYLOAD, client=RecordingClient(summary_with()))
    assert "answer" not in run_info["attempt_log"][0]   # it is saved once, as the summary


# ---------------------------------------------------------------------------
# Two question rules that contradicted the prompt, and one cue list that was too narrow
# ---------------------------------------------------------------------------

def question_problems_for(text, theme):
    item = Question(theme=theme, question=text)
    problems = validate_summary(summary_with(questions=questions(first=item)), FLAT_NRR)
    return [problem for problem in problems if "name a metric" in problem or "asks for an explanation" in problem]


@pytest.mark.parametrize("text, caught", [
    ("Who owns delivery of the missing Q1 2025 NRR (annualized), and when will it exist?", False),
    ("For the missing Rule of 40 input in Q2 2026, when will it exist and who owns it?", False),
    ("Who owns the missing Q1 2025 figures, and when will they exist?", True),      # no metric named
    ("Which NRR (annualized) is missing?", True),                                  # no quarter, no who or when
    ("Is the NRR (annualized) for Q1 2025 missing?", True),                        # says missing, asks no who or when
    ("Who is on the board in Q1 2025 for NRR (annualized)?", True),                # nothing is missing
])
def test_a_question_about_a_missing_input_has_no_value_to_quote(text, caught):
    # There is no figure to quote when the input is blank, and the prompt asks for exactly this
    # question: when will it exist and who owns it. The value and decomposition rules cannot apply.
    assert bool(question_problems_for(text, "Definitions and assumptions")) == caught


@pytest.mark.parametrize("text, caught", [
    ("What line items make up Net burn vs budget at 19.6%?", False),                       # from the live run
    ("Which segments make up Pipeline ($K) at 2,500?", False),                             # from the live run
    ("Which segments make up Pipeline at 2,500?", True),        # a bare number: no unit, none in the label
    ("Why is Runway at current burn of 6.0 mo under its 12.0 mo threshold?", True),        # still a "why"
    ("What makes up the gap?", True),                            # the new cue does not stand in for a metric
])
def test_what_the_first_live_run_rejected_that_was_a_fair_question(text, caught):
    assert bool(question_problems_for(text, "Burn and budget variance")) == caught


@pytest.mark.parametrize("text, caught", [
    # Rejected in live run 2, attempt 1 (the exact text, from the kept answers). The retry replaced the first
    # with "How far is runway from its threshold?", which the slide already answers: the check made it worse.
    ("What reduction in net burn would move runway at current burn from 11.0 mo above its 12.0 mo threshold?", False),
    ("What would move runway at current burn from 6.0 mo to its 12.0 mo threshold?", False),
    ("What would bring net burn vs budget at 19.6% back to plan?", False),
    ("Which definition does Burn multiple's ∞ value use given net new ARR is negative?", False),   # ∞ is a value
    ("Which definition does Burn multiple use for an infinite value?", True),     # no figure at all
    ("What would surprise you about runway at current burn of 6.0 mo?", True),   # "would" alone is no demand
])
def test_what_the_second_live_run_rejected_that_was_a_fair_question(text, caught):
    assert bool(question_problems_for(text, "Liquidity and runway")) == caught


def test_the_value_message_says_what_a_value_looks_like():
    item = Question(theme="Burn and budget variance", question="Which cost lines carry the net burn variance?")
    problems = validate_summary(summary_with(questions=questions(first=item)), FLAT_NRR)
    assert any("with its unit" in problem and "$2,500K" in problem for problem in problems)
