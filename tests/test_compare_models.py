"""Unit tests for compare_models.py's free helpers (no API calls): blind letters and score entry.

Run from the project folder:  pytest
"""

import pytest

from compare_models import assign_letters, parse_scores


@pytest.mark.parametrize("run_count", [6, 11, 26])
def test_assign_letters_keeps_every_run(run_count):
    # Bug found in the Task 7 review: letters were "ABCDEFGHIJ", so runs 11+ silently vanished
    # from the blind set and the stats (e.g. RUNS_PER_MODEL = 6 -> 12 paid runs, 10 kept).
    records = [{"run": number} for number in range(run_count)]
    lettered = assign_letters(records)
    assert len(lettered) == run_count
    assert sorted(r["run"] for r in lettered.values()) == list(range(run_count))  # each run once


def test_assign_letters_stops_past_z():
    with pytest.raises(ValueError, match="only 26 letters"):
        assign_letters([{"run": number} for number in range(27)])


def test_parse_scores_rejects_a_letter_scored_twice():
    # Bug found in the Task 7 review: "A=4 A=2" silently kept 2.
    runs = {"A": {"passed": True}, "B": {"passed": True}}
    with pytest.raises(ValueError, match="already scored"):
        parse_scores(["A=4", "B=3", "A=2"], runs)


def test_parse_scores_same_letter_different_case_counts_as_repeat():
    runs = {"A": {"passed": True}}
    with pytest.raises(ValueError, match="already scored"):
        parse_scores(["A=4", "a=2"], runs)
