"""Unit tests for main.py's batch outputs: the summary CSV (decision M) and the quarter warning (decision N).

Results are written out by hand in the shape run_company / run_batch return, so no workbook is needed.
Run from the project folder:  pytest
"""

import csv

import main
from main import quarter_mismatch_warning, write_summary_csv


def ok(company, quarter, tripped=(), cannot_evaluate=(), gap_count=0, blank_quarters=()):
    """A successful company's result."""
    return {"company": company, "quarter": quarter, "flags_total": 9, "tripped": list(tripped),
            "cannot_evaluate": list(cannot_evaluate), "gap_count": gap_count,
            "blank_quarters": list(blank_quarters), "error": None}


def failed(company, error):
    return {"company": company, "error": error}


def test_summary_csv_has_one_row_per_company(tmp_path, monkeypatch):
    # Pretend build_deck.py doesn't exist yet, so the OK text is fixed whatever step we're on.
    monkeypatch.setattr(main, "BUILD_DECK_PATH", tmp_path / "no_build_deck.py")
    results = [
        ok("Northwind", "Q2 2026", tripped=["a"] * 6, gap_count=19, blank_quarters=["Q1 2025"]),
        ok("Fernhollow", "Q2 2026", tripped=["a"] * 7, cannot_evaluate=["Rule of 40"], gap_count=20,
           blank_quarters=["Q2 2025"]),
        failed("Broken", "ValueError: row 3 (header) is missing columns: pipeline"),
    ]
    path = write_summary_csv(results, tmp_path / "batch_summary.csv")
    with open(path, newline="") as file:
        rows = list(csv.reader(file))
    assert rows == [
        ["Company", "Latest quarter", "Flags tripped", "Flags total", "Cannot evaluate", "Data gaps",
         "Blank quarters", "Result"],
        ["Northwind", "Q2 2026", "6", "9", "0", "19", "Q1 2025", "OK (AI + deck skipped)"],
        ["Fernhollow", "Q2 2026", "7", "9", "1", "20", "Q2 2025", "OK (AI + deck skipped)"],
        ["Broken", "", "", "", "", "", "", "FAILED: ValueError: row 3 (header) is missing columns: pipeline"],
    ]


def test_no_warning_when_every_company_ends_on_the_same_quarter():
    results = [ok("Alderpeak", "Q2 2026"), ok("Northwind", "Q2 2026"), failed("Broken", "ValueError: x")]
    assert quarter_mismatch_warning(results) is None


def test_no_warning_for_a_single_company():
    assert quarter_mismatch_warning([ok("Northwind", "Q2 2026")]) is None


def test_warning_names_each_quarter_and_its_companies():
    results = [ok("Alderpeak", "Q2 2026"), ok("Northwind", "Q1 2026"), ok("Fernhollow", "Q2 2026"),
               failed("Broken", "ValueError: x")]  # a failed company has no quarter and is left out
    assert quarter_mismatch_warning(results) == (
        "⚠ Companies end on different quarters (Q2 2026: Alderpeak, Fernhollow; Q1 2026: Northwind) "
        "- compare them with care")


def test_result_text_says_when_ai_was_skipped():
    # build_deck.py exists now, so a --skip-ai run built a deck without AI text.
    assert main.result_text({**ok("Northwind", "Q2 2026"), "ai_skipped": True}) == "OK (AI skipped)"
    assert main.result_text({**ok("Northwind", "Q2 2026"), "ai_skipped": False}) == "OK"
    assert main.result_text(failed("Broken", "ValueError: x")) == "FAILED: ValueError: x"
