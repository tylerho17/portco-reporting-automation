"""Tests for mapping.py: headers clean.py doesn't know get a proposal, and nothing runs until a person confirms it.

Each company's workbook is copied to pytest's temp folder with its headers renamed to words that
are in neither clean.HEADER_ALIASES nor the standard names ("Opening ARR", "GP", "Plan Burn"). The
three companies get different renames, so a proposal can't pass by copying one list. Every test uses
a temporary mappings/ folder, so the project's own mappings/ is never touched.

What is proved:
- proposal: every renamed header is proposed as the column it was renamed from, with a confidence,
  a reason and the cell values as written;
- required confirmation: an unconfirmed header stops clean_workbook with a message naming the header
  and its proposal, and nothing is saved on the way;
- identical metrics: once confirmed, the renamed copy gives exactly the original's numbers, metrics
  and flags, and next quarter's workbook runs with no question asked.

Run from the project folder:  python -m pytest -q
"""

import shutil
from pathlib import Path

import pandas as pd
import pytest
import yaml
from openpyxl import load_workbook

import mapping
from clean import HEADER_ALIASES, STANDARD_COLUMNS, UnconfirmedMappingError, clean_workbook, normalize_header, \
    standard_column
from metrics import compute_metrics, evaluate_flags, load_config, metric_reasons

PROJECT_DIR = Path(__file__).parent.parent

# Standard column -> the new header its copy gets. None of these is a known header.
RENAMES = {
    "northwind": {   # every header renamed: the hardest case
        "starting_arr": "Opening ARR", "new_arr": "Gross New ARR", "expansion_arr": "Upsell ARR",
        "contraction_arr": "Downgrades", "churned_arr": "Lost ARR", "revenue": "Total Revenue",
        "gross_profit": "GP", "net_burn": "Cash Burn", "ending_cash": "Closing Cash Balance",
        "sm_spend": "Sales & Marketing", "new_customers": "Customers Won", "headcount": "FTEs",
        "pipeline": "Sales Pipeline", "budget_new_arr": "Plan New ARR", "budget_arr": "Plan Ending ARR",
        "budget_net_burn": "Plan Burn",
    },
    "alderpeak": {   # a few renamed: the rest are known headers
        "starting_arr": "BoP ARR", "churned_arr": "Churn", "revenue": "Revenue (Total)",
        "ending_cash": "Cash Balance", "headcount": "Employees", "budget_arr": "ARR Target",
    },
    "fernhollow": {  # names that need the values too ("Gross Margin $" could be a percentage)
        "expansion_arr": "Expansion", "gross_profit": "Gross Margin $", "net_burn": "Burn",
        "budget_net_burn": "Burn Budget", "sm_spend": "S&M Expense", "new_customers": "New Clients",
        "pipeline": "Pipe ($K)",
    },
}
COMPANIES = list(RENAMES)


# ---------------------------------------------------------------------------
# Renamed copies
# ---------------------------------------------------------------------------

def header_row_cells(sheet):
    """The cells of the row holding the 'Quarter' header (first 10 rows, as clean.py looks)."""
    for row in sheet.iter_rows(max_row=10):
        if any(normalize_header(cell.value) == "quarter" for cell in row if cell.value is not None):
            return row
    raise AssertionError("no Quarter header")


def renamed_copy(company, folder, renames=None):
    """data/<company>.xlsx copied to folder/<company>.xlsx with RENAMES' headers written in."""
    renames = RENAMES[company] if renames is None else renames
    path = Path(folder) / f"{company}.xlsx"
    workbook = load_workbook(PROJECT_DIR / "data" / f"{company}.xlsx")
    for sheet in workbook.worksheets:
        if sheet.title.lower().startswith("notes"):
            continue
        for cell in header_row_cells(sheet):
            if cell.value is not None and normalize_header(cell.value) != "quarter":
                cell.value = renames.get(standard_column(cell.value), cell.value)
    workbook.save(path)
    return path


@pytest.fixture(autouse=True)
def mappings_dir(tmp_path, monkeypatch):
    """A temporary mappings/ folder for every test: the project's own is never read or written."""
    folder = tmp_path / "mappings"
    monkeypatch.setattr(mapping, "MAPPINGS_DIR", folder)
    return folder


@pytest.fixture
def copies(tmp_path):
    """{company: its renamed copy}."""
    folder = tmp_path / "data"
    folder.mkdir()
    return {company: renamed_copy(company, folder) for company in COMPANIES}


def confirm_all(workbook):
    """What a person does after reading the proposals: confirm each one as proposed."""
    proposals = mapping.review_workbook(workbook)
    return mapping.save_mapping(workbook, {p.header: p.column for p in proposals})


# ---------------------------------------------------------------------------
# Proposal
# ---------------------------------------------------------------------------

def test_the_renamed_headers_are_unknown_to_clean_py():
    # The starting point: if any rename were a known header, the tests below would prove nothing.
    for renames in RENAMES.values():
        for header in renames.values():
            with pytest.raises(ValueError):
                standard_column(header)


@pytest.mark.parametrize("company", COMPANIES)
def test_every_renamed_header_is_proposed_as_the_column_it_came_from(copies, company):
    proposals = mapping.review_workbook(copies[company])
    assert {p.header: p.column for p in proposals} == {new: old for old, new in RENAMES[company].items()}


@pytest.mark.parametrize("company", COMPANIES)
def test_each_proposal_has_a_confidence_a_reason_and_the_values_as_written(copies, company):
    for proposal in mapping.review_workbook(copies[company]):
        assert 0 < proposal.confidence < 1
        assert proposal.reason
        assert 1 <= len(proposal.samples) <= mapping.SAMPLE_COUNT
        assert proposal.column in proposal.choices   # Change offers the proposal too


def test_samples_are_the_cells_as_written():
    # Northwind's revenue is stored as text in places ("$5,090"-style); the page must show what the file says.
    sheet = pd.DataFrame([["Quarter", "Top Line"], ["Q1 2025", "$1.2M"], ["Q2 2025", 850], ["Q3 2025", None]])
    assert mapping.sample_values(sheet, 0, 1) == ["$1.2M", "850"]


def test_a_known_header_is_never_proposed(copies):
    # Alderpeak keeps most of its headers: only the six renamed ones are asked about.
    assert len(mapping.review_workbook(copies["alderpeak"])) == len(RENAMES["alderpeak"])


def test_choices_are_only_the_columns_still_without_a_header(copies):
    # "Churn" can't be starting ARR's second header: every choice is a column no known header has.
    for proposal in mapping.review_workbook(copies["alderpeak"]):
        assert proposal.column in proposal.choices and set(proposal.choices) <= set(RENAMES["alderpeak"])


def test_a_column_ruled_out_by_the_budget_row_is_not_a_choice(copies):
    # "ARR Target" has a value in the budget-only row: clean.py would stop if it were an actual column.
    [target] = [p for p in mapping.review_workbook(copies["alderpeak"]) if p.header == "ARR Target"]
    assert target.choices == ["budget_arr"]


def test_a_workbook_with_every_header_known_has_nothing_to_review():
    for company in COMPANIES:
        assert mapping.review_workbook(PROJECT_DIR / "data" / f"{company}.xlsx") == []


def test_an_extra_column_gets_no_proposal(tmp_path):
    # Every input already has a header, so "EBITDA" can't be one of them: say so, don't guess.
    path = tmp_path / "northwind.xlsx"
    workbook = load_workbook(PROJECT_DIR / "data" / "northwind.xlsx")
    sheet = next(s for s in workbook.worksheets if not s.title.lower().startswith("notes"))
    header = header_row_cells(sheet)
    sheet.cell(row=header[0].row, column=len(header) + 1, value="EBITDA")
    workbook.save(path)
    [proposal] = mapping.review_workbook(path)
    assert proposal.column is None and proposal.choices == []
    assert "every input column already has a header" in proposal.reason


# ---------------------------------------------------------------------------
# The heuristics, one at a time
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("header, words", [
    ("Opening ARR", ["starting", "arr"]),
    ("S&M Expense", ["sm", "spend"]),
    ("Sales & Marketing", ["sm"]),
    ("Total Revenue ($K)", ["revenue"]),
    ("GP", ["gross", "profit"]),
    ("Plan Burn", ["budget", "burn"]),
])
def test_header_words_drop_filler_and_use_the_usual_synonyms(header, words):
    assert mapping.header_words(header) == words


def test_a_typo_still_scores_close_to_its_column():
    score, _ = mapping.name_score(mapping.header_words("Revenu"), "revenue")
    assert score >= 0.9


def test_an_unrelated_name_scores_low():
    score, _ = mapping.name_score(mapping.header_words("EBITDA"), "starting_arr")
    assert score < mapping.MIN_CONFIDENCE


def test_a_value_in_the_budget_only_row_rules_out_every_actual_column():
    # clean.py stops if an actual column has a value in that row, so an actual can't be meant.
    column = mapping.ColumnValues(actuals=[100.0, 110.0], budget_filled=True)
    assert mapping.budget_row_evidence("revenue", column)[0] is None
    assert mapping.budget_row_evidence("budget_arr", column)[0] > 0


def test_cash_rolling_forward_supports_net_burn_and_a_break_counts_against():
    known = {"ending_cash": [1000.0, 900.0, 750.0]}
    fits = mapping.ColumnValues(actuals=[120.0, 100.0, 150.0], budget_filled=False)
    breaks = mapping.ColumnValues(actuals=[120.0, 30.0, 40.0], budget_filled=False)
    assert mapping.roll_forward_evidence("net_burn", fits, known)[0] > 0
    assert mapping.roll_forward_evidence("net_burn", breaks, known)[0] < 0


def test_arr_rolling_forward_supports_starting_arr():
    known = {"new_arr": [100.0, 100.0], "expansion_arr": [20.0, 20.0], "contraction_arr": [5.0, 5.0],
             "churned_arr": [15.0, 15.0]}
    column = mapping.ColumnValues(actuals=[1000.0, 1100.0], budget_filled=False)
    points, reason = mapping.roll_forward_evidence("starting_arr", column, known)
    assert points > 0 and "rolls forward" in reason


def test_gross_profit_above_revenue_counts_against():
    known = {"revenue": [100.0, 120.0]}
    below = mapping.ColumnValues(actuals=[70.0, 80.0], budget_filled=False)
    above = mapping.ColumnValues(actuals=[170.0, 80.0], budget_filled=False)
    assert mapping.gross_profit_evidence("gross_profit", below, known)[0] > 0
    assert mapping.gross_profit_evidence("gross_profit", above, known)[0] < 0


def test_a_blank_quarter_is_skipped_by_the_value_checks():
    # Fernhollow's Q2 2025 is blank: that quarter can neither support nor break a roll-forward.
    known = {"ending_cash": [1000.0, None, 750.0, 600.0]}
    column = mapping.ColumnValues(actuals=[120.0, None, 150.0, 150.0], budget_filled=False)
    assert mapping.roll_forward_evidence("net_burn", column, known)[0] > 0


def test_confidence_levels_are_words_a_reviewer_can_read():
    assert mapping.confidence_text(0.95) == "95% (high)"
    assert mapping.confidence_text(0.6) == "60% (medium)"
    assert mapping.confidence_text(0.41) == "41% (low)"


# ---------------------------------------------------------------------------
# Required confirmation
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("company", COMPANIES)
def test_an_unconfirmed_mapping_stops_naming_each_header_and_its_proposal(copies, company):
    with pytest.raises(UnconfirmedMappingError) as caught:
        clean_workbook(copies[company])
    message = str(caught.value)
    for column, header in RENAMES[company].items():
        assert f"Unknown column header {header!r}" in message
        assert f"proposed {column}" in message
    assert "Nothing is guessed" in message and "--confirm" in message and "Review mapping" in message


def test_the_stop_is_a_value_error_so_every_caller_shows_it_as_an_input_problem(copies):
    with pytest.raises(ValueError, match="Sheet '"):
        clean_workbook(copies["northwind"])


def test_the_stop_carries_the_proposals_for_the_web_page(copies):
    with pytest.raises(UnconfirmedMappingError) as caught:
        clean_workbook(copies["fernhollow"])
    assert {p.header for p in caught.value.proposals} == set(RENAMES["fernhollow"].values())


def test_nothing_is_saved_until_a_person_confirms(copies, mappings_dir):
    for workbook in copies.values():
        mapping.review_workbook(workbook)
        with pytest.raises(UnconfirmedMappingError):
            clean_workbook(workbook)
    assert not mappings_dir.exists() or not any(mappings_dir.iterdir())


def test_a_partly_confirmed_mapping_still_stops_on_the_rest(copies):
    proposals = mapping.review_workbook(copies["northwind"])
    mapping.save_mapping(copies["northwind"], {p.header: p.column for p in proposals[:5]})
    with pytest.raises(UnconfirmedMappingError) as caught:
        clean_workbook(copies["northwind"])
    assert {p.header for p in caught.value.proposals} == {p.header for p in proposals[5:]}


def test_a_high_confidence_proposal_still_needs_confirming(copies):
    # Even "Plan New ARR" (an exact word match) isn't used until confirmed: no confidence is enough.
    [only] = [p for p in mapping.review_workbook(copies["northwind"]) if p.header == "Plan New ARR"]
    assert only.confidence >= 0.8
    with pytest.raises(UnconfirmedMappingError):
        clean_workbook(copies["northwind"])


# ---------------------------------------------------------------------------
# Identical metrics after confirming
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("company", COMPANIES)
def test_after_confirming_the_numbers_metrics_and_flags_are_identical(copies, company):
    confirm_all(copies[company])
    renamed_actuals, renamed_budget = clean_workbook(copies[company])
    actuals, next_budget = clean_workbook(PROJECT_DIR / "data" / f"{company}.xlsx")
    pd.testing.assert_frame_equal(renamed_actuals, actuals)
    pd.testing.assert_series_equal(renamed_budget, next_budget)

    config = load_config()
    renamed_metrics, metrics = compute_metrics(renamed_actuals), compute_metrics(actuals)
    pd.testing.assert_frame_equal(renamed_metrics, metrics)
    # As tables: a flag with no value holds NaN, and NaN never equals NaN in a plain ==.
    pd.testing.assert_frame_equal(
        pd.DataFrame(evaluate_flags(renamed_metrics, metric_reasons(renamed_actuals, renamed_metrics), config)),
        pd.DataFrame(evaluate_flags(metrics, metric_reasons(actuals, metrics), config)))


def test_next_quarter_runs_unattended_once_confirmed(copies, tmp_path):
    # Next quarter's workbook has the same headers and new numbers: no question, no stop.
    confirm_all(copies["northwind"])
    next_quarter = load_workbook(copies["northwind"])
    sheet = next(s for s in next_quarter.worksheets if not s.title.lower().startswith("notes"))
    header = header_row_cells(sheet)
    revenue = next(cell.column for cell in header if cell.value == "Total Revenue")
    sheet.cell(row=header[0].row + 1, column=revenue, value=3300)
    next_quarter.save(copies["northwind"])
    actuals, _ = clean_workbook(copies["northwind"])
    assert actuals["revenue"].iloc[0] == 3300


def test_change_saves_the_column_the_person_chose_not_the_proposal(copies):
    # The reviewer swaps two budget columns: the saved mapping must follow them, not the proposal.
    proposals = mapping.review_workbook(copies["northwind"])
    choices = {p.header: p.column for p in proposals}
    choices["Plan New ARR"], choices["Plan Burn"] = "budget_net_burn", "budget_new_arr"
    mapping.save_mapping(copies["northwind"], choices)
    renamed, _ = clean_workbook(copies["northwind"])
    original, _ = clean_workbook(PROJECT_DIR / "data" / "northwind.xlsx")
    pd.testing.assert_series_equal(renamed["budget_new_arr"], original["budget_net_burn"], check_names=False)


# ---------------------------------------------------------------------------
# The saved file
# ---------------------------------------------------------------------------

def test_the_mapping_is_saved_as_mappings_company_yaml(copies, mappings_dir):
    path = confirm_all(copies["alderpeak"])
    assert path == mappings_dir / "alderpeak.yaml"
    saved = yaml.safe_load(path.read_text())
    assert saved["columns"] == {new: old for old, new in RENAMES["alderpeak"].items()}
    assert saved["company"] == "alderpeak" and saved["confirmed_at"]
    assert path.read_text().startswith("#")   # a comment says what the file is for


def test_a_second_confirmation_keeps_the_first(copies):
    proposals = mapping.review_workbook(copies["northwind"])
    mapping.save_mapping(copies["northwind"], {p.header: p.column for p in proposals[:3]})
    path = mapping.save_mapping(copies["northwind"], {p.header: p.column for p in proposals[3:]})
    assert len(yaml.safe_load(path.read_text())["columns"]) == len(proposals)


def test_the_saved_mapping_matches_headers_as_clean_py_does(copies):
    # "opening arr " in next quarter's file is the same header as "Opening ARR": normalized, like HEADER_ALIASES.
    mapping.save_mapping(copies["northwind"], {"Opening ARR": "starting_arr"})
    assert mapping.confirmed_aliases(copies["northwind"])["opening_arr"] == "starting_arr"


def test_a_mapping_to_a_column_that_does_not_exist_is_refused(copies):
    with pytest.raises(ValueError, match="'Opening ARR' -> 'arr_start' isn't one of the 16 input columns"):
        mapping.save_mapping(copies["northwind"], {"Opening ARR": "arr_start"})


def test_a_hand_edited_file_with_a_bad_column_stops_naming_the_file(copies, mappings_dir):
    mappings_dir.mkdir()
    (mappings_dir / "northwind.yaml").write_text("columns:\n  Opening ARR: arr_start\n")
    with pytest.raises(ValueError, match=r"northwind\.yaml: 'Opening ARR' -> 'arr_start'"):
        clean_workbook(copies["northwind"])


def test_a_known_header_wins_over_a_saved_mapping(tmp_path, mappings_dir):
    # A saved line can't redefine a header clean.py already knows ("Revenue" is revenue).
    workbook = tmp_path / "northwind.xlsx"
    shutil.copy(PROJECT_DIR / "data" / "northwind.xlsx", workbook)
    mapping.save_mapping(workbook, {"Revenue": "gross_profit"})
    actuals, _ = clean_workbook(workbook)
    original, _ = clean_workbook(PROJECT_DIR / "data" / "northwind.xlsx")
    pd.testing.assert_frame_equal(actuals, original)


def test_the_mapping_hash_is_none_without_a_file_and_changes_with_it(copies):
    assert mapping.mapping_sha256(copies["northwind"]) is None
    mapping.save_mapping(copies["northwind"], {"Opening ARR": "starting_arr"})
    first = mapping.mapping_sha256(copies["northwind"])
    mapping.save_mapping(copies["northwind"], {"GP": "gross_profit"})
    assert first and mapping.mapping_sha256(copies["northwind"]) != first


def test_no_mapping_file_exists_for_the_three_companies():
    # Their headers are all known, so the project needs no mapping for them.
    for company in COMPANIES:
        assert not (PROJECT_DIR / "mappings" / f"{company}.yaml").exists()


def test_every_standard_column_is_a_choice_somewhere():
    # The vocabulary covers all 16 inputs, so no column can only be reached by typing its name.
    assert set(mapping.column_names()) == set(STANDARD_COLUMNS)
    assert set(HEADER_ALIASES.values()) <= set(STANDARD_COLUMNS)


# ---------------------------------------------------------------------------
# The command line: python mapping.py <workbook> [--confirm]
# ---------------------------------------------------------------------------

def answers(*replies):
    """A fake keyboard: each call to ask() returns the next reply."""
    replies = list(replies)
    return lambda prompt: replies.pop(0)


def test_the_command_line_lists_the_proposals_and_saves_nothing(copies, capsys, mappings_dir):
    assert mapping.main([str(copies["alderpeak"])]) == 1   # 1: the workbook can't run yet
    printed = capsys.readouterr().out
    assert "BoP ARR" in printed and "starting_arr" in printed and "--confirm" in printed
    assert not mappings_dir.exists()


def test_confirm_on_the_command_line_accepts_with_enter(copies):
    count = len(RENAMES["alderpeak"])
    assert mapping.main([str(copies["alderpeak"]), "--confirm"], ask=answers(*[""] * count)) == 0
    clean_workbook(copies["alderpeak"])   # runs now


def test_confirm_on_the_command_line_can_change_a_column(copies):
    proposals = mapping.review_workbook(copies["alderpeak"])
    replies = {"Churn": "headcount", "Employees": "churned_arr"}   # the reviewer swaps two
    mapping.main([str(copies["alderpeak"]), "--confirm"], ask=answers(*[replies.get(p.header, "") for p in proposals]))
    saved = mapping.confirmed_aliases(copies["alderpeak"])
    assert saved["churn"] == "headcount" and saved["employees"] == "churned_arr"


def test_confirm_on_the_command_line_asks_again_after_a_typo(copies, capsys):
    count = len(RENAMES["alderpeak"])
    mapping.main([str(copies["alderpeak"]), "--confirm"], ask=answers("revenu", *[""] * count))
    assert "isn't one of" in capsys.readouterr().out
    clean_workbook(copies["alderpeak"])


def test_quitting_the_confirmation_saves_nothing(copies, mappings_dir):
    assert mapping.main([str(copies["alderpeak"]), "--confirm"], ask=answers("", "q")) == 1
    assert not mappings_dir.exists()


def test_the_command_line_says_so_when_there_is_nothing_to_confirm(capsys):
    assert mapping.main([str(PROJECT_DIR / "data" / "northwind.xlsx")]) == 0
    assert "every header is a known input column" in capsys.readouterr().out
