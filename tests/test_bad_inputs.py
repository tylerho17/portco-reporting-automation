"""Bad-input tests: small broken workbooks must stop with a message that says where to look.

Each test writes a tiny workbook to pytest's temp folder (never data/), breaks one thing, and
checks that clean_workbook stops with an error naming the sheet and the row, column or cell.

The workbook looks like a real one, so the Excel addresses in the messages are tested properly:

    Sheet "Notes"        junk tab, first (the KPI tab must still be found)
    Sheet "KPI Tracker"  row 1 title, row 2 empty, row 3 header, rows 4-7 Q1-Q4 2025,
                         row 8 budget-only row "Q1 2026 (Budget)"

    Columns: A Quarter | B starting_arr  C new_arr  D expansion_arr  E contraction_arr
             F churned_arr  G revenue  H gross_profit  I net_burn  J ending_cash  K sm_spend
             L new_customers  M headcount  N pipeline | O budget_new_arr  P budget_arr
             Q budget_net_burn

Expected messages are typed out by hand from that layout (e.g. Q2 2025 revenue is cell G5),
not copied from running the code.

Run from the project folder:  python -m pytest -q
"""

import datetime

import pytest
from openpyxl import Workbook

from clean import ACTUAL_COLUMNS, BUDGET_COLUMNS, STANDARD_COLUMNS, clean_workbook

SHEET = "KPI Tracker"
QUARTERS = ["Q1 2025", "Q2 2025", "Q3 2025", "Q4 2025"]
BUDGET_LABEL = "Q1 2026 (Budget)"


# ---------------------------------------------------------------------------
# Building workbooks
# ---------------------------------------------------------------------------

def good_table():
    """A valid KPI table as a list of rows: header, 4 quarters (every value 100), budget-only row."""
    rows = [["Quarter"] + STANDARD_COLUMNS]
    for quarter in QUARTERS:
        rows.append([quarter] + [100] * len(STANDARD_COLUMNS))
    rows.append([BUDGET_LABEL] + [None] * len(ACTUAL_COLUMNS) + [100] * len(BUDGET_COLUMNS))
    return rows


def set_cell(rows, label, column, value):
    """Change one value: the row whose label is `label`, the column whose header is `column`."""
    row = next(r for r in rows[1:] if r[0] == label)
    row[rows[0].index(column)] = value


def drop_column(rows, column):
    """Remove a column (its header and every value)."""
    position = rows[0].index(column)
    return [row[:position] + row[position + 1:] for row in rows]


def add_column(rows, header, value=None):
    """Add a column at the right end with `value` in every row under the header."""
    return [rows[0] + [header]] + [row + [value] for row in rows[1:]]


def write_workbook(path, rows, empty_columns_left=0):
    """Save a junk Notes tab, then the KPI tab: title, empty row, then `rows` from row 3.

    `empty_columns_left` shifts the table right, e.g. 2 starts it in column C.
    """
    workbook = Workbook()
    notes = workbook.active
    notes.title = "Notes"
    notes["A1"] = "scratch - ignore"
    kpi = workbook.create_sheet(SHEET)
    kpi.append(["Test Co KPIs ($K)"])
    kpi.append([])
    for row in rows:
        kpi.append([None] * empty_columns_left + row)
    workbook.save(path)
    return path


def error_from(tmp_path, rows, **options):
    """Write the workbook, run clean_workbook, and return the error message it stopped with."""
    path = write_workbook(tmp_path / "broken.xlsx", rows, **options)
    with pytest.raises(ValueError) as caught:
        clean_workbook(path)
    return str(caught.value)


def assert_stops_with(tmp_path, rows, expected_start, **options):
    """The error must start with `expected_start`: the sheet, the place, then the problem."""
    message = error_from(tmp_path, rows, **options)
    assert message.startswith(expected_start), f"\nExpected to start with: {expected_start}\nGot: {message}"


def test_good_workbook_cleans(tmp_path):
    # The starting point for every test below must itself be valid, or the tests prove nothing.
    actuals, next_budget = clean_workbook(write_workbook(tmp_path / "good.xlsx", good_table()))
    assert list(actuals.index) == QUARTERS
    assert (actuals == 100).all().all()
    assert next_budget.name == BUDGET_LABEL


# ---------------------------------------------------------------------------
# Missing required column
# ---------------------------------------------------------------------------

def test_missing_column_stops(tmp_path):
    rows = drop_column(good_table(), "pipeline")
    assert_stops_with(tmp_path, rows, "Sheet 'KPI Tracker', row 3 (header) is missing columns: pipeline")


def test_several_missing_columns_are_all_listed(tmp_path):
    # All of them at once, in the standard order, so one fix round is enough.
    rows = drop_column(drop_column(good_table(), "budget_arr"), "starting_arr")
    assert_stops_with(tmp_path, rows,
                      "Sheet 'KPI Tracker', row 3 (header) is missing columns: starting_arr, budget_arr")


def test_missing_quarter_column_stops(tmp_path):
    # Without a 'Quarter' header no tab counts as the KPI tab, so the message lists the tabs it checked.
    rows = drop_column(good_table(), "Quarter")
    message = error_from(tmp_path, rows)
    assert "has a 'Quarter' header in its first 10 rows (tabs: ['Notes', 'KPI Tracker'])" in message


# ---------------------------------------------------------------------------
# Two headers with the same meaning
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("first_header, second_header, expected_start", [
    # Column B holds the first header; the second one is added at the far right (column R).
    ("Starting ARR", "Beginning ARR",
     "Sheet 'KPI Tracker', row 3 (header): columns B ('Starting ARR') and R ('Beginning ARR') both mean "
     "'starting_arr'"),
    ("starting_arr", "starting_arr",   # the exact same header twice
     "Sheet 'KPI Tracker', row 3 (header): columns B ('starting_arr') and R ('starting_arr') both mean "
     "'starting_arr'"),
])
def test_two_headers_same_meaning_stop(tmp_path, first_header, second_header, expected_start):
    rows = good_table()
    rows[0][1] = first_header
    rows = add_column(rows, second_header, value=90)
    assert_stops_with(tmp_path, rows, expected_start)


def test_budget_alias_duplicate_stops(tmp_path):
    # "Net Burn (Bud.)" is an alias for budget_net_burn (column Q), which is already there.
    rows = add_column(good_table(), "Net Burn (Bud.)")
    assert_stops_with(tmp_path, rows,
                      "Sheet 'KPI Tracker', row 3 (header): columns Q ('budget_net_burn') and "
                      "R ('Net Burn (Bud.)') both mean 'budget_net_burn'")


def test_two_quarter_columns_stop(tmp_path):
    # Two label columns: which one holds the real quarters?
    rows = add_column(good_table(), "Quarter")
    assert_stops_with(tmp_path, rows,
                      "Sheet 'KPI Tracker', row 3 (header): columns A ('Quarter') and R ('Quarter') both mean "
                      "'quarter'")


# ---------------------------------------------------------------------------
# Unknown header
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("header", ["EBITDA", "Revenu", "ARR"])
def test_unknown_header_stops(tmp_path, header):
    # An extra column (R) that isn't one of the 16 inputs: never dropped silently.
    rows = add_column(good_table(), header, value=50)
    assert_stops_with(tmp_path, rows,
                      f"Sheet 'KPI Tracker', cell R3 (header): Unknown column header {header!r}")


def test_unknown_header_message_says_how_to_fix(tmp_path):
    message = error_from(tmp_path, add_column(good_table(), "EBITDA"))
    assert "add it to HEADER_ALIASES in clean.py; otherwise delete the column" in message


def test_values_under_blank_header_stop(tmp_path):
    # A column with numbers but no header would otherwise be skipped without a word.
    rows = add_column(good_table(), None)
    set_cell(rows, "Q2 2025", None, 75)  # first value in that column: row 5
    assert_stops_with(tmp_path, rows,
                      "Sheet 'KPI Tracker', column R has values (first in cell R5) but no header in row 3")


def test_unknown_header_address_follows_table_position(tmp_path):
    # Table shifted 2 columns right (starts in C): the extra column is T, not R.
    rows = add_column(good_table(), "EBITDA")
    assert_stops_with(tmp_path, rows, "Sheet 'KPI Tracker', cell T3 (header): Unknown column header 'EBITDA'",
                      empty_columns_left=2)


# ---------------------------------------------------------------------------
# Quarters out of order
# ---------------------------------------------------------------------------

def reorder_quarters(labels):
    """The good table with its quarter rows replaced by `labels`, in that order (budget row kept last)."""
    rows = good_table()
    return [rows[0]] + [[label] + [100] * len(STANDARD_COLUMNS) for label in labels] + [rows[-1]]


@pytest.mark.parametrize("labels, expected_start", [
    # Rows 4, 5, 6... hold the labels in the order given; the message names the first bad row.
    (["Q1 2025", "Q3 2025", "Q2 2025", "Q4 2025"],   # two quarters swapped
     "Sheet 'KPI Tracker', row 5: After Q1 2025 expected Q2 2025, found Q3 2025"),
    (["Q1 2025", "Q2 2025", "Q4 2025"],              # Q3 row missing
     "Sheet 'KPI Tracker', row 6: After Q2 2025 expected Q3 2025, found Q4 2025"),
    (["Q4 2025", "Q3 2025", "Q2 2025", "Q1 2025"],   # newest first
     "Sheet 'KPI Tracker', row 5: After Q4 2025 expected Q1 2026, found Q3 2025"),
    (["Q3 2025", "Q4 2025", "Q1 2025"],              # year rollover wrong (should be Q1 2026)
     "Sheet 'KPI Tracker', row 6: After Q4 2025 expected Q1 2026, found Q1 2025"),
])
def test_quarters_out_of_order_stop(tmp_path, labels, expected_start):
    assert_stops_with(tmp_path, reorder_quarters(labels), expected_start)


def test_out_of_order_message_says_how_to_fix(tmp_path):
    message = error_from(tmp_path, reorder_quarters(["Q1 2025", "Q3 2025"]))
    assert "a quarter with no data still needs its own row" in message


def test_repeated_quarter_stops(tmp_path):
    rows = reorder_quarters(["Q1 2025", "Q2 2025", "Q2 2025", "Q3 2025"])
    assert_stops_with(tmp_path, rows,
                      "Sheet 'KPI Tracker', row 6: quarter 'Q2 2025' appears twice (also in row 5)")


@pytest.mark.parametrize("label", ["Q2-2025", "2025 Q2", "Q5 2025", "Total"])
def test_unreadable_quarter_label_stops(tmp_path, label):
    rows = reorder_quarters(["Q1 2025", label, "Q3 2025"])
    assert_stops_with(tmp_path, rows,
                      f"Sheet 'KPI Tracker', row 5: Can't read quarter label {label!r} (expected e.g. 'Q2 2026')")


def test_values_without_quarter_label_stop(tmp_path):
    # A row with numbers but no label would otherwise be skipped, losing a quarter without a word.
    rows = reorder_quarters(["Q1 2025", None, "Q2 2025"])
    assert_stops_with(tmp_path, rows,
                      "Sheet 'KPI Tracker', row 5 has values but no quarter label in column A")


def test_header_without_quarter_rows_stops(tmp_path):
    assert_stops_with(tmp_path, [good_table()[0]],
                      "Sheet 'KPI Tracker', row 3 (header) has no quarter rows under it")


# ---------------------------------------------------------------------------
# Budget-only row containing actuals
# ---------------------------------------------------------------------------

def test_budget_row_with_one_actual_stops(tmp_path):
    rows = good_table()
    set_cell(rows, BUDGET_LABEL, "revenue", 500)
    assert_stops_with(tmp_path, rows,
                      "Sheet 'KPI Tracker', row 8 ('Q1 2026 (Budget)') is labelled as a budget-only row but also "
                      "has actual values in G8 (revenue)")


def test_budget_row_lists_every_actual_cell(tmp_path):
    rows = good_table()
    set_cell(rows, BUDGET_LABEL, "starting_arr", 900)
    set_cell(rows, BUDGET_LABEL, "pipeline", "$1.2M")  # text counts as a value too
    assert_stops_with(tmp_path, rows,
                      "Sheet 'KPI Tracker', row 8 ('Q1 2026 (Budget)') is labelled as a budget-only row but also "
                      "has actual values in B8 (starting_arr), N8 (pipeline)")


def test_budget_row_message_says_which_columns_are_allowed(tmp_path):
    rows = good_table()
    set_cell(rows, BUDGET_LABEL, "revenue", 500)
    message = error_from(tmp_path, rows)
    assert message.endswith("a budget-only row may fill only budget_new_arr, budget_arr, budget_net_burn")


def test_second_budget_row_stops(tmp_path):
    rows = good_table()
    rows.append(["Q2 2026 (Plan)"] + [None] * len(ACTUAL_COLUMNS) + [100] * len(BUDGET_COLUMNS))
    assert_stops_with(tmp_path, rows,
                      "Sheet 'KPI Tracker', row 9 ('Q2 2026 (Plan)') is a second budget-only row "
                      "(the first is 'Q1 2026 (Budget)')")


# ---------------------------------------------------------------------------
# Text that can't be read as a number
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("cell", [
    "TBD", "about 400",
    "n/a", "NA", "NULL", "None",  # pandas used to read these as empty cells, so they became "data missing"
    "(120)",                   # accounting negative: could be read either way, so stop
    "12.5%",                   # a percent in a $K column
    "1.250,5",                 # European style: would silently read as 1.2505
    "1,2,5",                   # commas not in groups of 3
    "−1.2M",                   # a typographic minus sign, not "-"
    "12.5B",                   # unknown unit
    "inf", "nan", "1e3",       # Python would read these as infinity, blank and 1000
    True,                      # a TRUE/FALSE cell: Python would count True as 1
])
def test_unreadable_number_names_the_cell(tmp_path, cell):
    # Q2 2025 revenue is cell G5.
    rows = good_table()
    set_cell(rows, "Q2 2025", "revenue", cell)
    assert_stops_with(tmp_path, rows,
                      f"Sheet 'KPI Tracker', cell G5 (Q2 2025, revenue): Can't read {cell!r} as a number")


@pytest.mark.parametrize("error_value", ["#DIV/0!", "#REF!", "#N/A", "#VALUE!"])
def test_excel_error_cell_stops(tmp_path, error_value):
    # A broken formula in the source workbook. pandas reads it as an empty cell, so without this
    # check it would show up on the deck as "data missing" instead of being fixed. Q3 2025 net_burn is I6.
    rows = good_table()
    set_cell(rows, "Q3 2025", "net_burn", error_value)  # openpyxl saves these strings as real error cells
    assert_stops_with(tmp_path, rows,
                      f"Sheet 'KPI Tracker', cell I6: Excel error value {error_value!r} - fix the formula")


def test_date_in_number_cell_stops(tmp_path):
    # Excel sometimes turns typed text into a date. Q4 2025 budget_arr is cell P7.
    rows = good_table()
    set_cell(rows, "Q4 2025", "budget_arr", datetime.datetime(2025, 3, 1))
    assert_stops_with(tmp_path, rows, "Sheet 'KPI Tracker', cell P7 (Q4 2025, budget_arr): Can't read ")


def test_unreadable_number_in_budget_row_names_the_cell(tmp_path):
    # Budget row, budget_net_burn: cell Q8.
    rows = good_table()
    set_cell(rows, BUDGET_LABEL, "budget_net_burn", "about 400")
    assert_stops_with(tmp_path, rows,
                      "Sheet 'KPI Tracker', cell Q8 (Q1 2026 (Budget), budget_net_burn): "
                      "Can't read 'about 400' as a number")


def test_unreadable_number_message_says_how_to_fix(tmp_path):
    rows = good_table()
    set_cell(rows, "Q2 2025", "revenue", "n/a")
    assert error_from(tmp_path, rows).endswith("leave the cell empty if there's no data)")
