"""Unit tests for clean.py: reading numbers, headers and quarter labels.

Run from the project folder:  pytest
"""

import math

import numpy as np
import pandas as pd
import pytest

from clean import (STANDARD_COLUMNS, check_quarters_in_order, clean_workbook, normalize_header,
                   parse_number, parse_quarter, standard_column)


# ---------------------------------------------------------------------------
# parse_number: one cell -> a number in $K
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("cell, expected", [
    ("$12.5M", 12500.0),     # millions -> thousands
    ("$4.03M", 4030.0),      # exact: plain float math gives 4.03 * 1000 = 4030.0000000000005
    ("120K", 120.0),         # already thousands
    ("850k", 850.0),         # lowercase unit
    ("5,090", 5090.0),       # thousands separator
    ("$1,250K", 1250.0),     # dollar sign, comma and unit together
    (" $1.2 M ", 1200.0),    # stray spaces
    ("-$1.2M", -1200.0),     # negative (e.g. cash-generating burn)
    ("$-1.2M", -1200.0),     # minus after the dollar sign
    ("1,250,000", 1250000.0),  # commas in groups of 3
    ("0", 0.0),
])
def test_parse_number_reads_text(cell, expected):
    assert parse_number(cell) == expected


@pytest.mark.parametrize("cell, expected", [
    (5090, 5090.0),
    (12.5, 12.5),
    (np.int64(300), 300.0),      # pandas hands over numpy numbers
    (np.float64(7.25), 7.25),
])
def test_parse_number_keeps_real_numbers(cell, expected):
    result = parse_number(cell)
    assert result == expected
    assert type(result) is float  # always a plain float, whatever came in


@pytest.mark.parametrize("cell", [None, float("nan"), np.nan, "", "   "])
def test_parse_number_blank_cell_is_nan(cell):
    # Blank stays blank: NaN, never 0.
    assert math.isnan(parse_number(cell))


@pytest.mark.parametrize("cell", ["n/a", "TBD", "(120)", "12.5%", "M", "$", "1.2.3M",
                                  # Task 5: Python's Decimal used to read these as infinity, blank and 1000
                                  "inf", "Infinity", "nan", "1e3",
                                  "1.250,5", "1,2,5", ".5M", "-", True, False])
def test_parse_number_unreadable_text_stops(cell):
    # Anything it can't read with certainty is an error that names the cell, never a guess.
    with pytest.raises(ValueError, match="Can't read"):
        parse_number(cell)


# ---------------------------------------------------------------------------
# normalize_header and standard_column: messy header -> standard column name
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("header, expected", [
    ("Net Burn (Bud.)", "net_burn_bud"),
    ("S&M Spend", "s_m_spend"),
    (" Churned ARR ", "churned_arr"),     # stray spaces
    ("New ARR ($K)", "new_arr_k"),
    ("Cash - End of Qtr", "cash_end_of_qtr"),
    ("GROSS PROFIT", "gross_profit"),
    ("Contraction_ARR", "contraction_arr"),
    ("starting_arr", "starting_arr"),     # already clean: unchanged
    ("Quarter", "quarter"),
])
def test_normalize_header(header, expected):
    assert normalize_header(header) == expected


@pytest.mark.parametrize("header, expected", [
    # Headers that need an alias
    ("Beginning ARR", "starting_arr"),
    ("New ARR ($K)", "new_arr"),
    ("Cash - End of Qtr", "ending_cash"),
    ("S&M Spend", "sm_spend"),
    ("New Logos", "new_customers"),
    ("Headcount (FTE)", "headcount"),
    ("Qualified Pipeline", "pipeline"),
    ("New ARR - Budget", "budget_new_arr"),
    ("Net Burn (Bud.)", "budget_net_burn"),
    # Headers that normalize straight to a standard name
    ("expansion arr", "expansion_arr"),
    (" Churned ARR ", "churned_arr"),
    ("Budget ARR", "budget_arr"),
])
def test_standard_column_maps_messy_headers(header, expected):
    assert standard_column(header) == expected


@pytest.mark.parametrize("header", ["EBITDA", "ARR", "Quarter", ""])
def test_standard_column_unknown_header_stops(header):
    # An unknown header is an error, never silently dropped or guessed.
    with pytest.raises(ValueError, match="Unknown column header"):
        standard_column(header)


# ---------------------------------------------------------------------------
# Quarter labels and quarter order
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("label, expected", [
    ("Q2 2026", (2026, 2)),
    ("Q4 2024", (2024, 4)),
    ("Q1  2025", (2025, 1)),  # extra space is fine
])
def test_parse_quarter(label, expected):
    assert parse_quarter(label) == expected


@pytest.mark.parametrize("label", ["Q5 2026", "Q0 2026", "2026 Q2", "Q2-2026", "Q2 26", "Q3 2026 (Budget)"])
def test_parse_quarter_bad_label_stops(label):
    with pytest.raises(ValueError, match="Can't read quarter label"):
        parse_quarter(label)


@pytest.mark.parametrize("labels", [
    ["Q3 2024", "Q4 2024", "Q1 2025", "Q2 2025"],  # Q4 rolls into next year's Q1
    ["Q2 2026"],                                   # one quarter: nothing to compare
    [],
])
def test_quarters_in_order_pass(labels):
    check_quarters_in_order(labels)  # no error = pass


@pytest.mark.parametrize("labels, message", [
    (["Q3 2024", "Q1 2025"], "After Q3 2024 expected Q4 2024, found Q1 2025"),  # skipped
    (["Q3 2024", "Q3 2024"], "After Q3 2024 expected Q4 2024, found Q3 2024"),  # repeated
    (["Q4 2024", "Q3 2024"], "After Q4 2024 expected Q1 2025, found Q3 2024"),  # newest first
    (["Q4 2024", "Q1 2026"], "After Q4 2024 expected Q1 2025, found Q1 2026"),  # a whole year skipped
])
def test_quarters_out_of_order_stop(labels, message):
    with pytest.raises(ValueError, match=message):
        check_quarters_in_order(labels)


def write_workbook(path, labels, blank=()):
    """Write a tiny KPI workbook: header row, then one row per label (every value 100).

    Labels in `blank` get a row with the label only, like the blank quarter in the real data.
    """
    rows = [["Quarter"] + STANDARD_COLUMNS]
    for label in labels:
        values = [None] * len(STANDARD_COLUMNS) if label in blank else [100] * len(STANDARD_COLUMNS)
        rows.append([label] + values)
    pd.DataFrame(rows).to_excel(path, header=False, index=False)


def test_workbook_blank_quarter_row_is_kept(tmp_path):
    # A blank row keeps its place as all-NaN, so the order check and look-backs still line up.
    path = tmp_path / "blank.xlsx"
    write_workbook(path, ["Q1 2025", "Q2 2025", "Q3 2025"], blank=["Q2 2025"])
    actuals, _ = clean_workbook(path)
    assert list(actuals.index) == ["Q1 2025", "Q2 2025", "Q3 2025"]
    assert actuals.loc["Q2 2025"].isna().all()


def test_workbook_missing_quarter_row_stops(tmp_path):
    path = tmp_path / "missing.xlsx"
    write_workbook(path, ["Q1 2025", "Q3 2025"])
    with pytest.raises(ValueError, match="expected Q2 2025, found Q3 2025"):
        clean_workbook(path)


def test_workbook_duplicate_quarter_row_stops(tmp_path):
    # Bug found while writing these tests: the second Q2 2025 row used to overwrite the first
    # without any error. It must stop instead.
    path = tmp_path / "duplicate.xlsx"
    write_workbook(path, ["Q1 2025", "Q2 2025", "Q2 2025", "Q3 2025"])
    with pytest.raises(ValueError, match="'Q2 2025' appears twice"):
        clean_workbook(path)
