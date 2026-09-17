"""Turn a messy KPI workbook into a clean table of numbers.

Main entry point: clean_workbook(path) -> (actuals, next_budget)
- actuals: one row per actual quarter, the 16 standard columns, all in $K.
  Missing values are NaN ("not a number", pandas' marker for "no value").
- next_budget: the 3 budget numbers from the budget-only forecast row, or None.

Nothing here guesses or fills in numbers. If something can't be read with
certainty, the script stops with an error that says what and where.

Run directly to print the clean table:  python clean.py data/northwind.xlsx
"""

import numbers
import re
import sys
from decimal import Decimal, InvalidOperation

import pandas as pd

# ---------------------------------------------------------------------------
# Standard column names (from CLAUDE.md)
# ---------------------------------------------------------------------------

ACTUAL_COLUMNS = [
    "starting_arr", "new_arr", "expansion_arr", "contraction_arr", "churned_arr",
    "revenue", "gross_profit", "net_burn", "ending_cash", "sm_spend",
    "new_customers", "headcount", "pipeline",
]
BUDGET_COLUMNS = ["budget_new_arr", "budget_arr", "budget_net_burn"]
STANDARD_COLUMNS = ACTUAL_COLUMNS + BUDGET_COLUMNS

# Normalized messy header -> standard name. Only needed when they differ:
# a header that already normalizes to a standard name (e.g. "expansion arr") needs no entry.
HEADER_ALIASES = {
    "beginning_arr": "starting_arr",
    "new_arr_k": "new_arr",
    "cash_end_of_qtr": "ending_cash",
    "s_m_spend": "sm_spend",
    "new_logos": "new_customers",
    "headcount_fte": "headcount",
    "qualified_pipeline": "pipeline",
    "new_arr_budget": "budget_new_arr",
    "net_burn_bud": "budget_net_burn",
}

# A row label containing any of these words is the budget-only forecast row.
BUDGET_LABEL_WORDS = ("budget", "bud", "plan")

# Quarter labels look like "Q2 2026": Q, a digit 1-4, a space, a 4-digit year.
QUARTER_PATTERN = re.compile(r"^Q([1-4])\s+(\d{4})$")


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------

def is_blank(value):
    """True for an empty cell: missing (None/NaN) or text that is only spaces."""
    if isinstance(value, str):
        return value.strip() == ""
    return pd.isna(value)


def normalize_header(text):
    """'Net Burn (Bud.)' -> 'net_burn_bud'. Lowercase; each run of symbols/spaces becomes one '_'."""
    return re.sub(r"[^a-z0-9]+", "_", str(text).lower()).strip("_")


def standard_column(header):
    """Map a messy header to its standard column name, or stop if it's unknown."""
    name = normalize_header(header)
    name = HEADER_ALIASES.get(name, name)  # use the alias if there is one, else keep as is
    if name not in STANDARD_COLUMNS:
        raise ValueError(f"Unknown column header {header!r} - add it to HEADER_ALIASES")
    return name


def parse_number(value):
    """Turn one cell into a number in $K.

    '$12.5M' -> 12500.0   '120K' -> 120.0   '5,090' -> 5090.0   blank -> NaN
    """
    if is_blank(value):
        return float("nan")
    if isinstance(value, numbers.Real):  # already a number (int, float, numpy number)
        return float(value)

    # Strip the decoration, then read the unit letter at the end.
    text = str(value).replace("$", "").replace(",", "").replace(" ", "").upper()
    multiplier = 1
    if text.endswith("M"):
        multiplier = 1000  # millions -> thousands
        text = text[:-1]
    elif text.endswith("K"):
        text = text[:-1]   # already thousands

    # Decimal does exact decimal math, so '21.85' * 1000 is exactly 21850 (no binary rounding).
    try:
        return float(Decimal(text) * multiplier)
    except InvalidOperation:
        raise ValueError(f"Can't read {value!r} as a number") from None


def parse_quarter(label):
    """'Q2 2026' -> (2026, 2). Stops if the label isn't in that format."""
    match = QUARTER_PATTERN.match(label)
    if not match:
        raise ValueError(f"Can't read quarter label {label!r} (expected e.g. 'Q2 2026')")
    return int(match.group(2)), int(match.group(1))


def check_quarters_in_order(labels):
    """Stop unless quarters run oldest -> newest with none skipped or repeated.

    QoQ looks back 1 row and YoY looks back 4 rows, so a missing row would
    silently compare the wrong quarters. (A blank row is fine; a missing row is not.)
    """
    quarters = [parse_quarter(label) for label in labels]
    for i in range(1, len(quarters)):
        year, q = quarters[i - 1]
        expected = (year, q + 1) if q < 4 else (year + 1, 1)  # Q4 rolls into next year's Q1
        if quarters[i] != expected:
            raise ValueError(f"After {labels[i - 1]} expected Q{expected[1]} {expected[0]}, found {labels[i]}")


# ---------------------------------------------------------------------------
# Reading the workbook
# ---------------------------------------------------------------------------

def find_kpi_sheet(path):
    """Return (sheet, header_row) for the tab whose header row contains 'Quarter'.

    Checks the first 10 rows of each tab, so title rows above the table are fine.
    Tabs without a 'Quarter' header (like Notes) are ignored.
    """
    sheets = pd.read_excel(path, sheet_name=None, header=None)  # every tab, raw, no header guessing
    for sheet in sheets.values():
        for row_number in range(min(10, len(sheet))):
            cells = [normalize_header(v) for v in sheet.iloc[row_number] if not is_blank(v)]
            if "quarter" in cells:
                return sheet, row_number
    raise ValueError(f"No tab in {path} has a 'Quarter' header")


def map_columns(headers):
    """Return (label_position, {position: standard name}) from the header row."""
    label_position = None
    column_map = {}
    for position, header in headers.items():
        if is_blank(header):
            continue
        if normalize_header(header) == "quarter":
            label_position = position
            continue
        name = standard_column(header)
        if name in column_map.values():
            raise ValueError(f"Two headers both mean {name!r} - check the workbook")
        column_map[position] = name

    missing = [c for c in STANDARD_COLUMNS if c not in column_map.values()]
    if missing:
        raise ValueError(f"Workbook is missing columns: {missing}")
    return label_position, column_map


def parse_row(label, raw_row, column_map):
    """Parse every cell in one row. Errors say which quarter and column failed."""
    row = {}
    for position, name in column_map.items():
        try:
            row[name] = parse_number(raw_row[position])
        except ValueError as error:
            raise ValueError(f"{label}, {name}: {error}") from None
    return row


def is_budget_only_row(label, row):
    """True for the forecast row, e.g. 'Q3 2026 (Budget)'. Stops if that row also has actuals."""
    if not any(word in label.lower() for word in BUDGET_LABEL_WORDS):
        return False
    filled_actuals = [c for c in ACTUAL_COLUMNS if not pd.isna(row[c])]
    if filled_actuals:
        raise ValueError(f"Budget row {label!r} also has actual values in: {filled_actuals}")
    return True


def clean_workbook(path):
    """Read a messy KPI workbook and return (actuals, next_budget)."""
    sheet, header_row = find_kpi_sheet(path)
    label_position, column_map = map_columns(sheet.iloc[header_row])

    actual_rows = {}
    next_budget = None
    for _, raw_row in sheet.iloc[header_row + 1:].iterrows():  # every row below the header
        if is_blank(raw_row[label_position]):
            continue  # skip empty rows under the table
        label = str(raw_row[label_position]).strip()
        row = parse_row(label, raw_row, column_map)

        if is_budget_only_row(label, row):
            if next_budget is not None:
                raise ValueError("Workbook has more than one budget-only row")
            next_budget = pd.Series({c: row[c] for c in BUDGET_COLUMNS}, name=label)
        else:
            actual_rows[label] = row  # blank quarters land here too, as all-NaN rows

    actuals = pd.DataFrame.from_dict(actual_rows, orient="index", columns=STANDARD_COLUMNS)
    actuals.index.name = "quarter"
    check_quarters_in_order(list(actuals.index))
    return actuals, next_budget


if __name__ == "__main__":
    workbook_path = sys.argv[1] if len(sys.argv) > 1 else "data/northwind.xlsx"
    actuals, next_budget = clean_workbook(workbook_path)
    print(actuals.T.to_string())  # .T flips the table: quarters across, columns down
    print(f"\nBudget-only row: {next_budget.name if next_budget is not None else 'none'}")
    if next_budget is not None:
        print(next_budget.to_string())
