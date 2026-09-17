"""Save the computed metrics, flags and data gaps as an Excel workbook (build step 4b).

Output: output/<company>_metrics.xlsx with three sheets:
- Metrics:   one row per quarter, one column per metric. Cells hold real numbers
             (ratios as decimals, e.g. 0.971); Excel number formats display them as 97.1%.
             A cell whose flag trips in that quarter is red; a data-missing cell is gray.
- Flags:     every flag for the latest quarter: value, threshold, status.
             Row color: red = tripped, green = passed, gray = cannot evaluate.
- Data gaps: every metric and flag that can't be shown because input data is missing.

No math happens here. metrics.py computes everything; this file only writes and formats.

Run: python excel_output.py data/northwind.xlsx
"""

import math
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from analyze import METRIC_LABELS
from clean import clean_workbook
from metrics import (DOLLAR_COLUMNS, FLAG_RULES, MISSING, MONTH_COLUMNS, PASS, TRIP, compute_metrics,
                     data_gaps, evaluate_flags, load_config, runway_at_next_budget)

OUTPUT_DIR = Path(__file__).parent / "output"
SHEET_NAMES = ["Metrics", "Flags", "Data gaps"]

# Excel can't store "no value" or infinity as a number, so those cells get words that say why.
# "Data missing" must never look like "not meaningful" (CLAUDE.md), so each case has its own text.
DATA_MISSING = "data missing"
NO_PRIOR_PERIOD = "n/a (no prior period)"      # e.g. YoY in the first 4 quarters: nothing to compare with
INFINITE_LABELS = {                            # the CLAUDE.md edge cases
    "burn_multiple": "∞ (ARR shrank)",
    "runway_months": "∞ (not burning)",
    "cac_payback_months": "∞ (never pays back)",
}
NO_BUDGET_ROW = "n/a (no budget row)"
BUDGET_NOT_BURNING = "∞ (budget not burning)"

STATUS_LABELS = {TRIP: "Tripped", PASS: "Passed", MISSING: "Cannot evaluate — data missing"}
KIND_LABELS = {"min": "below threshold", "max": "above threshold"}  # when a flag trips
FLAG_KINDS = {name: kind for name, _, _, kind in FLAG_RULES}      # flag name -> "min" or "max"

# Status -> (fill color, text color) as hex RGB. Same light red/green Excel uses for "Bad"/"Good".
STATUS_COLORS = {
    TRIP: ("FFC7CE", "9C0006"),
    PASS: ("C6EFCE", "006100"),
    MISSING: ("D9D9D9", "404040"),
}

RUNWAY_CONTEXT_LABEL = "Runway at next quarter's budgeted burn (context, not a flag)"
NO_GAPS_LABEL = "None — every metric and flag has the data it needs"


# ---------------------------------------------------------------------------
# Cell values and formats
# ---------------------------------------------------------------------------

def number_format(column):
    """Excel display format for a metric column. Same rules as metrics.format_value."""
    if column in DOLLAR_COLUMNS:
        return "#,##0"                  # $K with thousands commas: 27,470
    if column in MONTH_COLUMNS:
        return '0.0" mo"'               # 11.0 mo
    if column == "burn_multiple":
        return '0.00"x"'                # 2.35x
    return "0.0%"                       # decimal 0.971 shows as 97.1%


def cell_value(column, value, is_gap):
    """The number itself, or a text label when there is no usable number.

    NaN -> "data missing" if it's a gap, else "no prior period".
    Infinity -> "∞" plus the reason (e.g. burn multiple when ARR shrank).
    """
    if math.isnan(value):
        return DATA_MISSING if is_gap else NO_PRIOR_PERIOD
    if math.isinf(value):
        return INFINITE_LABELS.get(column, "∞") if value > 0 else "-∞"
    return float(value)  # plain Python float; openpyxl writes it exactly


def runway_context_value(runway):
    """Runway at next quarter's budgeted burn: a number, or a label saying why there isn't one."""
    if math.isnan(runway):
        return NO_BUDGET_ROW
    if math.isinf(runway):
        return BUDGET_NOT_BURNING
    return float(runway)


# ---------------------------------------------------------------------------
# Styling helpers
# ---------------------------------------------------------------------------

def color_cell(cell, status):
    """Fill one cell with its status color (red / green / gray) and matching text color."""
    fill_color, text_color = STATUS_COLORS[status]
    cell.fill = PatternFill(fill_type="solid", fgColor=fill_color)
    cell.font = Font(color=text_color)


def write_header(sheet, headers):
    """Write bold headers in row 1 and freeze them so they stay visible while scrolling."""
    sheet.append(headers)
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    sheet.freeze_panes = "B2"  # freezes row 1 and column A


def set_column_widths(sheet, widths):
    """Set each column's width, in characters, left to right."""
    for position, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(position)].width = width


# ---------------------------------------------------------------------------
# Sheet 1: Metrics
# ---------------------------------------------------------------------------

def tripped_cells(metrics, config):
    """Every (quarter, metric column) whose flag trips in that quarter.

    The combo rule has no single metric cell, so it only shows on the Flags sheet.
    """
    tripped = set()
    for quarter in metrics.index:
        for flag in evaluate_flags(metrics, config, quarter):
            if flag["status"] == TRIP and flag["metric"] is not None:
                tripped.add((quarter, flag["metric"]))
    return tripped


def style_metric_cell(cell, column, is_gap, is_tripped):
    """Number format, right alignment, and gray (data missing) or red (tripped) fill."""
    cell.number_format = number_format(column)
    cell.alignment = Alignment(horizontal="right")  # text labels line up with the numbers
    if is_gap:
        color_cell(cell, MISSING)
    elif is_tripped:
        color_cell(cell, TRIP)


def write_metrics_sheet(sheet, metrics, gaps, config):
    """One row per quarter, one column per metric, with formats and highlights."""
    columns = list(metrics.columns)
    write_header(sheet, ["Quarter"] + [METRIC_LABELS[column] for column in columns])
    tripped = tripped_cells(metrics, config)

    for row_number, quarter in enumerate(metrics.index, start=2):  # row 1 is the header
        sheet.cell(row=row_number, column=1, value=quarter)
        for column_number, column in enumerate(columns, start=2):  # column A is the quarter
            is_gap = quarter in gaps.get(column, [])
            value = cell_value(column, metrics.loc[quarter, column], is_gap)
            cell = sheet.cell(row=row_number, column=column_number, value=value)
            style_metric_cell(cell, column, is_gap, (quarter, column) in tripped)

    set_column_widths(sheet, [10] + [max(len(METRIC_LABELS[c]), 12) + 2 for c in columns])


# ---------------------------------------------------------------------------
# Sheet 2: Flags
# ---------------------------------------------------------------------------

FLAG_HEADERS = ["Flag", "Quarter", "Value", "Threshold", "Trips when", "Status"]


def flag_row(flag, gaps, config):
    """One flag as a list of cell values, in FLAG_HEADERS order."""
    status = STATUS_LABELS[flag["status"]]
    if flag["metric"] is None:  # the combo rule is a trend test with no single value
        size = config["combo_lookback_quarters"]
        return [flag["flag"], flag["quarter"], "see NRR and Pipeline on Metrics sheet",
                f"last {size} quarters", "NRR falls and pipeline rises at every step", status]
    is_gap = flag["quarter"] in gaps.get(flag["metric"], [])
    value = cell_value(flag["metric"], flag["value"], is_gap)
    return [flag["flag"], flag["quarter"], value, flag["threshold"],
            KIND_LABELS[FLAG_KINDS[flag["flag"]]], status]


def style_flag_row(sheet, row_number, flag):
    """Color the whole row by status; show value and threshold in the metric's format."""
    for cell in sheet[row_number]:
        color_cell(cell, flag["status"])
    if flag["metric"] is not None:
        for column_number in (3, 4):  # Value, Threshold
            cell = sheet.cell(row=row_number, column=column_number)
            cell.number_format = number_format(flag["metric"])
            cell.alignment = Alignment(horizontal="right")


def write_runway_context(sheet, runway, quarter):
    """Below the flag table, after one empty row: runway at next quarter's budgeted burn."""
    row_number = sheet.max_row + 2
    sheet.cell(row=row_number, column=1, value=RUNWAY_CONTEXT_LABEL)
    sheet.cell(row=row_number, column=2, value=quarter)
    cell = sheet.cell(row=row_number, column=3, value=runway_context_value(runway))
    cell.number_format = number_format("runway_months")
    cell.alignment = Alignment(horizontal="right")


def write_flags_sheet(sheet, flags, gaps, config, runway_at_budget):
    """Every flag for one quarter (the latest), colored by status, plus runway context."""
    write_header(sheet, FLAG_HEADERS)
    for flag in flags:
        sheet.append(flag_row(flag, gaps, config))
        style_flag_row(sheet, sheet.max_row, flag)
    write_runway_context(sheet, runway_at_budget, flags[0]["quarter"])
    set_column_widths(sheet, [56, 10, 36, 18, 42, 32])


# ---------------------------------------------------------------------------
# Sheet 3: Data gaps
# ---------------------------------------------------------------------------

def gap_label(name):
    """'nrr' -> 'NRR (annualized)'; 'flag: Rule of 40' -> 'Flag: Rule of 40'."""
    if name.startswith("flag: "):
        return "Flag: " + name.removeprefix("flag: ")
    return METRIC_LABELS[name]


def write_gaps_sheet(sheet, gaps):
    """One row per affected metric or flag, with the quarters it can't be shown for."""
    write_header(sheet, ["Metric or flag", "Quarters with data missing"])
    if not gaps:
        sheet.append([NO_GAPS_LABEL, ""])
    for name, quarters in gaps.items():
        sheet.append([gap_label(name), ", ".join(quarters)])
    set_column_widths(sheet, [52, 30])


# ---------------------------------------------------------------------------
# Putting it together
# ---------------------------------------------------------------------------

def build_workbook(actuals, next_budget, config):
    """Compute metrics, flags and gaps, then write all three sheets into a new workbook."""
    metrics = compute_metrics(actuals)
    flags = evaluate_flags(metrics, config)  # latest quarter
    gaps = data_gaps(actuals, metrics, flags)

    book = Workbook()
    metrics_sheet = book.active  # a new workbook starts with one empty sheet; reuse it
    metrics_sheet.title = SHEET_NAMES[0]
    write_metrics_sheet(metrics_sheet, metrics, gaps, config)
    write_flags_sheet(book.create_sheet(SHEET_NAMES[1]), flags, gaps, config,
                      runway_at_next_budget(actuals, next_budget))
    write_gaps_sheet(book.create_sheet(SHEET_NAMES[2]), gaps)
    return book


def output_path(workbook_path, output_dir=OUTPUT_DIR):
    """data/northwind.xlsx -> output/northwind_metrics.xlsx"""
    return Path(output_dir) / f"{Path(workbook_path).stem}_metrics.xlsx"


def save_metrics_workbook(workbook_path, config, output_dir=OUTPUT_DIR):
    """Clean one input workbook and save its metrics workbook. Returns the saved path."""
    actuals, next_budget = clean_workbook(workbook_path)
    path = output_path(workbook_path, output_dir)
    path.parent.mkdir(exist_ok=True)
    build_workbook(actuals, next_budget, config).save(path)
    return path


if __name__ == "__main__":
    input_path = sys.argv[1] if len(sys.argv) > 1 else "data/northwind.xlsx"
    print(f"Saved {save_metrics_workbook(input_path, load_config())}")
