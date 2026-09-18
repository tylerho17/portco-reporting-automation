"""Automated proof for Task 10: what changed since the last run, against each company's answer key.

For each company, "last quarter's workbook" is written from its make_data script without its
latest quarter (Q1 2026 is then the latest, and the budget-only row is Q2 2026's budget). That
workbook is run first, then today's, into a temporary output folder, as a quarter apart would be.

Run: python check_diff.py  -> prints "All checks passed" or stops at the first failure.
"""

import contextlib
import io
from pathlib import Path

from make_data_common import save_workbook


def earlier_budget_label(label, dropped):
    """The company's own budget-row wording, a quarter earlier: 'Q3 2026 (Budget)' -> 'Q2 2026 (Budget)'.

    The label starts with the quarter after the latest actual one ("Q3 2026", two words); with
    the latest quarter dropped, that's the dropped quarter.
    """
    next_quarter = " ".join(label.split(" ")[:2])
    return label.replace(next_quarter, dropped)


def last_quarter_workbook(answer_key, folder):
    """Write the company's workbook as it stood a quarter ago (its latest quarter left off) into folder.

    answer_key is the company's make_data module. The file keeps the company's name
    (folder/northwind.xlsx), so its manifest is the same file as today's workbook's. The dropped
    quarter's budget columns become the budget-only row, as they were before that quarter closed.
    The title line (Fernhollow) is kept; the Notes tab's position isn't (it doesn't change a number).
    """
    dropped = answer_key.QUARTERS[-1]
    path = Path(folder) / answer_key.OUTPUT_PATH.name
    company = {
        "output_path": path,
        "quarters": answer_key.QUARTERS[:-1],
        "blank_quarter": answer_key.BLANK_QUARTER,
        "budget_only_label": earlier_budget_label(answer_key.BUDGET_ONLY_LABEL, dropped),
        "true_data": {column: values[:-1] for column, values in answer_key.TRUE_DATA.items()},
        "next_quarter_budget": {column: answer_key.TRUE_DATA[column][-1] for column in answer_key.NEXT_QUARTER_BUDGET},
        "header_names": answer_key.HEADER_NAMES,
        "text_cells": {cell: style for cell, style in answer_key.TEXT_CELLS.items() if cell[1] != dropped},
        "notes": answer_key.NOTES,
        "title": getattr(answer_key, "TITLE", None),
    }
    with contextlib.redirect_stdout(io.StringIO()):   # save_workbook prints a line per file
        save_workbook(company)
    return path
