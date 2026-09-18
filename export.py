"""Export one company's metrics and flags for other tools, and a summary to paste into an email (Task 11).

Four files per company, in output/ (or --output-dir), next to its metrics workbook:
- <company>_metrics.csv:  one row per quarter and metric: the number (ratios as decimals, $K as
                          thousands), the text the deck and workbook show, its unit, why there is no
                          number when there isn't one, and whether its flag tripped (the red cell in Excel).
- <company>_flags.csv:    one row per flag for the latest quarter: value, threshold, status and reason.
- <company>_export.json:  both of the above, plus the flag count, runway at next quarter's budgeted
                          burn, the data gaps, and the SHA-256 hashes of the workbook, config.yaml and
                          column mapping it was built from (so a tool can tell which inputs it reflects).
- <company>_email.html:   slide 1's key metrics table, the flag count, runway context and data gaps,
                          laid out the way Outlook needs (see email_html). Open it in a browser, select
                          all, copy, and paste into a new email.

Rules:
- No math: every number comes from metrics.py (through build_deck.collect_deck_data, the same numbers
  the deck, the metrics workbook and the web page show), and every word from excel_output.py and
  build_deck.py. tests/test_export.py and check_export.py prove the values equal the workbook's.
- A value with no number is left empty (CSV) or null (JSON), never 0 and never "NaN" or "Infinity"
  (strict JSON has neither). Its status says why: missing input, no prior period, not meaningful, or infinite.
- No AI text, no run time: the same workbook and config.yaml always give byte-for-byte the same files.
- UTF-8 without a byte order mark; no em dash.

Run: python export.py data/northwind.xlsx        (one company)
     python export.py --all                      (every workbook in data/)
"""

import argparse
import csv
import io
import json
import math
import sys
from html import escape
from pathlib import Path

from build_deck import COMBO_TABLE_TEXT, column_widths, flag_count_text, gaps_lines, kpi_header, kpi_rows, value_text
from excel_output import KIND_LABELS, FLAG_KINDS, combo_rule_words, combo_window_text, gap_label, status_label, \
    tripped_cells
from main import DATA_DIR, OUTPUT_DIR, company_name, find_workbooks, shown_path
from mapping import mapping_sha256
from metrics import (CANNOT_EVALUATE, CONFIG_PATH, DOLLAR_COLUMNS, METRIC_LABELS, MONTH_COLUMNS, PASS, TRIP,
                     format_value, load_config, runway_context_label)
from portfolio import load_company, plain
from provenance import file_sha256
from theme import FONT, LINE, MID_GRAY, NAVY, SLATE, STATUS_COLORS, SURFACE, WHITE

EXPORT_KINDS = ("metrics_csv", "flags_csv", "json", "email")
FILE_ENDINGS = {"metrics_csv": "_metrics.csv", "flags_csv": "_flags.csv", "json": "_export.json",
                "email": "_email.html"}

STORED_DIGITS = 16   # significant digits the metrics workbook keeps (as_stored)

# A value's status: a number, infinite (a CLAUDE.md edge case), or metrics.py's reason for no number.
NUMBER = "number"
INFINITE = "infinite"

# What the "unit" column means, written into the JSON for whoever reads it.
UNITS = {
    "$K": "thousands of US dollars",
    "ratio": "a decimal: 0.971 means 97.1%",
    "months": "months",
    "multiple": "times: 2.35 means 2.35x",
}

METRIC_FIELDS = ["company", "quarter", "metric", "label", "unit", "value", "text", "status", "flag_tripped"]
FLAG_FIELDS = ["company", "quarter", "flag", "metric", "value", "text", "threshold", "trips_when", "status",
               "reason", "status_text"]
FORMAT_NAME = "board-pack-generator export"
FORMAT_VERSION = 1   # raise it when a field is renamed or removed, so a downstream tool can tell

# The email: one column 640 px wide (what most email clients show without scrolling).
EMAIL_WIDTH = 640
EMAIL_FONT = f"{FONT}, Helvetica, sans-serif"   # on every cell and paragraph: Outlook falls back to Times otherwise
TITLE_PX, BODY_PX, TABLE_PX, NOTE_PX = 20, 14, 13, 12
CELL_PADDING = 6
RUNWAY_LINE = "Runway at next quarter's budgeted burn: {text} (context, not a flag)"
EMAIL_NOTE = ("Fictional data. Every number is computed in Python from {source}, the same numbers as the "
              "board deck and the metrics workbook; no AI text. $K = thousands of dollars.")


# ---------------------------------------------------------------------------
# 1. Metric rows: one per quarter and metric
# ---------------------------------------------------------------------------

def unit_of(metric):
    """'$K', 'months', 'multiple' or 'ratio': the same split as the workbook's number formats."""
    if metric in DOLLAR_COLUMNS:
        return "$K"
    if metric in MONTH_COLUMNS:
        return "months"
    return "multiple" if metric == "burn_multiple" else "ratio"


def as_stored(value):
    """A number as the metrics workbook stores it: 16 significant digits (openpyxl writes "%.16g").

    2.3493975903614457 -> 2.349397590361446. The 17th digit is float noise (Excel itself keeps 15),
    and without this a tool comparing the CSV with the workbook would see two different numbers.
    """
    return float(f"{value:.{STORED_DIGITS}g}")


def value_and_status(data, metric, quarter):
    """(the number, NUMBER), or (None, why there is none): a reason from metrics.py, or INFINITE."""
    reason = data["reasons"].loc[quarter, metric]
    if isinstance(reason, str):
        return None, reason
    value = data["metrics"].loc[quarter, metric]
    if math.isinf(value):
        return None, INFINITE   # the text says which edge case: "∞ (ARR shrank)"
    return as_stored(value), NUMBER


def metric_row(data, tripped, metric, quarter):
    """One metric in one quarter, in METRIC_FIELDS order. tripped = the workbook's red cells."""
    value, status = value_and_status(data, metric, quarter)
    return {"company": data["company"], "quarter": quarter, "metric": metric, "label": METRIC_LABELS[metric],
            "unit": unit_of(metric), "value": value, "text": value_text(data, metric, quarter),
            "status": status, "flag_tripped": (quarter, metric) in tripped}


def metric_rows(data):
    """Every quarter (oldest first), and within it every metric in the workbook's column order."""
    tripped = tripped_cells(data["metrics"], data["reasons"], data["config"])
    return [metric_row(data, tripped, metric, quarter)
            for quarter in data["metrics"].index for metric in data["metrics"].columns]


# ---------------------------------------------------------------------------
# 2. Flag rows, runway context, data gaps
# ---------------------------------------------------------------------------

def combo_trips_when(config):
    """The combo rule in the workbook's words: 'NRR falls at least 1 pt and ..., last 3 quarters'."""
    return f"{combo_rule_words(config)}, {combo_window_text(config)}"


def flag_row(data, flag):
    """One flag, in FLAG_FIELDS order. The combo rule has no value or threshold, only what trips it."""
    row = {"company": data["company"], "quarter": flag["quarter"], "flag": flag["flag"], "metric": flag["metric"],
           "value": None, "text": None, "threshold": None, "trips_when": combo_trips_when(data["config"]),
           "status": flag["status"], "reason": flag["reason"], "status_text": status_label(flag)}
    if flag["metric"] is not None:
        value, _ = value_and_status(data, flag["metric"], flag["quarter"])
        row.update(value=value, text=value_text(data, flag["metric"], flag["quarter"]), threshold=flag["threshold"],
                   trips_when=KIND_LABELS[FLAG_KINDS[flag["flag"]]])
    return row


def flag_rows(data):
    """Every flag for the latest quarter, in the order metrics.py evaluates them."""
    return [flag_row(data, flag) for flag in data["flags"]]


def flag_summary(flags):
    """'6 of 9 flags tripped' and the counts behind it."""
    statuses = [flag["status"] for flag in flags]
    return {"text": flag_count_text(flags), "checked": len(flags), "tripped": statuses.count(TRIP),
            "passed": statuses.count(PASS), "cannot_evaluate": statuses.count(CANNOT_EVALUATE)}


def runway_context(data):
    """Runway at next quarter's budgeted burn (context, not a flag): the number (or None) and its text."""
    runway = data["runway_at_budget"]
    words = runway_context_label(runway, data["next_budget"] is not None)   # None when it's a normal number
    return {"quarter": data["latest"], "value": None if words else as_stored(runway),
            "text": words or format_value("runway_months", runway)}


def gap_rows(gaps):
    """Each metric or flag with an input missing: its key, its label, and the quarters."""
    return [{"name": name, "label": gap_label(name), "quarters": quarters} for name, quarters in gaps.items()]


# ---------------------------------------------------------------------------
# 3. CSV and JSON
# ---------------------------------------------------------------------------

def export_record(data, workbook_path, config_path=CONFIG_PATH):
    """Everything the JSON holds, as a dict (no NaN or infinity anywhere: those are None)."""
    return {
        "format": FORMAT_NAME, "format_version": FORMAT_VERSION,
        "company": data["company"], "workbook": Path(workbook_path).name,
        "sources": {"workbook_sha256": file_sha256(workbook_path), "config_sha256": file_sha256(config_path),
                    "mapping_sha256": mapping_sha256(workbook_path)},
        "quarters": list(data["metrics"].index), "latest_quarter": data["latest"], "units": UNITS,
        "flag_summary": flag_summary(data["flags"]), "flags": flag_rows(data),
        "runway_at_next_budget": runway_context(data), "data_gaps": gap_rows(data["gaps"]),
        "metrics": metric_rows(data),
    }


def csv_cell(value):
    """One CSV cell: empty for no value, true/false (as in JSON), a float's exact digits (repr)."""
    if value is None:
        return ""
    if isinstance(value, bool):   # before numbers: True is also the number 1 in Python
        return "true" if value else "false"
    if isinstance(value, float):
        return repr(value)
    return value


def csv_text(rows, fields):
    """A header row, then one line per row, in `fields` order."""
    buffer = io.StringIO()
    writer = csv.writer(buffer)   # quotes a cell only when it holds a comma or quote; lines end \r\n (the CSV standard)
    writer.writerow(fields)
    for row in rows:
        writer.writerow([csv_cell(row[field]) for field in fields])
    return buffer.getvalue()


def json_text(record):
    """Indented JSON. allow_nan=False stops with an error rather than write NaN, which strict readers reject."""
    return json.dumps(record, indent=2, ensure_ascii=False, allow_nan=False) + "\n"


# ---------------------------------------------------------------------------
# 4. The email summary
# ---------------------------------------------------------------------------
# Outlook draws email with Word's engine, so the HTML keeps to what Word understands:
# - inline styles only (it drops <style> sheets and classes), and a font on every cell and paragraph;
# - tables for layout, each with width, cellpadding, cellspacing and border set as attributes;
# - fills as bgcolor as well as background-color, colors as 6-digit hex; no images, flexbox or scripts.
# check_export.outlook_problems tests every one of these.

def css(**declarations):
    """css(font_size="14px") -> 'font-size: 14px' (underscores become hyphens)."""
    return "; ".join(f"{name.replace('_', '-')}: {value}" for name, value in declarations.items())


def paragraph_html(text, size, color=SLATE, bold=False, margin="0 0 8px 0"):
    """One paragraph with its font, size, color and spacing written on it."""
    style = css(font_family=EMAIL_FONT, font_size=f"{size}px", font_weight="bold" if bold else "normal",
                color=f"#{color}", margin=margin)
    return f'<p style="{style}">{escape(plain(text), quote=False)}</p>'


def cell_html(tag, text, fill, color, width=None, bold=False):
    """One table cell ("th" or "td"), filled and colored the Outlook way (bgcolor and background-color)."""
    width_part = f'width="{width}" ' if width else ""
    style = css(font_family=EMAIL_FONT, font_size=f"{TABLE_PX}px", font_weight="bold" if bold else "normal",
                color=f"#{color}", background_color=f"#{fill}", border_bottom=f"1px solid #{LINE}")
    return (f'<{tag} {width_part}align="left" valign="top" bgcolor="#{fill}" style="{style}">'
            f'{escape(plain(text), quote=False)}</{tag}>')


def email_table_rows(data):
    """Slide 1's rows, with the combo rule's own words in place of the deck's pointer to slide 3."""
    combo = combo_trips_when(data["config"])
    return [([combo if cell == COMBO_TABLE_TEXT else cell for cell in cells], status)
            for cells, status in kpi_rows(data)]


def table_row_html(cells, status, stripe):
    """A body row: striped like the deck's table, the status cell of a flag in its status colors."""
    parts = []
    for column, text in enumerate(cells):
        fill, color = STATUS_COLORS[status] if status is not None and column == len(cells) - 1 else (stripe, SLATE)
        parts.append(cell_html("td", text, fill, color))
    return "<tr>" + "".join(parts) + "</tr>"


def table_html(data):
    """Slide 1's key metrics table: navy header, striped rows, a status cell per flag."""
    header = "".join(cell_html("th", text, NAVY, WHITE, width, bold=True)
                     for text, width in zip(kpi_header(data), column_widths(EMAIL_WIDTH)))
    rows = [table_row_html(cells, status, SURFACE if number % 2 == 0 else WHITE)
            for number, (cells, status) in enumerate(email_table_rows(data), start=1)]
    style = css(width=f"{EMAIL_WIDTH}px", border_collapse="collapse")
    return (f'<table width="{EMAIL_WIDTH}" cellpadding="{CELL_PADDING}" cellspacing="0" border="0" style="{style}">\n'
            f"<tr>{header}</tr>\n" + "\n".join(rows) + "\n</table>")


def email_html(data):
    """The whole email: title, flag count, key metrics table, runway context, data gaps, a note on the source."""
    title = f"{data['company']}: board update, {data['latest']}"
    body = [paragraph_html(title, TITLE_PX, NAVY, bold=True, margin="0 0 4px 0"),
            paragraph_html(flag_count_text(data["flags"]), BODY_PX, margin="0 0 12px 0"),
            table_html(data),
            paragraph_html(RUNWAY_LINE.format(text=runway_context(data)["text"]), BODY_PX, margin="12px 0 12px 0"),
            paragraph_html("Data gaps", BODY_PX, NAVY, bold=True, margin="0 0 4px 0")]
    body += [paragraph_html(f"• {line}", BODY_PX, margin="0 0 4px 0") for line in gaps_lines(data["gaps"])]
    body.append(paragraph_html(EMAIL_NOTE.format(source=data["source_name"]), NOTE_PX, MID_GRAY, margin="16px 0 0 0"))
    return ('<!DOCTYPE html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n'
            f"<title>{escape(plain(title), quote=False)}</title>\n</head>\n"
            f'<body style="{css(margin="0", padding="16px")}">\n' + "\n".join(body) + "\n</body>\n</html>\n")


# ---------------------------------------------------------------------------
# 5. Files: the web page's downloads and the command line
# ---------------------------------------------------------------------------

def export_paths(workbook_path, output_dir=OUTPUT_DIR):
    """data/northwind.xlsx -> {"metrics_csv": output/northwind_metrics.csv, ...} in EXPORT_KINDS order."""
    stem = Path(workbook_path).stem
    return {kind: Path(output_dir) / f"{stem}{FILE_ENDINGS[kind]}" for kind in EXPORT_KINDS}


def export_text(kind, data, workbook_path):
    """One export's text: "metrics_csv", "flags_csv", "json" or "email"."""
    if kind == "metrics_csv":
        return csv_text(metric_rows(data), METRIC_FIELDS)
    if kind == "flags_csv":
        return csv_text(flag_rows(data), FLAG_FIELDS)
    if kind == "json":
        return json_text(export_record(data, workbook_path))
    return email_html(data)


def export_bytes(kind, data, workbook_path):
    """One export as the bytes the web page's download button sends (UTF-8, no byte order mark)."""
    return export_text(kind, data, workbook_path).encode("utf-8")


def save_exports(workbook_path, config, output_dir=OUTPUT_DIR):
    """Write the four exports for one workbook. Stops with clean.py's message if it can't be read."""
    data, problem = load_company(workbook_path, config)
    if problem:
        raise ValueError(problem)
    paths = export_paths(workbook_path, output_dir)
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    for kind, path in paths.items():
        path.write_bytes(export_bytes(kind, data, workbook_path))   # bytes: the CSV's \r\n stays as written
    return paths


def main(argv=None):
    parser = argparse.ArgumentParser(description="Export a company's metrics and flags as CSV, JSON and an "
                                                 "email-ready HTML summary.")
    parser.add_argument("workbook", nargs="?", help="path to a KPI workbook, e.g. data/northwind.xlsx")
    parser.add_argument("--all", action="store_true", help="every workbook in data/")
    parser.add_argument("--data-dir", default=DATA_DIR, help="where --all looks (default: data/)")
    parser.add_argument("--output-dir", default=OUTPUT_DIR, help="where the files go (default: output/)")
    args = parser.parse_args(argv)
    if bool(args.workbook) == args.all:
        parser.error("give one workbook (data/northwind.xlsx) or --all")
    config = load_config()
    failed = 0
    for path in find_workbooks(args.data_dir) if args.all else [Path(args.workbook)]:
        try:
            saved = save_exports(path, config, args.output_dir)
        except ValueError as error:   # one company that can't be read never stops the others
            print(f"{company_name(path)}: {error}")
            failed += 1
            continue
        for saved_path in saved.values():
            print(f"Saved {shown_path(saved_path)}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
