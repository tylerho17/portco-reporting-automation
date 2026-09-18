"""A web page for people who don't use the command line: drag in a KPI workbook, get the board pack.

What the page does:
1. You drag in one .xlsx workbook (the same kind main.py reads).
2. It runs the same steps as main.py: clean -> metrics and flags -> Excel -> AI commentary (only if
   the box is ticked) -> deck.
3. It shows the flags and the metrics table in the Excel workbook's colors: red = tripped,
   green = passed, gray = data missing / cannot evaluate.
4. Two download buttons: the deck (.pptx) and the metrics workbook (.xlsx).

Rules:
- No new math: every number and every word comes from metrics.py, excel_output.py and build_deck.py.
- A workbook clean.py can't read shows clean.py's own message, never a traceback.
- Files are built in a temporary folder and handed to the browser, so output/ (main.py's folder)
  is never touched.
- AI commentary: a saved analysis in output/ made from exactly the same numbers is reused (free).
  Otherwise Claude is asked, which costs money - the checkbox says roughly how much.

Run: streamlit run app.py      (or double-click run_app.command on a Mac)
"""

import json
import tempfile
import traceback
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl.utils.exceptions import InvalidFileException

from analyze import build_payload
from build_deck import (PLACEHOLDER_TEXT, analysis_path, collect_deck_data, flag_count_text, gaps_lines,
                        load_analysis, points_text, save_deck, threshold_text, value_text)
from clean import clean_workbook
from excel_output import STATUS_COLORS, save_metrics_workbook, status_label, tripped_cells
from main import INPUT_ERRORS, OUTPUT_DIR, ai_step, api_key_problem, company_name
from metrics import CANNOT_EVALUATE, METRIC_LABELS, MISSING_INPUT, TRIP, load_config

PAGE_TITLE = "Board Pack Generator"
INTRO = ("Drag in a portfolio company's KPI workbook (.xlsx). You'll see its flags and metrics here, "
         "and can download the board deck and the metrics workbook. Fictional data only.")

# Typical cost of the AI step, from README.md's Cost section (latest live run: $0.0911 and about
# 70 seconds per company with claude-sonnet-5). A label, not a calculation: update it with the README.
TYPICAL_AI_COST_USD = 0.09
TYPICAL_AI_SECONDS = 70

# What the page says about the AI commentary.
AI_OFF_NOTE = f"AI commentary not requested: slide 4 of the deck says \"{PLACEHOLDER_TEXT}\"."
AI_REUSED_START = "AI commentary reused from a saved analysis of exactly these numbers"
AI_NEW_NOTE = "AI commentary written by Claude and checked; it is on slide 4. Review it before use."

# Error messages for files clean.py never gets to read.
NOT_A_WORKBOOK = "This file isn't a readable Excel workbook. Save it from Excel as .xlsx and try again."
UNEXPECTED_ERROR_START = "Something unexpected went wrong - please send this file to whoever looks after the tool"
NOT_A_WORKBOOK_ERRORS = (zipfile.BadZipFile, InvalidFileException)

FLAG_COLUMNS = ["Flag", "Value", "Threshold", "Status"]
LEGEND = "Red = tripped · green = passed · gray = data missing or cannot evaluate"


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------

def ai_checkbox_label():
    """'Include AI commentary (typically about $0.09 and 70 seconds per workbook)'."""
    return (f"Include AI commentary (typically about ${TYPICAL_AI_COST_USD:.2f} and "
            f"{TYPICAL_AI_SECONDS} seconds per workbook)")


def save_upload(file_name, data, folder):
    """Write the uploaded bytes to folder/<file name>. Any folder part of the name is dropped."""
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / Path(file_name).name   # "../x.xlsx" -> "x.xlsx": stays inside the folder
    path.write_bytes(data)
    return path


def status_css(status):
    """The Excel workbook's colors for a status, as a style the on-screen table understands."""
    fill, text = STATUS_COLORS[status]
    return f"background-color: #{fill}; color: #{text}"


# ---------------------------------------------------------------------------
# The two tables on the page (formatting only - no math)
# ---------------------------------------------------------------------------

def metrics_table(data):
    """One row per metric, one column per quarter, as the same text the deck shows."""
    quarters = list(data["metrics"].index)
    rows = {METRIC_LABELS[metric]: [value_text(data, metric, quarter) for quarter in quarters]
            for metric in data["metrics"].columns}
    return pd.DataFrame.from_dict(rows, orient="index", columns=quarters)


def metric_cell_css(data, tripped, metric, quarter):
    """Gray if an input is blank, red if its flag trips that quarter, else no color (as in Excel)."""
    if data["reasons"].loc[quarter, metric] == MISSING_INPUT:
        return status_css(CANNOT_EVALUATE)
    return status_css(TRIP) if (quarter, metric) in tripped else ""


def metrics_colors(data):
    """A table the same shape as metrics_table, holding each cell's style."""
    tripped = tripped_cells(data["metrics"], data["reasons"], data["config"])
    quarters = list(data["metrics"].index)
    rows = {METRIC_LABELS[metric]: [metric_cell_css(data, tripped, metric, quarter) for quarter in quarters]
            for metric in data["metrics"].columns}
    return pd.DataFrame.from_dict(rows, orient="index", columns=quarters)


def combo_rule_text(config):
    """What the combo rule tests, with the settings from config.yaml."""
    return (f"NRR falls by at least {points_text(config['combo_min_nrr_drop'])} and pipeline rises "
            f"at every step, last {config['combo_lookback_quarters']} quarters")


def flag_row(data, flag):
    """One flag as [Flag, Value, Threshold, Status]. The combo rule has no single value."""
    if flag["metric"] is None:
        return [flag["flag"], "see NRR and pipeline", combo_rule_text(data["config"]), status_label(flag)]
    return [flag["flag"], value_text(data, flag["metric"], flag["quarter"]), threshold_text(flag),
            status_label(flag)]


def flags_table(data):
    """Every flag for the latest quarter, in the order metrics.py evaluates them."""
    return pd.DataFrame([flag_row(data, flag) for flag in data["flags"]], columns=FLAG_COLUMNS)


def flags_colors(data):
    """Each flag's whole row in its status color, like the Excel Flags sheet."""
    return pd.DataFrame([[status_css(flag["status"])] * len(FLAG_COLUMNS) for flag in data["flags"]],
                        columns=FLAG_COLUMNS)


# ---------------------------------------------------------------------------
# AI commentary: reuse a saved analysis, or ask Claude
# ---------------------------------------------------------------------------

def reusable_analysis(workbook_path, saved_dir, config):
    """The saved analysis in saved_dir if it was made from exactly these numbers and still passes, else None.

    "Exactly these numbers": the facts Claude saw (the payload) must equal today's, not just the
    company and quarter - so an edited workbook is never described by an old analysis.
    """
    path = analysis_path(workbook_path, saved_dir)
    if not path.exists():
        return None
    try:
        saved = json.loads(path.read_text())
    except json.JSONDecodeError:
        return None
    actuals, next_budget = clean_workbook(workbook_path)
    payload = build_payload(company_name(workbook_path), actuals, next_budget, config)
    if not isinstance(saved, dict) or saved.get("payload") != json.loads(json.dumps(payload)):
        return None
    summary, _ = load_analysis(path, payload)  # the same checks the deck makes
    return path if summary else None


def ai_commentary(workbook_path, config, folder, include_ai, saved_dir, client):
    """(analysis file for the deck or None, a note for the page saying what happened)."""
    if not include_ai:
        return None, AI_OFF_NOTE
    saved = reusable_analysis(workbook_path, saved_dir, config)
    if saved:
        return saved, f"{AI_REUSED_START} ({saved.name}): no API call, no cost. Review it before use."
    problem = None if client else api_key_problem()  # a test's fake client needs no key
    if problem:
        return None, f"{PLACEHOLDER_TEXT}: {problem}"
    actuals, next_budget = clean_workbook(workbook_path)
    ai = ai_step(workbook_path, actuals, next_budget, config, folder, client)
    if ai["analysis_file"] is None:
        return None, f"{PLACEHOLDER_TEXT}: the AI {ai['validation']}"
    return ai["analysis_file"], AI_NEW_NOTE


# ---------------------------------------------------------------------------
# One upload -> everything the page shows
# ---------------------------------------------------------------------------

def build_in_folder(folder, file_name, file_bytes, include_ai, saved_dir, client):
    """Run every step on one uploaded workbook inside folder. Raises if a step fails."""
    workbook_path = save_upload(file_name, file_bytes, folder / "input")
    if not zipfile.is_zipfile(workbook_path):  # every .xlsx is a zip file; else pandas' message is cryptic
        raise ValueError(NOT_A_WORKBOOK)
    output_dir = folder / "output"
    config = load_config()
    actuals, next_budget = clean_workbook(workbook_path)   # stops here on a workbook it can't read
    data = collect_deck_data(company_name(workbook_path), workbook_path.name, actuals, next_budget, config)
    excel_path = save_metrics_workbook(workbook_path, config, output_dir)
    analysis_file, ai_note = ai_commentary(workbook_path, config, output_dir, include_ai, saved_dir, client)
    deck, why_unavailable = save_deck(workbook_path, config, analysis_file, output_dir=output_dir)
    if analysis_file is not None and why_unavailable is not None:  # the deck rejected the analysis
        ai_note = f"{PLACEHOLDER_TEXT}: {why_unavailable}"
    return {
        "error": None, "company": data["company"], "quarter": data["latest"],
        "flag_count": flag_count_text(data["flags"]), "gaps": gaps_lines(data["gaps"]),
        "flags": flags_table(data), "flags_colors": flags_colors(data),
        "metrics": metrics_table(data), "metrics_colors": metrics_colors(data),
        "deck_name": deck.name, "deck_bytes": deck.read_bytes(),
        "excel_name": excel_path.name, "excel_bytes": excel_path.read_bytes(),
        "ai_note": ai_note,
    }


def build_outputs(file_name, file_bytes, include_ai, saved_dir=OUTPUT_DIR, client=None):
    """Everything the page shows for one upload, or {"error": a plain message}. Never raises.

    saved_dir is where a reusable analysis may be (main.py's output/). client is for tests.
    The temporary folder is deleted afterwards: the files live on as bytes for the download buttons.
    """
    with tempfile.TemporaryDirectory() as folder:
        try:
            return build_in_folder(Path(folder), file_name, file_bytes, include_ai, saved_dir, client)
        except NOT_A_WORKBOOK_ERRORS:
            return {"error": NOT_A_WORKBOOK}
        except INPUT_ERRORS as error:        # clean.py's message says what's wrong and where
            return {"error": str(error)}
        except Exception as error:  # noqa: BLE001 - a bug: plain words on the page, details in the terminal
            traceback.print_exc()
            return {"error": f"{UNEXPECTED_ERROR_START} ({type(error).__name__}: {error})"}


# ---------------------------------------------------------------------------
# The page
# ---------------------------------------------------------------------------

def styled(table, colors):
    """The table with each cell styled by the matching cell of colors."""
    return table.style.apply(lambda _: colors, axis=None)


def show_downloads(result):
    """Two buttons, side by side: the deck and the metrics workbook."""
    left, right = st.columns(2)
    left.download_button("Download the deck (.pptx)", result["deck_bytes"], file_name=result["deck_name"],
                         mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    right.download_button("Download the metrics workbook (.xlsx)", result["excel_bytes"],
                          file_name=result["excel_name"],
                          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def show_result(result):
    """The error message, or the downloads, flags, data gaps and metrics."""
    if result["error"]:
        st.error(result["error"])
        return
    st.header(f"{result['company']} — {result['quarter']}")
    show_downloads(result)
    st.info(result["ai_note"])
    st.subheader(result["flag_count"])
    st.caption(LEGEND)
    st.dataframe(styled(result["flags"], result["flags_colors"]), hide_index=True, width="stretch")
    st.markdown("**Data gaps**\n\n" + "\n".join(f"- {line}" for line in result["gaps"]))
    st.subheader("Metrics")
    st.dataframe(styled(result["metrics"], result["metrics_colors"]), width="stretch")


def main():
    st.set_page_config(page_title=PAGE_TITLE, layout="wide")
    st.title(PAGE_TITLE)
    st.write(INTRO)
    include_ai = st.checkbox(ai_checkbox_label())
    upload = st.file_uploader("KPI workbook (.xlsx)", type=["xlsx"])
    if upload is None:
        return
    key = (upload.file_id, include_ai)  # build once per file and choice, not on every click
    if st.session_state.get("built_for") != key:
        with st.spinner("Building the board pack..."):
            st.session_state["result"] = build_outputs(upload.name, upload.getvalue(), include_ai)
        st.session_state["built_for"] = key
    show_result(st.session_state["result"])


if __name__ == "__main__":  # Streamlit runs this file as __main__; importing it (tests) shows nothing
    main()
