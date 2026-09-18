"""The portfolio behind the web page (app.py): every company in data/, the state of its files, and what the buttons do.

There is no Streamlit here, so every function runs (and is tested) without a browser.
- A row per workbook in data/: latest quarter, flags tripped and data gaps (worked out from the
  workbook now, by metrics.py), last run and deck status (read from output/<company>_manifest.json).
- Generate: main.run_company into output/, the same files `python main.py` writes for one company.
- Add a company: an uploaded workbook goes into data/ only if clean.py can read it.
- Review mapping (Task 5): headers clean.py doesn't know get mapping.py's proposals; a person's
  confirmed choices are saved to mappings/<company>.yaml, only once the workbook reads with them.
- Approve: approve.approve, exactly what `python approve.py` records.
- What changed since the last run (Task 10): diff_runs.py's comparison of today's workbook with
  the last run on record whose results were different.

Rules:
- No new math and no numbers of its own: every figure comes from metrics.py through main.py and build_deck.py.
- Never raises to the page: a problem becomes plain words (clean.py's own message for a workbook it
  can't read), never a traceback. One company failing never stops the others.
- AI: a saved analysis of exactly these numbers is always reused (free). Claude is asked only when
  the page's box is ticked.
- No em dash in anything the page shows (plain()).
"""

import json
import re
import tempfile
import traceback
import zipfile
from pathlib import Path

from openpyxl.utils.exceptions import InvalidFileException

from analyze import BoardSummary
from approve import approve
from build_deck import PLACEHOLDER_TEXT, collect_deck_data, deck_path, flag_count_text
from clean import UnconfirmedMappingError, clean_workbook
from diff_runs import changes_since_last_run, move_settings
from excel_output import output_path as excel_path
from main import (AI_FAILED, AI_REUSED, AI_SKIPPED, DATA_DIR, INPUT_ERRORS, OUTPUT_DIR, SUMMARY_CSV_PATH,
                  api_key_problem, blank_quarters, company_name, find_workbooks, gaps_text, reusable_analysis,
                  run_company, write_summary_csv)
from mapping import mapping_sha256, review_workbook, save_mapping, saved_columns
from memo import memo_paths, no_em_dash
from metrics import CONFIG_PATH
from provenance import approval_status, deck_status, file_sha256, manifest_path, read_manifest

# Files that aren't Excel workbooks at all: clean.py never gets to read them.
NOT_A_WORKBOOK = "This file isn't a readable Excel workbook. Save it from Excel as .xlsx and try again."
NOT_A_WORKBOOK_ERRORS = (zipfile.BadZipFile, InvalidFileException)
EXCEL_WORKBOOK_PART = "xl/workbook.xml"   # inside every .xlsx (which is a zip file)
UNEXPECTED_ERROR_START = "Something unexpected went wrong - please send this file to whoever looks after the tool"

# What a row says before or instead of a number.
NO_VALUE = "-"
NEVER_RUN = "never"
NOT_GENERATED = "Not generated yet"
OUT_OF_DATE = "Out of date: {what} has changed since the last run. Generate again."

# Search and Add a company.
NO_WORKBOOK_FOUND = 'No KPI workbook found for "{name}". Upload one under "Add a company".'
NAME_PATTERN = re.compile(r"[A-Za-z][A-Za-z0-9 -]{0,39}")   # a letter first, then up to 39 more
NAME_RULE = ("A company name starts with a letter and may have letters, digits, spaces and hyphens "
             "(at most 40 characters).")
ALREADY_THERE = 'A workbook for {company} is already in data/. Tick "Replace its workbook" to swap it for this one.'
ADDED = "Added {company} (latest quarter {quarter}). Click Generate on its row to build its files."

# Review mapping.
MAPPING_NEEDED = ("{count} column header{s} aren't known input columns: {headers}. Nothing is guessed: open "
                  "the company's page and confirm or change what each one means under Review mapping.")
MAPPING_INCOMPLETE = "Choose a column for every header first: {headers}."
MAPPING_SAVED = ("Saved the column mapping for {company} ({count} headers). Next quarter's workbook with the "
                 "same headers uses it without asking.")

# Approve.
REVIEWER_NEEDED = "Type the reviewer's name first: an approval needs a person's name on it."
NOT_APPROVABLE = ("Not approved: a reviewer approves files built from today's workbook and config.yaml "
                  "(now: {status})")
APPROVED = ("Approved {company} by {reviewer} on {when}. Click Generate to rebuild the deck and memo, "
            "so their footers say reviewed.")

# What happened to the AI commentary on Generate.
AI_REUSED_NOTE = "AI commentary reused from a saved analysis of exactly these numbers (no API call). Review it before use."
AI_NEW_NOTE = "AI commentary written by Claude and checked. Review it before use."
AI_NOT_ASKED_NOTE = (f"No AI commentary: no saved analysis matches these numbers and the AI box isn't ticked, "
                     f"so slide 4 says \"{PLACEHOLDER_TEXT}\".")
NO_KEY_NOTE = "No AI commentary: no API key is set (ANTHROPIC_API_KEY in .env), so Claude wasn't asked."


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------

def plain(text):
    """Text as the page shows it: no em dash ('Cannot evaluate: missing input'). The memo's own rule."""
    return no_em_dash(str(text))


def last_run_text(run_at):
    """'2026-09-17T14:03:11' -> '2026-09-17 14:03'; no run -> 'never'."""
    return run_at.replace("T", " ")[:16] if run_at else NEVER_RUN


def is_excel_workbook(path):
    """True if the file is a zip holding an Excel workbook part (every .xlsx has xl/workbook.xml).

    Being a zip isn't enough: a .pptx or .docx renamed .xlsx is a zip too, and pandas' error on
    one is cryptic.
    """
    if not zipfile.is_zipfile(path):
        return False
    with zipfile.ZipFile(path) as archive:
        return EXCEL_WORKBOOK_PART in archive.namelist()


def quoted(headers):
    """['GP', 'FTEs'] -> '"GP", "FTEs"' (header names as a person reads them in a sentence)."""
    return ", ".join(f'"{header}"' for header in headers)


def mapping_needed(proposals):
    """The page's words for an unconfirmed mapping: which headers, and where to confirm them.

    Not clean.py's own message: that one says how to confirm on the command line, and for an
    upload it would name a temporary file.
    """
    count = len(proposals)
    return MAPPING_NEEDED.format(count=count, s="s" if count != 1 else "", headers=quoted(p.header for p in proposals))


def error_message(error):
    """Plain words for an error. A bug also prints its traceback in the Terminal window, never on the page."""
    if isinstance(error, NOT_A_WORKBOOK_ERRORS):
        return NOT_A_WORKBOOK
    if isinstance(error, UnconfirmedMappingError):
        return plain(mapping_needed(error.proposals))
    if isinstance(error, INPUT_ERRORS):     # clean.py's message says what's wrong and where
        return plain(error)
    traceback.print_exception(error)
    return plain(f"{UNEXPECTED_ERROR_START} ({type(error).__name__}: {error})")


def load_company(workbook_path, config, mappings_dir=None):
    """(everything the deck shows, None), or (None, why the workbook can't be read). Never raises.

    mappings_dir: where the confirmed column mappings are (None: mappings/). Review mapping passes a
    temporary one, to try a mapping before saving it.
    """
    try:
        if not is_excel_workbook(workbook_path):  # else pandas' message is cryptic
            raise ValueError(NOT_A_WORKBOOK)
        actuals, next_budget = clean_workbook(workbook_path, mappings_dir)
        return collect_deck_data(company_name(workbook_path), Path(workbook_path).name, actuals, next_budget,
                                 config), None
    except Exception as error:  # noqa: BLE001 - the page shows plain words instead
        return None, error_message(error)


# ---------------------------------------------------------------------------
# Where each company's files are, and whether they match today's inputs
# ---------------------------------------------------------------------------

def output_files(workbook_path, output_dir=OUTPUT_DIR):
    """The files Generate writes for one company, by the name the page's buttons use."""
    memo_docx, memo_pdf = memo_paths(workbook_path, output_dir)
    return {"deck": deck_path(workbook_path, output_dir), "memo": memo_pdf, "memo_docx": memo_docx,
            "excel": excel_path(workbook_path, output_dir)}


def run_state(workbook_path, output_dir=OUTPUT_DIR, config_path=CONFIG_PATH):
    """{"last_run", "status", "current"} from the company's manifest.

    current = the files were built from today's workbook, config.yaml and column mapping (their
    hashes match), so they may be downloaded and approved. The status is provenance.py's, so the page, main.py and
    approve.py can't disagree about whether a deck is reviewed.
    """
    manifest = read_manifest(manifest_path(workbook_path, output_dir))
    if manifest is None:
        return {"last_run": NEVER_RUN, "status": NOT_GENERATED, "current": False}
    last_run = last_run_text(manifest.get("run_at"))
    input_hash, config_hash = file_sha256(workbook_path), file_sha256(config_path)
    if (manifest.get("input") or {}).get("sha256") != input_hash:
        return {"last_run": last_run, "status": OUT_OF_DATE.format(what="the workbook"), "current": False}
    if (manifest.get("config") or {}).get("sha256") != config_hash:
        return {"last_run": last_run, "status": OUT_OF_DATE.format(what="config.yaml"), "current": False}
    mapping_hash = mapping_sha256(workbook_path)
    if (manifest.get("mapping") or {}).get("sha256") != mapping_hash:
        return {"last_run": last_run, "status": OUT_OF_DATE.format(what="the column mapping"), "current": False}
    approval, _ = approval_status(manifest, input_hash, config_hash, mapping_hash)
    return {"last_run": last_run, "status": deck_status(approval), "current": True}


def download(workbook_path, output_dir, kind, current):
    """(file name, bytes) for a download button, or None if the file is missing or out of date.

    kind: "deck", "memo" (PDF), "memo_docx" or "excel". An out-of-date file is never offered: it
    shows last run's numbers, not today's workbook.
    """
    path = output_files(workbook_path, output_dir)[kind]
    if not current or not path.exists():
        return None
    return path.name, path.read_bytes()


# ---------------------------------------------------------------------------
# The portfolio table and search
# ---------------------------------------------------------------------------

def gaps_summary(data):
    """'none', or '19 metrics/flags (blank: Q1 2025)': main.py's summary-table words."""
    return gaps_text({"gap_count": len(data["gaps"]), "blank_quarters": blank_quarters(data["actuals"])})


def company_row(workbook_path, config, output_dir=OUTPUT_DIR):
    """One row of the portfolio table. A workbook clean.py can't read shows "-" and clean.py's message."""
    workbook_path = Path(workbook_path)
    data, problem = load_company(workbook_path, config)
    row = {"stem": workbook_path.stem, "company": company_name(workbook_path), "workbook": workbook_path,
           "problem": problem, "latest": NO_VALUE, "flags": NO_VALUE, "gaps": NO_VALUE,
           **run_state(workbook_path, output_dir)}
    if data is not None:
        row.update(latest=data["latest"], flags=flag_count_text(data["flags"]), gaps=gaps_summary(data))
    return row


def portfolio_rows(config, data_dir=DATA_DIR, output_dir=OUTPUT_DIR):
    """One row per workbook in data/, sorted by file name."""
    return [company_row(path, config, output_dir) for path in find_workbooks(data_dir)]


def search_rows(rows, query):
    """The rows whose company name contains the query, ignoring case. An empty query keeps every row."""
    words = query.strip().lower()
    return [row for row in rows if words in row["company"].lower()]


def search_message(rows, query):
    """'No KPI workbook found for "X". Upload one ...' when a search matches nothing, else None."""
    name = query.strip()
    if not name or search_rows(rows, name):
        return None
    return NO_WORKBOOK_FOUND.format(name=name)


def find_workbook(stem, data_dir=DATA_DIR):
    """data/<stem>.xlsx if it's one of the portfolio's workbooks, else None (so a name can't point outside data/)."""
    return next((path for path in find_workbooks(data_dir) if path.stem == stem), None)


# ---------------------------------------------------------------------------
# Generate
# ---------------------------------------------------------------------------

def ai_note(result, manifest, no_key):
    """One sentence on what happened to the AI commentary, from main.py's result and the manifest it wrote."""
    validation = ((manifest or {}).get("ai") or {}).get("validation") or ""
    if result["ai"] == AI_SKIPPED:
        return NO_KEY_NOTE if no_key else AI_NOT_ASKED_NOTE
    if result["ai"] == AI_FAILED:
        return plain(f"{PLACEHOLDER_TEXT}: the AI step {validation}")
    return AI_REUSED_NOTE if validation == AI_REUSED else AI_NEW_NOTE


def generate_company(workbook_path, config, ask_claude, output_dir=OUTPUT_DIR, client=None):
    """Build one company's files, as `python main.py <workbook>` does. Never raises.

    Returns {"company", "ok", "message" (for the page), "result" (main.py's, for batch_summary.csv)}.
    ask_claude=False never calls the API. client is for tests (a fake Claude client).
    """
    workbook_path = Path(workbook_path)
    company = company_name(workbook_path)
    try:
        if not is_excel_workbook(workbook_path):
            raise ValueError(NOT_A_WORKBOOK)
        no_key = bool(ask_claude and client is None and api_key_problem())   # a fake client needs no key
        result = run_company(workbook_path, config, skip_ai=not ask_claude or no_key, client=client,
                             output_dir=output_dir, reuse_saved=True)
    except Exception as error:  # noqa: BLE001 - one company failing must not stop the others
        message = error_message(error)
        return {"company": company, "ok": False, "message": f"{company}: {message}",
                "result": {"company": company, "error": f"{type(error).__name__}: {message}"}}
    result["error"] = None
    manifest = read_manifest(manifest_path(workbook_path, output_dir))
    return {"company": company, "ok": True, "result": result,
            "message": f"{company}: generated. {ai_note(result, manifest, no_key)}"}


def generate_all(config, ask_claude, data_dir=DATA_DIR, output_dir=OUTPUT_DIR, client=None, on_progress=None):
    """Generate every company in data/ in turn, then save batch_summary.csv as main.py --all does.

    on_progress(done, total, company about to start or None at the end) moves the page's progress bar.
    """
    paths = find_workbooks(data_dir)
    outcomes = []
    for done, path in enumerate(paths):
        if on_progress:
            on_progress(done, len(paths), company_name(path))
        outcomes.append(generate_company(path, config, ask_claude, output_dir, client))
    if on_progress:
        on_progress(len(paths), len(paths), None)
    write_summary_csv([outcome["result"] for outcome in outcomes], Path(output_dir) / SUMMARY_CSV_PATH.name)
    return outcomes


# ---------------------------------------------------------------------------
# Add a company
# ---------------------------------------------------------------------------

def company_stem(name):
    """The file name (without .xlsx) for a company name: '  Blue  River ' -> 'blue river'.

    Only letters, digits, spaces and hyphens, so a name can't reach outside data/ ("../x").
    """
    words = " ".join(str(name).split())
    if not NAME_PATTERN.fullmatch(words):
        raise ValueError(NAME_RULE)
    return words.lower()


def suggested_name(file_name):
    """A starting company name from the upload's file name: other characters become spaces, 40 at most."""
    words = re.sub(r"[^A-Za-z0-9 -]", " ", Path(file_name).stem)
    return " ".join(words.split())[:40]


def add_company(file_name, file_bytes, name, config, data_dir=DATA_DIR, replace=False, columns=None):
    """Save an uploaded workbook as data/<name>.xlsx, only if clean.py can read it. Never raises.

    Returns {"ok", "message"}: clean.py's own message for a workbook it can't read. An existing
    company's workbook is replaced only when asked (its old files then show as out of date).
    columns = the confirmed {header: column} from Review mapping, or None: they are saved to
    mappings/<name>.yaml together with the workbook, and neither is saved unless the workbook reads.
    """
    try:
        stem = company_stem(name)
    except ValueError as error:
        return {"ok": False, "message": str(error)}
    target = Path(data_dir) / f"{stem}.xlsx"
    if target.exists() and not replace:
        return {"ok": False, "message": ALREADY_THERE.format(company=company_name(target))}
    with tempfile.TemporaryDirectory() as folder:   # check it before it goes anywhere near data/
        upload = Path(folder) / target.name
        upload.write_bytes(file_bytes)
        data, problem = try_mapping(upload, columns, config) if columns else load_company(upload, config)
    if problem:
        return {"ok": False, "message": problem}
    message = ADDED.format(company=company_name(target), quarter=data["latest"])
    if columns:
        save_mapping(target, columns)
        message += " " + MAPPING_SAVED.format(company=company_name(target), count=len(columns))
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_bytes(file_bytes)
    return {"ok": True, "message": message}


# ---------------------------------------------------------------------------
# Review mapping: headers clean.py doesn't know
# ---------------------------------------------------------------------------

def mapping_proposals(workbook_path):
    """mapping.py's proposals for the workbook's unknown headers; [] if there are none or it can't be read.

    A workbook that can't be read for another reason shows that reason instead (load_company).
    """
    try:
        return review_workbook(workbook_path) if is_excel_workbook(workbook_path) else []
    except (*INPUT_ERRORS, *NOT_A_WORKBOOK_ERRORS):
        return []


def upload_proposals(file_bytes, name):
    """The proposals for an uploaded workbook, as it would be saved under this company name.

    A mapping already confirmed for that name counts, so a new quarter's upload asks only about new
    headers. A bad name gives [] (add_company then says what's wrong with it).
    """
    try:
        stem = company_stem(name)
    except ValueError:
        return []
    with tempfile.TemporaryDirectory() as folder:
        upload = Path(folder) / f"{stem}.xlsx"
        upload.write_bytes(file_bytes)
        return mapping_proposals(upload)


def try_mapping(workbook_path, columns, config):
    """(data, None) if the workbook reads with these confirmed columns, else (None, why). Saves nothing.

    columns = {header: chosen column}. Every unknown header needs a column. The try runs on a
    temporary mappings folder holding the company's saved mapping plus these choices.
    """
    missing = [p.header for p in mapping_proposals(workbook_path) if not columns.get(p.header)]
    if missing:
        return None, MAPPING_INCOMPLETE.format(headers=quoted(missing))
    try:
        saved = saved_columns(workbook_path)
    except ValueError as error:   # a hand-edited mappings file that doesn't read
        return None, plain(error)
    with tempfile.TemporaryDirectory() as folder:
        save_mapping(workbook_path, {**saved, **columns}, mappings_dir=folder)
        return load_company(workbook_path, config, mappings_dir=folder)


def confirm_mapping(workbook_path, columns, config):
    """Save a person's confirmed columns for a company already in data/, if the workbook then reads. Never raises."""
    data, problem = try_mapping(workbook_path, columns, config)
    if problem:
        return {"ok": False, "message": problem}
    save_mapping(workbook_path, columns)
    return {"ok": True, "message": MAPPING_SAVED.format(company=company_name(workbook_path), count=len(columns))}


# ---------------------------------------------------------------------------
# Approve, and the company page's AI commentary
# ---------------------------------------------------------------------------

def approve_company(stem, reviewer, data_dir=DATA_DIR, output_dir=OUTPUT_DIR, now=None):
    """Record the reviewer in the manifest with approve.approve (what approve.py does). Never raises.

    The page asks for a typed name instead of falling back to git's user name: whoever is at the
    browser must say who they are. Nothing is built: approving and producing stay separate acts.
    """
    name = (reviewer or "").strip()
    if not name:
        return {"ok": False, "message": REVIEWER_NEEDED}
    workbook_path = Path(data_dir) / f"{stem}.xlsx"
    state = run_state(workbook_path, output_dir)
    if not state["current"]:
        return {"ok": False, "message": NOT_APPROVABLE.format(status=state["status"].rstrip("."))}
    try:
        manifest = approve(stem, name, data_dir, output_dir, now=now)
    except (ValueError, OSError) as error:
        return {"ok": False, "message": plain(f"Not approved: {error}")}
    approval = manifest["approval"]
    return {"ok": True, "message": APPROVED.format(company=manifest["company"], reviewer=approval["reviewer"],
                                                   when=last_run_text(approval["approved_at"]))}


def run_changes(workbook_path, data, config, output_dir=OUTPUT_DIR):
    """(what changed since the last run, the move settings, None), or (None, None, why not). Never raises.

    The report is diff_runs.changes_since_last_run's: today's workbook against the last run with
    different results, so it's the same comparison the memo built from these numbers shows. None
    when there's no earlier run on record. A bad diff_min_* setting in config.yaml is plain words.
    """
    try:
        settings = move_settings(config)
        manifest = read_manifest(manifest_path(workbook_path, output_dir))
        return changes_since_last_run(data, manifest, settings), settings, None
    except Exception as error:  # noqa: BLE001 - plain words on the page, never a traceback
        return None, None, error_message(error)


def saved_commentary(workbook_path, config, output_dir=OUTPUT_DIR):
    """The saved AI commentary (a BoardSummary) if it was made from exactly these numbers and still passes, else None."""
    path = reusable_analysis(workbook_path, output_dir, config)
    if path is None:
        return None
    return BoardSummary.model_validate(json.loads(path.read_text())["summary"])
