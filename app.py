"""The web page for people who don't use the command line: the whole portfolio, and a page per company.

Two pages:
1. Portfolio: a table of every company in data/ (latest quarter, flags tripped, data gaps, last run,
   deck status), with Generate, Download deck, Download memo and Download Excel on each row; a
   search box; Generate all with a progress bar; Download rollup (Task 9: rollup.py's one deck and
   one workbook across every company, built when clicked); an "Add a company" panel for a new workbook;
   and Recent runs (Task 15: the newest five run logs from output/logs, each step with its seconds,
   result and any error, whether the run came from this page or the command line).
2. Company (click a company's name): its flags with thresholds and reasons, what changed since the
   last run (Task 10: flags that flipped, metrics that moved, new and resolved data gaps), data gaps, metrics
   table in the status colors (red = tripped, green = passed, gray = data missing / cannot
   evaluate), both charts, the AI commentary when a saved one matches these numbers, and the
   buttons Generate, the downloads, Export and Approve. Export (Task 11) holds export.py's metrics and
   flags as CSV and JSON and the email summary, built from today's workbook when clicked.

Review mapping (Task 5): a workbook with headers clean.py doesn't know shows, on its company page
and in "Add a company", each header with mapping.py's proposed column, how sure it is, why, and the
first values under it. Each pair has Change (pick another column) and Confirm; nothing is saved
until every pair is confirmed, and then only if the workbook reads with it.

Look (theme.py, final Task 3): one column about 1100 px wide, white cards on the surface color,
tables with a navy header and 40 px rows, Arial. One primary (navy) button per page: Generate all
on the portfolio, Generate on a company's page. Every other button is white with a navy border;
none is ever red or green, since those colors mean a flag's status.

Rules:
- No new math: every number and every word comes from metrics.py, excel_output.py, build_deck.py
  and main.py. portfolio.py does the work behind the buttons; this file only draws.
- Generate writes to output/: the same files and manifest `python main.py` writes, so the table,
  approve.py and the command line all see the same state.
- A workbook clean.py can't read shows clean.py's own message, never a traceback; one company
  failing never stops the others. No em dash on the page (portfolio.plain).
- AI commentary: a saved analysis of exactly the same numbers is always reused (free). Claude is
  asked only when the box is ticked, which costs money - the checkbox says roughly how much.

Run: streamlit run app.py      (or double-click run_app.command on a Mac)
"""

from html import escape

import pandas as pd
import streamlit as st
from matplotlib import pyplot as plt

from analyze import grouped_questions
from build_deck import AI_DRAFTED_LINE, QUESTIONS_HEADING, flag_count_text, gaps_lines, points_text, runway_lines, \
    threshold_text, value_text
from charts import arr_chart, cash_chart
from diff_runs import HEADING, NO_EARLIER_RUN, change_sections, compared_with_text
from excel_output import status_label, tripped_cells
from export import export_bytes, export_paths
from main import DATA_DIR, OUTPUT_DIR, company_name
from mapping import confidence_text
from metrics import CANNOT_EVALUATE, METRIC_LABELS, MISSING_INPUT, TRIP, load_config
from portfolio import (NO_WORKBOOK_FOUND, add_company, approve_company, confirm_mapping, download, error_message,
                       find_workbook, generate_all, generate_company, load_company, mapping_proposals, plain,
                       portfolio_rows, run_changes, run_state, saved_commentary, search_message, search_rows,
                       suggested_name, upload_proposals)
from rollup import rollup_download, rollup_paths
from run_log import recent_runs, run_label, step_rows
from theme import STATUS_COLORS, streamlit_css

PAGE_TITLE = "Board Pack Generator"
INTRO = ("Every portfolio company with a KPI workbook in data/. Click a name for its flags, metrics and "
         "charts; Generate builds its deck, memo and metrics workbook. Fictional data only.")

# Typical cost of the AI step, from README.md's Cost section (latest live run: $0.0911 and about
# 70 seconds per company with claude-sonnet-5). A label, not a calculation: update it with the README.
TYPICAL_AI_COST_USD = 0.09
TYPICAL_AI_SECONDS = 70
AI_CAPTION = ("Unticked, nothing is sent to Claude. Either way, a saved analysis of exactly the same numbers "
              "is reused for free.")

LEGEND = "Red = tripped · green = passed · gray = data missing or cannot evaluate"
FLAG_COLUMNS = ["Flag", "Value", "Threshold", "Status"]
NO_COMPANIES = "No KPI workbooks in data/ yet. Add one below."
NOT_AVAILABLE = "Generate first: there is no file built from today's workbook."
NO_COMMENTARY = ("No AI commentary: no saved analysis matches these numbers. Tick the AI box and click Generate "
                 "to ask Claude.")
ADD_INTRO = ("Drop in a company's KPI workbook (.xlsx). It is added to data/ only if it can be read; "
             "otherwise you'll see what to fix.")
REVIEW_INTRO = ("These headers aren't known input columns. Each shows the column it most likely means, how sure "
                "that is and why, and the first values under it. Confirm each pair, or Change the column and "
                "then confirm. Nothing is used until every pair is confirmed; the mapping is then saved, so next "
                "quarter's workbook with the same headers runs without asking.")
CONFIRM_ALL_FIRST = "Confirm every pair first"
APPROVE_INTRO = ("Approving records your name against today's workbook, as approve.py does. It builds nothing: "
                 "click Generate afterwards so the deck and memo footers say reviewed.")

# The portfolio table: a column per piece of a row, widths relative to each other. The last two
# columns hold Generate and one "Download" button that opens the three downloads (four buttons
# side by side don't fit a 1100 px page beside six columns of text).
TABLE_HEADERS = ["Company", "Latest quarter", "Flags tripped", "Data gaps", "Last run", "Deck status", "", ""]
ROW_TEXT_KEYS = ["latest", "flags", "gaps", "last_run", "status"]
ROW_WIDTHS = [1.3, 1.0, 1.5, 1.8, 1.3, 1.8, 1.1, 1.2]

# The download buttons: (portfolio.output_files name, label, file type for the browser).
PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
ROW_DOWNLOADS = [("deck", "Download deck", PPTX), ("memo", "Download memo", "application/pdf"),
                 ("excel", "Download Excel", XLSX)]
ROLLUP_DOWNLOADS = [("deck", "Download rollup deck", PPTX), ("excel", "Download rollup Excel", XLSX)]
ROLLUP_HELP = ("One deck and one workbook across every company: ranked by flags tripped, each company's worst "
               "flag, companies by status and a runway chart. Built from today's workbooks when you click; "
               "no AI, and nothing is written to output/.")
PAGE_DOWNLOADS = [("deck", "Download deck", PPTX), ("memo", "Download memo (PDF)", "application/pdf"),
                  ("memo_docx", "Download memo (Word)", DOCX), ("excel", "Download Excel", XLSX)]
# Task 11: the company page's Export popover: (export.EXPORT_KINDS name, label, file type for the browser).
EXPORT_DOWNLOADS = [("metrics_csv", "Metrics (CSV)", "text/csv"), ("flags_csv", "Flags (CSV)", "text/csv"),
                    ("json", "Metrics and flags (JSON)", "application/json"),
                    ("email", "Email summary (HTML)", "text/html")]
EXPORT_HELP = ("Metrics and flags for other tools (CSV, JSON), and an email summary: open the HTML file in a "
               "browser, select all, copy, and paste into Outlook. Built from today's workbook when you click; "
               "no AI text, and nothing is written to output/.")

# Task 15: the portfolio's Recent runs card, read from output/logs (run_log.py).
NO_RUNS = "No runs yet. Generate a company here, or run python main.py, and each step will be listed."
RUNS_CAPTION = ("The newest five runs from here or the command line, newest first. Open one for each company's "
                "steps, how long each took and any error. The full logs are in output/logs.")
RUN_COLUMNS = ["Company", "Step", "Seconds", "Result", "Error"]

CHART_INCHES = (6.4, 4.4)       # each chart's size on the page
COMPANY_KEY = "company"         # st.session_state: the company page being shown, or none (the portfolio)
MESSAGES_KEY = "messages"       # st.session_state: what the last button did, shown once


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------

def ai_checkbox_label():
    """'Ask Claude for AI commentary when no saved analysis matches (typically about $0.09 and 70 seconds per company)'."""
    return (f"Ask Claude for AI commentary when no saved analysis matches (typically about "
            f"${TYPICAL_AI_COST_USD:.2f} and {TYPICAL_AI_SECONDS} seconds per company)")


def status_css(status):
    """A status's colors (theme.py: red, green or gray, each on its light fill) as a table cell's style."""
    fill, text = STATUS_COLORS[status]
    return f"background-color: #{fill}; color: #{text}"


def md(text):
    """Text for st.markdown: no em dash, and "$" kept as a dollar sign (two of them would start a formula)."""
    return plain(text).replace("$", "\\$")


# ---------------------------------------------------------------------------
# The company page's tables and charts (formatting only - no math)
# ---------------------------------------------------------------------------

def metrics_table(data):
    """One row per metric, one column per quarter, as the same text the deck shows."""
    quarters = list(data["metrics"].index)
    rows = {METRIC_LABELS[metric]: [plain(value_text(data, metric, quarter)) for quarter in quarters]
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
    """One flag as [Flag, Value, Threshold, Status]; the status carries the reason it can't be evaluated."""
    if flag["metric"] is None:  # the combo rule has no single value
        return [flag["flag"], "see NRR and pipeline", combo_rule_text(data["config"]), plain(status_label(flag))]
    return [flag["flag"], plain(value_text(data, flag["metric"], flag["quarter"])), threshold_text(flag),
            plain(status_label(flag))]


def flags_table(data):
    """Every flag for the latest quarter, in the order metrics.py evaluates them."""
    return pd.DataFrame([flag_row(data, flag) for flag in data["flags"]], columns=FLAG_COLUMNS)


def flags_colors(data):
    """Each flag's whole row in its status color, like the Excel Flags sheet."""
    return pd.DataFrame([[status_css(flag["status"])] * len(FLAG_COLUMNS) for flag in data["flags"]],
                        columns=FLAG_COLUMNS)


def chart_figures(data):
    """The deck's two charts (ARR with net new ARR; ending cash with runway), drawn by charts.py."""
    quarters = list(data["metrics"].index)
    return [arr_chart(quarters, data["metrics"]["ending_arr"], data["metrics"]["net_new_arr"], CHART_INCHES),
            cash_chart(quarters, data["actuals"]["ending_cash"], runway_lines(data), CHART_INCHES)]


def cell_html(tag, text, style=""):
    """One table cell; the text is escaped, so a "<" in a company's data can never become markup."""
    style_part = f' style="{style}"' if style else ""
    return f"<{tag}{style_part}>{escape(str(text))}</{tag}>"


def table_html(table, colors, index_header=None):
    """A table as HTML in theme.py's look (class "bp-table": navy header, white text, 40 px rows).

    colors holds each cell's style (same shape as table). index_header names a first column
    holding the table's row names (the metrics table's metric labels); None leaves them out.
    """
    header = ([index_header] if index_header else []) + list(table.columns)
    rows = []
    for (name, values), (_, styles) in zip(table.iterrows(), colors.iterrows()):
        cells = [cell_html("td", name)] if index_header else []
        cells += [cell_html("td", value, style) for value, style in zip(values, styles)]
        rows.append("<tr>" + "".join(cells) + "</tr>")
    head = "<tr>" + "".join(cell_html("th", text) for text in header) + "</tr>"
    return (f'<div class="bp-table-wrap"><table class="bp-table"><thead>{head}</thead>'
            f'<tbody>{"".join(rows)}</tbody></table></div>')


# ---------------------------------------------------------------------------
# Pieces both pages use
# ---------------------------------------------------------------------------

def say(ok, text):
    """Remember a button's outcome, to show at the top of the page after it redraws."""
    st.session_state.setdefault(MESSAGES_KEY, []).append((ok, plain(text)))


def show_messages():
    """Show (once) what the last button did: green for done, red for a problem."""
    for ok, text in st.session_state.pop(MESSAGES_KEY, []):
        (st.success if ok else st.error)(text)


def ai_checkbox():
    """The AI box (same on both pages) and what ticking it means."""
    ask_claude = st.checkbox(ai_checkbox_label(), key="ask_claude")
    st.caption(AI_CAPTION)
    return ask_claude


def download_button(workbook, output_dir, current, kind, label, mime, key):
    """A download button for one of the company's files; greyed out if it's missing or out of date."""
    found = download(workbook, output_dir, kind, current)
    if found is None:
        st.download_button(label, b"", disabled=True, help=NOT_AVAILABLE, key=key)
    else:
        st.download_button(label, found[1], file_name=found[0], mime=mime, key=key)


def generate_button(workbook, ask_claude, config, output_dir, key, primary=False):
    """Generate one company's files, then redraw the page with what happened. primary=True: the page's navy button."""
    if st.button("Generate", key=key, type="primary" if primary else "secondary"):
        with st.spinner(f"Generating {company_name(workbook)}..."):
            outcome = generate_company(workbook, config, ask_claude, output_dir)
        say(outcome["ok"], outcome["message"])
        st.rerun()


def page_config():
    """config.yaml's thresholds, or None after showing why they can't be read."""
    try:
        return load_config()
    except Exception as error:  # noqa: BLE001 - plain words on the page, never a traceback
        st.error(error_message(error))
        return None


# ---------------------------------------------------------------------------
# Page 1: the portfolio
# ---------------------------------------------------------------------------

def open_company(stem):
    """Show a company's page on the next redraw."""
    st.session_state[COMPANY_KEY] = stem
    st.rerun()


def pair_summary(proposal):
    """The header, where it is, the proposal and its confidence, and the values under it (markdown lines)."""
    proposed = (f"Proposed {proposal.column}, confidence {confidence_text(proposal.confidence)}"
                if proposal.column else "No proposal")
    values = ", ".join(proposal.samples) or "none"
    return md(f"**{proposal.header}** (cell {proposal.cell})  \n{proposed}  \nValues: {values}")


def pair_reason(proposal, column):
    """Why the column was proposed, and whether the person changed it."""
    reason = f"Why: {proposal.reason}"
    return reason if column == proposal.column else f"Changed from the proposal. {reason}"


def mapping_pair(proposal, key):
    """One header and its column: the summary, Change (a list of columns) and Confirm. Returns (column, confirmed).

    Confirm's key includes the chosen column, so changing the column clears its tick: a changed
    pair has to be confirmed again.
    """
    with st.container(key=f"mapping-{key}"):
        about, change, confirm = st.columns([2.4, 1.6, 0.8], vertical_alignment="center")
        about.markdown(pair_summary(proposal))
        index = proposal.choices.index(proposal.column) if proposal.column else None
        column = change.selectbox("Change", proposal.choices, index=index, placeholder="Choose a column",
                                  key=f"map_{key}")
        confirmed = confirm.checkbox("Confirm", key=f"confirm_{key}_{column}", disabled=column is None)
        st.caption(md(pair_reason(proposal, column)))
    return column, confirmed


def review_mapping(proposals, key):
    """The Review mapping step: every pair. Returns ({header: chosen column}, every pair confirmed)."""
    st.subheader("Review mapping")
    st.caption(REVIEW_INTRO)
    columns, confirmed = {}, []
    for number, proposal in enumerate(proposals):
        column, ticked = mapping_pair(proposal, f"{key}_{number}")
        columns[proposal.header] = column
        confirmed.append(ticked)
    return columns, all(confirmed)


def portfolio_row(row, ask_claude, config, output_dir):
    """One company: its name (click to open), five text cells, Generate, and Download (deck, memo, Excel)."""
    with st.container(key=f"portfolio-row-{row['stem']}"):   # theme.py: 40 px tall, a line under it
        cells = st.columns(ROW_WIDTHS, vertical_alignment="center")
        if cells[0].button(row["company"], key=f"open_{row['stem']}", type="tertiary"):
            open_company(row["stem"])
        for cell, key in zip(cells[1:6], ROW_TEXT_KEYS):
            cell.text(plain(row[key]))
        with cells[6]:
            generate_button(row["workbook"], ask_claude, config, output_dir, key=f"generate_{row['stem']}")
        with cells[7].popover("Download", key=f"downloads_{row['stem']}"):
            for kind, label, mime in ROW_DOWNLOADS:
                download_button(row["workbook"], output_dir, row["current"], kind, label, mime,
                                key=f"{kind}_{row['stem']}")
        if row["problem"]:
            st.error(f"{row['company']}: {row['problem']}")


def portfolio_table(rows, ask_claude, config, output_dir):
    """The navy header line, then one line per company."""
    with st.container(key="portfolio-header"):
        for cell, text in zip(st.columns(ROW_WIDTHS, vertical_alignment="center"), TABLE_HEADERS):
            cell.markdown(text)
    for row in rows:
        portfolio_row(row, ask_claude, config, output_dir)


def generate_all_button(config, ask_claude, data_dir, output_dir):
    """Generate every company in turn under a progress bar; a failure is reported and the rest carry on."""
    if not st.button("Generate all", key="generate_all", type="primary"):
        return
    bar = st.progress(0.0, text="Starting")

    def progress(done, total, company):
        text = f"Generating {company} ({done + 1} of {total})" if company else f"Done: {total} of {total}"
        bar.progress(done / total if total else 1.0, text=text)

    outcomes = generate_all(config, ask_claude, data_dir, output_dir, on_progress=progress)
    built = sum(outcome["ok"] for outcome in outcomes)
    say(built == len(outcomes), f"Generate all: {built} of {len(outcomes)} companies generated.")
    for outcome in outcomes:
        say(outcome["ok"], outcome["message"])
    st.rerun()


def rollup_file(kind, config, data_dir, output_dir):
    """The function a rollup download button calls when clicked: it builds the file then (rollup.py), not on every redraw.

    A build that fails is printed in the Terminal window; the page says only that the file couldn't be made.
    """
    return lambda: rollup_download(kind, config, data_dir, output_dir)[1]


def rollup_button(config, data_dir, output_dir, has_companies):
    """Download rollup: a popover with the rollup deck and workbook (off when data/ has no workbook)."""
    with st.popover("Download rollup", key="rollup_downloads", help=ROLLUP_HELP, disabled=not has_companies):
        for kind, label, mime in ROLLUP_DOWNLOADS:
            st.download_button(label, rollup_file(kind, config, data_dir, output_dir),
                               file_name=rollup_paths()[kind].name, mime=mime, key=f"rollup_{kind}")


def add_company_panel(config, data_dir, expanded):
    """Upload a new company's workbook; it joins the table only if clean.py can read it."""
    with st.expander("Add a company", expanded=expanded):
        st.write(ADD_INTRO)
        upload = st.file_uploader("KPI workbook (.xlsx)", type=["xlsx"], key="new_workbook")
        if upload is None:
            return
        name = st.text_input("Company name", value=suggested_name(upload.name), key=f"new_name_{upload.file_id}")
        replace = st.checkbox("Replace its workbook", key="replace",
                              help="Only if a company of this name is already in the table")
        proposals = upload_proposals(upload.getvalue(), name)
        columns, ready = review_mapping(proposals, f"new_{upload.file_id}") if proposals else (None, True)
        if st.button("Add company", key="add_company", disabled=not ready, help=None if ready else CONFIRM_ALL_FIRST):
            outcome = add_company(upload.name, upload.getvalue(), name, config, data_dir, replace, columns)
            say(outcome["ok"], outcome["message"])
            st.rerun()


def portfolio_page(data_dir, output_dir):
    """Page 1: search, Generate all, the table, and Add a company, each in a white card."""
    st.title(PAGE_TITLE)
    st.write(INTRO)
    with st.container(key="card-ai"):
        ask_claude = ai_checkbox()
    show_messages()
    config = page_config()
    if config is None:
        return
    rows = portfolio_rows(config, data_dir, output_dir)
    with st.container(key="card-portfolio"):
        search_cell, button_cell, rollup_cell = st.columns([3, 1, 1.2], vertical_alignment="bottom")
        query = search_cell.text_input("Search companies", key="search", placeholder="Company name")
        with button_cell:
            generate_all_button(config, ask_claude, data_dir, output_dir)
        with rollup_cell:
            rollup_button(config, data_dir, output_dir, has_companies=bool(rows))
        not_found = search_message(rows, query)
        if not_found:
            st.warning(not_found)
        elif not rows:
            st.info(NO_COMPANIES)
        else:
            portfolio_table(search_rows(rows, query), ask_claude, config, output_dir)
    with st.container(key="card-add"):
        add_company_panel(config, data_dir, expanded=bool(not_found or not rows))
    with st.container(key="card-runs"):
        recent_runs_panel(output_dir)


def recent_runs_panel(output_dir):
    """Recent runs (Task 15): the newest run logs in output/logs, one expander per run with a line per step."""
    st.subheader("Recent runs")
    runs = recent_runs(output_dir)
    if not runs:
        st.info(NO_RUNS)
        return
    st.caption(RUNS_CAPTION)
    for run in runs:
        with st.expander(md(run_label(run))):
            for problem in run["problems"]:
                st.error(problem)
            table = pd.DataFrame(step_rows(run), columns=RUN_COLUMNS)
            st.html(table_html(table, pd.DataFrame("", index=table.index, columns=table.columns)))


# ---------------------------------------------------------------------------
# Page 2: one company
# ---------------------------------------------------------------------------

def export_file(kind, data, workbook):
    """The function an export button calls when clicked: it builds the file then (export.py), not on every redraw."""
    return lambda: export_bytes(kind, data, workbook)


def export_button(workbook, data):
    """Export: a popover with the four exports, built from today's workbook (off when it can't be read)."""
    with st.popover("Export", key="exports", help=EXPORT_HELP, disabled=data is None):
        for kind, label, mime in EXPORT_DOWNLOADS:
            st.download_button(label, export_file(kind, data, workbook), file_name=export_paths(workbook)[kind].name,
                               mime=mime, key=f"export_{kind}")


def company_buttons(workbook, state, ask_claude, config, output_dir, data):
    """Generate, the four downloads and Export, side by side."""
    cells = st.columns(2 + len(PAGE_DOWNLOADS))
    with cells[0]:
        generate_button(workbook, ask_claude, config, output_dir, key="generate_page", primary=True)
    for cell, (kind, label, mime) in zip(cells[1:], PAGE_DOWNLOADS):
        with cell:
            download_button(workbook, output_dir, state["current"], kind, label, mime, key=f"{kind}_page")
    with cells[-1]:
        export_button(workbook, data)


def mapping_panel(workbook, config):
    """Review mapping for a company in data/ whose headers need confirming; Save mapping once every pair is confirmed."""
    proposals = mapping_proposals(workbook)
    if not proposals:
        return
    with st.container(key="card-mapping"):
        columns, ready = review_mapping(proposals, workbook.stem)
        if st.button("Save mapping", key="save_mapping", disabled=not ready, help=None if ready else CONFIRM_ALL_FIRST):
            outcome = confirm_mapping(workbook, columns, config)
            say(outcome["ok"], outcome["message"])
            st.rerun()


def approve_panel(workbook, state, data_dir, output_dir):
    """The reviewer's name and the Approve button (only for files built from today's workbook)."""
    st.subheader("Approve")
    st.caption(APPROVE_INTRO)
    name_cell, button_cell = st.columns([3, 1], vertical_alignment="bottom")
    reviewer = name_cell.text_input("Reviewer name", key=f"reviewer_{workbook.stem}")
    if button_cell.button("Approve", key="approve", disabled=not state["current"],
                          help=None if state["current"] else NOT_AVAILABLE):
        outcome = approve_company(workbook.stem, reviewer, data_dir, output_dir)
        say(outcome["ok"], outcome["message"])
        st.rerun()


def show_flags(data):
    """The flag count, the colored flags table, and runway at next quarter's budget (context, not a flag)."""
    st.subheader(flag_count_text(data["flags"]))
    st.caption(LEGEND)
    st.html(table_html(flags_table(data), flags_colors(data)))
    st.caption(md(runway_lines(data).splitlines()[-1] + " (context, not a flag)"))


def show_changes(workbook, data, config, output_dir):
    """What changed since the last run (diff_runs.py): which run, then flags flipped, metrics moved and data gaps."""
    st.subheader(HEADING)
    report, settings, problem = run_changes(workbook, data, config, output_dir)
    if problem:
        st.error(problem)
        return
    if report is None:
        st.caption(NO_EARLIER_RUN)
        return
    st.caption(md(compared_with_text(report)))
    for title, lines in change_sections(report, settings):
        st.markdown(f"**{md(title)}**\n\n" + "\n".join(f"- {md(line)}" for line in lines))


def show_gaps(data):
    """Every metric and flag with an input missing, by quarter."""
    st.markdown("**Data gaps**\n\n" + "\n".join(f"- {md(line)}" for line in gaps_lines(data["gaps"])))


def show_metrics(data):
    """Every metric (rows) in every quarter (columns), colored like the flags."""
    st.subheader("Metrics")
    st.html(table_html(metrics_table(data), metrics_colors(data), index_header="Metric"))


def show_charts(data):
    """The deck's two charts, side by side."""
    st.subheader("ARR and cash")
    for cell, figure in zip(st.columns(2), chart_figures(data)):
        cell.pyplot(figure)
        plt.close(figure)  # free its memory: the page redraws often


def question_markdown(summary):
    """The questions as markdown: a bold theme line, then that theme's questions as bullets."""
    groups = [f"**{md(theme)}**\n\n" + "\n".join(f"- {md(question)}" for question in questions)
              for theme, questions in grouped_questions(summary.questions)]
    return f"**{QUESTIONS_HEADING}**\n\n" + "\n\n".join(groups)


def show_commentary(workbook, config, output_dir):
    """The AI headline, diagnosis, risks and questions (as on the deck) if a saved analysis matches these numbers."""
    st.subheader("AI commentary")
    summary = saved_commentary(workbook, config, output_dir)
    if summary is None:
        st.info(NO_COMMENTARY)
        return
    st.caption(AI_DRAFTED_LINE)
    st.markdown(f"**{md(summary.headline)}**")
    st.markdown(md(summary.diagnosis))
    risks, questions = st.columns(2)
    risks.markdown("**Risks**\n\n" + "\n".join(f"- **{md(point.title)}:** {md(point.detail)}"
                                               for point in summary.risks))
    questions.markdown(question_markdown(summary))


def company_page(stem, data_dir, output_dir):
    """Page 2: everything about one company, and its buttons."""
    if st.button("Back to portfolio", key="back"):
        st.session_state.pop(COMPANY_KEY, None)
        st.rerun()
    workbook = find_workbook(stem, data_dir)
    if workbook is None:
        st.error(NO_WORKBOOK_FOUND.format(name=stem.title()))
        return
    config = page_config()
    if config is None:
        return
    data, problem = load_company(workbook, config)
    state = run_state(workbook, output_dir)
    st.title(f"{company_name(workbook)}: {data['latest']}" if data else company_name(workbook))
    st.caption(f"Last run: {state['last_run']} · Deck status: {plain(state['status'])}")
    ask_claude = ai_checkbox()
    show_messages()
    with st.container(key="card-buttons"):
        company_buttons(workbook, state, ask_claude, config, output_dir, data)
    if problem:
        st.error(problem)
        mapping_panel(workbook, config)   # shown only when the problem is headers to confirm
        return
    sections = [("flags", show_flags, (data,)), ("changes", show_changes, (workbook, data, config, output_dir)),
                ("gaps", show_gaps, (data,)), ("metrics", show_metrics, (data,)),
                ("charts", show_charts, (data,)), ("commentary", show_commentary, (workbook, config, output_dir)),
                ("approve", approve_panel, (workbook, state, data_dir, output_dir))]
    for name, show, arguments in sections:
        with st.container(key=f"card-{name}"):   # theme.py: white, 1 px border, on the surface color
            show(*arguments)


def main(data_dir=DATA_DIR, output_dir=OUTPUT_DIR):
    """The page Streamlit draws: a company's page if one was clicked, else the portfolio. Tests pass temporary folders."""
    st.set_page_config(page_title=PAGE_TITLE, layout="wide")   # theme.py's style sheet caps it at about 1100 px
    st.html(f"<style>{streamlit_css()}</style>")
    stem = st.session_state.get(COMPANY_KEY)
    if stem:
        company_page(stem, data_dir, output_dir)
    else:
        portfolio_page(data_dir, output_dir)


if __name__ == "__main__":  # Streamlit runs this file as __main__; importing it (tests) shows nothing
    main()
