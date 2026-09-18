"""Smoke tests for app.py (the Streamlit page): its table helpers, and both pages drawn with Streamlit's AppTest.

Every page is drawn on a temporary data/ folder (copies of the three workbooks) and a temporary
output/ folder, so clicking Generate or Approve never touches the real ones. A guard makes creating
a real Anthropic client fail, and the AI box is never ticked, so no test can reach the API.
The work behind the buttons is tested in tests/test_portfolio.py.
Run from the project folder:  pytest
"""

import shutil
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import Workbook
from streamlit.testing.v1 import AppTest

import analyze
import app
import mapping
import portfolio
from analyze import BoardSummary, build_payload, save_analysis
from clean import clean_workbook
import theme
from metrics import CANNOT_EVALUATE, PASS, TRIP, load_config
from provenance import NOT_REVIEWED, manifest_path, read_manifest
from test_mapping import RENAMES, renamed_copy
from theme import STATUS_COLORS

PROJECT_DIR = Path(__file__).parent.parent
COMPANIES = ["alderpeak", "fernhollow", "northwind"]
EM_DASH = chr(0x2014)   # by its Unicode number, so this file shows none
TIMEOUT = 120   # seconds: Generate builds a deck, a memo and a workbook


@pytest.fixture(autouse=True)
def no_real_client(monkeypatch):
    """Creating a real Anthropic client fails the test, so no test here can reach the API."""
    def refuse(*args, **kwargs):
        raise AssertionError("a test tried to create a real Anthropic client")
    monkeypatch.setattr(analyze.anthropic, "Anthropic", refuse)


@pytest.fixture(autouse=True)
def mappings_dir(tmp_path, monkeypatch):
    """A temporary mappings/ folder: the project's own is never read or written."""
    folder = tmp_path / "mappings"
    monkeypatch.setattr(mapping, "MAPPINGS_DIR", folder)
    return folder


@pytest.fixture
def folders(tmp_path):
    """(data folder with copies of the three workbooks, empty output folder)."""
    data_dir, output_dir = tmp_path / "data", tmp_path / "output"
    data_dir.mkdir()
    for name in COMPANIES:
        shutil.copy(PROJECT_DIR / "data" / f"{name}.xlsx", data_dir)
    return data_dir, output_dir


def company_data(name):
    data, problem = portfolio.load_company(PROJECT_DIR / "data" / f"{name}.xlsx", load_config())
    assert problem is None
    return data


# ---------------------------------------------------------------------------
# Small helpers and the company page's tables
# ---------------------------------------------------------------------------

def test_the_ai_checkbox_label_names_the_typical_cost():
    label = app.ai_checkbox_label()
    assert "AI commentary" in label and "$0.09" in label


def test_status_colors_are_the_palette_s():
    # Typed from the Task 3 brief: red C0392B on FDE8E6, green 1E8449 on EAF6EF, gray fill EDF0F3.
    assert app.status_css(TRIP) == "background-color: #FDE8E6; color: #C0392B"
    assert app.status_css(PASS) == "background-color: #EAF6EF; color: #1E8449"
    assert app.status_css(CANNOT_EVALUATE) == "background-color: #EDF0F3; color: #334155"
    for status in (TRIP, PASS, CANNOT_EVALUATE):
        fill, text = STATUS_COLORS[status]
        assert app.status_css(status) == f"background-color: #{fill}; color: #{text}"


def test_a_table_is_html_with_the_theme_s_class_and_escaped_text():
    table = pd.DataFrame([["<b>NRR</b>", "97.1%"]], columns=["Flag", "Value"])
    colors = pd.DataFrame([["", app.status_css(TRIP)]], columns=["Flag", "Value"])
    html = app.table_html(table, colors)
    assert html.startswith('<div class="bp-table-wrap"><table class="bp-table">')
    assert "<th>Flag</th><th>Value</th>" in html
    assert "<td>&lt;b&gt;NRR&lt;/b&gt;</td>" in html                     # text, never markup
    assert f'<td style="{app.status_css(TRIP)}">97.1%</td>' in html


def test_the_metrics_table_names_its_rows_in_the_first_column():
    data = company_data("northwind")
    html = app.table_html(app.metrics_table(data), app.metrics_colors(data), index_header="Metric")
    assert "<th>Metric</th><th>Q3 2024</th>" in html and "<td>NRR (annualized)</td>" in html


def test_markdown_text_keeps_dollar_signs_and_drops_em_dashes():
    assert app.md(f"$1.2M {EM_DASH} up from $0.9M") == "\\$1.2M: up from \\$0.9M"


def test_the_flag_table_matches_northwinds_story():
    data = company_data("northwind")
    flags, colors = app.flags_table(data), app.flags_colors(data)
    assert list(flags.columns) == app.FLAG_COLUMNS and len(flags) == 9
    tripped_rows = flags["Status"] == "Tripped"
    assert tripped_rows.sum() == 6
    assert (colors[tripped_rows] == app.status_css(TRIP)).all().all()      # whole row red
    assert (colors[~tripped_rows] == app.status_css(PASS)).all().all()     # the rest green
    assert flags.loc[0, "Threshold"].startswith("trips below")


def test_a_flag_that_cannot_be_evaluated_gives_its_reason_without_an_em_dash():
    flags = app.flags_table(company_data("fernhollow"))
    statuses = list(flags["Status"])
    assert "Cannot evaluate: missing input" in statuses
    assert not any(EM_DASH in text for text in flags.to_numpy().ravel())


def test_the_metrics_table_is_colored_like_the_excel_sheet():
    data = company_data("northwind")
    table, colors = app.metrics_table(data), app.metrics_colors(data)
    assert table.shape == colors.shape
    assert list(table.columns)[-1] == "Q2 2026"
    assert table.loc["NRR (annualized)", "Q1 2025"] == "data missing"                  # the blank quarter
    assert colors.loc["NRR (annualized)", "Q1 2025"] == app.status_css(CANNOT_EVALUATE)  # gray
    assert colors.loc["NRR (annualized)", "Q2 2026"] == app.status_css(TRIP)             # below 100%: red
    assert colors.loc["Ending ARR ($K)", "Q2 2026"] == ""                                # no flag: no color


def test_both_charts_are_drawn():
    figures = app.chart_figures(company_data("northwind"))
    assert len(figures) == 2
    for figure in figures:
        app.plt.close(figure)


# ---------------------------------------------------------------------------
# The pages, drawn by AppTest
# ---------------------------------------------------------------------------

def render(data_dir, output_dir):
    """Draw the page on the test's folders (AppTest runs this function as a Streamlit script)."""
    from pathlib import Path

    import app
    app.main(Path(data_dir), Path(output_dir))


def page(folders, company=None):
    """The page after its first draw; company= opens that company's page."""
    test = AppTest.from_function(render, args=tuple(str(folder) for folder in folders), default_timeout=TIMEOUT)
    if company:
        test.session_state[app.COMPANY_KEY] = company
    return test.run()


def downloads(test):
    """Every download button on the page as (label, disabled)."""
    return [(button.proto.label, button.proto.disabled) for button in test.get("download_button")]


def texts(test):
    return [element.value for element in test.text]


def html_bodies(test):
    """Every piece of HTML the page wrote with st.html."""
    return [element.proto.body for element in test.get("html")]


def tables(test):
    return [body for body in html_bodies(test) if 'class="bp-table"' in body]


def primary_buttons(test):
    """The key of every primary button on the page (buttons and download buttons)."""
    return [button.proto.id for button in list(test.button) + list(test.get("download_button"))
            if button.proto.type == "primary"]


def test_both_pages_carry_the_theme_s_style_sheet(folders):
    for company in (None, "northwind"):
        assert f"<style>{theme.streamlit_css()}</style>" in html_bodies(page(folders, company=company))


def test_each_page_has_exactly_one_primary_button(folders):
    portfolio_page = page(folders)
    assert [button.key for button in portfolio_page.button if button.proto.type == "primary"] == ["generate_all"]
    assert len(primary_buttons(portfolio_page)) == 1
    company_page = page(folders, company="northwind")
    assert [button.key for button in company_page.button if button.proto.type == "primary"] == ["generate_page"]
    assert len(primary_buttons(company_page)) == 1
    assert company_page.button(key="approve").proto.type == "secondary"   # approve is never red or green


def test_the_flags_table_is_drawn_with_status_fills(folders):
    flags_html = tables(page(folders, company="northwind"))[0]
    assert "<th>Flag</th>" in flags_html
    assert flags_html.count(app.status_css(TRIP)) == 6 * len(app.FLAG_COLUMNS)   # six tripped rows, every cell


def test_the_portfolio_lists_every_company_with_nothing_to_download_yet(folders):
    test = page(folders)
    assert not test.exception and not test.error
    assert test.checkbox[0].label == app.ai_checkbox_label()
    assert [test.button(key=f"open_{name}").label for name in COMPANIES] == ["Alderpeak", "Fernhollow", "Northwind"]
    assert "6 of 9 flags tripped" in texts(test) and portfolio.NOT_GENERATED in texts(test)
    assert len(downloads(test)) == 9 and all(disabled for _, disabled in downloads(test))


def test_search_narrows_the_table_and_a_missing_name_says_upload_one(folders):
    test = page(folders)
    test.text_input(key="search").set_value("north").run()
    assert [button.key for button in test.button if button.key.startswith("open_")] == ["open_northwind"]
    test.text_input(key="search").set_value("Bluefin").run()
    assert [warning.value for warning in test.warning] == [
        'No KPI workbook found for "Bluefin". Upload one under "Add a company".']
    assert not [button for button in test.button if button.key.startswith("open_")]


def test_generate_on_a_row_builds_its_files_and_offers_the_downloads(folders):
    test = page(folders)
    test.button(key="generate_northwind").click().run()
    assert not test.exception
    assert [message.value for message in test.success] == [
        f"Northwind: generated. {portfolio.AI_NOT_ASKED_NOTE}"]
    assert NOT_REVIEWED in texts(test)
    enabled = [label for label, disabled in downloads(test) if not disabled]
    assert enabled == ["Download deck", "Download memo", "Download Excel"]


def test_generate_all_builds_every_company(folders):
    test = page(folders)
    test.button(key="generate_all").click().run()
    assert not test.exception and not test.error
    assert test.success[0].value == "Generate all: 3 of 3 companies generated."
    assert all(not disabled for _, disabled in downloads(test))
    assert (folders[1] / "batch_summary.csv").exists()


def test_an_unreadable_workbook_shows_clean_py_s_message_and_the_others_still_work(folders):
    book = Workbook()
    book.active.append(["Quarter", "Starting ARR"])   # a known header: an unknown one stops at the mapping step
    book.active.append(["Q1 2025", 100])
    book.save(folders[0] / "broken.xlsx")
    _, message = portfolio.load_company(folders[0] / "broken.xlsx", load_config())
    test = page(folders)
    assert not test.exception
    assert [error.value for error in test.error] == [f"Broken: {message}"]
    assert "Traceback" not in message
    test.button(key="generate_all").click().run()
    assert test.error[0].value == "Generate all: 3 of 4 companies generated."


def test_clicking_a_name_opens_the_company_page(folders):
    test = page(folders)
    test.button(key="open_northwind").click().run()
    assert not test.exception
    assert test.title[0].value == "Northwind: Q2 2026"


def test_the_company_page_shows_flags_gaps_metrics_charts_and_no_commentary_yet(folders):
    test = page(folders, company="northwind")
    assert not test.exception and not test.error
    assert test.subheader[0].value == "6 of 9 flags tripped"
    assert len(tables(test)) == 2                                     # flags, then metrics
    assert len(test.get("image")) == 2                                 # the two charts
    assert any(info.value == app.NO_COMMENTARY for info in test.info)
    assert [label for label, _ in downloads(test)] == [label for _, label, _ in app.PAGE_DOWNLOADS]
    assert test.button(key="approve").disabled                        # nothing to approve yet


def test_a_company_page_for_a_name_with_no_workbook_says_upload_one(folders):
    test = page(folders, company="bluefin")
    assert [error.value for error in test.error] == [
        'No KPI workbook found for "Bluefin". Upload one under "Add a company".']


def test_approve_on_the_company_page_records_the_reviewer(folders):
    test = page(folders, company="northwind")
    test.button(key="generate_page").click().run()
    test.text_input(key="reviewer_northwind").set_value("Dana Reviewer").run()
    test.button(key="approve").click().run()
    assert not test.exception
    assert test.success[0].value.startswith("Approved Northwind by Dana Reviewer on ")
    approval = read_manifest(manifest_path(folders[0] / "northwind.xlsx", folders[1]))["approval"]
    assert approval["reviewer"] == "Dana Reviewer"


def test_the_company_page_shows_saved_commentary_for_these_numbers(folders):
    actuals, next_budget = clean_workbook(folders[0] / "northwind.xlsx")
    payload = build_payload("Northwind", actuals, next_budget, load_config())
    point = {"title": "Retention", "detail": "Churn is rising."}
    answer = BoardSummary.model_validate({"headline": "Retention needs a plan.", "wins": [point] * 3,
                                          "risks": [point] * 3, "questions": ["Why?", "Who?", "When?"]})
    save_analysis(folders[1] / "northwind_analysis.json", payload, answer, {"model": "claude-sonnet-5"})
    test = page(folders, company="northwind")
    assert not test.exception
    shown = [markdown.value for markdown in test.markdown]
    assert "**Retention needs a plan.**" in shown
    assert any("Churn is rising." in text for text in shown) and any("When?" in text for text in shown)
    assert app.AI_DRAFTED_LINE in [caption.value for caption in test.caption]


def test_no_em_dash_anywhere_on_either_page(folders):
    for company in (None, "fernhollow"):
        test = page(folders, company=company)
        shown = [element.value for kind in ("title", "subheader", "caption", "markdown", "text", "info", "error")
                 for element in getattr(test, kind)] + html_bodies(test)   # the tables are HTML
        assert shown and not any(EM_DASH in str(text) for text in shown)


# ---------------------------------------------------------------------------
# Review mapping (Task 5)
# ---------------------------------------------------------------------------

def renamed_alderpeak_page(folders, tmp_path):
    """Alderpeak's page after its workbook is swapped for a copy with six headers renamed."""
    renamed = renamed_copy("alderpeak", tmp_path)
    (folders[0] / "alderpeak.xlsx").write_bytes(renamed.read_bytes())
    return page(folders, company="alderpeak")


def test_review_mapping_shows_each_pair_its_confidence_and_sample_values(folders, tmp_path):
    test = renamed_alderpeak_page(folders, tmp_path)
    assert not test.exception
    assert "Review mapping" in [subheader.value for subheader in test.subheader]
    changes = [box for box in test.selectbox if box.key.startswith("map_")]
    assert {box.value for box in changes} == set(RENAMES["alderpeak"])            # each proposal preselected
    assert all(box.label == "Change" for box in changes)
    shown = " ".join(markdown.value for markdown in test.markdown)
    for header in RENAMES["alderpeak"].values():
        assert f"**{header}**" in shown
    assert "99% (high)" in shown and "8000, 9000, 10090, 11250" in shown   # BoP ARR's confidence and values


def test_review_mapping_needs_every_pair_confirmed_before_saving(folders, tmp_path, mappings_dir):
    test = renamed_alderpeak_page(folders, tmp_path)
    confirms = [box for box in test.checkbox if box.key.startswith("confirm_")]
    assert len(confirms) == len(RENAMES["alderpeak"]) and all(box.label == "Confirm" for box in confirms)
    assert test.button(key="save_mapping").disabled
    for box in confirms[:-1]:
        box.check()
    test.run()
    assert test.button(key="save_mapping").disabled                            # one still unconfirmed
    assert not mappings_dir.exists()


def test_confirming_every_pair_saves_the_mapping_and_the_page_shows_the_numbers(folders, tmp_path, mappings_dir):
    test = renamed_alderpeak_page(folders, tmp_path)
    for box in [box for box in test.checkbox if box.key.startswith("confirm_")]:
        box.check()
    test.run()
    test.button(key="save_mapping").click().run()
    assert not test.exception
    assert test.success[0].value == portfolio.MAPPING_SAVED.format(company="Alderpeak", count=6)
    assert (mappings_dir / "alderpeak.yaml").exists()
    assert test.subheader[0].value == "0 of 9 flags tripped"                   # Alderpeak's own story


def test_change_resets_that_pair_s_confirmation_and_saves_the_new_column(folders, tmp_path, mappings_dir):
    test = renamed_alderpeak_page(folders, tmp_path)
    for box in [box for box in test.checkbox if box.key.startswith("confirm_")]:
        box.check()
    test.run()
    churn = next(box for box in test.selectbox if box.key.startswith("map_") and box.value == "churned_arr")
    churn.set_value("headcount").run()
    assert test.button(key="save_mapping").disabled                            # the changed pair needs a new tick
    employees = next(box for box in test.selectbox if box.key.startswith("map_") and box.value == "headcount"
                     and box.key != churn.key)
    employees.set_value("churned_arr").run()
    for box in [box for box in test.checkbox if box.key.startswith("confirm_") and not box.value]:
        box.check()
    test.run()
    test.button(key="save_mapping").click().run()
    saved = mapping.confirmed_aliases(folders[0] / "alderpeak.xlsx")
    assert saved["churn"] == "headcount" and saved["employees"] == "churned_arr"


def test_a_company_with_known_headers_has_no_review_step(folders):
    test = page(folders, company="northwind")
    assert "Review mapping" not in [subheader.value for subheader in test.subheader]


# ---------------------------------------------------------------------------
# The Mac launcher
# ---------------------------------------------------------------------------

def test_the_mac_launcher_is_double_clickable_and_starts_this_app():
    launcher = PROJECT_DIR / "run_app.command"
    assert launcher.stat().st_mode & 0o111                     # executable, so Finder will run it
    assert "streamlit run app.py" in launcher.read_text()
    assert "streamlit" in (PROJECT_DIR / "requirements.txt").read_text().split()


def test_the_launcher_installs_packages_when_the_venv_has_no_streamlit():
    # A .venv made before the web page existed has Python but no streamlit. Checking only for
    # .venv/bin/python skipped the install, and the double-click ended in "streamlit: command not found".
    assert "[ ! -x .venv/bin/streamlit ]" in (PROJECT_DIR / "run_app.command").read_text()
