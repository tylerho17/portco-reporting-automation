"""Tests for export.py (Task 11): a company's metrics and flags as CSV and JSON, and an email-ready HTML summary.

The main promise: the exports carry the same values as the metrics workbook. For each of the three
companies the workbook (excel_output.py) and the four exports are saved into a temporary folder,
read back from disk, and compared cell by cell (check_export.py's comparisons, also used by
python check_export.py). Spot values are typed here by hand from each company's story.
The email summary must follow the rules that keep a paste into Outlook clean (check_export.outlook_problems).
No API calls: creating a real Anthropic client fails the test, and nothing here has an AI step.
Run from the project folder:  pytest
"""

import csv
import io
import json
import math
import shutil
from pathlib import Path

import pytest

import analyze
import check_export
import check_northwind
import export
import mapping
from build_deck import collect_deck_data, kpi_header, kpi_rows
from check_export import email_tables, outlook_problems, read_workbook
from clean import clean_workbook
from excel_output import save_metrics_workbook
from metrics import CANNOT_EVALUATE, MISSING_INPUT, NO_PRIOR_PERIOD, PASS, TRIP, load_config
from provenance import file_sha256
from theme import STATUS_COLORS

PROJECT_DIR = Path(__file__).parent.parent
COMPANIES = ["alderpeak", "fernhollow", "northwind"]
EM_DASH = chr(0x2014)   # by its Unicode number, so this file shows none


@pytest.fixture(autouse=True)
def no_real_client(monkeypatch):
    """Creating a real Anthropic client fails the test: exports never ask Claude."""
    def refuse(*args, **kwargs):
        raise AssertionError("a test tried to create a real Anthropic client")
    monkeypatch.setattr(analyze.anthropic, "Anthropic", refuse)


@pytest.fixture(autouse=True)
def mappings_dir(tmp_path, monkeypatch):
    """A temporary mappings/ folder: the project's own is never read or written."""
    folder = tmp_path / "mappings"
    monkeypatch.setattr(mapping, "MAPPINGS_DIR", folder)
    return folder


def workbook(name):
    return PROJECT_DIR / "data" / f"{name}.xlsx"


def company_data(name):
    actuals, next_budget = clean_workbook(workbook(name))
    return collect_deck_data(name.title(), f"{name}.xlsx", actuals, next_budget, load_config())


def row_for(rows, quarter, metric):
    return next(row for row in rows if row["quarter"] == quarter and row["metric"] == metric)


def flag_named(rows, name):
    return next(row for row in rows if row["flag"] == name)


# ---------------------------------------------------------------------------
# File names and units
# ---------------------------------------------------------------------------

def test_each_export_has_its_own_file_next_to_the_metrics_workbook(tmp_path):
    paths = export.export_paths(workbook("northwind"), tmp_path)
    assert list(paths) == list(export.EXPORT_KINDS)
    assert {kind: path.name for kind, path in paths.items()} == {
        "metrics_csv": "northwind_metrics.csv", "flags_csv": "northwind_flags.csv",
        "json": "northwind_export.json", "email": "northwind_email.html"}
    assert all(path.parent == tmp_path for path in paths.values())


def test_every_metric_has_a_unit_a_downstream_tool_can_read():
    assert export.unit_of("ending_arr") == "$K" and export.unit_of("pipeline") == "$K"
    assert export.unit_of("nrr") == "ratio" and export.unit_of("rule_of_40") == "ratio"
    assert export.unit_of("runway_months") == "months" and export.unit_of("cac_payback_months") == "months"
    assert export.unit_of("burn_multiple") == "multiple"
    assert set(export.UNITS) == {"$K", "ratio", "months", "multiple"}


# ---------------------------------------------------------------------------
# Metric rows: one per quarter and metric
# ---------------------------------------------------------------------------

def test_one_metric_row_per_quarter_and_metric_quarter_by_quarter():
    rows = export.metric_rows(company_data("northwind"))
    assert len(rows) == 8 * 20
    assert all(list(row) == export.METRIC_FIELDS for row in rows)
    assert [row["quarter"] for row in rows[:20]] == ["Q3 2024"] * 20
    assert rows[0]["metric"] == "ending_arr" and rows[0]["label"] == "Ending ARR ($K)"
    assert {row["company"] for row in rows} == {"Northwind"}


def test_a_ratio_is_the_exact_decimal_with_the_deck_s_text_beside_it():
    data = company_data("northwind")
    row = row_for(export.metric_rows(data), "Q2 2026", "nrr")
    assert row["value"] == 0.9705540488182874                            # the decimal, not 97.1
    assert math.isclose(row["value"], check_northwind.EXPECTED_LATEST["nrr"], rel_tol=1e-9)   # the hand formula
    assert row["text"] == "97.1%" and row["unit"] == "ratio" and row["status"] == export.NUMBER
    assert row["flag_tripped"] is True                                   # NRR below 100%: red in Excel


def test_a_value_has_the_16_significant_digits_the_workbook_stores():
    # Python's burn multiple is 2.3493975903614457; openpyxl saves "%.16g" of it, 2.349397590361446.
    data = company_data("northwind")
    assert data["metrics"].loc["Q2 2026", "burn_multiple"] == 2.3493975903614457
    row = row_for(export.metric_rows(data), "Q2 2026", "burn_multiple")
    assert row["value"] == 2.349397590361446 and row["text"] == "2.35x" and row["unit"] == "multiple"
    assert flag_named(export.flag_rows(data), "Burn multiple")["value"] == 2.349397590361446


def test_a_dollar_metric_is_in_thousands():
    row = row_for(export.metric_rows(company_data("northwind")), "Q2 2026", "ending_arr")
    assert row["value"] == 27470.0 and row["text"] == "27,470" and row["unit"] == "$K"
    assert row["flag_tripped"] is False                                  # no flag on ending ARR


def test_a_blank_quarter_is_data_missing_with_no_number():
    row = row_for(export.metric_rows(company_data("northwind")), "Q1 2025", "nrr")
    assert row["value"] is None and row["text"] == "data missing" and row["status"] == MISSING_INPUT


def test_the_first_quarter_s_growth_has_no_prior_period():
    row = row_for(export.metric_rows(company_data("northwind")), "Q3 2024", "arr_qoq")
    assert row["value"] is None and row["text"] == "n/a (no prior period)" and row["status"] == NO_PRIOR_PERIOD


def test_an_infinite_burn_multiple_is_no_number_with_its_reason_in_words():
    row = row_for(export.metric_rows(company_data("fernhollow")), "Q2 2026", "burn_multiple")
    assert row["value"] is None and row["status"] == export.INFINITE
    assert row["text"] == "∞ (ARR shrank)" and row["flag_tripped"] is True


# ---------------------------------------------------------------------------
# Flag rows: the latest quarter's flags
# ---------------------------------------------------------------------------

def test_northwind_s_flags_as_rows():
    rows = export.flag_rows(company_data("northwind"))
    assert len(rows) == 9 and all(list(row) == export.FLAG_FIELDS for row in rows)
    assert [row["status"] for row in rows].count(TRIP) == 6
    runway = flag_named(rows, "Runway at current burn")
    assert runway == {"company": "Northwind", "quarter": "Q2 2026", "flag": "Runway at current burn",
                      "metric": "runway_months", "value": 11.0, "text": "11.0 mo", "threshold": 12,
                      "trips_when": "below threshold", "status": TRIP, "reason": None, "status_text": "Tripped"}


def test_the_combo_rule_has_no_single_value_and_says_what_it_tests():
    combo = flag_named(export.flag_rows(company_data("northwind")), "NRR falling while pipeline rising")
    assert combo["metric"] is None and combo["value"] is None and combo["threshold"] is None
    assert combo["trips_when"] == "NRR falls at least 1 pt and pipeline rises at every step, last 3 quarters"
    assert combo["status"] == TRIP and combo["status_text"] == "Tripped"


def test_a_flag_that_cannot_be_evaluated_carries_its_reason():
    rule_of_40 = flag_named(export.flag_rows(company_data("fernhollow")), "Rule of 40")
    assert rule_of_40["status"] == CANNOT_EVALUATE and rule_of_40["reason"] == MISSING_INPUT
    assert rule_of_40["status_text"] == "Cannot evaluate: missing input"
    assert rule_of_40["value"] is None and rule_of_40["text"] == "data missing"


# ---------------------------------------------------------------------------
# CSV and JSON text
# ---------------------------------------------------------------------------

def test_csv_has_a_header_row_empty_cells_for_no_value_and_true_or_false():
    rows = [{"a": None, "b": True, "c": False, "d": 0.1 + 0.2, "e": "∞ (ARR shrank)"}]
    text = export.csv_text(rows, ["a", "b", "c", "d", "e"])
    assert text == "a,b,c,d,e\r\n,true,false,0.30000000000000004,∞ (ARR shrank)\r\n"


def test_every_csv_number_reads_back_exactly():
    data = company_data("northwind")
    rows = export.metric_rows(data)
    read = list(csv.DictReader(io.StringIO(export.csv_text(rows, export.METRIC_FIELDS))))
    for row, back in zip(rows, read):
        expected = "" if row["value"] is None else row["value"]
        assert (float(back["value"]) if back["value"] else "") == expected


def test_the_json_names_its_sources_and_has_no_nan_or_infinity():
    data = company_data("fernhollow")
    record = export.export_record(data, workbook("fernhollow"))
    text = export.json_text(record)
    json.loads(text, parse_constant=lambda name: pytest.fail(f"{name} in the JSON"))
    assert record["company"] == "Fernhollow" and record["workbook"] == "fernhollow.xlsx"
    assert record["sources"]["workbook_sha256"] == file_sha256(workbook("fernhollow"))
    assert record["sources"]["config_sha256"] == file_sha256(PROJECT_DIR / "config.yaml")
    assert record["sources"]["mapping_sha256"] is None                   # no column mapping for this company
    assert record["latest_quarter"] == "Q2 2026" and len(record["quarters"]) == 8
    assert record["flag_summary"] == {"text": "7 of 9 flags tripped, 1 cannot evaluate", "checked": 9,
                                      "tripped": 7, "passed": 1, "cannot_evaluate": 1}
    assert record["units"] == export.UNITS


def test_the_json_keeps_runway_at_budget_and_the_data_gaps():
    record = export.export_record(company_data("fernhollow"), workbook("fernhollow"))
    runway = record["runway_at_next_budget"]
    assert runway["quarter"] == "Q2 2026" and runway["text"] == "7.6 mo"
    assert math.isclose(runway["value"], 3300 / (1300 / 3), rel_tol=1e-9)
    gaps = {gap["name"]: gap for gap in record["data_gaps"]}
    assert gaps["rule_of_40"] == {"name": "rule_of_40", "label": "Rule of 40", "quarters": ["Q2 2025", "Q2 2026"]}
    assert gaps["flag: Rule of 40"]["label"] == "Flag: Rule of 40"


def test_a_company_with_no_gaps_has_an_empty_gap_list():
    assert export.export_record(company_data("alderpeak"), workbook("alderpeak"))["data_gaps"] == []


def test_the_same_numbers_always_give_the_same_files():
    # No run time in any export: a downstream tool can compare two exports byte for byte.
    first = {kind: export.export_bytes(kind, company_data("northwind"), workbook("northwind"))
             for kind in export.EXPORT_KINDS}
    second = {kind: export.export_bytes(kind, company_data("northwind"), workbook("northwind"))
              for kind in export.EXPORT_KINDS}
    assert first == second


# ---------------------------------------------------------------------------
# The main promise: the same values as the metrics workbook
# ---------------------------------------------------------------------------

@pytest.fixture(scope="module")
def saved(tmp_path_factory):
    """{company: (workbook read back, the four exports read back)}, all saved to disk first."""
    folder = tmp_path_factory.mktemp("exports")
    config = load_config()
    result = {}
    for name in COMPANIES:
        excel = read_workbook(save_metrics_workbook(workbook(name), config, folder))
        result[name] = (excel, check_export.read_exports(export.save_exports(workbook(name), config, folder)))
    return result


@pytest.mark.parametrize("name", COMPANIES)
def test_the_metrics_csv_carries_every_workbook_cell(saved, name):
    excel, exports = saved[name]
    assert check_export.metric_problems(exports["metrics_csv"], excel) == []
    assert len(exports["metrics_csv"]) == len(excel["metrics"]) == 8 * 20


@pytest.mark.parametrize("name", COMPANIES)
def test_the_json_metrics_carry_every_workbook_cell(saved, name):
    excel, exports = saved[name]
    assert check_export.metric_problems(exports["json"]["metrics"], excel) == []


@pytest.mark.parametrize("name", COMPANIES)
def test_the_flags_csv_and_json_match_the_flags_sheet(saved, name):
    excel, exports = saved[name]
    assert check_export.flag_problems(exports["flags_csv"], excel) == []
    assert check_export.flag_problems(exports["json"]["flags"], excel) == []


@pytest.mark.parametrize("name", COMPANIES)
def test_the_json_runway_context_and_gaps_match_the_workbook(saved, name):
    excel, exports = saved[name]
    assert check_export.record_problems(exports["json"], excel) == []


@pytest.mark.parametrize("name", COMPANIES)
def test_the_email_shows_the_workbook_s_values(saved, name):
    excel, exports = saved[name]
    assert check_export.email_problems(exports["email"], excel) == []


def test_a_changed_export_value_is_caught(saved):
    # The comparison itself: one number off in the 12th significant digit is a mismatch.
    excel, exports = saved["northwind"]
    rows = [dict(row) for row in exports["metrics_csv"]]
    row = next(row for row in rows if row["metric"] == "nrr" and row["quarter"] == "Q2 2026")
    row["value"] = repr(float(row["value"]) * (1 + 1e-12))
    assert check_export.metric_problems(rows, excel) == ["Q2 2026 NRR (annualized): value "
                                                         f"{row['value']} but the workbook has "
                                                         f"{excel['metrics'][('Q2 2026', 'NRR (annualized)')]['value']!r}"]


def test_a_row_the_workbook_lacks_a_repeated_row_and_a_missing_row_are_each_caught(saved):
    # Found by a planted bug: with "exported but not in the workbook" switched off, every test passed.
    excel, exports = saved["northwind"]
    rows = [dict(row) for row in exports["metrics_csv"]]
    extra = dict(rows[0], quarter="Q3 2026")
    assert check_export.metric_problems(rows + [extra], excel) == [
        "Q3 2026 Ending ARR ($K): exported but not in the workbook"]
    assert check_export.metric_problems(rows + [rows[0]], excel) == ["Q3 2024 Ending ARR ($K): exported twice"]
    assert check_export.metric_problems(rows[1:], excel) == ["Q3 2024 Ending ARR ($K): in the workbook but not exported"]


# ---------------------------------------------------------------------------
# The email summary
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("name", COMPANIES)
def test_the_email_follows_the_rules_for_a_clean_paste_into_outlook(name):
    assert outlook_problems(export.email_html(company_data(name))) == []


def test_the_email_s_table_is_slide_one_s_table():
    data = company_data("northwind")
    table = email_tables(export.email_html(data))[0]
    assert table[0] == kpi_header(data)
    combo = "NRR falls at least 1 pt and pipeline rises at every step, last 3 quarters"
    expected = [[combo if cell == "rule on Risks and flags slide" else cell for cell in cells]
                for cells, _ in kpi_rows(data)]
    assert table[1:] == expected


def test_the_email_s_status_cells_take_the_status_colors():
    html = export.email_html(company_data("northwind"))
    fill, text = STATUS_COLORS[TRIP]
    assert html.count(f'bgcolor="#{fill}"') == 6                        # six tripped flags
    assert f"color: #{text}" in html
    assert html.count(f'bgcolor="#{STATUS_COLORS[PASS][0]}"') == 3


def test_the_email_says_what_it_is_and_that_it_has_no_ai_text():
    html = export.email_html(company_data("fernhollow"))
    assert "Fernhollow: board update, Q2 2026" in html
    assert "7 of 9 flags tripped, 1 cannot evaluate" in html
    assert "Runway at next quarter's budgeted burn: 7.6 mo (context, not a flag)" in html
    assert "Q2 2026: Flag: Rule of 40" in html                           # the Data gaps lines
    assert "Fictional data" in html and "no AI text" in html
    assert EM_DASH not in html


def test_the_email_escapes_text_from_the_workbook():
    data = company_data("alderpeak")
    data["company"] = "<b>Alder & Peak</b>"
    html = export.email_html(data)
    assert "<b>Alder" not in html and "&lt;b&gt;Alder &amp; Peak&lt;/b&gt;" in html


def test_the_outlook_check_catches_what_outlook_mangles():
    good = export.email_html(company_data("alderpeak"))
    assert outlook_problems(good) == []
    broken = {
        "a style sheet": good.replace("<head>", "<head><style>td {color: red}</style>"),
        "a class": good.replace("<table ", '<table class="x" ', 1),
        "a cell without its font": good.replace("font-family: Arial, Helvetica, sans-serif; ", "", 1),
        "a table without cellpadding": good.replace('cellpadding="', 'data-x="', 1),
        "a fill without bgcolor": good.replace(' bgcolor="#', ' data-bg="#', 1),
        "an rgba color": good.replace("color: #", "color: rgba(0,0,0,1); x: #", 1),
        "flexbox": good.replace('style="', 'style="display: flex; ', 1),
        "an image": good.replace("</body>", '<img src="https://example.com/x.png"></body>'),
        "too wide": good.replace('width="640"', 'width="900"', 1),
    }
    for what, html in broken.items():
        assert outlook_problems(html), what


# ---------------------------------------------------------------------------
# Saving, the web page's bytes, and the command line
# ---------------------------------------------------------------------------

def test_save_exports_writes_the_four_files_the_page_offers(tmp_path):
    config = load_config()
    paths = export.save_exports(workbook("northwind"), config, tmp_path)
    assert sorted(path.name for path in tmp_path.iterdir()) == sorted(path.name for path in paths.values())
    data = company_data("northwind")
    for kind, path in paths.items():
        assert path.read_bytes() == export.export_bytes(kind, data, workbook("northwind")), kind


def test_the_files_are_utf8_without_a_byte_order_mark(tmp_path):
    paths = export.save_exports(workbook("fernhollow"), load_config(), tmp_path)
    for path in paths.values():
        raw = path.read_bytes()
        assert not raw.startswith(b"\xef\xbb\xbf") and "∞" in raw.decode("utf-8"), path.name


def test_the_command_line_exports_one_company(tmp_path, capsys):
    assert export.main([str(workbook("northwind")), "--output-dir", str(tmp_path)]) == 0
    printed = capsys.readouterr().out
    assert printed.count("Saved ") == 4 and "northwind_email.html" in printed
    assert len(list(tmp_path.iterdir())) == 4


def test_the_command_line_exports_every_company_and_carries_on_past_a_bad_one(tmp_path, capsys):
    data_dir, output_dir = tmp_path / "data", tmp_path / "output"
    data_dir.mkdir()
    for name in COMPANIES:
        shutil.copy(workbook(name), data_dir)
    (data_dir / "broken.xlsx").write_bytes(b"not a workbook")
    code = export.main(["--all", "--data-dir", str(data_dir), "--output-dir", str(output_dir)])
    assert code == 1                                                     # one company failed
    printed = capsys.readouterr().out
    assert "Broken: " in printed and "Traceback" not in printed
    assert len(list(output_dir.iterdir())) == 3 * 4


def test_the_command_line_needs_a_workbook_or_all(capsys):
    with pytest.raises(SystemExit):
        export.main([])
