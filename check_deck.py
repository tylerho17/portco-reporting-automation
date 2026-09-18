"""Automated proof for build step 4: the board decks are complete, match the numbers, and fit.

For each of the three companies, build the deck (with output/<company>_analysis.json if it exists,
otherwise with the placeholder), open the saved file, and check:
1. 4 slides, titles in order (typed by hand here).
2. Slide 1: every number on the slide appears in the metrics table - the saved
   output/<company>_metrics.xlsx (Metrics sheet, plus thresholds on its Flags sheet), shown in
   Excel's own number formats. Each row's latest and prior cells equal that metric's Excel cells;
   thresholds and statuses match the Flags sheet; status cells are red / green / gray.
3. Slide 2: two chart pictures; a blank quarter has no bar and no line through it.
4. Slide 3: the flag count matches the company's story; every tripped flag, the combo rule, and
   every metric with data missing (or "None").
5. Slide 4: the JSON's headline, 3 risks and 3 questions under "AI-drafted from computed metrics -
   review before use", and none of its wins; or exactly "AI summary unavailable" with no AI-drafted
   line. Slides 1 to 3 carry no AI text either way.
6. Nothing overflows: every text box and table cell is re-measured from the saved file
   (text_fit.py), every font is at least 12 pt, and every shape sits inside the slide.
   The overflow check is itself proven by breaking a deck on purpose.
7. A footer on every slide: fictional-data note, source file, today's date, and the review status
   read from output/<company>_manifest.json ("AI-drafted | reviewed by NAME on DATE" or
   "AI-drafted | not reviewed"), measured to fit one line. No DRAFT watermark by default.
Plus: a tampered analysis (one number changed) or a stale one (another quarter) gets the placeholder;
--draft stamps every slide of an unreviewed deck and never an approved one.

No API calls. Run: python check_deck.py  -> prints "All checks passed" or stops at the first failure.
"""

import copy
import datetime
import json
import math
import re
import shutil
import tempfile
from pathlib import Path

from openpyxl import load_workbook
from pptx import Presentation
from pptx.util import Emu

from build_deck import FICTIONAL_NOTE, PLACEHOLDER_TEXT, analysis_path, save_deck
from charts import arr_chart, cash_chart
from check_companies import COMPANIES
from clean import clean_workbook
from excel_output import save_metrics_workbook
from make_template import FOOTER_RULE
from metrics import CANNOT_EVALUATE, COMBO_FLAG_NAME, CONFIG_PATH, TRIP, compute_metrics, load_config
from provenance import NOT_REVIEWED, approval_status, file_sha256, manifest_path, read_manifest
from text_fit import MIN_FONT_PT, paragraph, text_height_pt, text_width_pt

PROJECT_DIR = Path(__file__).parent
LATEST, PRIOR, FIRST = "Q2 2026", "Q1 2026", "Q3 2024"
TOLERANCE_PT = 0.5   # rounding when sizes are saved in EMU
PICTURE = 13         # MSO_SHAPE_TYPE.PICTURE

# Status colors by the words in the status cell (Excel's "Bad" red, "Good" green, and gray).
STATUS_FILLS = {"Tripped": "FFC7CE", "Passed": "C6EFCE", "Cannot evaluate": "D9D9D9"}

# Excel number format -> the same display written as a Python format (what Excel shows in the cell).
EXCEL_FORMATS = {"0.0%": "{:.1%}", '0.00"x"': "{:.2f}x", '0.0" mo"': "{:.1f} mo", "#,##0": "{:,.0f}"}

# A number as shown on a slide: optional minus, digits with commas and decimals, optional unit.
NUMBER = re.compile(r"(?<![0-9A-Za-z])([-−–])?(\d[\d,]*(?:\.\d+)?)(%|x| mo)?")


# ---------------------------------------------------------------------------
# Reading the metrics workbook
# ---------------------------------------------------------------------------

def excel_display(cell):
    """What Excel shows in a cell: text as is, numbers in the cell's number format."""
    if cell.value is None or isinstance(cell.value, str):
        return cell.value
    if cell.number_format not in EXCEL_FORMATS:
        raise AssertionError(f"Unexpected number format {cell.number_format!r} in {cell.coordinate}")
    return EXCEL_FORMATS[cell.number_format].format(cell.value)


def number_tokens(text):
    """Every number in a text with its sign and unit, e.g. '-19.0%', '11.0 mo', '2026'."""
    return {("-" if sign else "") + digits + (unit or "") for sign, digits, unit in NUMBER.findall(text)}


def read_metrics_workbook(path):
    """{(quarter, label): shown text} from the Metrics sheet, and {flag: row dict} from the Flags sheet."""
    book = load_workbook(path)
    metrics_sheet, flags_sheet = book["Metrics"], book["Flags"]
    labels = [cell.value for cell in metrics_sheet[1]]
    table = {}
    for row in metrics_sheet.iter_rows(min_row=2):
        for label, cell in zip(labels[1:], row[1:]):
            table[(row[0].value, label)] = excel_display(cell)
    headers = [cell.value for cell in flags_sheet[1]]
    flags = {}
    for row in flags_sheet.iter_rows(min_row=2):
        shown = {header: excel_display(cell) for header, cell in zip(headers, row)}
        if shown["Status"]:  # flag rows only (skips the empty row and the runway context row below them)
            flags[shown["Flag"]] = shown
    return table, flags


def allowed_numbers(table, flags):
    """Every number shown anywhere in the metrics workbook: metric cells, quarter labels, flag thresholds."""
    texts = [text for text in table.values() if text] + [quarter for quarter, _ in table]
    texts += [str(value) for flag in flags.values() for value in flag.values() if value]
    return set().union(*(number_tokens(text) for text in texts))


# ---------------------------------------------------------------------------
# Reading the deck
# ---------------------------------------------------------------------------

def shape(slide, name):
    matches = [s for s in slide.shapes if s.name == name]
    assert len(matches) == 1, f"Expected one shape named {name!r}, found {len(matches)}"
    return matches[0]


def slide_text(slide):
    return "\n".join(s.text_frame.text for s in slide.shapes if s.has_text_frame)


def table_rows(slide):
    """Slide 1's table as lists of cell texts (header first)."""
    return [[cell.text for cell in row.cells] for row in shape(slide, "KPI table").table.rows]


# ---------------------------------------------------------------------------
# Checks
# ---------------------------------------------------------------------------

def check_titles(slides, company):
    expected = [f"{company}: key metrics — {LATEST} vs {PRIOR}", f"ARR and cash — {FIRST} to {LATEST}",
                f"Risks and flags — {LATEST}", f"AI commentary — {LATEST}"]
    assert len(slides) == 4, f"{company}: {len(slides)} slides, expected 4"
    titles = [slide.shapes.title.text for slide in slides]
    assert titles == expected, f"{company}: titles\nExpected: {expected}\nGot:      {titles}"


def expected_flag_count(company):
    """'6 of 9 flags tripped' counted from the company's story in check_companies.py."""
    statuses = list(company["expected_flags"].values())
    text = f"{statuses.count(TRIP)} of {len(statuses)} flags tripped"
    if statuses.count(CANNOT_EVALUATE):
        text += f", {statuses.count(CANNOT_EVALUATE)} cannot evaluate"
    return text


def check_ai_slide(slides, summary, name):
    """Slide 4: the analysis text under the AI-drafted line, or the placeholder (never a mix).

    Also: the wins are nowhere on the deck, and slides 1 to 3 carry no AI text or placeholder.
    """
    slides = list(slides)   # python-pptx's slide list can't be sliced
    slide = slides[3]
    headline = shape(slide, "Headline").text_frame.text
    names = [item.name for item in slide.shapes]
    for other in slides[:3]:
        assert "Headline" not in [item.name for item in other.shapes], f"{name}: AI text off slide 4"
        assert PLACEHOLDER_TEXT not in slide_text(other), f"{name}: placeholder off slide 4"
    if summary is None:
        assert headline == PLACEHOLDER_TEXT, f"{name}: headline should be the placeholder, got {headline!r}"
        assert "AI-drafted line" not in names, f"{name}: 'AI-drafted' line on a slide with no AI text"
        return
    line = shape(slide, "AI-drafted line").text_frame.text
    assert line == "AI-drafted from computed metrics - review before use", f"{name}: AI-drafted line {line!r}"
    assert headline == summary["headline"], f"{name}: headline {headline!r} != JSON {summary['headline']!r}"
    risks, questions = shape(slide, "Risks").text_frame.text, shape(slide, "Questions").text_frame.text
    for point in summary["risks"]:
        assert point["title"] in risks and point["detail"] in risks, f"{name}: risk missing: {point}"
    for question in summary["questions"]:
        assert question in questions, f"{name}: question missing from slide 4: {question!r}"
    deck_text = "\n".join(slide_text(each) for each in slides)
    for point in summary["wins"]:
        assert point["detail"] not in deck_text, f"{name}: a win is on the deck: {point}"
    assert PLACEHOLDER_TEXT not in slide_text(slide), f"{name}: placeholder shown"


def check_kpi_numbers(slide, table, flags, name):
    """Every number on slide 1 is in the metrics workbook. (The footer's run date is checked in check_footers.)"""
    allowed = allowed_numbers(table, flags)
    texts = [item.text_frame.text for item in slide.shapes if item.has_text_frame and item.name != "Footer"]
    shown = set().union(*(number_tokens(text) for text in texts),
                        *(number_tokens(cell) for row in table_rows(slide) for cell in row))
    invented = shown - allowed
    assert not invented, f"{name}: numbers on slide 1 that aren't in the metrics table: {sorted(invented)}"
    return len(shown)


def check_kpi_rows(slide, table, flags, name):
    """Row by row: latest/prior cells, budget cell, threshold and status (with its color)."""
    rows = table_rows(slide)
    assert rows[0][1:3] == [LATEST, PRIOR], f"{name}: table header {rows[0]}"
    kpi_table = shape(slide, "KPI table").table
    for row_number, cells in enumerate(rows[1:], start=1):
        label, latest, prior, budget, status = cells
        if (LATEST, label) in table:  # a metric row (flag names are the metric labels)
            assert [latest, prior] == [table[(LATEST, label)], table[(PRIOR, label)]], \
                f"{name}: row {label!r} shows {latest!r}, {prior!r}; metrics table has " \
                f"{table[(LATEST, label)]!r}, {table[(PRIOR, label)]!r}"
        if label == "Ending ARR ($K)":
            assert budget == f"vs budget: {table[(LATEST, 'Ending ARR vs budget')]}", f"{name}: budget cell {budget!r}"
        if label in flags:
            flag = flags[label]
            assert status == flag["Status"], f"{name}: {label} status {status!r}, Flags sheet {flag['Status']!r}"
            if label != COMBO_FLAG_NAME:
                assert budget.endswith(flag["Threshold"]), f"{name}: {label} threshold {budget!r} vs {flag['Threshold']!r}"
            fill = str(kpi_table.cell(row_number, 4).fill.fore_color.rgb)
            expected_fill = STATUS_FILLS[status.split(" — ")[0]]
            assert fill == expected_fill, f"{name}: {label} status cell is {fill}, expected {expected_fill}"
    flag_rows = [cells[0] for cells in rows[1:] if cells[0] in flags]
    assert len(flag_rows) == len(flags), f"{name}: table shows {len(flag_rows)} of {len(flags)} flags"


def check_charts(slide, company):
    """Two pictures; the blank quarter has no bar and breaks the cash line."""
    for name in ("ARR chart", "Cash chart"):
        assert shape(slide, name).shape_type == PICTURE, f"{company['name']}: {name} is not a picture"
    blank = company["answer_key"].BLANK_QUARTER
    if not blank:
        return
    actuals, _ = clean_workbook(company["answer_key"].OUTPUT_PATH)
    metrics = compute_metrics(actuals)
    quarters = list(metrics.index)
    position = quarters.index(blank)
    arr = arr_chart(quarters, metrics["ending_arr"], metrics["net_new_arr"], (6, 5))
    bars = [round(bar.get_x() + bar.get_width() / 2) for bar in arr.axes[0].patches]
    assert position not in bars, f"{company['name']}: a bar is drawn for blank {blank}"
    cash = cash_chart(quarters, actuals["ending_cash"], "", (6, 5))
    assert math.isnan(cash.axes[0].lines[0].get_ydata()[position]), f"{company['name']}: cash line joins over {blank}"


def check_risks_slide(slide, company, table_flags, gap_labels):
    text = shape(slide, "Risks and flags").text_frame.text
    count = expected_flag_count(company)
    assert f"Tripped flags ({count})" in text, f"{company['name']}: flag count should read {count!r} on slide 3"
    for flag, status in company["expected_flags"].items():
        if status == TRIP and flag != COMBO_FLAG_NAME:
            assert f"• {flag}: " in text, f"{company['name']}: tripped flag {flag!r} missing from slide 3"
    combo_status = table_flags[COMBO_FLAG_NAME]["Status"]
    assert f"{COMBO_FLAG_NAME}: {combo_status}" in text, f"{company['name']}: combo result missing"
    text = shape(slide, "Data gaps").text_frame.text
    assert text.startswith("Data gaps"), f"{company['name']}: no Data gaps line"
    for label in gap_labels:
        assert label in text, f"{company['name']}: data gap {label!r} missing from slide 3"
    if not gap_labels:
        assert "None — every metric and flag has the data it needs" in text, f"{company['name']}: gaps should say None"


def expected_review(workbook, output_dir):
    """What the footer must end with, worked out here from the manifest (not from build_deck's code)."""
    manifest = read_manifest(manifest_path(workbook, output_dir))
    approval, _ = approval_status(manifest, file_sha256(workbook), file_sha256(CONFIG_PATH))
    if approval is None:
        return "AI-drafted | not reviewed"
    return f"AI-drafted | reviewed by {approval['reviewer']} on {approval['approved_at'][:10]}"


def check_footers(slides, source_name, name, review):
    """Every footer: the note, source file, today's date, ends with the review status, and fits one line."""
    today = datetime.date.today().isoformat()
    for number, slide in enumerate(slides, start=1):
        box = shape(slide, "Footer")
        footer = box.text_frame.text
        for part in (FICTIONAL_NOTE, source_name, today):
            assert part in footer, f"{name}: slide {number} footer lacks {part!r}: {footer!r}"
        assert footer.endswith(" | " + review), f"{name}: slide {number} footer should end {review!r}: {footer!r}"
        room = Emu(box.width - box.text_frame.margin_left - box.text_frame.margin_right).pt
        width = text_width_pt(footer, MIN_FONT_PT)
        assert width <= room, f"{name}: slide {number} footer is {width:.0f} pt wide, one line holds {room:.0f} pt"
    return width, room


def watermark_count(slides):
    return sum(item.name == "Watermark" for slide in slides for item in slide.shapes)


# ---------------------------------------------------------------------------
# Overflow: re-measure everything in the saved file
# ---------------------------------------------------------------------------

def frame_paragraphs(frame, where):
    """The paragraphs of a saved text frame, in the form text_fit.py measures. Every run needs a size."""
    result = []
    for item in frame.paragraphs:
        sizes = [run.font.size.pt for run in item.runs if run.font.size is not None]
        assert len(sizes) == len(item.runs), f"{where}: a run has no font size, so its fit can't be proven"
        assert all(size >= MIN_FONT_PT for size in sizes), f"{where}: font below {MIN_FONT_PT} pt: {sizes}"
        space_after = Emu(item.space_after).pt if item.space_after is not None else 0
        result.append(paragraph(item.text, max(sizes, default=MIN_FONT_PT),
                                bold=any(run.font.bold for run in item.runs), space_after=space_after))
    return result


def check_text_fits(frame, width, height, where):
    """The text needs no more height than its box has (after the frame's margins)."""
    inner_width = Emu(width - frame.margin_left - frame.margin_right).pt
    inner_height = Emu(height - frame.margin_top - frame.margin_bottom).pt
    needed = text_height_pt(frame_paragraphs(frame, where), inner_width)
    assert needed <= inner_height + TOLERANCE_PT, f"{where}: text overflows ({needed:.1f} pt needed, {inner_height:.1f} pt)"


def check_table_fits(frame, where):
    """Every cell's text fits its row height and column width."""
    table = frame.table
    for row_number, row in enumerate(table.rows):
        for column, cell in enumerate(row.cells):
            width = table.columns[column].width - cell.margin_left - cell.margin_right
            needed = (text_height_pt(frame_paragraphs(cell.text_frame, where), Emu(width).pt)
                      + Emu(cell.margin_top + cell.margin_bottom).pt)
            assert needed <= Emu(row.height).pt + TOLERANCE_PT, \
                f"{where}: cell ({row_number}, {column}) overflows ({needed:.1f} pt needed, {Emu(row.height).pt:.1f} pt)"
    table_bottom = frame.top + sum(row.height for row in table.rows)
    assert table_bottom <= FOOTER_RULE[1], f"{where}: table runs into the footer"


def check_no_overflow(presentation, name):
    """Every shape inside the slide (and above the footer rule, except the footer); every text fits."""
    width, height = presentation.slide_width, presentation.slide_height
    for number, slide in enumerate(presentation.slides, start=1):
        for item in slide.shapes:
            where = f"{name} slide {number}, {item.name}"
            assert item.left >= 0 and item.top >= 0, f"{where}: starts off the slide"
            assert item.left + item.width <= width and item.top + item.height <= height, f"{where}: ends off the slide"
            if item.name != "Footer":
                assert item.top + item.height <= FOOTER_RULE[1], f"{where}: runs into the footer"
            if item.has_text_frame:
                check_text_fits(item.text_frame, item.width, item.height, where)
            elif item.has_table:
                check_table_fits(item, where)


def check_overflow_check_catches_overflow(path):
    """Break a saved deck on purpose (long text, then a tiny font); the overflow check must fail both times."""
    for break_deck in (lambda frame: setattr(frame.paragraphs[0].runs[0], "text", "overflow " * 400),
                       lambda frame: setattr(frame.paragraphs[0].runs[0].font, "size", Emu(MIN_FONT_PT * 12700 - 12700))):
        presentation = Presentation(path)
        break_deck(shape(presentation.slides[3], "Headline").text_frame)
        try:
            check_no_overflow(presentation, "broken deck")
        except AssertionError:
            continue
        raise AssertionError("The overflow check didn't catch a deliberately broken deck")


# ---------------------------------------------------------------------------
# One company, and the tampered analyses
# ---------------------------------------------------------------------------

def check_company(company, config, output_dir):
    name, workbook = company["name"], Path(company["answer_key"].OUTPUT_PATH)
    json_path = analysis_path(workbook)
    summary = json.loads(json_path.read_text())["summary"] if json_path.exists() else None

    deck, why_unavailable = save_deck(workbook, config, json_path, output_dir=output_dir)
    assert (summary is None) == (why_unavailable is not None), f"{name}: analysis use unexpected: {why_unavailable}"
    table, table_flags = read_metrics_workbook(save_metrics_workbook(workbook, config))
    presentation = Presentation(deck)
    slides = presentation.slides

    check_titles(slides, name)
    numbers = check_kpi_numbers(slides[0], table, table_flags, name)
    check_kpi_rows(slides[0], table, table_flags, name)
    check_charts(slides[1], company)
    gap_labels = [label for (quarter, label), text in table.items() if text == "data missing"]
    check_risks_slide(slides[2], company, table_flags, sorted(set(gap_labels)))
    check_ai_slide(slides, summary, name)
    count = expected_flag_count(company)
    review = expected_review(workbook, output_dir)
    width, room = check_footers(slides, workbook.name, name, review)
    assert watermark_count(slides) == 0, f"{name}: a DRAFT watermark without --draft"
    check_no_overflow(presentation, name)
    ai = "AI text from the JSON" if summary else f"placeholder ({why_unavailable})"
    print(f"✓ {name}: 4 slides, {ai}, {count}; {numbers} distinct numbers on slide 1 all in the metrics table; "
          f"nothing overflows; footer '{review}' on one line ({width:.0f} of {room:.0f} pt); no watermark")
    return deck


def tampered_analysis(folder, change):
    """A copy of Northwind's analysis with one change applied."""
    saved = json.loads(analysis_path(COMPANIES[0]["answer_key"].OUTPUT_PATH).read_text())
    changed = copy.deepcopy(saved)
    change(changed)
    path = Path(folder) / "tampered_analysis.json"
    path.write_text(json.dumps(changed))
    return path


def check_bad_analysis_gets_placeholder(config, folder):
    """An analysis with a number that isn't in the data, or for another quarter, must not reach the deck."""
    def invent_number(saved):
        saved["summary"]["headline"] = saved["summary"]["headline"].replace("11.0 mo", "11.5 mo")

    def older_quarter(saved):
        saved["payload"]["latest_quarter"] = PRIOR

    workbook = COMPANIES[0]["answer_key"].OUTPUT_PATH
    for change, expected_reason in ((invent_number, "11.5"), (older_quarter, PRIOR)):
        path = tampered_analysis(folder, change)
        deck, why_unavailable = save_deck(workbook, config, path, output_dir=folder)
        assert why_unavailable and expected_reason in why_unavailable, f"Tampered analysis accepted: {why_unavailable}"
        slides = Presentation(deck).slides
        check_ai_slide(slides, None, f"Northwind ({change.__name__})")
        print(f"✓ Northwind with a tampered analysis ({change.__name__}) shows the placeholder: {why_unavailable[:70]}...")


def check_draft_option(config, folder):
    """--draft: an unreviewed deck gets the watermark on all 4 slides; an approved one gets none.

    Built in a temporary folder, so output/ keeps its default (unwatermarked) decks. Approved means
    a copy of a manifest that approve.py signed, for the same workbook and thresholds.
    """
    folder = Path(folder)
    for company in COMPANIES:
        workbook = Path(company["answer_key"].OUTPUT_PATH)
        real_manifest = manifest_path(workbook, PROJECT_DIR / "output")
        if real_manifest.exists():
            shutil.copy(real_manifest, manifest_path(workbook, folder))
        review = expected_review(workbook, folder)
        deck, _ = save_deck(workbook, config, None, output_dir=folder, draft=True)
        slides = Presentation(deck).slides
        count = watermark_count(slides)
        expected = 0 if "reviewed by" in review else len(slides)
        assert count == expected, f"{company['name']} --draft: {count} watermarks, expected {expected} ({review})"
        if count:
            assert shape(slides[0], "Watermark").text_frame.text == NOT_REVIEWED
        check_footers(slides, workbook.name, company["name"], review)
        print(f"✓ {company['name']} with --draft: {count} of {len(slides)} slides watermarked ('{review}')")


def main():
    config = load_config()
    decks = [check_company(company, config, PROJECT_DIR / "output") for company in COMPANIES]
    check_overflow_check_catches_overflow(decks[0])
    print("✓ The overflow check fails on a deck broken on purpose (long headline; font below the floor)")
    with tempfile.TemporaryDirectory() as folder:
        check_bad_analysis_gets_placeholder(config, folder)
    with tempfile.TemporaryDirectory() as folder:
        check_draft_option(config, folder)
    print("All checks passed")


if __name__ == "__main__":
    main()
