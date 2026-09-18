"""Task 19: the two slide-2 charts look the same and nothing on them collides, at any data range.

Each chart is rendered (as save_chart would) for the three companies and for two made-up extremes:
- huge and long: 16 quarters, ARR in the billions of $K, two blank quarters in a row, a big negative
  latest net new ARR and cash run down to almost nothing;
- tiny and short: 2 quarters, every value 0 or 1 ($K), and a runway with no number.
Then every piece of text on the figure is measured in pixels, and the test fails if two overlap, if one
sticks out of the figure, or if one sits on a bar or a point.

Run from the project folder:  python -m pytest tests/test_chart_layout.py -q
"""

import math
from itertools import combinations
from pathlib import Path

import pytest
from matplotlib.colors import to_hex
from matplotlib.patches import Rectangle
from matplotlib.text import Annotation
from matplotlib.transforms import Bbox

import theme
from build_deck import chart_size_inches, collect_deck_data, runway_lines
from charts import DPI, SHRANK_LABEL, arr_chart, cash_chart, runway_chart
from clean import clean_workbook
from metrics import format_value, load_config

NAN = math.nan
PROJECT_DIR = Path(__file__).parent.parent
COMPANIES = ["northwind", "alderpeak", "fernhollow"]
TOLERANCE_PX = 1     # two boxes that only touch (anti-aliasing) don't count as overlapping


def quarter_names(first_year, count):
    """count quarter names from Q1 of first_year: Q1 2023, Q2 2023, ..."""
    return [f"Q{index % 4 + 1} {first_year + index // 4}" for index in range(count)]


def huge_and_long():
    """16 quarters, values in the billions of $K, blanks at positions 5 and 6, big swings either way."""
    arr = [1_000_000_000.0 + 550_000_000 * index for index in range(16)]
    net_new = [(-1) ** index * 400_000_000.0 for index in range(16)]
    net_new[-1] = -2_345_678_901.0
    cash = [5_000_000_000.0 - 330_000_000 * index for index in range(16)]
    cash[-1] = 1.0
    for series in (arr, net_new, cash):
        series[5] = series[6] = NAN
    runway = "Runway at current burn: 0.0 mo\nAt next quarter's budgeted burn: ∞ (budget not burning)"
    return quarter_names(2023, 16), arr, net_new, cash, runway


def tiny_and_short():
    """2 quarters, every value 0 or 1 ($K), and a runway with no number."""
    runway = "Runway at current burn: data missing\nAt next quarter's budgeted burn: n/a (no budget row)"
    return ["Q1 2026", "Q2 2026"], [1.0, 0.0], [0.0, -1.0], [1.0, 0.0], runway


def company_case(name):
    """A real company's chart inputs, exactly as build_deck.charts_slide passes them."""
    actuals, next_budget = clean_workbook(PROJECT_DIR / "data" / f"{name}.xlsx")
    data = collect_deck_data(name.title(), f"{name}.xlsx", actuals, next_budget, load_config())
    return (list(data["metrics"].index), list(data["metrics"]["ending_arr"]), list(data["metrics"]["net_new_arr"]),
            list(data["actuals"]["ending_cash"]), runway_lines(data))


CASES = {name: (lambda name=name: company_case(name)) for name in COMPANIES}
CASES["huge and long"] = huge_and_long
CASES["tiny and short"] = tiny_and_short


def draw(kind, case):
    """Build one chart at its slide size and render it the way save_chart does."""
    quarters, arr, net_new, cash, runway = CASES[case]()
    if kind == "ARR":
        figure = arr_chart(quarters, arr, net_new, chart_size_inches())
    else:
        figure = cash_chart(quarters, cash, runway, chart_size_inches())
    return render(figure)


def render(figure):
    """Draw a figure at save_chart's resolution and leave it that way, so what's measured is what's saved.

    (savefig would put the resolution back afterwards, and the legend keeps pixel positions from the
    last drawing, so measuring after savefig would compare two resolutions.)
    """
    figure.set_dpi(DPI)
    figure.canvas.draw()
    return figure


# ---------------------------------------------------------------------------
# What's on the figure, in pixels
# ---------------------------------------------------------------------------

def in_view(location, limits):
    """True if a tick location is inside the axis limits (matplotlib doesn't draw the rest)."""
    low, high = sorted(limits)
    return low - 1e-9 <= location <= high + 1e-9


def drawn_ticks(axis):
    """The tick labels matplotlib actually draws: visible, not empty, inside the limits."""
    ticks = [(tick, axis.get_xlim()) for tick in axis.xaxis.get_major_ticks()]
    ticks += [(tick, axis.get_ylim()) for tick in axis.yaxis.get_major_ticks()]
    return [tick.label1 for tick, limits in ticks if in_view(tick.get_loc(), limits)]


def drawn_texts(figure):
    """Every text drawn: titles, tick labels, value labels, gap labels, legend words."""
    texts = []
    for axis in figure.axes:
        texts += [axis.title, axis._left_title, axis._right_title, *axis.texts, *drawn_ticks(axis)]
        if axis.get_legend():
            texts += axis.get_legend().get_texts()
    return [text for text in texts if text.get_visible() and text.get_text().strip()]


def box(artist, renderer):
    """An artist's box in pixels, shrunk by the tolerance so touching isn't overlapping."""
    return artist.get_window_extent(renderer).padded(-TOLERANCE_PX)


def mark_boxes(figure, renderer):
    """Every bar and every line point on the figure, as boxes in pixels."""
    marks = []
    for axis in figure.axes:
        marks += [box(bar, renderer) for bar in axis.patches if isinstance(bar, Rectangle)]
        for line in axis.lines:
            size = line.get_markersize() * figure.dpi / 72 / 2   # points -> pixels, half a marker
            for x, y in axis.transData.transform(line.get_xydata()):
                if not (math.isnan(x) or math.isnan(y)) and line.get_marker() not in (None, "None", ""):
                    marks.append(Bbox.from_extents(x - size, y - size, x + size, y + size))
    return marks


def label(text):
    return repr(text.get_text().replace("\n", " "))


def layout_problems(figure):
    """Every text that overlaps another text, sticks out of the figure, or sits on a bar or a point."""
    renderer = figure.canvas.get_renderer()
    texts = drawn_texts(figure)
    boxes = {id(text): box(text, renderer) for text in texts}
    problems = [f"{label(a)} overlaps {label(b)}" for a, b in combinations(texts, 2)
                if boxes[id(a)].overlaps(boxes[id(b)])]
    problems += [f"{label(text)} sticks out of the figure" for text in texts
                 if not figure.bbox.padded(TOLERANCE_PX).contains(boxes[id(text)].x0, boxes[id(text)].y0)
                 or not figure.bbox.padded(TOLERANCE_PX).contains(boxes[id(text)].x1, boxes[id(text)].y1)]
    marks = mark_boxes(figure, renderer)
    problems += [f"{label(text)} sits on a bar or point" for text in texts
                 if any(boxes[id(text)].overlaps(mark) for mark in marks)]
    return problems


# ---------------------------------------------------------------------------
# The test the task asks for: no overlap, for every chart and every case
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("case", list(CASES))
@pytest.mark.parametrize("kind", ["ARR", "cash"])
def test_no_text_overlaps_anything_on_either_chart(kind, case):
    figure = draw(kind, case)
    assert layout_problems(figure) == []


def test_the_overlap_check_catches_two_labels_on_top_of_each_other():
    """The check itself: a second copy of the latest label, drawn in the same place, must be caught."""
    figure = draw("cash", "northwind")
    axis = figure.axes[0]
    latest = [text for text in axis.texts if isinstance(text, Annotation)][0]
    axis.annotate(latest.get_text(), latest.xy, xytext=latest.xyann, textcoords="offset points",
                  fontsize=latest.get_fontsize())
    assert layout_problems(render(figure)) == [f"{label(latest)} overlaps {label(latest)}"]


# ---------------------------------------------------------------------------
# One axis style for both charts
# ---------------------------------------------------------------------------

def money_axes(figure):
    """The $K axes on a chart: every panel of the ARR chart, the cash chart's one axis."""
    return figure.axes


@pytest.mark.parametrize("case", list(CASES))
@pytest.mark.parametrize("kind", ["ARR", "cash"])
def test_every_money_axis_has_the_same_tick_rule(kind, case):
    """Whole-$K ticks, 3 to 6 of them, zero among them, a labelled gridline at the top and bottom,
    and each label written by format_value (so "27,470" on the axis reads like "27,470" on slide 1)."""
    for axis in money_axes(draw(kind, case)):
        ticks = list(axis.get_yticks())
        assert 3 <= len(ticks) <= 6, ticks
        assert all(tick == int(tick) for tick in ticks), ticks
        assert 0 in ticks
        assert list(axis.get_ylim()) == [ticks[0], ticks[-1]]
        labels = [text.get_text() for text in axis.get_yticklabels()]
        assert labels == [format_value("ending_arr", tick) for tick in ticks]
        assert len(set(labels)) == len(labels)   # no "0, 0, 1": two ticks never read the same


@pytest.mark.parametrize("case", list(CASES))
def test_both_charts_give_each_quarter_the_same_slot(case):
    """Same x range on every panel of both charts: half a quarter of room either side."""
    quarters = CASES[case]()[0]
    for kind in ("ARR", "cash"):
        for axis in draw(kind, case).axes:
            assert list(axis.get_xlim()) == [-0.5, len(quarters) - 0.5]


def test_the_latest_quarter_is_always_labelled_even_when_labels_are_thinned():
    figure = draw("ARR", "huge and long")
    for axis in figure.axes:
        shown = [text.get_text() for text in axis.get_xticklabels() if text.get_visible() and text.get_text()]
        assert shown[-1] == "Q4\n2026"
        assert len(shown) < 16   # 16 two-line labels don't fit in the width: some are thinned


def test_quarter_labels_are_not_thinned_when_they_fit():
    for axis in draw("ARR", "northwind").axes:
        assert all(text.get_visible() and text.get_text() for text in axis.get_xticklabels())


# ---------------------------------------------------------------------------
# Colorblind-safe: a shrinking quarter differs by lightness and pattern, not by hue alone
# ---------------------------------------------------------------------------

def net_new_bars():
    """The net new ARR panel of a chart with a positive and a negative quarter."""
    figure = arr_chart(["Q1 2026", "Q2 2026"], [100.0, 90.0], [20.0, -10.0], chart_size_inches())
    return figure.axes[1]


def test_a_quarter_where_arr_shrank_is_amber_and_hatched_the_rest_navy_and_solid():
    grew, shrank = sorted(net_new_bars().patches, key=lambda bar: bar.get_height(), reverse=True)
    assert to_hex(grew.get_facecolor()) == f"#{theme.NAVY.lower()}" and not grew.get_hatch()
    assert to_hex(shrank.get_facecolor()) == f"#{theme.AMBER.lower()}" and shrank.get_hatch()


def test_the_shrank_pattern_is_named_in_a_legend_only_when_a_quarter_shrank():
    assert [text.get_text() for text in net_new_bars().get_legend().get_texts()] == [SHRANK_LABEL]
    growing = arr_chart(["Q1 2026", "Q2 2026"], [100.0, 120.0], [20.0, 20.0], chart_size_inches())
    assert growing.axes[1].get_legend() is None


@pytest.mark.parametrize("first, second, where", theme.SERIES_PAIRS, ids=[p[2] for p in theme.SERIES_PAIRS])
def test_two_series_colors_differ_in_lightness_by_3_to_1(first, second, where):
    """Told apart in grayscale, so by any kind of color blindness: lightness survives where hue doesn't."""
    assert theme.contrast_ratio(first, second) >= theme.AA_NON_TEXT_RATIO


def test_red_and_green_are_never_the_only_difference_between_series():
    """Neither color of a series pair is the status red or green (they mean a flag and nothing else)."""
    for first, second, _ in theme.SERIES_PAIRS:
        assert {first, second}.isdisjoint({theme.RED, theme.GREEN})


def test_a_tripped_runway_bar_is_hatched_as_well_as_red():
    figure = runway_chart(["A", "B"], [6.0, 30.0], ["6.0 mo", "30.0 mo"], [True, False], 12.0, "trips below 12.0 mo",
                          (6, 4))
    tripped, passed = figure.axes[0].patches
    assert tripped.get_hatch() and not passed.get_hatch()
