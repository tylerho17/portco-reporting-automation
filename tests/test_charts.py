"""Unit tests for charts.py: a blank quarter is a visible gap, never a bar or a line through it.

Run from the project folder:  pytest
"""

import math

import matplotlib.pyplot as plt
from matplotlib.colors import to_hex
from matplotlib.text import Text

import theme
from charts import arr_chart, cash_chart

NAN = math.nan
QUARTERS = ["Q3 2024", "Q4 2024", "Q1 2025", "Q2 2025"]


def texts(axis):
    """Every text drawn on an axis, with line breaks read as spaces ("data\nmissing" -> "data missing")."""
    return [text.get_text().replace("\n", " ") for text in axis.texts]


def bar_positions(axis):
    """The x position (centre) of every bar drawn on an axis."""
    return sorted(round(bar.get_x() + bar.get_width() / 2, 6) for bar in axis.patches)


def test_arr_chart_has_no_bar_for_the_blank_quarter_and_says_data_missing():
    figure = arr_chart(QUARTERS, [100.0, 120.0, NAN, 150.0], [10.0, 20.0, NAN, -5.0], (6, 5))
    arr_axis, net_new_axis = figure.axes
    assert bar_positions(arr_axis) == [0, 1, 3]        # nothing at position 2 (Q1 2025)
    assert bar_positions(net_new_axis) == [0, 1, 3]
    assert "data missing" in texts(arr_axis)
    assert "data missing" in texts(net_new_axis)
    plt.close(figure)


def test_arr_chart_labels_the_latest_values():
    figure = arr_chart(QUARTERS, [100.0, 120.0, 130.0, 1500.0], [10.0, 20.0, 10.0, -240.0], (6, 5))
    arr_axis, net_new_axis = figure.axes
    assert "1,500" in texts(arr_axis)
    assert "-240" in texts(net_new_axis)
    plt.close(figure)


def test_cash_line_keeps_the_blank_so_matplotlib_breaks_the_line():
    figure = cash_chart(QUARTERS, [900.0, 800.0, NAN, 600.0], "Runway at current burn: 9.0 mo", (6, 5))
    axis = figure.axes[0]
    y_values = list(axis.lines[0].get_ydata())
    # NaN stays in place: matplotlib draws no segment into or out of a NaN point.
    assert math.isnan(y_values[2])
    assert y_values[:2] == [900.0, 800.0] and y_values[3] == 600.0
    assert "data missing" in texts(axis)
    plt.close(figure)


def test_cash_chart_shows_the_runway_text_and_starts_at_zero():
    figure = cash_chart(QUARTERS, [900.0, 800.0, 700.0, 600.0], "Runway at current burn: 9.0 mo", (6, 5))
    axis = figure.axes[0]
    assert "Runway at current burn: 9.0 mo" in axis.get_title(loc="left")
    assert axis.get_ylim()[0] == 0   # cash running out means reaching zero, so zero is on the chart
    plt.close(figure)


def test_figure_is_the_size_it_will_have_on_the_slide():
    figure = cash_chart(QUARTERS, [900.0, 800.0, 700.0, 600.0], "x", (6.1, 5.2))
    assert list(figure.get_size_inches()) == [6.1, 5.2]
    plt.close(figure)


# ---------------------------------------------------------------------------
# Task 3: the look comes from theme.py
# ---------------------------------------------------------------------------

def every_text(figure):
    """Every piece of text in a figure: titles, tick labels, value labels and gap labels."""
    return [text for text in figure.findobj(Text) if text.get_text().strip()]


def test_chart_text_is_arial_or_the_next_installed_font_at_caption_size():
    figure = arr_chart(QUARTERS, [100.0, 120.0, NAN, 150.0], [10.0, 20.0, NAN, -5.0], (6, 5))
    found = every_text(figure)
    assert found
    assert {text.get_fontname() for text in found} == {theme.first_installed_font()}
    assert {text.get_fontsize() for text in found} == {13}          # the caption size
    plt.close(figure)


def test_bars_and_line_are_navy_titles_slate_and_gap_labels_mid_gray():
    arr = arr_chart(QUARTERS, [100.0, 120.0, 130.0, 150.0], [10.0, 20.0, 10.0, -5.0], (6, 5))
    assert {to_hex(bar.get_facecolor()) for bar in arr.axes[0].patches} == {"#0b2545"}
    assert to_hex(arr.axes[0]._left_title.get_color()) == "#334155"
    cash = cash_chart(QUARTERS, [900.0, 800.0, NAN, 600.0], "x", (6, 5))
    assert to_hex(cash.axes[0].lines[0].get_color()) == "#0b2545"
    gap = [text for text in cash.axes[0].texts if "data" in text.get_text()][0]
    assert to_hex(gap.get_color()) == "#64748b"
    plt.close(arr)
    plt.close(cash)
