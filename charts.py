"""The two charts on the deck's Charts slide, drawn with matplotlib and saved as PNG images.

- ARR chart:  ending ARR by quarter (top) and net new ARR by quarter (bottom). Two panels
              with their own scales instead of one chart with two y-axes: net new ARR is
              a small fraction of ARR and can go negative, so it gets its own zero line.
- Cash chart: ending cash by quarter as a line, with runway in the title.

A blank quarter is never hidden or bridged:
- bars: no bar is drawn, and "data missing" is written where the bar would be;
- line: the NaN stays in the data, so matplotlib stops the line at the gap and restarts after it.

No math and no hand-typed numbers here: values arrive computed, and every label is
formatted by metrics.format_value.
"""

import math

import matplotlib

matplotlib.use("Agg")  # draw to files only; no window (must be set before importing pyplot)
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.ticker import FuncFormatter  # noqa: E402

from make_template import DARK_GRAY, MID_GRAY, NAVY  # noqa: E402
from metrics import INPUT_LABELS, METRIC_LABELS, MISSING_INPUT, REASON_DISPLAY, format_value  # noqa: E402

FONT_SIZE = 12                 # same floor as the slide text: charts are read on the same screen
DPI = 200                      # sharp enough for a projector
BAR_WIDTH = 0.6
LINE_WIDTH = 2
MARKER_SIZE = 7
LABEL_OFFSET_PT = 8            # gap between a bar or point and its value label
X_PADDING = 0.5                # half a quarter of space left of the first and right of the last point
HEADROOM = 1.2                 # top of the cash axis = 1.2 x the highest cash, room for the label
GAP_LABEL = REASON_DISPLAY[MISSING_INPUT]   # "data missing"
LABEL_BACKGROUND = {"facecolor": "white", "edgecolor": "none", "pad": 1}  # keeps a label readable over lines


def hex_color(value):
    """'1F2A44' -> '#1F2A44' (matplotlib wants the #)."""
    return f"#{value}"


NAVY_HEX, MID_GRAY_HEX, DARK_GRAY_HEX = hex_color(NAVY), hex_color(MID_GRAY), hex_color(DARK_GRAY)


def style_axis(axis, quarters, money_column):
    """Quiet axes: no top/right border, light horizontal grid, quarter labels on two lines."""
    for side in ("top", "right"):
        axis.spines[side].set_visible(False)
    axis.spines["left"].set_color(MID_GRAY_HEX)
    axis.spines["bottom"].set_color(MID_GRAY_HEX)
    axis.grid(axis="y", color=MID_GRAY_HEX, alpha=0.25, linewidth=0.8)
    axis.set_axisbelow(True)
    axis.set_xticks(range(len(quarters)))
    axis.set_xticklabels([quarter.replace(" ", "\n") for quarter in quarters])
    axis.yaxis.set_major_formatter(FuncFormatter(lambda value, _: format_value(money_column, value)))
    axis.tick_params(labelsize=FONT_SIZE, colors=DARK_GRAY_HEX)


def label_gaps(axis, values):
    """Write "data missing" at every blank quarter, halfway up the panel."""
    for position, value in enumerate(values):
        if math.isnan(value):
            axis.text(position, 0.5, GAP_LABEL.replace(" ", "\n"), transform=axis.get_xaxis_transform(),
                      ha="center", va="center", fontsize=FONT_SIZE, color=MID_GRAY_HEX, style="italic",
                      bbox=LABEL_BACKGROUND)


def label_latest(axis, values, column, below=None):
    """Write the latest value next to its bar or point: above it, or below when `below` is true.

    Bars: below only if negative (the label sits past the end of the bar).
    """
    position, value = len(values) - 1, values[-1]
    if math.isnan(value):
        return  # the gap label already says why
    if below is None:
        below = value < 0
    offset = (0, -LABEL_OFFSET_PT if below else LABEL_OFFSET_PT)  # in points, away from the bar end
    axis.annotate(format_value(column, value), (position, value), xytext=offset, textcoords="offset points",
                  ha="center", va="top" if below else "bottom", fontsize=FONT_SIZE, color=DARK_GRAY_HEX,
                  fontweight="bold", bbox=LABEL_BACKGROUND)


def bar_panel(axis, quarters, values, column):
    """Bars for every quarter that has a value; nothing (plus a label) for a blank one."""
    present = [position for position, value in enumerate(values) if not math.isnan(value)]
    axis.bar(present, [values[position] for position in present], width=BAR_WIDTH, color=NAVY_HEX)
    axis.axhline(0, color=MID_GRAY_HEX, linewidth=0.8)
    style_axis(axis, quarters, column)
    axis.set_title(METRIC_LABELS[column], loc="left", fontsize=FONT_SIZE, color=DARK_GRAY_HEX, fontweight="bold")
    axis.margins(y=0.35)  # room for the value label past the tallest (or most negative) bar
    label_gaps(axis, values)
    label_latest(axis, values, column)


def arr_chart(quarters, ending_arr, net_new_arr, size_inches):
    """Figure with two panels: ending ARR bars (taller) above net new ARR bars."""
    figure, (arr_axis, net_new_axis) = plt.subplots(
        2, 1, figsize=size_inches, gridspec_kw={"height_ratios": [3, 2]}, constrained_layout=True)
    bar_panel(arr_axis, quarters, list(ending_arr), "ending_arr")
    bar_panel(net_new_axis, quarters, list(net_new_arr), "net_new_arr")
    return figure


def cash_chart(quarters, ending_cash, runway_text, size_inches):
    """Figure with ending cash as a line that breaks at a blank quarter, and runway in the title."""
    figure, axis = plt.subplots(figsize=size_inches, constrained_layout=True)
    values = list(ending_cash)
    # plot() gets the NaN too: matplotlib leaves a gap there instead of joining the neighbours.
    axis.plot(range(len(values)), values, color=NAVY_HEX, linewidth=LINE_WIDTH, marker="o", markersize=MARKER_SIZE)
    style_axis(axis, quarters, "ending_cash")
    # Cash running out means reaching zero, so zero stays on the chart.
    present = [value for value in values if not math.isnan(value)]
    lowest, highest = min(present + [0]), max(present + [0])
    axis.set_ylim(lowest * HEADROOM, highest * HEADROOM if highest > 0 else 1)
    axis.set_title(f"{INPUT_LABELS['ending_cash']}\n{runway_text}", loc="left", fontsize=FONT_SIZE,
                   color=DARK_GRAY_HEX, fontweight="bold")
    axis.set_xlim(-X_PADDING, len(values) - 1 + X_PADDING)  # room for the last point's label
    label_gaps(axis, values)
    # The line reaches the last point from the left: put the label on the side the line isn't.
    came_down = len(values) > 1 and not math.isnan(values[-2]) and values[-2] > values[-1]
    label_latest(axis, values, "ending_cash", below=came_down)
    return figure


def save_chart(figure, path):
    """Save a figure as PNG at its own size (so 12 pt in the figure is 12 pt on the slide) and free its memory."""
    figure.savefig(path, dpi=DPI, facecolor="white")
    plt.close(figure)
    return path
