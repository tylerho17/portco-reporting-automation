"""The two charts on the deck's Charts slide (and the rollup's runway chart), drawn with matplotlib and saved as PNG.

- ARR chart:  ending ARR by quarter (top) and net new ARR by quarter (bottom). Two panels
              with their own scales instead of one chart with two y-axes: net new ARR is
              a small fraction of ARR and can go negative, so it gets its own zero line.
              A quarter where ARR shrank is amber AND hatched, named in a small legend.
- Cash chart: ending cash by quarter as a line, with runway in the title.
- Runway chart (the portfolio rollup, rollup.py): runway at current burn by company, one bar each,
              red and hatched where the runway flag trips, with the flag's threshold as a dashed line.

One axis style for both slide charts (Task 19):
- every $K axis: whole-$K ticks, 3 to 6 of them, zero always among them, and the axis starts and
  ends on a labelled gridline; the labels are written by metrics.format_value, like the tables;
- every quarter axis: one slot per quarter, half a slot of room either side; if the labels would
  touch, every second one is hidden (the latest always stays);
- the latest value is written just right of the latest bar or point, where nothing else is drawn.
tests/test_chart_layout.py renders both charts for every company and two made-up extremes and fails
if any two labels overlap, if a label sits on a bar or point, or if one sticks out of the picture.

A blank quarter is never hidden or bridged:
- bars: no bar is drawn, and "data missing" is written where the bar would be (turned on its side
  when the quarter's slot is too narrow for it);
- line: the NaN stays in the data, so matplotlib stops the line at the gap and restarts after it.

No math and no hand-typed numbers here: values arrive computed, and every label is
formatted by metrics.format_value.

The look comes from theme.py: navy bars and line, amber for a shrinking quarter, slate text, mid gray
axes and gap labels, Arial (else Helvetica, else DejaVu Sans) at the caption size.
"""

import math
import textwrap

import matplotlib

matplotlib.use("Agg")  # draw to files only; no window (must be set before importing pyplot)
# Figures are made with Figure(), not pyplot: pyplot keeps one shared list of open figures, which
# isn't safe when several companies draw charts at the same moment (main.py --workers).
from matplotlib.backends.backend_agg import FigureCanvasAgg  # noqa: E402
from matplotlib.figure import Figure  # noqa: E402
from matplotlib.patches import Patch  # noqa: E402
from matplotlib.ticker import FixedLocator, FuncFormatter, MaxNLocator  # noqa: E402

from metrics import INPUT_LABELS, METRIC_LABELS, MISSING_INPUT, REASON_DISPLAY, format_value  # noqa: E402
from theme import AMBER, CAPTION_PT, FONT_STACK, MID_GRAY, NAVY, RED, SLATE, WHITE, css_color  # noqa: E402

# Every chart is drawn in the first of Arial, Helvetica, DejaVu Sans this computer has.
matplotlib.rcParams["font.family"] = "sans-serif"
matplotlib.rcParams["font.sans-serif"] = list(FONT_STACK)

FONT_SIZE = CAPTION_PT         # 13: above the 12 pt floor, the size of every caption on the slides
DPI = 200                      # sharp enough for a projector
BAR_WIDTH = 0.6
LINE_WIDTH = 2
MARKER_SIZE = 7
LABEL_OFFSET_PT = 8            # gap between a bar or point and its value label
X_PADDING = 0.5                # half a quarter of space left of the first and right of the last point
HEADROOM = 1.2                 # runway chart: the axis runs to 1.2 x the longest bar, room for its label
Y_TICK_GAPS = 5                # a $K axis has at most 5 gaps between labelled gridlines (so 6 labels)
Y_MIN_TICKS = 3                # ...and at least 3 labels
Y_MARGIN = 0.05                # 5% of the range past the highest and lowest value, so no bar ends on the frame
LABEL_GAP_PX = 4               # quarter labels closer than this count as touching
GAP_LABEL = REASON_DISPLAY[MISSING_INPUT]   # "data missing"
NAVY_HEX, MID_GRAY_HEX, SLATE_HEX, WHITE_HEX = css_color(NAVY), css_color(MID_GRAY), css_color(SLATE), css_color(WHITE)
AMBER_HEX, RED_HEX = css_color(AMBER), css_color(RED)
HATCH = "///"                  # diagonal stripes: a shrinking quarter or a tripped bar, never color alone
SHRANK_LABEL = "ARR shrank"    # the legend's words for an amber, hatched bar
TRIPPED_WORD = "tripped"       # the runway chart says it in words as well as red
LABEL_BACKGROUND = {"facecolor": WHITE_HEX, "edgecolor": "none", "pad": 1}  # keeps a label readable over lines


# ---------------------------------------------------------------------------
# One style for every axis
# ---------------------------------------------------------------------------

def new_figure(size_inches):
    """An empty figure at the size it will have on the slide, with its own drawing canvas.

    The canvas lets finish_layout measure labels before the file is saved.
    """
    figure = Figure(figsize=size_inches, constrained_layout=True)
    FigureCanvasAgg(figure)
    return figure


def quiet_axes(axis, grid_axis):
    """No top/right border, mid gray left/bottom border, a light grid along `grid_axis`, slate tick labels."""
    for side in ("top", "right"):
        axis.spines[side].set_visible(False)
    axis.spines["left"].set_color(MID_GRAY_HEX)
    axis.spines["bottom"].set_color(MID_GRAY_HEX)
    axis.grid(axis=grid_axis, color=MID_GRAY_HEX, alpha=0.25, linewidth=0.8)
    axis.set_axisbelow(True)
    axis.tick_params(labelsize=FONT_SIZE, colors=SLATE_HEX)


def panel_title(axis, text):
    """A chart's title: bold slate, at the top left."""
    axis.set_title(text, loc="left", fontsize=FONT_SIZE, color=SLATE_HEX, fontweight="bold")


def quarter_axis(axis, quarters):
    """One slot per quarter, labels on two lines ("Q2\\n2026"), half a slot of room at either end."""
    axis.set_xticks(range(len(quarters)))
    axis.set_xticklabels([quarter.replace(" ", "\n") for quarter in quarters])
    axis.set_xlim(-X_PADDING, len(quarters) - 1 + X_PADDING)


def value_range(values):
    """(lowest, highest) a $K axis must show: every value and zero, plus a 5% margin past the data.

    A range under 3 ($K), e.g. all zero, is widened to 3 above the lowest value: whole-$K ticks need
    at least 3 whole numbers inside it, or matplotlib falls back to 0.25 steps that print as "0".
    """
    present = [value for value in values if not math.isnan(value)] + [0]
    low, high = min(present), max(present)
    margin = (high - low) * Y_MARGIN
    low, high = (low - margin if low < 0 else low), high + margin
    return low, max(high, low + Y_MIN_TICKS)


def money_ticks(values):
    """The tick values for a $K axis: whole $K, 3 to 6 round numbers, zero among them, covering every value."""
    locator = MaxNLocator(nbins=Y_TICK_GAPS, integer=True, min_n_ticks=Y_MIN_TICKS, steps=[1, 2, 2.5, 5, 10])
    return list(locator.tick_values(*value_range(values)))


def money_axis(axis, values, column):
    """The $K axis every chart shares: its ticks are its limits, so it starts and ends on a labelled gridline."""
    ticks = money_ticks(values)
    axis.yaxis.set_major_locator(FixedLocator(ticks))
    axis.set_ylim(ticks[0], ticks[-1])
    axis.yaxis.set_major_formatter(FuncFormatter(lambda value, _: format_value(column, value)))


# ---------------------------------------------------------------------------
# Labels
# ---------------------------------------------------------------------------

def label_gaps(axis, values):
    """Write "data missing" at every blank quarter, halfway up the panel (finish_layout may turn it on its side)."""
    for position, value in enumerate(values):
        if math.isnan(value):
            axis.text(position, 0.5, GAP_LABEL.replace(" ", "\n"), transform=axis.get_xaxis_transform(),
                      ha="center", va="center", fontsize=FONT_SIZE, color=MID_GRAY_HEX, style="italic",
                      bbox=LABEL_BACKGROUND, gid="gap")


def label_latest(axis, values, column, x_offset):
    """Write the latest value just right of the latest bar or point: nothing is ever drawn there.

    x_offset: how far right of the quarter's centre the bar or point ends (half a bar, or 0 for a point).
    """
    position, value = len(values) - 1, values[-1]
    if math.isnan(value):
        return  # the gap label already says why
    axis.annotate(format_value(column, value), (position + x_offset, value), xytext=(LABEL_OFFSET_PT, 0),
                  textcoords="offset points", ha="left", va="center", fontsize=FONT_SIZE, color=SLATE_HEX,
                  fontweight="bold", bbox=LABEL_BACKGROUND)


def quarter_slot_px(axis):
    """The width of one quarter's slot, in pixels."""
    left, right = axis.transData.transform([(0, 0), (1, 0)])[:, 0]
    return abs(right - left)


def labels_fit(boxes, step):
    """True if, showing every `step`-th label counted back from the latest, no two shown labels touch."""
    shown = boxes[::-1][::step]
    return all(right.x0 - left.x1 >= LABEL_GAP_PX for right, left in zip(shown, shown[1:]))


def thin_quarter_labels(axis, renderer):
    """Hide quarter labels, every second then every third..., until the ones left don't touch. The latest stays."""
    labels = axis.get_xticklabels()
    boxes = [label.get_window_extent(renderer) for label in labels]
    step = next(step for step in range(1, len(labels) + 1) if labels_fit(boxes, step))
    for index, label in enumerate(labels):
        label.set_visible((len(labels) - 1 - index) % step == 0)


def turn_narrow_gap_labels(axis, renderer):
    """A "data missing" wider than its quarter's slot is written on one line, turned on its side."""
    for text in axis.texts:
        if text.get_gid() == "gap" and text.get_window_extent(renderer).width > quarter_slot_px(axis):
            text.set_text(GAP_LABEL)
            text.set_rotation(90)


def title_fits(axis, renderer):
    """True if the panel's title ends before the right edge of the figure."""
    return axis._left_title.get_window_extent(renderer).x1 <= axis.get_figure().bbox.x1


def wrap_title(axis, renderer):
    """Break a title that runs off the figure into shorter lines, a few characters narrower each try.

    The cash chart's runway lines can be long ("At next quarter's budgeted burn: ∞ (budget not burning)").
    """
    title = axis._left_title   # the title set with loc="left" (matplotlib keeps it apart from the centre one)
    lines = title.get_text().split("\n")
    width = max(len(line) for line in lines)
    while not title_fits(axis, renderer) and width > 10:
        width -= 5
        title.set_text("\n".join(textwrap.fill(line, width) for line in lines))
        axis.get_figure().draw_without_rendering()   # lay out again: the title now takes more rows


def finish_layout(figure):
    """Lay the figure out once, then fix what only measuring can tell: crowded quarter labels, gap labels, long titles."""
    figure.draw_without_rendering()
    renderer = figure.canvas.get_renderer()
    for axis in figure.axes:
        wrap_title(axis, renderer)
        turn_narrow_gap_labels(axis, renderer)
        thin_quarter_labels(axis, renderer)
    return figure


# ---------------------------------------------------------------------------
# The charts
# ---------------------------------------------------------------------------

def shrank_legend(axis):
    """A one-entry legend above the panel's top right: the amber, hatched box means "ARR shrank"."""
    key = Patch(facecolor=AMBER_HEX, hatch=HATCH, edgecolor=WHITE_HEX, label=SHRANK_LABEL)
    legend = axis.legend(handles=[key], loc="lower right", bbox_to_anchor=(1, 1), frameon=False,
                         fontsize=FONT_SIZE, borderaxespad=0.2, handlelength=1.5)
    for text in legend.get_texts():
        text.set_color(SLATE_HEX)


def bar_panel(axis, quarters, values, column):
    """Bars for every quarter that has a value; nothing (plus a label) for a blank one.

    A bar below zero is amber and hatched (ARR shrank), so it differs from navy by lightness and
    pattern, never by hue alone; the legend names it only when there is one.
    """
    present = [position for position, value in enumerate(values) if not math.isnan(value)]
    shrank = [values[position] < 0 for position in present]
    axis.bar(present, [values[position] for position in present], width=BAR_WIDTH,
             color=[AMBER_HEX if below else NAVY_HEX for below in shrank],
             hatch=[HATCH if below else None for below in shrank], hatchcolor=WHITE_HEX, linewidth=0)
    axis.axhline(0, color=MID_GRAY_HEX, linewidth=0.8)
    quiet_axes(axis, "y")
    quarter_axis(axis, quarters)
    money_axis(axis, values, column)
    panel_title(axis, METRIC_LABELS[column])
    if any(shrank):
        shrank_legend(axis)
    label_gaps(axis, values)
    label_latest(axis, values, column, BAR_WIDTH / 2)


def arr_chart(quarters, ending_arr, net_new_arr, size_inches):
    """Figure with two panels: ending ARR bars (taller) above net new ARR bars."""
    figure = new_figure(size_inches)
    arr_axis, net_new_axis = figure.subplots(2, 1, gridspec_kw={"height_ratios": [3, 2]})
    bar_panel(arr_axis, quarters, list(ending_arr), "ending_arr")
    bar_panel(net_new_axis, quarters, list(net_new_arr), "net_new_arr")
    return finish_layout(figure)


def cash_chart(quarters, ending_cash, runway_text, size_inches):
    """Figure with ending cash as a line that breaks at a blank quarter, and runway in the title."""
    figure = new_figure(size_inches)
    axis = figure.subplots()
    values = list(ending_cash)
    # plot() gets the NaN too: matplotlib leaves a gap there instead of joining the neighbours.
    axis.plot(range(len(values)), values, color=NAVY_HEX, linewidth=LINE_WIDTH, marker="o", markersize=MARKER_SIZE)
    quiet_axes(axis, "y")
    quarter_axis(axis, quarters)
    money_axis(axis, values, "ending_cash")   # zero always on it: cash running out means reaching zero
    panel_title(axis, f"{INPUT_LABELS['ending_cash']}\n{runway_text}")
    label_gaps(axis, values)
    label_latest(axis, values, "ending_cash", 0)
    return finish_layout(figure)


def runway_chart(companies, runways, texts, tripped, threshold, threshold_text, size_inches):
    """The rollup's chart (rollup.py): one horizontal bar per company, months of runway, first company on top.

    companies, runways (months; NaN = no number, inf = not burning), texts (what each bar says, from
    build_deck.value_text) and tripped (the runway flag's result) are in the same order. Only a real
    number gets a bar: an infinite or missing runway gets its words instead, at the start of the row.
    Tripped bars are red and hatched (a flag's status, the only thing red means), the rest navy. The
    dashed line is the runway flag's threshold from config.yaml.
    """
    figure = new_figure(size_inches)
    axis = figure.subplots()
    rows = range(len(companies))
    drawn = [row for row in rows if math.isfinite(runways[row])]
    axis.barh(drawn, [runways[row] for row in drawn], height=BAR_WIDTH,
              color=[RED_HEX if tripped[row] else NAVY_HEX for row in drawn],
              hatch=[HATCH if tripped[row] else None for row in drawn], hatchcolor=WHITE_HEX, linewidth=0)
    for row in rows:   # the value past the end of its bar, or the words where a bar would start
        at = runways[row] if row in drawn else 0
        text = f"{texts[row]}, {TRIPPED_WORD}" if tripped[row] else texts[row]   # never color alone
        axis.annotate(text, (at, row), xytext=(LABEL_OFFSET_PT, 0), textcoords="offset points", va="center",
                      fontsize=FONT_SIZE, color=SLATE_HEX, fontweight="bold" if tripped[row] else None,
                      bbox=LABEL_BACKGROUND, zorder=3)   # above the threshold line
    axis.axvline(threshold, color=SLATE_HEX, linestyle="--", linewidth=1, zorder=2)
    axis.text(threshold, 1.0, f" {threshold_text}", transform=axis.get_xaxis_transform(), va="bottom",
              fontsize=FONT_SIZE, color=SLATE_HEX)
    axis.set_yticks(list(rows), companies)
    axis.invert_yaxis()   # first company at the top, as in a ranked list
    longest = max([runways[row] for row in drawn] + [threshold])
    axis.set_xlim(0, longest * HEADROOM)   # room for the last label past the longest bar
    axis.xaxis.set_major_formatter(FuncFormatter(lambda value, _: format_value("runway_months", value)))
    quiet_axes(axis, "x")
    return figure


def save_chart(figure, path):
    """Save a figure as PNG at its own size (so 12 pt in the figure is 12 pt on the slide).

    The figure isn't registered with pyplot, so there's nothing to close: Python frees it once unused.
    """
    figure.savefig(path, dpi=DPI, facecolor=WHITE_HEX)
    return path
