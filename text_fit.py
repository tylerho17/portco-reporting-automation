"""Measure slide text, shrink it until it fits its box, and fail loudly when it can't.

PowerPoint doesn't tell python-pptx how much room text takes, so this estimates it:
- Width: measured with a real font file (DejaVu Sans, which ships with matplotlib).
  DejaVu Sans is wider than Arial, the deck's font, so the estimate errs on the side of
  "needs more room" - text that fits here fits on the slide.
- Lines: words are placed left to right, starting a new line when the next word doesn't fit.
- Height: lines x font size x LINE_SPACING, plus the space after each paragraph.

Rule (CLAUDE.md / Task 2): shrink to a 12 pt floor, then stop with an error naming the slide
and the box. A board deck with text running off the page is worse than no deck.

build_deck.py uses this to size text before writing it; check_deck.py uses it to re-measure
the saved deck.
"""

import math
from functools import lru_cache

from matplotlib import font_manager
from PIL import ImageFont

MIN_FONT_PT = 12          # never shrink below this: smaller is hard to read in a board room
LINE_SPACING = 1.2        # one line of text is 1.2 x the font size tall (PowerPoint single spacing ~1.15)
MEASURE_SIZE = 100        # fonts are loaded once at this size and scaled (width grows in step with size)


class TextDoesNotFitError(Exception):
    """Text doesn't fit its box even at the minimum font size."""


# ---------------------------------------------------------------------------
# Measuring
# ---------------------------------------------------------------------------

@lru_cache(maxsize=None)  # load each font file once, not once per word
def measuring_font(bold):
    """DejaVu Sans (regular or bold) from matplotlib's own font folder."""
    weight = "bold" if bold else "normal"
    path = font_manager.findfont(font_manager.FontProperties(family="DejaVu Sans", weight=weight))
    return ImageFont.truetype(path, MEASURE_SIZE)


def text_width_pt(text, size_pt, bold=False):
    """How wide `text` is, in points, at `size_pt`."""
    return measuring_font(bold).getlength(text) * size_pt / MEASURE_SIZE


def count_lines(text, width_pt, size_pt, bold=False):
    """How many lines `text` wraps to in a box `width_pt` wide.

    Greedy, like PowerPoint: keep adding words to the line until the next one doesn't fit.
    A single word wider than the box is broken across as many lines as it needs.
    """
    lines, current = 1, ""
    for word in text.split(" "):
        candidate = f"{current} {word}" if current else word
        if text_width_pt(candidate, size_pt, bold) <= width_pt:
            current = candidate
            continue
        if current:
            lines += 1  # the word starts a new line
        lines += math.ceil(text_width_pt(word, size_pt, bold) / width_pt) - 1
        current = word  # conservative: an over-wide word leaves no room on its last line
    return lines


def paragraph(text, size, bold=False, color=None, space_after=0):
    """One paragraph to write: its text, font size (pt), boldness, color (hex) and space after (pt)."""
    return {"text": text, "size": size, "bold": bold, "color": color, "space_after": space_after}


def text_height_pt(paragraphs, width_pt):
    """Total height of a list of paragraphs in a box `width_pt` wide."""
    height = 0
    for item in paragraphs:
        lines = count_lines(item["text"], width_pt, item["size"], item["bold"])
        height += lines * item["size"] * LINE_SPACING + item["space_after"]
    return height


# ---------------------------------------------------------------------------
# Fitting
# ---------------------------------------------------------------------------

def preview(text):
    """The start of a text, for error messages."""
    return text if len(text) <= 60 else text[:57] + "..."


def shrink_to_fit(paragraphs, width_pt, height_pt, where):
    """Return the paragraphs with every font size lowered by the same amount until they fit.

    Sizes keep their differences (a heading stays bigger than its body text). Stops with
    TextDoesNotFitError once the smallest size would go below MIN_FONT_PT.
    """
    smallest = min(item["size"] for item in paragraphs)
    for shrink in range(max(smallest - MIN_FONT_PT, 0) + 1):
        sized = [{**item, "size": item["size"] - shrink} for item in paragraphs]
        if text_height_pt(sized, width_pt) <= height_pt:
            return sized
    first_text = " ".join(item["text"] for item in paragraphs)
    raise TextDoesNotFitError(f"{where}: text doesn't fit even at the {MIN_FONT_PT} pt minimum: "
                              f"'{preview(first_text)}'")


def row_heights_pt(rows, column_widths_pt, size_pt, cell_padding_pt):
    """Height of each table row: its tallest cell (header row bold) plus the cell's top and bottom padding."""
    heights = []
    for row_number, row in enumerate(rows):
        bold = row_number == 0
        tallest = max(count_lines(text, width, size_pt, bold) for text, width in zip(row, column_widths_pt))
        heights.append(tallest * size_pt * LINE_SPACING + cell_padding_pt)
    return heights


def fit_table(rows, column_widths_pt, height_pt, where, start_size, cell_padding_pt=0):
    """Largest font size (from start_size down to MIN_FONT_PT) at which the table fits `height_pt`.

    `column_widths_pt` is the room for text in each column (column width minus its side padding).
    Returns (size, row heights in pt). Stops with TextDoesNotFitError if it can't fit at the floor.
    """
    for size in range(start_size, MIN_FONT_PT - 1, -1):
        heights = row_heights_pt(rows, column_widths_pt, size, cell_padding_pt)
        if sum(heights) <= height_pt:
            return size, heights
    raise TextDoesNotFitError(f"{where}: {len(rows)} table rows don't fit even at the {MIN_FONT_PT} pt minimum")
