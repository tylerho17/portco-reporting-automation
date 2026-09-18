"""Unit tests for text_fit.py: measuring text, shrinking it to fit, failing loudly when it can't.

Run from the project folder:  pytest
"""

import pytest

from text_fit import (LINE_SPACING, MIN_FONT_PT, TextDoesNotFitError, count_lines, fit_table, paragraph,
                      shrink_to_fit, text_height_pt, text_width_pt)


def test_wider_text_measures_wider_and_bold_is_wider_still():
    assert text_width_pt("ARR", 12) < text_width_pt("Ending ARR", 12)
    assert text_width_pt("Ending ARR", 12) < text_width_pt("Ending ARR", 12, bold=True)


def test_width_scales_with_font_size():
    # Twice the font size -> twice the width.
    assert text_width_pt("Runway", 24) == pytest.approx(2 * text_width_pt("Runway", 12))


def test_short_text_is_one_line_and_empty_text_still_takes_a_line():
    assert count_lines("Wins", 500, 14) == 1
    assert count_lines("", 500, 14) == 1


def test_text_wraps_when_wider_than_the_box():
    text = "word " * 40
    one_line_width = text_width_pt(text.strip(), 14)
    # A box a quarter as wide needs at least 4 lines.
    assert count_lines(text.strip(), one_line_width / 4, 14) >= 4


def test_a_single_word_wider_than_the_box_breaks_across_lines():
    word = "x" * 60
    assert count_lines(word, text_width_pt(word, 14) / 3, 14) >= 3


def test_height_adds_lines_and_spacing():
    paragraphs = [paragraph("One", 20, space_after=6), paragraph("Two", 10)]
    # 1 line at 20 pt + 6 pt after + 1 line at 10 pt
    assert text_height_pt(paragraphs, 1000) == pytest.approx(20 * LINE_SPACING + 6 + 10 * LINE_SPACING)


def test_shrink_keeps_sizes_when_the_text_already_fits():
    paragraphs = [paragraph("Headline", 24, bold=True), paragraph("Detail", 16)]
    assert [p["size"] for p in shrink_to_fit(paragraphs, 1000, 1000, "test box")] == [24, 16]


def test_shrink_lowers_every_size_together_until_it_fits():
    paragraphs = [paragraph("Heading", 16, bold=True), paragraph("word " * 30, 14)]
    width = text_width_pt("word " * 12, 14)
    needed_at_12 = text_height_pt([paragraph("Heading", 14, bold=True), paragraph("word " * 30, 12)], width)
    fitted = shrink_to_fit(paragraphs, width, needed_at_12, "test box")
    assert [p["size"] for p in fitted][1] < 14          # it had to shrink
    assert fitted[0]["size"] - fitted[1]["size"] == 2   # the heading stays 2 pt bigger
    assert text_height_pt(fitted, width) <= needed_at_12


def test_text_that_cannot_fit_at_the_floor_fails_loudly_and_names_the_box():
    paragraphs = [paragraph("word " * 200, 18)]
    with pytest.raises(TextDoesNotFitError, match="Slide 9, headline"):
        shrink_to_fit(paragraphs, 100, 30, "Slide 9, headline")


def test_shrink_never_goes_below_the_floor():
    assert MIN_FONT_PT == 12
    paragraphs = [paragraph("word " * 20, 13)]
    width = text_width_pt("word " * 20, 12)  # fits on one line only at 12 pt
    fitted = shrink_to_fit(paragraphs, width, 12 * LINE_SPACING, "test box")
    assert fitted[0]["size"] == 12


def test_text_that_starts_below_the_floor_fails_loudly():
    # Review finding 3: shrinking can only make text smaller, so text handed in below the floor
    # can never be made to comply. It used to be returned unchanged, at a size no one can read.
    paragraphs = [paragraph("short", MIN_FONT_PT - 2)]
    with pytest.raises(TextDoesNotFitError, match=f"below the {MIN_FONT_PT} pt minimum"):
        shrink_to_fit(paragraphs, 1000, 1000, "Slide 9, headline")


def test_fit_table_returns_one_size_and_a_height_per_row():
    rows = [["Metric", "Value"], ["NRR (annualized)", "97.1%"], ["GRR (annualized)", "88.1%"]]
    size, heights = fit_table(rows, [200, 100], 1000, "test table", start_size=14)
    assert size == 14 and len(heights) == 3
    assert all(height > 0 for height in heights)


def test_fit_table_shrinks_then_fails_loudly():
    rows = [["a long metric name " * 3, "value"]] * 30
    with pytest.raises(TextDoesNotFitError, match="test table"):
        fit_table(rows, [100, 50], 200, "test table", start_size=14)
