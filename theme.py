"""The one palette, font and type scale: every deck, memo, chart and web page takes its look from here.

No other file types a color. The brand is the fictional "Example Capital" (navy and grays);
red and green mean a flag's status and nothing else, so no button is ever red or green.

- Palette: navy, navy dark, slate, mid gray, line, surface, white; red and green each with a light fill,
  and a gray fill, for tripped / passed / cannot evaluate.
- The Excel workbook keeps Excel's own "Bad" / "Good" fills (EXCEL_STATUS_COLORS), as before.
- Font: Arial, else Helvetica, else DejaVu Sans (it ships with matplotlib, so it is always there).
- Sizes (pt on slides, px on the web page): title 28, section 20, body 15, caption 13, table 14,
  and never below the 12 pt floor.
- The web page: streamlit_css() styles buttons, cards and tables; streamlit_theme() is what
  .streamlit/config.toml must say (tests/test_theme.py checks the two match).
"""

from functools import lru_cache
from pathlib import Path

from matplotlib import font_manager

from metrics import CANNOT_EVALUATE, PASS, TRIP

BRAND_NAME = "Example Capital"   # fictional

# ---------------------------------------------------------------------------
# Palette (6-digit hex, no "#": python-pptx, python-docx and openpyxl take it this way)
# ---------------------------------------------------------------------------

NAVY = "0B2545"         # titles, table headers, primary buttons, chart bars
NAVY_DARK = "08192F"    # a primary button under the mouse
SLATE = "334155"        # body text
MID_GRAY = "64748B"     # captions, footers, notes, disabled text
LINE = "E2E8F0"         # borders and rules
SURFACE = "F8FAFC"      # page background, table stripes, disabled buttons
WHITE = "FFFFFF"        # cards, table rows, secondary buttons
RED, RED_FILL = "C0392B", "FDE8E6"        # tripped
GREEN, GREEN_FILL = "1E8449", "EAF6EF"    # passed
GRAY_FILL = "EDF0F3"                      # cannot evaluate / data missing

# Status -> (fill, text): the deck, the memo and the web page.
STATUS_COLORS = {
    TRIP: (RED_FILL, RED),
    PASS: (GREEN_FILL, GREEN),
    CANNOT_EVALUATE: (GRAY_FILL, SLATE),
}

# Status -> (fill, text) in the metrics workbook: Excel's own light red / green for "Bad" / "Good".
EXCEL_STATUS_COLORS = {
    TRIP: ("FFC7CE", "9C0006"),
    PASS: ("C6EFCE", "006100"),
    CANNOT_EVALUATE: ("D9D9D9", "404040"),
}


def css_color(value):
    """'0B2545' -> '#0B2545' (web pages and matplotlib want the #)."""
    return f"#{value}"


# ---------------------------------------------------------------------------
# Font and sizes
# ---------------------------------------------------------------------------

FONT = "Arial"
FONT_STACK = ("Arial", "Helvetica", "DejaVu Sans")   # first one installed wins

TITLE_PT = 28
SECTION_PT = 20
BODY_PT = 15
CAPTION_PT = 13
TABLE_PT = 14
MIN_PT = 12             # the floor: text_fit.py never shrinks below it


def css_font_stack():
    """"Arial, Helvetica, 'DejaVu Sans', sans-serif": names with a space are quoted on a web page."""
    names = [f"'{name}'" if " " in name else name for name in FONT_STACK]
    return ", ".join(names + ["sans-serif"])


def installed(family, bold=False):
    """The font file for a family if matplotlib can find it on this computer, else None."""
    properties = font_manager.FontProperties(family=family, weight="bold" if bold else "normal")
    try:
        return font_manager.findfont(properties, fallback_to_default=False)
    except ValueError:
        return None


def first_installed_font():
    """The family charts are drawn in: the first of FONT_STACK installed here (matplotlib picks the same way)."""
    return next(family for family in FONT_STACK if installed(family))


@lru_cache(maxsize=None)  # look each one up once
def font_file(bold=False):
    """(family, path) of the first font in FONT_STACK installed as a single .ttf file.

    For the PDF memo, which embeds the font. A Mac keeps Helvetica in a .ttc (several fonts in one
    file), which the PDF library can't embed without knowing which one is bold, so it is skipped there.
    DejaVu Sans ships with matplotlib, so the loop always finds something.
    """
    for family in FONT_STACK:
        path = installed(family, bold)
        if path and Path(path).suffix.lower() == ".ttf":
            return family, path
    raise FileNotFoundError(f"None of {', '.join(FONT_STACK)} is installed")


# ---------------------------------------------------------------------------
# The web page (Streamlit)
# ---------------------------------------------------------------------------

PAGE_WIDTH_PX = 1100
ROW_HEIGHT_PX = 40
BUTTON_RADIUS_PX = 6
BUTTON_PADDING = "10px 18px"       # top and bottom 10, left and right 18

# Streamlit marks each button with its kind: <button data-testid="stBaseButton-primary">.
PRIMARY = 'button[data-testid="stBaseButton-primary"]'
SECONDARY = 'button[data-testid="stBaseButton-secondary"]'
TERTIARY = 'button[data-testid="stBaseButton-tertiary"]'


def css_rule(selector, **declarations):
    """One CSS rule: css_rule("h1", font_size="28px") -> 'h1 { font-size: 28px; }'."""
    body = " ".join(f"{name.replace('_', '-')}: {value};" for name, value in declarations.items())
    return f"{selector} {{ {body} }}"


def page_rules():
    """Arial everywhere, the type sizes, one column about 1100 px wide on the surface color."""
    return [
        css_rule("html, body, .stApp, button, input, textarea", font_family=css_font_stack()),
        css_rule(".stApp", background_color=css_color(SURFACE), color=css_color(SLATE)),
        css_rule('[data-testid="stMainBlockContainer"], .block-container', max_width=f"{PAGE_WIDTH_PX}px",
                 padding_top="2rem"),
        css_rule(".stApp h1", font_size=f"{TITLE_PT}px", color=css_color(NAVY)),
        css_rule(".stApp h3", font_size=f"{SECTION_PT}px", color=css_color(NAVY)),
        css_rule(".stApp p, .stApp li, .stApp label", font_size=f"{BODY_PT}px"),
        css_rule('[data-testid="stCaptionContainer"], [data-testid="stCaptionContainer"] p',
                 font_size=f"{CAPTION_PT}px", color=css_color(MID_GRAY)),
    ]


def button_rules():
    """Primary: navy, white text. Secondary: white, navy text and border. Never red or green."""
    shape = {"border_radius": f"{BUTTON_RADIUS_PX}px", "padding": BUTTON_PADDING}
    disabled = {"background_color": css_color(SURFACE), "color": css_color(MID_GRAY),
                "border": f"1px solid {css_color(LINE)}"}
    return [
        css_rule(PRIMARY, background_color=css_color(NAVY), color=css_color(WHITE),
                 border=f"1px solid {css_color(NAVY)}", **shape),
        css_rule(f"{PRIMARY}:hover", background_color=css_color(NAVY_DARK), color=css_color(WHITE),
                 border_color=css_color(NAVY_DARK)),
        css_rule(f"{PRIMARY}:disabled", **disabled),
        css_rule(SECONDARY, background_color=css_color(WHITE), color=css_color(NAVY),
                 border=f"1px solid {css_color(NAVY)}", **shape),
        css_rule(f"{SECONDARY}:hover", background_color=css_color(SURFACE), color=css_color(NAVY_DARK),
                 border_color=css_color(NAVY_DARK)),
        css_rule(f"{SECONDARY}:disabled", **disabled),
        css_rule(f"{TERTIARY}, {TERTIARY} p", color=css_color(NAVY), font_weight="bold"),
    ]


def card_and_table_rules():
    """White cards with a 1 px border; tables and the portfolio's rows 40 px tall under a navy header."""
    return [
        css_rule('[class*="st-key-card"]', background_color=css_color(WHITE), border=f"1px solid {css_color(LINE)}",
                 border_radius="8px", padding="16px 20px"),
        css_rule(".bp-table", width="100%", border_collapse="collapse", font_size=f"{TABLE_PT}px",
                 color=css_color(SLATE), background_color=css_color(WHITE)),
        css_rule(".bp-table th", background_color=css_color(NAVY), color=css_color(WHITE), height=f"{ROW_HEIGHT_PX}px",
                 padding="0 10px", text_align="left", font_weight="bold"),
        css_rule(".bp-table td", height=f"{ROW_HEIGHT_PX}px", padding="0 10px",
                 border_bottom=f"1px solid {css_color(LINE)}"),
        css_rule('[class*="st-key-portfolio-header"]', background_color=css_color(NAVY), color=css_color(WHITE),
                 min_height=f"{ROW_HEIGHT_PX}px", padding="0 8px", border_radius="4px"),
        css_rule('[class*="st-key-portfolio-header"] p', color=css_color(WHITE), font_weight="bold",
                 font_size=f"{TABLE_PT}px"),
        css_rule('[class*="st-key-portfolio-row"]', min_height=f"{ROW_HEIGHT_PX}px", padding="0 8px",
                 border_bottom=f"1px solid {css_color(LINE)}", font_size=f"{TABLE_PT}px"),
    ]


def streamlit_css():
    """The web page's style sheet, built from the palette above."""
    return "\n".join(page_rules() + button_rules() + card_and_table_rules())


def streamlit_theme():
    """What .streamlit/config.toml's [theme] must say: Streamlit's own accents (checkbox, progress bar, links)."""
    return {
        "base": "light",
        "primaryColor": css_color(NAVY),
        "backgroundColor": css_color(SURFACE),
        "secondaryBackgroundColor": css_color(WHITE),
        "textColor": css_color(SLATE),
        "linkColor": css_color(NAVY),
        "borderColor": css_color(LINE),
        "font": css_font_stack(),
        "baseFontSize": BODY_PT,
        "buttonRadius": f"{BUTTON_RADIUS_PX}px",
    }
