"""The evaluation set: 12 fictional companies, each a messy workbook with its own answer key.

Output: eval/data/<company>.xlsx (all money figures in $K). Run:  python eval/make_eval_data.py

Each company is one edge case the pipeline has to get right. The three companies in data/ tell
stories; these twelve test cases:

| Company      | Case                                                                                   |
|--------------|----------------------------------------------------------------------------------------|
| Larkspur     | healthy: no flag trips in any quarter                                                  |
| Quillmoor    | distressed: all 9 flags trip, including the combo (NRR falling, pipeline rising)        |
| Tidewell     | exactly at every threshold: all 8 threshold flags pass; NRR falls exactly 1 point a     |
|              | quarter, so the combo trips                                                             |
| Brackenfield | blank quarter first (Q3 2024): its own YoY is "missing input", not "no prior period"    |
| Copperlane   | blank quarter second to last (Q1 2026): the latest QoQ flag and the combo can't run     |
| Duskhaven    | blank quarter last (Q2 2026): every latest flag is "cannot evaluate: missing input"     |
| Emberfall    | zero revenue for a year before launch: 0 / 0 and growth from zero are "not meaningful"  |
| Glenmarsh    | negative budget: budgeted burn of 0 and below, and a budget that plans ARR to shrink     |
| Hollowmere   | NRR falling with pipeline falling too: the combo passes (not retention alone)            |
| Ivywick      | short history, 2 quarters: YoY and the combo are "no prior period", never a data gap     |
| Kestrelwood  | a quarter's row pasted twice: clean.py must stop and name both rows                     |
| Lanternreach | a quarter's row missing: clean.py must stop and name the quarter it expected            |

Northwind (blank in position 3) and Fernhollow (position 4) already cover a blank in the middle.
The blank-quarter and stop companies share Larkspur's numbers, so the only difference between them
is the one thing each tests.

The answer key, per company (worked out by hand from the numbers, never by running metrics.py):
- true_data, blank_quarter, next_quarter_budget: what clean.py must read back
- expected_latest: the 8 flag metrics in the latest quarter, as hand formulas (NaN = no number)
- expected_runway_at_budget: runway at next quarter's budgeted burn
- expected_flags: every flag's status in the latest quarter ("cannot evaluate: <reason>" says why)
- not_meaningful: {metric: [quarters]} where every input is there but the math is undefined.
  "Missing input" and "no prior period" are not listed: run_eval.py works them out from the
  blank quarter and the quarter's position, with the CLAUDE.md rules.
- never_trips (optional): no flag may trip in any quarter
- stop (optional): the words clean.py's error must contain; row_edit: the damage done to the workbook

The data gaps are not listed either: CLAUDE.md's rules predict them from the blank quarter and the
flags expected to be "cannot evaluate: missing input" (run_eval.predicted_gaps).
"""

import math
import sys
from pathlib import Path

from openpyxl import load_workbook

PROJECT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_DIR))   # so the project's own files import when this runs as a script

import make_data as northwind              # noqa: E402  (after the path line on purpose)
import make_data_alderpeak as alderpeak    # noqa: E402
import make_data_fernhollow as fernhollow  # noqa: E402
from make_data_common import save_workbook  # noqa: E402

DATA_DIR = Path(__file__).parent / "data"

EIGHT_QUARTERS = [
    "Q3 2024", "Q4 2024", "Q1 2025", "Q2 2025",
    "Q3 2025", "Q4 2025", "Q1 2026", "Q2 2026",
]

# The 9 flag names, typed by hand (not read from metrics.py), so a wrong label can't pass its own check.
FLAG_NAMES = [
    "NRR (annualized)", "GRR (annualized)", "Burn multiple", "Net burn vs budget", "Runway at current burn",
    "CAC payback", "Net new ARR vs budget", "Rule of 40", "NRR falling while pipeline rising",
]
ALL_PASS = {name: "pass" for name in FLAG_NAMES}
ALL_TRIP = {name: "trip" for name in FLAG_NAMES}
ALL_MISSING = {name: "cannot evaluate: missing input" for name in FLAG_NAMES}

NAN = math.nan


def notes(title, line):
    """A junk Notes tab: a title and one remark."""
    return {"A1": f"{title} - notes", "A3": line}


# ---------------------------------------------------------------------------
# Larkspur Analytics: healthy. NRR ~110%, GRR ~95%, burn shrinking and under budget, pipeline rising.
# Its numbers are shared by the blank-quarter and stop companies further down.
# ---------------------------------------------------------------------------

LARKSPUR_DATA = {
    "starting_arr":    [10000, 11250, 12580, 14000, 15500, 17080, 18750, 20520],
    "new_arr":         [1000, 1050, 1100, 1150, 1200, 1250, 1300, 1350],
    "expansion_arr":   [380, 420, 470, 520, 580, 640, 700, 770],
    "contraction_arr": [50, 60, 60, 70, 80, 90, 90, 100],
    "churned_arr":     [80, 80, 90, 100, 120, 130, 140, 150],
    "revenue":         [2660, 2980, 3320, 3690, 4070, 4480, 4910, 5360],
    "gross_profit":    [2070, 2320, 2590, 2880, 3170, 3490, 3830, 4180],
    "net_burn":        [500, 450, 400, 350, 300, 250, 200, 150],
    "ending_cash":     [13100, 12650, 12250, 11900, 11600, 11350, 11150, 11000],
    "sm_spend":        [1100, 1160, 1210, 1260, 1320, 1380, 1430, 1490],
    "new_customers":   [20, 21, 22, 23, 24, 25, 26, 27],
    "headcount":       [70, 73, 76, 79, 82, 85, 88, 91],
    "pipeline":        [4000, 4300, 4600, 4900, 5200, 5500, 5800, 6100],
    "budget_new_arr":  [980, 1030, 1080, 1130, 1180, 1230, 1280, 1330],
    "budget_arr":      [11150, 12460, 13860, 15340, 16900, 18550, 20300, 22150],
    "budget_net_burn": [540, 490, 440, 390, 340, 290, 240, 190],
}
LARKSPUR_NEXT_BUDGET = {"budget_new_arr": 1380, "budget_arr": 24100, "budget_net_burn": 120}

# Q2 2026 by hand. Net new ARR = 1350 + 770 - 100 - 150 = 1870; a year earlier (Q2 2025) revenue was 3690.
LARKSPUR_LATEST = {
    "nrr": 1 + 4 * (770 - 100 - 150) / 20520,
    "grr": 1 - 4 * (100 + 150) / 20520,
    "burn_multiple": 150 / 1870,
    "burn_vs_budget": 150 / 190 - 1,
    "runway_months": 11000 / (150 / 3),
    "cac_payback_months": 1490 / (1350 * 4180 / 5360) * 12,
    "net_new_arr_vs_budget": 1870 / (22150 - 20300) - 1,
    "rule_of_40": 5360 / 3690 - 1 - 150 / 5360,
}
LARKSPUR_RUNWAY_AT_BUDGET = 11000 / (120 / 3)

LARKSPUR = {
    "name": "Larkspur",
    "story": "healthy: no flag trips in any quarter",
    "quarters": EIGHT_QUARTERS,
    "blank_quarter": None,
    "budget_only_label": "Q3 2026 Plan",
    "true_data": LARKSPUR_DATA,
    "next_quarter_budget": LARKSPUR_NEXT_BUDGET,
    "header_names": alderpeak.HEADER_NAMES,
    "text_cells": {("pipeline", "Q3 2024"): "M", ("ending_cash", "Q2 2026"): "M",
                   ("revenue", "Q4 2025"): "comma", ("churned_arr", "Q1 2026"): "K"},
    "notes": notes("Larkspur Analytics", "Board pack due two weeks after quarter end"),
    "notes_first": True,
    "expected_latest": LARKSPUR_LATEST,
    "expected_runway_at_budget": LARKSPUR_RUNWAY_AT_BUDGET,
    "expected_flags": ALL_PASS,
    "not_meaningful": {},
    "never_trips": True,
}

# ---------------------------------------------------------------------------
# Quillmoor Software: distressed. Every flag trips. Unlike Fernhollow, net new ARR stays just above
# zero (burn multiple 28x, finite) and pipeline keeps rising, so the combo trips too.
# ---------------------------------------------------------------------------

QUILLMOOR = {
    "name": "Quillmoor",
    "story": "distressed: all 9 flags trip, the combo too (NRR falling, pipeline rising)",
    "quarters": EIGHT_QUARTERS,
    "blank_quarter": None,
    "budget_only_label": "Q3 2026 (Budget)",
    "true_data": {
        "starting_arr":    [9000, 9660, 10240, 10730, 11130, 11430, 11630, 11710],
        "new_arr":         [700, 680, 650, 620, 580, 540, 500, 560],
        "expansion_arr":   [120, 110, 100, 90, 80, 70, 60, 50],
        "contraction_arr": [60, 80, 100, 120, 140, 160, 190, 220],
        "churned_arr":     [100, 130, 160, 190, 220, 250, 290, 330],
        "revenue":         [2300, 2400, 2480, 2540, 2580, 2600, 2610, 2600],
        "gross_profit":    [1560, 1600, 1620, 1630, 1620, 1600, 1570, 1530],
        "net_burn":        [1300, 1350, 1400, 1450, 1500, 1560, 1620, 1700],
        "ending_cash":     [14780, 13430, 12030, 10580, 9080, 7520, 5900, 4200],
        "sm_spend":        [1300, 1350, 1400, 1450, 1500, 1550, 1600, 1650],
        "new_customers":   [25, 24, 23, 22, 20, 18, 16, 14],
        "headcount":       [120, 124, 128, 131, 133, 134, 134, 133],
        "pipeline":        [5000, 5200, 5400, 5700, 6000, 6400, 6900, 7500],
        "budget_new_arr":  [700, 720, 740, 760, 780, 800, 820, 840],
        "budget_arr":      [9700, 10400, 11100, 11800, 12500, 13200, 13900, 14600],
        "budget_net_burn": [1300, 1300, 1300, 1300, 1300, 1300, 1300, 1300],
    },
    "next_quarter_budget": {"budget_new_arr": 600, "budget_arr": 15300, "budget_net_burn": 1300},
    "header_names": northwind.HEADER_NAMES,
    "text_cells": {("ending_cash", "Q2 2026"): "M", ("pipeline", "Q1 2026"): "M", ("sm_spend", "Q2 2026"): "comma"},
    "notes": notes("Quillmoor Software", "Largest reseller did not renew"),
    # Q2 2026: net new ARR = 560 + 50 - 220 - 330 = 60; Q2 2025 revenue was 2540.
    "expected_latest": {
        "nrr": 1 + 4 * (50 - 220 - 330) / 11710,
        "grr": 1 - 4 * (220 + 330) / 11710,
        "burn_multiple": 1700 / 60,
        "burn_vs_budget": 1700 / 1300 - 1,
        "runway_months": 4200 / (1700 / 3),
        "cac_payback_months": 1650 / (560 * 1530 / 2600) * 12,
        "net_new_arr_vs_budget": 60 / (14600 - 13900) - 1,
        "rule_of_40": 2600 / 2540 - 1 - 1700 / 2600,
    },
    "expected_runway_at_budget": 4200 / (1300 / 3),
    "expected_flags": ALL_TRIP,
    "not_meaningful": {},
}

# ---------------------------------------------------------------------------
# Tidewell Systems: exactly at every threshold in Q2 2026 (config.yaml: exactly at the threshold passes).
# NRR is 1.02, 1.01, 1.00 over the last 3 quarters: each step falls exactly combo_min_nrr_drop (1 point),
# which counts as falling, and pipeline rises, so the combo trips.
# ---------------------------------------------------------------------------

TIDEWELL = {
    "name": "Tidewell",
    "story": "exactly at every threshold: all 8 pass; NRR falls exactly 1 point a quarter, so the combo trips",
    "quarters": EIGHT_QUARTERS,
    "blank_quarter": None,
    "budget_only_label": "Q3 2026 - Bud",
    "true_data": {
        "starting_arr":    [35120, 38820, 42760, 46940, 51360, 56000, 60000, 64000],
        "new_arr":         [3000, 3200, 3400, 3600, 3800, 3720, 3850, 4600],
        "expansion_arr":   [1500, 1600, 1700, 1800, 1900, 2000, 1600, 2400],
        "contraction_arr": [300, 320, 340, 360, 400, 420, 450, 800],
        "churned_arr":     [500, 540, 580, 620, 660, 1300, 1000, 1600],
        "revenue":         [10400, 11200, 11800, 12100, 15000, 17000, 19500, 22000],
        "gross_profit":    [7800, 8400, 8850, 9080, 11250, 12750, 14630, 16500],
        "net_burn":        [6000, 6400, 6800, 7200, 7600, 8000, 8600, 9200],
        "ending_cash":     [90600, 84200, 77400, 70200, 62600, 54600, 46000, 36800],
        "sm_spend":        [4000, 4200, 4400, 4600, 5000, 5400, 6000, 6900],
        "new_customers":   [60, 62, 64, 66, 68, 66, 68, 80],
        "headcount":       [300, 310, 320, 330, 340, 350, 360, 380],
        "pipeline":        [15000, 16000, 17000, 18000, 19000, 20000, 21000, 22000],
        "budget_new_arr":  [3000, 3200, 3400, 3600, 3800, 4000, 4200, 4400],
        "budget_arr":      [39000, 43000, 47000, 51000, 55000, 59000, 63000, 68750],
        "budget_net_burn": [6000, 6200, 6400, 6600, 6800, 7000, 7500, 8000],
    },
    "next_quarter_budget": {"budget_new_arr": 4800, "budget_arr": 74500, "budget_net_burn": 9600},
    "header_names": fernhollow.HEADER_NAMES,
    "text_cells": {("starting_arr", "Q2 2026"): "M", ("ending_cash", "Q2 2026"): "M",
                   ("revenue", "Q2 2026"): "comma", ("budget_arr", "Q2 2026"): "M"},
    "notes": notes("Tidewell Systems", "Plan assumes the Series C closes in Q4"),
    "title": "Tidewell Systems - KPI tracker ($K)",
    # Q2 2026: net new ARR = 4600 + 2400 - 800 - 1600 = 4600; Q2 2025 revenue was 12100.
    "expected_latest": {
        "nrr": 1 + 4 * (2400 - 800 - 1600) / 64000,          # 1.00
        "grr": 1 - 4 * (800 + 1600) / 64000,                 # 0.85
        "burn_multiple": 9200 / 4600,                        # 2.0
        "burn_vs_budget": 9200 / 8000 - 1,                   # 0.15
        "runway_months": 36800 / (9200 / 3),                 # 12
        "cac_payback_months": 6900 / (4600 * 16500 / 22000) * 12,  # 24
        "net_new_arr_vs_budget": 4600 / (68750 - 63000) - 1,       # -0.20
        "rule_of_40": 22000 / 12100 - 1 - 9200 / 22000,            # 0.40
    },
    "expected_runway_at_budget": 36800 / (9600 / 3),
    "expected_flags": {**ALL_PASS, "NRR falling while pipeline rising": "trip"},
    "not_meaningful": {},
}

# ---------------------------------------------------------------------------
# A blank quarter in three positions, all on Larkspur's numbers.
# ---------------------------------------------------------------------------

BRACKENFIELD = {
    "name": "Brackenfield",
    "story": "blank quarter first (Q3 2024): its own YoY is missing input, the next quarter's QoQ and "
             "Q3 2025's YoY too",
    "quarters": EIGHT_QUARTERS,
    "blank_quarter": "Q3 2024",
    "budget_only_label": "Q3 2026 Budget",
    "true_data": LARKSPUR_DATA,
    "next_quarter_budget": LARKSPUR_NEXT_BUDGET,
    "header_names": northwind.HEADER_NAMES,
    "text_cells": {("pipeline", "Q4 2024"): "M", ("gross_profit", "Q2 2026"): "comma"},
    "notes": notes("Brackenfield Data", "First quarter's close was lost in the system move"),
    "expected_latest": LARKSPUR_LATEST,   # the blank is too far back to reach the latest quarter
    "expected_runway_at_budget": LARKSPUR_RUNWAY_AT_BUDGET,
    "expected_flags": ALL_PASS,
    "not_meaningful": {},
}

COPPERLANE = {
    "name": "Copperlane",
    "story": "blank quarter second to last (Q1 2026): the latest net new ARR vs budget and the combo "
             "can't be evaluated",
    "quarters": EIGHT_QUARTERS,
    "blank_quarter": "Q1 2026",
    "budget_only_label": "Q3 2026 Plan",
    "true_data": LARKSPUR_DATA,
    "next_quarter_budget": LARKSPUR_NEXT_BUDGET,
    "header_names": alderpeak.HEADER_NAMES,
    "text_cells": {("ending_cash", "Q4 2025"): "M", ("new_arr", "Q2 2026"): "K"},
    "notes": notes("Copperlane Cloud", "Q1 2026 numbers held back by the auditors"),
    # Net new ARR vs budget needs Q1 2026's budget ARR; the rest use Q2 2026 (or Q2 2025) only.
    "expected_latest": {**LARKSPUR_LATEST, "net_new_arr_vs_budget": NAN},
    "expected_runway_at_budget": LARKSPUR_RUNWAY_AT_BUDGET,
    "expected_flags": {**ALL_PASS, "Net new ARR vs budget": "cannot evaluate: missing input",
                       "NRR falling while pipeline rising": "cannot evaluate: missing input"},
    "not_meaningful": {},
}

DUSKHAVEN = {
    "name": "Duskhaven",
    "story": "blank quarter last (Q2 2026): every latest flag is cannot evaluate: missing input",
    "quarters": EIGHT_QUARTERS,
    "blank_quarter": "Q2 2026",
    "budget_only_label": "Q3 2026 (Budget)",
    "true_data": LARKSPUR_DATA,
    "next_quarter_budget": LARKSPUR_NEXT_BUDGET,
    "header_names": fernhollow.HEADER_NAMES,
    "text_cells": {("pipeline", "Q1 2026"): "M", ("revenue", "Q3 2024"): "comma"},
    "notes": notes("Duskhaven Labs", "Latest quarter not closed yet"),
    "title": "Duskhaven Labs - quarterly KPIs",
    "expected_latest": {metric: NAN for metric in LARKSPUR_LATEST},
    "expected_runway_at_budget": NAN,   # the latest cash is blank, so there is nothing to divide
    "expected_flags": ALL_MISSING,
    "not_meaningful": {},
}

# ---------------------------------------------------------------------------
# Emberfall Robotics: zero revenue and zero ARR for a year before launch (Q3 2025).
# 0 / 0 (margins, NRR on no starting ARR, ARR vs a budget of 0) and growth from a zero base are
# "not meaningful". Burning with no ARR growth is the CLAUDE.md edge case: burn multiple infinite.
# ---------------------------------------------------------------------------

PRE_LAUNCH = ["Q3 2024", "Q4 2024", "Q1 2025", "Q2 2025"]
YEAR_AFTER_LAUNCH = ["Q3 2025", "Q4 2025", "Q1 2026", "Q2 2026"]

EMBERFALL = {
    "name": "Emberfall",
    "story": "zero revenue for a year before launch: 0 / 0 and growth from zero are not meaningful",
    "quarters": EIGHT_QUARTERS,
    "blank_quarter": None,
    "budget_only_label": "Q3 2026 Plan",
    "true_data": {
        "starting_arr":    [0, 0, 0, 0, 0, 400, 1000, 1810],
        "new_arr":         [0, 0, 0, 0, 400, 600, 800, 1000],
        "expansion_arr":   [0, 0, 0, 0, 0, 0, 20, 40],
        "contraction_arr": [0, 0, 0, 0, 0, 0, 0, 10],
        "churned_arr":     [0, 0, 0, 0, 0, 0, 10, 20],
        "revenue":         [0, 0, 0, 0, 50, 200, 350, 600],
        "gross_profit":    [0, 0, 0, 0, 30, 120, 210, 360],
        "net_burn":        [800, 850, 900, 950, 1000, 1050, 1100, 1150],
        "ending_cash":     [16200, 15350, 14450, 13500, 12500, 11450, 10350, 9200],
        "sm_spend":        [200, 200, 250, 300, 400, 500, 600, 700],
        "new_customers":   [0, 0, 0, 0, 8, 11, 14, 17],
        "headcount":       [12, 14, 16, 18, 22, 26, 30, 34],
        "pipeline":        [300, 500, 800, 1100, 1400, 1800, 2200, 2700],
        "budget_new_arr":  [0, 0, 0, 0, 500, 600, 900, 1100],
        "budget_arr":      [0, 0, 0, 0, 500, 1100, 2000, 3100],
        "budget_net_burn": [850, 900, 950, 1000, 1050, 1100, 1150, 1200],
    },
    "next_quarter_budget": {"budget_new_arr": 1200, "budget_arr": 4400, "budget_net_burn": 1250},
    "header_names": alderpeak.HEADER_NAMES,
    "text_cells": {("revenue", "Q3 2024"): "comma", ("starting_arr", "Q4 2024"): "M",
                   ("ending_cash", "Q2 2026"): "M"},
    "notes": notes("Emberfall Robotics", "Launch slipped from Q1 2025 to Q3 2025"),
    # Q2 2026: net new ARR = 1000 + 40 - 10 - 20 = 1010. A year earlier revenue was 0, so no Rule of 40.
    "expected_latest": {
        "nrr": 1 + 4 * (40 - 10 - 20) / 1810,
        "grr": 1 - 4 * (10 + 20) / 1810,
        "burn_multiple": 1150 / 1010,
        "burn_vs_budget": 1150 / 1200 - 1,
        "runway_months": 9200 / (1150 / 3),
        "cac_payback_months": 700 / (1000 * 360 / 600) * 12,
        "net_new_arr_vs_budget": 1010 / (3100 - 2000) - 1,
        "rule_of_40": NAN,
    },
    "expected_runway_at_budget": 9200 / (1250 / 3),
    "expected_flags": {**ALL_PASS, "Rule of 40": "cannot evaluate: not meaningful"},
    "not_meaningful": {
        "nrr": PRE_LAUNCH + ["Q3 2025"],            # starting ARR is 0 up to and including launch
        "grr": PRE_LAUNCH + ["Q3 2025"],
        "gross_margin": PRE_LAUNCH,                 # 0 gross profit / 0 revenue
        "fcf_margin": PRE_LAUNCH,                   # burn / 0 revenue
        "cac_payback_months": PRE_LAUNCH,           # uses the gross margin, which is 0 / 0
        "arr_vs_budget": PRE_LAUNCH,                # 0 ARR / 0 budget
        "arr_qoq": PRE_LAUNCH[1:] + ["Q3 2025"],    # 0 / 0, then 400 / 0 at launch
        "revenue_qoq": PRE_LAUNCH[1:] + ["Q3 2025"],
        "net_new_arr_vs_budget": PRE_LAUNCH[1:],    # budgeted net new ARR of 0
        "arr_yoy": YEAR_AFTER_LAUNCH,               # growth from a zero base a year earlier
        "revenue_yoy": YEAR_AFTER_LAUNCH,
        "rule_of_40": YEAR_AFTER_LAUNCH,
    },
}

# ---------------------------------------------------------------------------
# Glenmarsh Payroll: negative budget. The plan had burn falling to 0 in Q4 2025 and going negative
# (cash coming in) after, and ARR shrinking in Q2 2026 as a large customer left. Actual burn stayed
# positive. CLAUDE.md: a % of a budget of 0 or less means nothing, so both vs-budget flags are
# "cannot evaluate: not meaningful", and next quarter's budgeted burn of -150 gives infinite runway.
# ---------------------------------------------------------------------------

GLENMARSH = {
    "name": "Glenmarsh",
    "story": "negative budget: budgeted burn of 0 and below, and budgeted ARR shrinking",
    "quarters": EIGHT_QUARTERS,
    "blank_quarter": None,
    "budget_only_label": "Q3 2026 (Budget)",
    "true_data": {
        "starting_arr":    [30000, 30150, 30290, 30450, 30590, 30750, 30900, 31060],
        "new_arr":         [700, 700, 720, 720, 740, 740, 760, 760],
        "expansion_arr":   [200, 200, 210, 210, 220, 220, 230, 230],
        "contraction_arr": [150, 150, 150, 160, 160, 160, 170, 170],
        "churned_arr":     [600, 610, 620, 630, 640, 650, 660, 670],
        "revenue":         [7500, 7520, 7540, 7560, 7590, 7620, 7650, 7680],
        "gross_profit":    [5630, 5640, 5660, 5670, 5690, 5720, 5740, 5760],
        "net_burn":        [300, 280, 260, 240, 220, 200, 150, 270],
        "ending_cash":     [6620, 6340, 6080, 5840, 5620, 5420, 5270, 5000],
        "sm_spend":        [900, 900, 910, 910, 920, 920, 930, 930],
        "new_customers":   [14, 14, 15, 15, 15, 15, 16, 16],
        "headcount":       [210, 210, 208, 206, 204, 202, 200, 198],
        "pipeline":        [3000, 3100, 3000, 3100, 3000, 3100, 3000, 3100],
        "budget_new_arr":  [700, 700, 700, 700, 700, 700, 700, 700],
        "budget_arr":      [30100, 30200, 30300, 30400, 30500, 30600, 30700, 30500],
        "budget_net_burn": [300, 250, 200, 150, 100, 0, -100, -200],
    },
    "next_quarter_budget": {"budget_new_arr": 700, "budget_arr": 30600, "budget_net_burn": -150},
    "header_names": northwind.HEADER_NAMES,
    "text_cells": {("budget_net_burn", "Q2 2026"): "M", ("budget_net_burn", "Q1 2026"): "K",
                   ("starting_arr", "Q2 2026"): "comma"},   # "$-0.2M" and "-100K": negative text
    "notes": notes("Glenmarsh Payroll", "Plan: cash flow positive from Q1 2026"),
    # Q2 2026: net new ARR = 760 + 230 - 170 - 670 = 150; Q2 2025 revenue was 7560.
    "expected_latest": {
        "nrr": 1 + 4 * (230 - 170 - 670) / 31060,
        "grr": 1 - 4 * (170 + 670) / 31060,
        "burn_multiple": 270 / 150,
        "burn_vs_budget": NAN,              # budget -200
        "runway_months": 5000 / (270 / 3),
        "cac_payback_months": 930 / (760 * 5760 / 7680) * 12,
        "net_new_arr_vs_budget": NAN,       # budgeted net new ARR = 30500 - 30700 = -200
        "rule_of_40": 7680 / 7560 - 1 - 270 / 7680,
    },
    "expected_runway_at_budget": math.inf,
    "expected_flags": {**ALL_PASS, "NRR (annualized)": "trip", "Rule of 40": "trip",
                       "Net burn vs budget": "cannot evaluate: not meaningful",
                       "Net new ARR vs budget": "cannot evaluate: not meaningful"},
    "not_meaningful": {
        "burn_vs_budget": ["Q4 2025", "Q1 2026", "Q2 2026"],   # budgets of 0, -100, -200
        "net_new_arr_vs_budget": ["Q2 2026"],
    },
}

# ---------------------------------------------------------------------------
# Hollowmere Health IT: NRR falls every quarter (by 1.4 to 2.5 points in the last two steps) while
# pipeline falls too. The combo is "NRR falling WHILE pipeline rising", so it passes: this is a
# sales slowdown as well as a retention one, as at Fernhollow.
# ---------------------------------------------------------------------------

HOLLOWMERE = {
    "name": "Hollowmere",
    "story": "NRR falling with pipeline falling too: the combo passes",
    "quarters": EIGHT_QUARTERS,
    "blank_quarter": None,
    "budget_only_label": "Q3 2026 - Bud",
    "true_data": {
        "starting_arr":    [12000, 13240, 14540, 15850, 17140, 18410, 19660, 20880],
        "new_arr":         [900, 950, 1000, 1050, 1100, 1150, 1200, 1250],
        "expansion_arr":   [500, 520, 500, 460, 420, 380, 330, 280],
        "contraction_arr": [60, 60, 70, 80, 90, 100, 110, 120],
        "churned_arr":     [100, 110, 120, 140, 160, 180, 200, 230],
        "revenue":         [3100, 3380, 3680, 3990, 4300, 4620, 4940, 5270],
        "gross_profit":    [2360, 2570, 2800, 3030, 3270, 3510, 3750, 4000],
        "net_burn":        [700, 700, 720, 740, 760, 780, 800, 820],
        "ending_cash":     [20320, 19620, 18900, 18160, 17400, 16620, 15820, 15000],
        "sm_spend":        [1000, 1050, 1100, 1150, 1200, 1250, 1300, 1350],
        "new_customers":   [20, 21, 22, 23, 24, 25, 26, 27],
        "headcount":       [90, 94, 98, 102, 106, 110, 114, 118],
        "pipeline":        [6000, 6200, 6300, 6100, 5800, 5500, 5100, 4700],
        "budget_new_arr":  [900, 950, 1000, 1050, 1100, 1150, 1200, 1250],
        "budget_arr":      [13300, 14700, 16100, 17500, 18900, 20300, 21700, 23100],
        "budget_net_burn": [700, 720, 740, 760, 780, 800, 820, 840],
    },
    "next_quarter_budget": {"budget_new_arr": 1300, "budget_arr": 24500, "budget_net_burn": 860},
    "header_names": fernhollow.HEADER_NAMES,
    "text_cells": {("pipeline", "Q2 2026"): "K", ("expansion_arr", "Q2 2026"): "K"},
    "notes": notes("Hollowmere Health IT", "Two account managers left in Q4"),
    # Q2 2026: net new ARR = 1250 + 280 - 120 - 230 = 1180; Q2 2025 revenue was 3990.
    "expected_latest": {
        "nrr": 1 + 4 * (280 - 120 - 230) / 20880,
        "grr": 1 - 4 * (120 + 230) / 20880,
        "burn_multiple": 820 / 1180,
        "burn_vs_budget": 820 / 840 - 1,
        "runway_months": 15000 / (820 / 3),
        "cac_payback_months": 1350 / (1250 * 4000 / 5270) * 12,
        "net_new_arr_vs_budget": 1180 / (23100 - 21700) - 1,
        "rule_of_40": 5270 / 3990 - 1 - 820 / 5270,
    },
    "expected_runway_at_budget": 15000 / (860 / 3),
    "expected_flags": {**ALL_PASS, "NRR (annualized)": "trip", "Rule of 40": "trip"},
    "not_meaningful": {},
}

# ---------------------------------------------------------------------------
# Ivywick Security: only 2 quarters of history. YoY needs 4 quarters back and the combo needs 3
# quarters, so both are "no prior period": not a data gap, and never a pass.
# ---------------------------------------------------------------------------

IVYWICK = {
    "name": "Ivywick",
    "story": "short history (2 quarters): YoY and the combo have no prior period, which is never a data gap",
    "quarters": ["Q1 2026", "Q2 2026"],
    "blank_quarter": None,
    "budget_only_label": "Q3 2026 Budget",
    "true_data": {
        "starting_arr":    [5000, 5540],
        "new_arr":         [500, 560],
        "expansion_arr":   [100, 120],
        "contraction_arr": [20, 30],
        "churned_arr":     [40, 50],
        "revenue":         [1300, 1440],
        "gross_profit":    [1000, 1110],
        "net_burn":        [400, 380],
        "ending_cash":     [6380, 6000],
        "sm_spend":        [600, 640],
        "new_customers":   [10, 11],
        "headcount":       [40, 43],
        "pipeline":        [2000, 2300],
        "budget_new_arr":  [480, 520],
        "budget_arr":      [5500, 6100],
        "budget_net_burn": [420, 400],
    },
    "next_quarter_budget": {"budget_new_arr": 580, "budget_arr": 6750, "budget_net_burn": 380},
    "header_names": alderpeak.HEADER_NAMES,
    "text_cells": {("ending_cash", "Q2 2026"): "M"},
    "notes": notes("Ivywick Security", "Acquired January 2026, history before that not available"),
    # Q2 2026: net new ARR = 560 + 120 - 30 - 50 = 600.
    "expected_latest": {
        "nrr": 1 + 4 * (120 - 30 - 50) / 5540,
        "grr": 1 - 4 * (30 + 50) / 5540,
        "burn_multiple": 380 / 600,
        "burn_vs_budget": 380 / 400 - 1,
        "runway_months": 6000 / (380 / 3),
        "cac_payback_months": 640 / (560 * 1110 / 1440) * 12,
        "net_new_arr_vs_budget": 600 / (6100 - 5500) - 1,
        "rule_of_40": NAN,
    },
    "expected_runway_at_budget": 6000 / (380 / 3),
    "expected_flags": {**ALL_PASS, "Rule of 40": "cannot evaluate: no prior period",
                       "NRR falling while pipeline rising": "cannot evaluate: no prior period"},
    "not_meaningful": {},
}

# ---------------------------------------------------------------------------
# Two workbooks clean.py must refuse. Both are Larkspur's numbers with one row damaged after writing.
# ---------------------------------------------------------------------------

KESTRELWOOD = {
    "name": "Kestrelwood",
    "story": "a quarter's row pasted twice (Q4 2025): clean.py must stop and name both rows",
    "quarters": EIGHT_QUARTERS,
    "blank_quarter": None,
    "budget_only_label": "Q3 2026 (Budget)",
    "true_data": LARKSPUR_DATA,
    "next_quarter_budget": LARKSPUR_NEXT_BUDGET,
    "header_names": northwind.HEADER_NAMES,
    "text_cells": {},
    "notes": notes("Kestrelwood Logistics", "Pasted Q4 from the old tracker"),
    "row_edit": ("repeat", "Q4 2025"),
    # Header in row 1, Q3 2024 in row 2, so Q4 2025 is row 7 and its copy row 8.
    "stop": ["Sheet 'KPI Tracker', row 8: quarter 'Q4 2025' appears twice (also in row 7)"],
}

LANTERNREACH = {
    "name": "Lanternreach",
    "story": "a quarter's row missing (Q1 2026): clean.py must stop and name the quarter it expected",
    "quarters": EIGHT_QUARTERS,
    "blank_quarter": None,
    "budget_only_label": "Q3 2026 - Bud",
    "true_data": LARKSPUR_DATA,
    "next_quarter_budget": LARKSPUR_NEXT_BUDGET,
    "header_names": fernhollow.HEADER_NAMES,
    "text_cells": {},
    "notes": notes("Lanternreach Learning", "Deleted a row by mistake?"),
    "title": "Lanternreach Learning - KPIs",
    "row_edit": ("drop", "Q1 2026"),
    # Title in row 1, empty row 2, header row 3; Q3 2024 is row 4, so Q2 2026 moves up to row 10.
    "stop": ["Sheet 'KPI Tracker', row 10: After Q4 2025 expected Q1 2026, found Q2 2026"],
}

COMPANIES = [
    LARKSPUR, QUILLMOOR, TIDEWELL, BRACKENFIELD, COPPERLANE, DUSKHAVEN,
    EMBERFALL, GLENMARSH, HOLLOWMERE, IVYWICK, KESTRELWOOD, LANTERNREACH,
]
for _company in COMPANIES:   # each workbook is named after its company, e.g. eval/data/larkspur.xlsx
    _company["output_path"] = DATA_DIR / f"{_company['name'].lower()}.xlsx"


# ---------------------------------------------------------------------------
# Writing the workbooks
# ---------------------------------------------------------------------------

def row_of(sheet, quarter):
    """The Excel row number whose first cell is this quarter label."""
    return next(row[0].row for row in sheet.iter_rows() if row[0].value == quarter)


def damage_row(path, edit):
    """Repeat or drop one quarter's row in a saved workbook: ("repeat", "Q4 2025") or ("drop", "Q1 2026")."""
    action, quarter = edit
    workbook = load_workbook(path)
    sheet = workbook["KPI Tracker"]
    row = row_of(sheet, quarter)
    if action == "repeat":
        values = [cell.value for cell in sheet[row]]
        sheet.insert_rows(row + 1)   # an empty row right under it, then the same values
        for column, value in enumerate(values, start=1):
            sheet.cell(row=row + 1, column=column, value=value)
    elif action == "drop":
        sheet.delete_rows(row)
    else:
        raise ValueError(f"Unknown row edit: {action!r}")
    workbook.save(path)


def write_company(company, folder):
    """Check the answer key ties out, write the messy workbook into `folder`, then damage a row if asked."""
    path = Path(folder) / company["output_path"].name
    save_workbook({**company, "output_path": path})
    if company.get("row_edit"):
        damage_row(path, company["row_edit"])
    return path


def write_all(folder=DATA_DIR):
    """Write all 12 workbooks into `folder` (default eval/data/)."""
    for company in COMPANIES:
        write_company(company, folder)


if __name__ == "__main__":
    write_all()
