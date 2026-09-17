# Study Guide: Board Pack Generator

For you, a finance student learning Python, to understand this project well enough to explain every part of it in an interview.

It matches the code as of 2026-09-17 (commit `aa8838b`). Done so far: build steps 1, 2, 3, 3b, 4b and 5, plus the unit tests and bad-input tests. **Not built yet:** step 4 (`build_deck.py`, the PowerPoint deck) and step 6 (the full README).

## Contents

1. [How to use this guide](#1-how-to-use-this-guide)
2. [The data flow in plain English](#2-the-data-flow-in-plain-english)
3. [Python words you'll see](#3-python-words-youll-see)
4. [Every file, function by function](#4-every-file-function-by-function)
5. [25 interview questions](#5-25-interview-questions)
6. [10 trace-this-number exercises](#6-10-trace-this-number-exercises)
7. [5 functions to rewrite yourself](#7-5-functions-to-rewrite-yourself)
8. [Answers to the exercises](#8-answers-to-the-exercises)

---

## 1. How to use this guide

- **Read section 2 first.** Everything else hangs off it.
- **Keep the code open next to section 4.** Read a function, then read its row in the table.
- **Run things yourself.** None of these commands call the Claude API or cost money:

  ```bash
  source .venv/bin/activate
  python clean.py data/northwind.xlsx        # the cleaned input table
  python metrics.py data/northwind.xlsx      # every metric, the flags, the data gaps
  python main.py --all --skip-ai             # the batch run, writes output/*_metrics.xlsx
  python -m pytest -q                        # 212 unit tests
  python check_companies.py                  # end-to-end proof for all 3 companies
  ```

  **These two DO call the API and cost money:** `python analyze.py ...` and `python compare_models.py run`. You don't need them to study.
- **Do the exercises in section 6 before you look at section 8.**

---

## 2. The data flow in plain English

### The one-sentence version

A messy Excel file from a portfolio company goes in. Python tidies it, calculates the KPIs, checks them against investor thresholds and lists what's missing. Then it writes an Excel summary and asks Claude to word the commentary, checking that Claude didn't invent or calculate any number.

### The finance analogy

It's the quarterly process a portfolio analyst already runs by hand:

| Analyst by hand | This project |
|---|---|
| Receive the company's KPI file, with its own column names and typing habits | `data/northwind.xlsx` |
| Copy it into the fund's standard template, fixing "$14.3M" typed as text | `clean.py` |
| Build the KPI formulas in the model | `metrics.py` (`compute_metrics`) |
| Compare each KPI to the watch-list thresholds | `metrics.py` (`evaluate_flags`) + `config.yaml` |
| Note which numbers the company didn't send | `metrics.py` (`data_gaps`) |
| Write the commentary for the board | `analyze.py` (Claude words it, Python checks it) |
| Build the deck and the backup Excel | `build_deck.py` (not built yet) and `excel_output.py` |
| Do this for every company in the fund | `main.py --all` |

### The picture

```
make_data.py ──writes──▶ data/northwind.xlsx            fake messy input (answer key kept in make_data.py)
                              │
                              ▼
clean.py        clean_workbook()
                  ├─▶ actuals       8 quarters × 16 input columns, all $K, blanks = NaN
                  └─▶ next_budget   the "Q3 2026 (Budget)" row: 3 budget numbers
                              │
                              ▼
metrics.py      compute_metrics()        ─▶ metrics table: 8 quarters × 19 metrics (ratios as decimals)
                evaluate_flags()         ─▶ 9 flags for the latest quarter: trip / pass / cannot evaluate
                data_gaps()              ─▶ which metrics and flags have data missing, and in which quarters
                runway_at_next_budget()  ─▶ one context number (runway if burn returns to plan)
                              │
          ┌───────────────────┼──────────────────────────────┐
          ▼                   ▼                              ▼
excel_output.py         analyze.py                      build_deck.py
output/northwind_       numbers → text → Claude →        (build step 4, not built yet)
metrics.xlsx            check → retry once →
                        output/northwind_analysis.json

main.py      runs the chain for one workbook or every workbook in data/ (AI + deck skipped for now)
config.yaml  the flag thresholds            .env  the API key (never committed)
check_*.py and tests/   prove each step gives the right answer
```

### Step by step, with Northwind

**Step 0: make the fake input** (`make_data.py`). The script holds Northwind's true numbers, then messes them up on purpose the way real company files are messy. Headers are inconsistent (`" Churned ARR "` with stray spaces, `"Cash - End of Qtr"`). Some numbers are typed as text (`"$14.3M"`, `"260K"`, `"5,090"`). The Q1 2025 row has a label but no values. There's a budget-only row for next quarter and a junk Notes tab. Because the script keeps the true numbers, the checks can later prove that cleaning got every one of them back.

**Step 1: clean** (`clean.py`). Finds the tab that has a "Quarter" header and ignores the Notes tab. Translates each messy header to one of the 16 standard names from CLAUDE.md. Turns every cell into a plain number in $K ("$14.3M" becomes 14300). Leaves an empty cell empty (NaN), never 0. Splits off the budget-only row and checks that the quarters run in order with none skipped. **If anything is ambiguous, it stops** with a message naming the sheet and cell, e.g. `Sheet 'KPI Tracker', cell G5 (Q2 2025, revenue): Can't read 'TBD' as a number`. The output is a table with one row per quarter.

**Step 2: calculate** (`metrics.py`). Each metric is a small function that does column math on the whole table at once, like writing a formula in row 2 of Excel and filling it down. Ratios stay decimals (0.971, not 97.1). A blank input gives a blank result, so a missing quarter automatically spreads to every metric that uses it.

**Step 3: judge** (`metrics.py` + `config.yaml`). `evaluate_flags` compares the latest quarter with each threshold. Each flag gets one of 3 answers: **trip**, **pass**, or **cannot evaluate — data missing**. The combo rule looks at the last 3 quarters: NRR falling while pipeline is rising. `data_gaps` lists every metric that is blank *because data is missing*. That's different from a metric that's blank because there's no earlier period to compare with (YoY growth in the first year).

**Step 4: output.**
- `excel_output.py` writes a 3-sheet workbook: Metrics (red = tripped, gray = data missing), Flags, and Data gaps. It does no math.
- `analyze.py` turns the numbers into display text ("97.1%") and sends it to Claude with writing rules. It checks the answer: the right JSON shape, exactly 3 wins/risks/questions, short enough for a slide, and **every number Claude wrote must appear in the data**. If a check fails, it retries once and tells Claude what was wrong. The saved file keeps the data Claude saw next to what it wrote, so any claim can be traced back.
- `build_deck.py` doesn't exist yet.

**Step 5: run the batch** (`main.py`). For each workbook it runs steps 1–4, prints a ✓ line per step, and catches failures so one broken company doesn't stop the rest. It finishes with a summary table and an exit code (0 = all OK, 1 = a company failed, 2 = wrong arguments). The AI and deck steps are skipped with a message until `build_deck.py` exists.

**The proof layer.** Two kinds of proof:
- **`check_*.py` scripts** run the real workbooks end to end and compare the results with the answer keys and hand formulas.
- **`tests/`** (pytest) test one function at a time with tiny hand-worked tables, including edge cases and broken workbooks.

### The golden rule

**Python computes every number. Claude only interprets.** Two reasons you should be able to say out loud:
1. Language models can make arithmetic mistakes, and one wrong number in a board deck destroys trust in all the others.
2. Python's math can be tested; a prompt can't be tested the same way. So the math lives where it can be proven, and Claude's answer is checked by code before anyone sees it.

---

## 3. Python words you'll see

| Word | Plain English |
|---|---|
| **DataFrame** | A pandas table, like an Excel sheet. Rows are labelled by quarter ("Q2 2026"), columns by name ("nrr"). |
| **Series** | One column of a DataFrame: one value per quarter. |
| **NaN** | "Not a number": pandas' empty cell. Any math with NaN gives NaN. Excel is different: there, a blank often counts as 0. |
| **`math.inf`** | Infinity. Used on purpose for edge cases, e.g. CAC payback when new ARR is 0 ("never pays back"). |
| **dict** (dictionary) | A lookup table of key → value: `{"nrr_min": 1.00}`. Like VLOOKUP with an exact match. |
| **list** | An ordered collection: `["Q3 2024", "Q4 2024"]`. |
| **function** (`def`) | A named, reusable formula. The `"""text"""` right under it is its **docstring**: a note saying what it does. |
| **constant** | A name in CAPITALS at the top of a file (`MAX_ATTEMPTS = 2`). A setting that doesn't change while the program runs. |
| **`raise ValueError(...)`** | Stop the program with a message. |
| **`try` / `except`** | "Try this; if that kind of error happens, do this instead." |
| **`.shift(1)`** | Line each quarter up with the quarter above it, like a formula pointing at the cell one row up. `.shift(4)` = same quarter last year. |
| **`.mask(condition, x)`** | Replace values with x wherever the condition is true, like `IF(condition, x, value)`. |
| **`.diff()`** | This row minus the row above. |
| **`.iloc[-1]`** / **`.loc["Q2 2026", "nrr"]`** | The last row by position / one cell by row and column label. |
| **`import`** | Use functions from another file. `from metrics import nrr` brings in `nrr`. |
| **`if __name__ == "__main__":`** | Code that runs only when you run the file directly (`python metrics.py`), not when another file imports it. |
| **regex** | A text pattern. `^Q([1-4])\s+(\d{4})$` means "Q, a digit 1–4, spaces, 4 digits". |
| **f-string** | `f"{value:.1%}"` puts a value into text with a format: 0.971 → "97.1%". |
| **`assert`** | "This must be true, otherwise stop." The check scripts are built from these. |
| **pytest** | Runs every function whose name starts with `test_` and reports what passed. |
| **pydantic `BaseModel`** | A declared shape for data (these fields, these types). Used to force the shape of Claude's JSON. |
| **Decimal** | Exact decimal math, the way you'd do it on paper. Normal Python floats are stored in binary and can be very slightly off. |

---

## 4. Every file, function by function

Order: config, then the 5 pipeline files in the order data flows, then the data generators, the check scripts, the model comparison, the tests and the small files.

### `config.yaml`: the thresholds

Not code, just settings. Every flag threshold lives here with a comment saying why an investor cares, so a partner can change "runway under 12 months" to 18 without touching Python. Ratios are decimals (`nrr_min: 1.00` = 100%).

| Setting | Value | Trips when |
|---|---|---|
| `nrr_min` | 1.00 | NRR below 100% |
| `grr_min` | 0.85 | GRR below 85% |
| `burn_multiple_max` | 2.0 | burn multiple above 2.0x |
| `burn_over_budget_max` | 0.15 | burn more than 15% over budget |
| `runway_min_months` | 12 | runway under 12 months |
| `cac_payback_max_months` | 24 | CAC payback over 24 months |
| `net_new_arr_vs_budget_min` | -0.20 | net new ARR more than 20% under budget |
| `rule_of_40_min` | 0.40 | Rule of 40 below 40% |
| `nrr_falling_pipeline_rising_flag` | true | turns the combo rule on |
| `combo_lookback_quarters` | 3 | the combo rule's window (3 quarters = 2 quarter-over-quarter steps) |

---

### `clean.py`: messy workbook → clean table

**What it's for:** reads a messy KPI workbook and returns `(actuals, next_budget)`. It never guesses: anything it can't read with certainty stops the run with a message that says what and where.

**Constants worth knowing:**
- `ACTUAL_COLUMNS` (13), `BUDGET_COLUMNS` (3), `STANDARD_COLUMNS` (all 16): the standard names from CLAUDE.md.
- `HEADER_ALIASES`: messy name → standard name, only where they differ (`"beginning_arr"` → `"starting_arr"`).
- `BUDGET_LABEL_WORDS`: a row label containing "budget", "bud" or "plan" is the budget-only row.
- `QUARTER_PATTERN`: what a quarter label must look like ("Q2 2026").
- `NUMBER_TEXT`: the only number text it accepts: optional minus, digits (commas only between groups of 3), optional decimals, optional K or M.

**The call order inside `clean_workbook`:** `find_kpi_sheet` → `check_no_error_cells` → `clean_sheet`. Inside `clean_sheet`: `map_columns` → `check_no_headerless_values` → for each row (`parse_row` → `is_budget_only_row` → duplicate check) → build the table → `check_quarters_in_order`.

| Function | What it does, in plain English | Example / why it exists |
|---|---|---|
| `is_blank(value)` | True if a cell is empty: nothing in it, or only spaces. | `"   "` → True. A cell with only a space is treated as empty. |
| `normalize_header(text)` | Lowercases a header and turns every run of symbols/spaces into one `_`. | `"Net Burn (Bud.)"` → `"net_burn_bud"`. Makes different spellings of the same header match. |
| `standard_column(header)` | Normalizes a header, applies `HEADER_ALIASES`, and stops if the result isn't one of the 16 standard names. | `"Beginning ARR"` → `"starting_arr"`. `"EBITDA"` → error: add it to HEADER_ALIASES or delete the column. |
| `parse_number(value)` | Turns one cell into a number in $K. Blank → NaN. TRUE/FALSE → error. A real number → float. Text: removes `$` and spaces, checks it matches `NUMBER_TEXT`, reads K or M, and multiplies using `Decimal`. | `"$14.3M"` → 14300.0, `"260K"` → 260.0, `"5,090"` → 5090.0, `"n/a"` → error. `Decimal` makes `"$4.03M"` exactly 4030 (plain float math gives 4030.0000000000005). |
| `parse_quarter(label)` | Reads "Q2 2026" into `(2026, 2)` so quarters can be compared, or stops if the label is in another format. | `"Q2 2026"` → `(2026, 2)`. `"2Q26"` → error. |
| `at_row(rows, i)` | Builds the "row 7: " prefix for an error message, if Excel row numbers are known. | Helper for readable errors. |
| `check_quarters_in_order(labels, rows)` | Stops unless quarters run oldest → newest, none skipped or repeated. Handles Q4 → next year's Q1. | Why it matters: QoQ looks 1 row up and YoY 4 rows up. A missing row would silently compare the wrong quarters. |
| `excel_row(row_index)` | pandas row number (counts from 0) → Excel row number (counts from 1). | Index 6 → row 7. |
| `excel_column(position)` | pandas column position → Excel column letter. | Position 5 → "F". |
| `find_kpi_sheet(path)` | Reads every tab and returns the first one with a "Quarter" header in its first 10 rows, plus which row that header is on. | Skips the Notes tab and allows title rows above the table. Reads with `keep_default_na=False` so text like "n/a" isn't silently blanked (a bug found in Task 5). |
| `check_no_error_cells(path, sheet_name)` | Opens the KPI tab with openpyxl and stops on any Excel error like `#DIV/0!` or `#REF!`. | pandas reads error cells as empty, which would turn a broken formula into "data missing". openpyxl can still see them. |
| `header_name(header)` | "Quarter" → `"quarter"` (the label column); any other header → its standard name. | Lets the "Quarter" column go through the same duplicate check as the others. |
| `name_headers(headers, header_row)` | Names every filled header cell. Stops on an unknown header, or on two columns that mean the same thing. | "Starting ARR" in B and "Beginning ARR" in R → error: which one is right? |
| `map_columns(headers, header_row)` | Uses `name_headers`, stops if any of the 16 columns is missing, and returns where the Quarter column is plus a position → name map. | Missing columns are all listed at once so one fix round is enough. |
| `check_no_headerless_values(sheet, header_row, known_positions)` | Stops if a column without a header has values under it. | Those values would otherwise be skipped without a word. |
| `check_unlabelled_row_is_empty(...)` | Stops if a row has numbers but no quarter label. | Skipping that row would silently lose a quarter's data. |
| `parse_row(label, row_index, raw_row, column_map)` | Runs `parse_number` on every cell in one row. If one fails, the error names the cell, quarter and column. | `cell G5 (Q2 2025, revenue): Can't read 'TBD' as a number`. |
| `is_budget_only_row(label, row_index, row, column_map)` | True if the label contains budget/bud/plan. Stops if that row also has actual values. | `"Q3 2026 (Budget)"` → True. A budget row with revenue filled in is a mistake, so it stops. |
| `clean_sheet(sheet, header_row)` | The main loop. Maps the columns, then goes row by row: skips empty rows, parses each labelled row, separates the budget-only row, stops on a repeated quarter or a second budget row, builds the table and checks the quarter order. | The blank Q1 2025 row is kept as a row of NaN, so later look-backs still line up. The duplicate-quarter stop was a bug fix (Task 4). |
| `clean_workbook(path)` | **The entry point.** Finds the KPI tab, checks for error cells, runs `clean_sheet`, and puts the tab name (e.g. `Sheet 'KPI Tracker', `) in front of any error from that tab. The "no tab has a 'Quarter' header" error names the file instead, because no tab qualified. | Every other file calls this one function to read a workbook. |
| `if __name__ == "__main__":` block | `python clean.py data/northwind.xlsx` prints the clean table sideways (quarters across) and the budget row. | For looking at the data yourself. |

---

### `metrics.py`: the math, the flags, the data gaps

**What it's for:** every number on the board pack is calculated here. It also decides trip/pass for each flag and lists the data gaps. It never fills in a missing number.

**Constants worth knowing:**
- `TRIP = "trip"`, `PASS = "pass"`, `MISSING = "cannot evaluate — data missing"`: the three flag answers.
- `METRIC_LOOKBACK`: how many quarters back a metric looks (QoQ metrics 1, YoY metrics and Rule of 40 4, net new ARR vs budget 1). `data_gaps` uses it.
- `FLAG_RULES`: one line per flag: (display name, metric column, config.yaml key, "min" or "max"). Adding a flag means adding a line here plus a threshold in config.yaml.
- `DOLLAR_COLUMNS`, `MONTH_COLUMNS`: which columns are $K and which are months, for display.

**Metric functions.** Each takes the clean table (`df`) and returns one value per quarter.

| Function | Formula (from CLAUDE.md) | Northwind Q2 2026 |
|---|---|---|
| `net_new_arr(df)` | new + expansion − contraction − churn | 1850 + 580 − 260 − 510 = **1,660** |
| `ending_arr(df)` | starting_arr + net new ARR | 25810 + 1660 = **27,470** |
| `nrr(df)` | 1 + 4 × (expansion − contraction − churn) / starting_arr. Annualized from one quarter. | **97.1%** |
| `grr(df)` | 1 − 4 × (contraction + churn) / starting_arr. Annualized. | 1 − 4 × 770 / 25810 = **88.1%** |
| `gross_margin(df)` | gross_profit / revenue | 5000 / 6660 = **75.1%** |
| `growth(series, quarters_back)` | value / value N quarters earlier − 1. `.shift(N)` does the "N quarters earlier". The first N quarters come out NaN. | ARR YoY: 27470 / 19230 − 1 = **42.8%** |
| `fcf_margin(df)` | −net_burn / revenue (burning cash = negative margin) | −3900 / 6660 = **−58.6%** |
| `rule_of_40(df)` | revenue growth YoY + FCF margin | **−12.2%** |
| `burn_multiple(df)` | net_burn / net new ARR. Edge cases with `.mask`: not burning → 0; burning while net new ARR ≤ 0 → ∞. | 3900 / 1660 = **2.35x** |
| `burn_vs_budget(df)` | net_burn / budget_net_burn − 1 | 3900 / 3250 − 1 = **20.0%** |
| `budget_net_new_arr(df)` | budget_arr this quarter − budget_arr last quarter | 26800 − 24750 = **2,050** |
| `net_new_arr_vs_budget(df)` | net new ARR / budgeted net new ARR − 1 | 1660 / 2050 − 1 = **−19.0%** |
| `arr_vs_budget(df)` | ending ARR / budget_arr − 1 | 27470 / 26800 − 1 = **2.5%** |
| `cac_payback_months(df)` | sm_spend / (new_arr × gross margin) × 12. Edge case: new_arr × margin ≤ 0 → ∞. | 2400 / (1850 × 0.751) × 12 = **20.7 mo** |
| `runway_months(df)` | ending_cash / (net_burn / 3). Edge case: not burning → ∞. | 14300 / 1300 = **11.0 mo** |
| `runway_at_next_budget(actuals, next_budget)` | latest cash / (next quarter's budgeted burn / 3). NaN if there's no budget row, ∞ if the budget has no burn. Shown as context, never flagged. | 14300 / 1100 = **13.0 mo** |
| `compute_metrics(actuals)` | Calls every metric function above except `runway_at_next_budget` (a single number, which callers compute separately) and puts the results in one table: a row per quarter, 19 metric columns. `pipeline` is copied in for the combo rule. | The table `python metrics.py` prints. |

**Flags and gaps.**

| Function | What it does, in plain English | Example / why it exists |
|---|---|---|
| `load_config(path)` | Reads config.yaml into a dictionary. | `config["nrr_min"]` → 1.0. |
| `check_threshold(value, threshold, kind)` | NaN → MISSING. Otherwise rounds to 6 decimals, then "min" trips below the threshold and "max" trips above it. **Exactly at the threshold passes.** | Rounding matters: 3900/3250 − 1 is 0.19999999999999996 in Python. Without rounding, a burn exactly 15% over budget could trip on noise. |
| `check_combo(metrics, config, quarter)` | Takes the last 3 quarters up to `quarter`. Not enough history, or any NRR/pipeline value missing → MISSING. Otherwise trips only if NRR fell at **every** step and pipeline rose at **every** step. | Northwind: NRR 108.0% → 102.0% → 97.1% while pipeline 10,100 → 11,200 → 12,500 → trip. Missing data is never quietly treated as a pass. |
| `evaluate_flags(metrics, config, quarter=None)` | Loops over `FLAG_RULES`, runs `check_threshold` for each, then adds the combo rule. Defaults to the latest quarter but accepts any. Returns a list of dicts (flag, metric, quarter, value, threshold, status). | Northwind Q2 2026: 6 trip, 3 pass. |
| `data_gaps(actuals, metrics, flags)` | For every metric, a blank value counts as a gap if the quarter **has enough history** for that metric (position ≥ lookback) **or** the quarter's own inputs are incomplete. Then adds every flag that came out MISSING. Returns `{metric or flag: [quarters]}`. | Separates "data missing" from "no prior period". ARR YoY in Q4 2024 is blank but not a gap (no year-ago quarter exists). ARR YoY in Q1 2026 **is** a gap (Q1 2025 exists but is blank). |
| `format_value(column, value)` | Display text: $K with commas, "11.0 mo", "2.35x", or a %. NaN → "n/a", ∞ → "∞". | **The only place ratios become %** (plus analyze.py and excel_output.py, which reuse the same rules). |
| `print_report(actuals, next_budget, config)` | Prints the metrics table, runway at budget, flags and data gaps. | What you see with `python metrics.py data/northwind.xlsx`. |

---

### `analyze.py`: Claude writes the commentary, Python checks it

**What it's for:** builds a text "payload" of already-computed facts, sends it to Claude with writing rules, validates the answer, retries once if needed, and saves the result. Run on its own with `python analyze.py data/northwind.xlsx` (**calls the API**). `main.py` doesn't call it yet.

**Constants worth knowing:** `DEFAULT_MODEL = "claude-sonnet-5"` (the model comparison's pick), `MAX_ATTEMPTS = 2` (first try + one retry), `MAX_HEADLINE_WORDS = 30`, `MAX_DETAIL_WORDS = 45`, `MAX_DETAIL_SENTENCES = 2`. `METRIC_LABELS` maps column names to display names and is reused by `excel_output.py`. `SYSTEM_PROMPT` holds Claude's instructions: role, number rules, what to write, framing rules.

| Function / class | What it does, in plain English | Example / why it exists |
|---|---|---|
| `Point` (class) | The shape of one win or risk: a `title` and a `detail`. | "Retention slipping" / "NRR went from 108.0% to 97.1%…" |
| `BoardSummary` (class) | The shape of the whole answer: headline, wins, risks, questions. | Passed to the API so Claude's reply must be this JSON. |
| `input_trend(actuals, column)` | One raw input (net burn, ending cash) across all quarters, as display text. Blank → "data missing". | `{"Q3 2024": "2,400", ...}` |
| `metric_trend(metrics, gaps, column)` | One metric across all quarters as text. Uses the gap list to say "data missing" (a gap) or "n/a (no prior period)" (nothing to compare with). | Keeps "missing" and "not meaningful" apart for Claude too. |
| `describe_flag(flag, config)` | One flag as text: name, status (TRIPPED/passed/…), value, threshold. The combo rule gets a description instead of a value. | `{"flag": "NRR (annualized)", "status": "TRIPPED", "value": "97.1%", "threshold": "100.0%"}` |
| `build_payload(company, actuals, next_budget, config)` | Runs the metrics, flags and gaps, then packages every fact Claude may use as text: flag counts, flag details, runway if burn returns to plan, all trends, data gaps. | Flag counts are computed here so Claude quotes "6 of 9" instead of counting. |
| `payload_to_text(payload)` | Turns the payload into the exact JSON text sent to Claude. | The number check uses this same text, so "allowed numbers" = exactly what Claude saw. |
| `numbers_in(text)` | Finds every number in a piece of text, ignoring $, %, x, K, minus signs and commas. | `"$27,470K and 97.1%"` → {27470.0, 97.1}. |
| `count_sentences(text)` | Counts sentences by splitting after . ! or ? followed by a space. | The decimal point in "97.1%" isn't followed by a space, so it isn't a sentence end. Known limit: "vs. " counts as one. |
| `summary_texts(summary)` | Collects every piece of text Claude wrote into one list. | So the checks can scan everything at once. |
| `find_ungrounded_numbers(summary, payload_text)` | Numbers in Claude's answer that don't appear anywhere in the payload. | "NRR fell 11.8 points" → 11.8 isn't in the data (Claude subtracted) → caught. |
| `validate_summary(summary, payload_text)` | Returns a list of problems: not exactly 3 wins/risks/questions, empty text, headline too long, a detail over 45 words or 2 sentences, numbers not in the data. An empty list = pass. | The code-enforced part of "Claude never does math". |
| `AnalysisError` (class) | The error raised when both attempts fail. Carries `run_info` (tokens, timing, problems) so failures can still be logged. | Used by `compare_models.py` to count failed runs. |
| `build_messages(payload_text, previous)` | The conversation to send. On a retry it adds Claude's last answer and the list of problems. | The retry tells Claude *what* was wrong, not just "try again". |
| `call_claude(client, model, payload_text, previous)` | One API call using structured outputs (`output_format=BoardSummary`). Handles a malformed reply, a refusal or a cut-off answer. Otherwise runs `validate_summary`. Returns the summary, its problems and a log entry. | Timing and token counts are logged for the cost comparison. |
| `make_run_info(model, attempts, passed)` | Adds up tokens and seconds across attempts. | Feeds the step 3b cost table. |
| `analyze(payload_text, model, client)` | **The loop:** up to 2 attempts. Returns the first answer with no problems; if both fail, raises `AnalysisError`. Nothing unvalidated is ever returned. | |
| `print_commentary(summary)` | Prints headline, wins, risks, questions, with no model name or stats. | Also used for blind scoring in compare_models.py. |
| `print_summary(summary, run_info)` | `print_commentary` plus the run stats. | |
| `main()` | Command line: loads `.env`, cleans the workbook, builds the payload, calls `analyze`, and saves `output/<company>_analysis.json` holding the summary, run info **and the payload**. A failure is saved too, then the script exits with code 1. | Saving the payload next to the answer makes every claim traceable. |

---

### `excel_output.py`: the metrics as a spreadsheet (build step 4b)

**What it's for:** writes `output/<company>_metrics.xlsx` with 3 sheets. **It does no math.** It calls the same clean → metrics → flags → gaps chain as everything else, then only writes and formats.

**Constants worth knowing:** `DATA_MISSING`, `NO_PRIOR_PERIOD` and `INFINITE_LABELS` (`"∞ (ARR shrank)"`, `"∞ (not burning)"`, `"∞ (never pays back)"`) are the words written when there's no usable number. `STATUS_COLORS` holds Excel's own "Bad" red, "Good" green, plus gray.

| Function | What it does, in plain English | Example / why it exists |
|---|---|---|
| `number_format(column)` | The Excel display format for a column: `#,##0`, `0.0" mo"`, `0.00"x"` or `0.0%`. | The cell holds 0.971; Excel shows 97.1%. The number stays usable in formulas. |
| `cell_value(column, value, is_gap)` | The number, or a word: NaN → "data missing" or "n/a (no prior period)"; ∞ → "∞ (reason)". | Excel can't store NaN or ∞, and openpyxl would silently save both as an empty cell, so "missing" and "infinite" would look the same. |
| `runway_context_value(runway)` | Same idea for runway at budget: a number, "n/a (no budget row)" or "∞ (budget not burning)". | |
| `color_cell(cell, status)` | Fills a cell red/green/gray with matching text color. | |
| `write_header(sheet, headers)` | Bold header row, freezes row 1 and column A. | |
| `set_column_widths(sheet, widths)` | Sets column widths left to right. | |
| `tripped_cells(metrics, config)` | Runs `evaluate_flags` for **every** quarter and collects each (quarter, metric) that trips. | So a Q1 2026 burn-vs-budget breach is red too, not just the latest quarter. |
| `style_metric_cell(cell, column, is_gap, is_tripped)` | Format, right alignment, gray if data missing, else red if tripped. | Passed cells stay uncolored so the red stands out. |
| `write_metrics_sheet(sheet, metrics, gaps, config)` | One row per quarter, one column per metric, with values, formats and highlights. | |
| `flag_row(flag, gaps, config)` | One flag as a row: Flag, Quarter, Value, Threshold, Trips when, Status. The combo rule gets text instead of a value. | "Trips when" says "below threshold" or "above threshold", so −20.0% makes sense. |
| `style_flag_row(sheet, row_number, flag)` | Colors the whole row by status and formats Value and Threshold. | |
| `write_runway_context(sheet, runway, quarter)` | Adds "Runway at next quarter's budgeted burn (context, not a flag)" below the table, uncolored. | CLAUDE.md: shown as context, not flagged. |
| `write_flags_sheet(sheet, flags, gaps, config, runway_at_budget)` | Writes every flag row plus the runway context line. | |
| `gap_label(name)` | `"nrr"` → `"NRR (annualized)"`; `"flag: Rule of 40"` → `"Flag: Rule of 40"`. | |
| `write_gaps_sheet(sheet, gaps)` | One row per affected metric or flag, or "None — every metric and flag has the data it needs". | |
| `build_workbook(actuals, next_budget, config)` | Computes metrics, flags and gaps, then writes all 3 sheets into a new workbook. | |
| `output_path(workbook_path, output_dir)` | `data/northwind.xlsx` → `output/northwind_metrics.xlsx`. | The name comes from the input file, so `--all` gets unique names for free. |
| `save_metrics_workbook(workbook_path, config, output_dir)` | Cleans the input, builds the workbook, saves it, returns the path. | **The one function `main.py` calls.** It cleans the workbook itself, so in a `main.py` run each file is read twice. Harmless at this size. |

---

### `main.py`: the batch runner (build step 5)

**What it's for:** `python main.py data/northwind.xlsx` runs one company; `python main.py --all` runs every workbook in `data/`. One failing company never stops the batch.

| Function / class | What it does, in plain English | Example / why it exists |
|---|---|---|
| `NotWiredError` (class) | Raised if `build_deck.py` exists but main.py hasn't been connected to it. | So nobody gets an "OK" with no deck once step 4 starts. |
| `find_workbooks(data_dir)` | Every `.xlsx` in `data/`, sorted by name, skipping `~$` files. | `~$northwind.xlsx` is the lock file Excel creates while a workbook is open; it would show up as a failing company. |
| `company_name(workbook_path)` | `data/northwind.xlsx` → "Northwind". | |
| `ai_step(skip_ai)` | For now only returns why AI was skipped: "--skip-ai" or "build_deck.py doesn't exist yet". Without `--skip-ai`, raises `NotWiredError` if the deck file exists. With `--skip-ai`, it skips even then. | No API money is spent while there's no deck to put the commentary in. |
| `deck_step()` | Same for the deck. | |
| `blank_quarters(actuals)` | Quarters with at least one blank input. | Northwind → `["Q1 2025"]`. |
| `run_company(workbook_path, config, skip_ai)` | Clean → metrics → flags → gaps → save Excel → AI step → deck step, printing a ✓ line for each. Returns a result dict for the summary table. | |
| `describe_error(error)` | Prints `✗ FAILED: <type>: <message>`. Bad-input errors (`ValueError`, `OSError`) and `NotWiredError` get one line; anything else also gets a full traceback, because it's probably a bug. | A person fixing a workbook doesn't need a traceback; a developer fixing a bug does. |
| `run_batch(workbook_paths, config, skip_ai)` | Loops over the workbooks with `try`/`except` around each company, records the error and moves on. | The key reliability feature: company 2 breaking doesn't stop company 3. |
| `flags_text(result)` | "6 of 9", or "7 of 9, 1 cannot evaluate". | Without the second part, Fernhollow's "7 of 9" would hide a flag that had no answer. |
| `gaps_text(result)` | "none", or "19 metrics/flags (blank: Q1 2025)". | |
| `result_text(result)` | "OK", "OK (AI + deck skipped)" or "FAILED: …". | |
| `summary_rows(results)` | One row of text per company; a failed company shows "-" for flags and gaps. | |
| `print_summary(results)` | Prints the table with padded columns, then "3 of 3 companies succeeded". | |
| `parse_args(argv)` | Reads the command line. Exactly one of a file path or `--all`; both or neither → usage error (exit code 2). | |
| `main(argv)` | Parses the arguments, finds the files, runs the batch, prints the summary, returns 0 if all OK, else 1. | Exit codes let a scheduler tell from the code alone whether a run worked. |

---

### `make_data_common.py`, `make_data.py`, `make_data_alderpeak.py`, `make_data_fernhollow.py`: the fake inputs

**What they're for:** each company script holds only its own numbers (`TRUE_DATA`, the answer key) and its mess settings, then calls `save_workbook` from the shared file. Three companies prove the flags aren't hard-coded:

| Company | Story | Q2 2026 flags |
|---|---|---|
| **Northwind** (`make_data.py`) | Strong ARR growth; NRR sliding; pipeline rising; burn ~20% over budget; 11 months runway. Q1 2025 blank. | 6 trip, 3 pass |
| **Alderpeak** (`make_data_alderpeak.py`) | Healthy: ~47% ARR growth, NRR ~110%, burn falling and under budget. No blank quarter. Notes tab placed first (`"notes_first": True`). | 0 trip in **any** quarter |
| **Fernhollow** (`make_data_fernhollow.py`) | Distressed: ARR shrinking, NRR 77.9%, 6 months runway, pipeline falling. Q2 2025 blank, so Rule of 40 can't be evaluated. Title row above the table (`"title"`). | 7 trip, 1 pass, 1 cannot evaluate |

Each company script has constants (`OUTPUT_PATH`, `QUARTERS`, `BLANK_QUARTER`, `BUDGET_ONLY_LABEL`, `TRUE_DATA`, `NEXT_QUARTER_BUDGET`, `HEADER_NAMES`, `TEXT_CELLS`, `NOTES`) and one function, `main()`, which packs them into a dictionary and calls `save_workbook`.

**`make_data_common.py` functions:**

| Function | What it does, in plain English | Example / why it exists |
|---|---|---|
| `ending_arr_at(true_data, i)` | Ending ARR for quarter number i, from the answer key. | Used by the roll-forward check. |
| `check_true_data(company)` | Stops if the answer key doesn't tie out: one value per quarter, a header for every column, **ARR rolls forward** (ending = next starting), **cash rolls forward** (cash = last cash − burn), money values are multiples of 10, and the mess settings point at real columns and quarters. | Fake data that doesn't tie out would make every later check meaningless. |
| `to_text(value, style)` | Turns a $K number into messy text. | 14300 → `"$14.3M"`, 260 → `"260K"`, 5090 → `"5,090"`. |
| `quarter_row(i, quarter, columns, true_data, text_cells)` | One quarter's row: the label, then each value as a number or as messy text. | |
| `write_kpi_sheet(ws, company)` | Writes the optional title, the header row, each quarter (the blank quarter as a label only), then the budget-only row. | |
| `write_notes_sheet(ws, notes)` | Writes the junk Notes tab. | Northwind's has "scratch 3900 3250 20% over??", so cleaning must ignore it. |
| `save_workbook(company)` | Runs `check_true_data`, builds the workbook (KPI Tracker + Notes), optionally moves Notes first, saves it and prints a one-line summary. | |

---

### The check scripts: end-to-end proof

Each one prints ✓ lines and ends with "All checks passed", or stops at the first failure. They use `assert`.

#### `check_northwind.py` (step 2 + 3 proof)

| Function | What it checks |
|---|---|
| `check_cleaning(actuals, next_budget)` | Every cleaned value equals `TRUE_DATA`; Q1 2025 is entirely NaN; the budget row is right. |
| `check_latest_metrics(metrics, runway_budget)` | Q2 2026 metrics equal hand formulas typed like Excel (`EXPECTED_LATEST`, e.g. `1 + 4 * (580 - 260 - 510) / 25810`). |
| `check_flags(flags)` | Each flag's status matches the story (`EXPECTED_FLAGS`: 6 trip, 3 pass). |
| `check_gaps(gaps, metrics)` | QoQ gaps = blank quarter + next; YoY gaps = blank quarter + 4 later; everything else = blank quarter only. |
| `good_summary()` | A hand-written answer that uses only numbers from the payload. |
| `check_analysis_validation(payload_text)` | `validate_summary` passes the good answer and catches a calculated number (11.8), a rounded one (97), a 3-sentence detail and only 2 risks. **No API call.** |
| `check_recommendation_rule()` | The Sonnet/Haiku rule picks correctly at the edges, and score entry rejects bad scores. |
| `main()` | Runs all of the above. |

#### `check_companies.py` (step 5 proof, all 3 companies)

| Function | What it checks |
|---|---|
| `same_number(got, expected)` | Two values match: both NaN, or equal to 9 significant digits (∞ = ∞). |
| `check_cleaning(answer_key, actuals, next_budget)` | Same as above, for any company. |
| `check_latest_metrics(company, metrics, runway_budget)` | Q2 2026 metrics equal each company's hand formulas. |
| `check_flags(company, flags)` | Statuses match each company's story. |
| `expected_gaps(company, metric_columns)` | **Builds** the gap list the CLAUDE.md rules predict, independently of `data_gaps`. |
| `check_gaps(company, gaps, metric_columns)` | Real gaps equal the predicted ones, no more and no fewer. |
| `check_never_trips(metrics, config)` | No flag trips in any quarter (used for Alderpeak). |
| `count_statuses(flags)` | Text like "7 trip, 1 pass, 1 cannot evaluate". |
| `check_company(company, config)` | Runs the four checks for one company. |
| `main()` | All 3 companies, plus Alderpeak's every-quarter check. |

#### `check_excel_output.py` (step 4b proof)

Builds each Excel file, **reads it back from disk**, and compares it with the metrics table. Expected labels, formats and colors are typed out here rather than imported, so a wrong constant in `excel_output.py` can't make its own check pass.

| Function | What it does |
|---|---|
| `fill_color(cell)` | A cell's fill color as 6-digit hex, or None. |
| `is_number(value)` | True for a real number cell (not TRUE/FALSE). |
| `same_as_saved(got, expected)` | Equal to 15 significant digits: the precision Excel itself keeps (see LEARNINGS). |
| `excel_number(cell)` | A cell back to a metric value: number → float, "data missing" → NaN, "∞ …" → ∞. |
| `expected_format(column)` | The number format each kind of metric must have. |
| `check_file(path, answer_key)` | File name and sheet order. |
| `check_value_cell(cell, column, value, is_gap)` | One cell holds the right number or the right label, with the right format. |
| `tripped_by_quarter(metrics, config)` | Every (quarter, metric) that trips, in every quarter. |
| `check_metrics_sheet(sheet, metrics, gaps, config)` | Headers, quarters, every value, format and fill. |
| `check_latest_against_hand_formulas(sheet, company, metrics)` | Latest-quarter cells equal the hand formulas. |
| `check_story_highlights(sheet, company, metrics)` | Latest-quarter red cells are exactly the flags the story says trip. |
| `check_flag_row(...)` / `check_status_and_color(row, name, status)` | One Flags row's values, threshold, status text and row color. |
| `check_flags_sheet(...)` | Every flag row, the combo row, and the runway context line. |
| `check_gaps_sheet(sheet, gaps)` | The Data gaps sheet lists exactly `data_gaps()`. |
| `check_company(company, config)` / `main()` | Runs it all for each company. |

#### `check_main.py` (batch runner proof)

| Function | What it does |
|---|---|
| `run_main(*args)` | Runs `python main.py ...` as a real separate process and captures its output. |
| `summary_table(stdout)` | Reads the printed summary table back into rows. |
| `story_flags_text(company)` / `story_gaps_text(company)` | The expected table cells, built from each company's story, **not** from main.py. |
| `write_broken_workbook(folder)` | A workbook with most columns missing, in a temp folder. |
| `run_quietly(function, *args)` | Calls a function and captures what it prints. |
| `check_finds_the_three_companies()` | `--all` finds exactly the 3 workbooks. |
| `check_batch_run()` | Exit 0, both skip messages 3 times, correct summary rows, fresh Excel files. |
| `check_failure_does_not_stop_batch(folder, config)` | A broken workbook and a missing file between good companies: the good ones still succeed, and no traceback is printed for input errors. |
| `check_code_bug_does_not_stop_batch(config)` | Temporarily swaps in a `compute_metrics` that fails once (`buggy_first_call`): the error is recorded **with** a traceback and the batch continues. The real function is always put back. |
| `check_exit_codes(broken_path)` | 1 when a company fails; 2 for no arguments or both a file and `--all`. |
| `check_unwired_deck_fails_loudly(folder, config)` | A fake `build_deck.py` in a temp folder makes the run fail with `NotWiredError`. |
| `check_ignores_lock_and_other_files(folder)` | `~$` files and non-.xlsx files are skipped; results are sorted. |
| `main_check()` | Runs everything. |

---

### `compare_models.py`: Sonnet vs Haiku, scored blind (build step 3b)

**What it's for:** runs the same Northwind payload through `claude-sonnet-5` and `claude-haiku-4-5` 3 times each, shuffles the answers, and lets you score them without knowing which model wrote which. `run` calls the API (costs money); `score` is free. **Already done; the results are in README.md.**

| Function | What it does, in plain English |
|---|---|
| `run_cost(model, input_tokens, output_tokens)` | Dollar cost: tokens × price per million. |
| `make_record(...)` | One run's result in a fixed shape, passed or failed. |
| `run_once(client, model, run_number, payload_text)` | Calls `analyze` once. A failure is **recorded, never rerun**: a failure is data. |
| `assign_letters(records)` | Shuffles the runs and labels them A–F. |
| `save_blind_files(lettered)` | `answer_A.json` … holds commentary only; `key_DO_NOT_OPEN.json` holds models and stats. |
| `print_blind_answers(lettered)` | Prints the answers by letter with the rubric. |
| `command_run(force)` | Runs S, H, S, H, S, H, then saves and prints the blind set. Refuses to overwrite an existing set without `--force`. |
| `parse_scores(score_args, runs)` | `["A=4", "B=3"]` → `{"A": 4, "B": 3}`. Every passed answer needs a whole-number score from 1 to 5. |
| `average(values)` | Mean of the values that exist, or None. |
| `model_stats(model, runs, scores)` | Pass rate, average attempts, cost, seconds, score, and cost for 275 companies. |
| `recommend(haiku)` | The rule agreed **before** running: Haiku only if its average score is ≥ 4.0 **and** it passed 100% of runs; otherwise Sonnet. |
| `fmt(value, pattern, missing)` | Formats a number, or "n/a". |
| `recommendation_text(sonnet, haiku)` | The one-line recommendation with the reason. |
| `readme_section(stats, runs, scores)` | The Markdown table for the README. |
| `write_readme_section(section)` | Replaces the text between the README's marker comments (or appends it). |
| `append_learning(stats)` | Adds the result to LEARNINGS.md. (`command_score` makes sure this happens only the first time scores are entered.) |
| `command_score(score_args)` | Checks the scores, reveals the key, writes the README and LEARNINGS. |
| `main()` | The `run` / `score` command line. |

---

### `tests/`: unit tests (pytest)

Run with `python -m pytest -q` (212 tests, about 2 seconds). Expected values are **worked out by hand** in comments, not copied from running the code. Tests with `@pytest.mark.parametrize` run the same test on many inputs, each inputs line counting as one test.

| File | Helper functions | What the tests cover |
|---|---|---|
| `test_clean.py` (81) | `write_workbook(path, labels, blank)`: a tiny workbook in pytest's temp folder. | `parse_number` (good text, real numbers, blanks → NaN, 17 kinds of unreadable text stop), `normalize_header` and `standard_column` on most of the Northwind headers, quarter labels and order (Q4 → Q1 rollover, skipped, repeated), and whole workbooks (blank row kept, missing row stops, duplicate row stops). |
| `test_metrics.py` (78) | `table(**columns)`: a small table with only the needed columns. `values(series)`: compare with NaN allowed. `burn_table`, `cac_table`: tables for one metric. `full_actuals(blank, blank_cells)`: 8 realistic quarters with optional blanks. `combo_metrics`, `latest`, `gaps_for`: shortcuts for combo and gap tests. `TEST_CONFIG`: thresholds typed into the test file, so editing config.yaml never breaks a test. | Every metric against hand math; every CLAUDE.md edge case (∞, 0, −0.0); a missing input never becomes 0 or ∞; `check_threshold` exactly at the threshold, float noise, real misses, NaN, ∞; `check_combo` trip/pass/cannot evaluate; `data_gaps` following the QoQ/YoY rules. |
| `test_bad_inputs.py` (53) | `good_table()`: a valid table. `set_cell`, `drop_column`, `add_column`: break one thing. `write_workbook(path, rows, empty_columns_left)`: Notes tab, title, empty row, table from row 3. `error_from`, `assert_stops_with`: run `clean_workbook` and check how the error message starts. `reorder_quarters(labels)`: quarter rows in a given order. | Broken workbooks stop with the sheet name and Excel address: missing columns, two headers with one meaning, unknown headers, quarters out of order, a budget row with actuals, unreadable text, Excel error cells. `test_good_workbook_cleans` proves the starting workbook is valid, so each failure comes from the one thing that was broken. |

---

### Small files

| File | What it is |
|---|---|
| `requirements.txt` | The packages: pandas (tables), openpyxl (Excel), python-pptx and matplotlib (for the deck, step 4), anthropic (Claude API), python-dotenv (.env), pyyaml (config.yaml), pydantic (answer shape), pytest (tests). |
| `pytest.ini` | Tells pytest to look only in `tests/`, and lets tests `import clean` from the project folder. |
| `.env` / `.env.example` | `.env` holds `ANTHROPIC_API_KEY` and is never committed. `.env.example` shows the variable name with no key. |
| `.gitignore` | Keeps `.env`, `.venv/`, `__pycache__/`, `output/` and `.DS_Store` out of git. |
| `CLAUDE.md` | The project spec: goal, rules, metric definitions, edge cases, build order. |
| `LEARNINGS.md` | Everything that broke and how it was fixed, plus the prompt iterations and model comparison. **Interview gold.** |
| `OVERNIGHT_REPORT.md` | What each overnight task built, the decisions made, what failed, what's unresolved. |
| `README.md` | For now, just the model comparison table. The full README is build step 6. |
| `templates/` | Empty for now; `base.pptx` for the deck goes here (step 4). |
| `output/` | Generated files (git-ignored): `*_metrics.xlsx`, `northwind_analysis*.json`, `compare/`. |

---

## 5. 25 interview questions

Short answers you can say out loud in 30–60 seconds. **"Point to"** says where to look in the code or docs if they ask for detail. Practise saying the answers in your own words; don't memorize them.

### The project

**Q1. Walk me through the project in one minute.**
A messy portfolio-company KPI workbook goes in; a board pack comes out. `clean.py` reads the messy Excel: odd headers, "$14.3M" typed as text, a blank quarter. It produces a standard table in $K. `metrics.py` calculates NRR, GRR, burn multiple, runway, Rule of 40, CAC payback and budget variances, then checks 9 flags against thresholds in `config.yaml` and lists any data gaps. `excel_output.py` writes a highlighted Excel summary, and `analyze.py` has Claude write the headline, wins, risks and questions for management. Code then checks that Claude used only numbers from the data. `main.py` runs it all for one company or a whole folder. The deck itself is the next build step.
*Point to:* section 2 of this guide.

**Q2. Why would a PE fund want this?**
Portfolio companies report KPIs in their own formats, so analysts spend time re-typing and reconciling before they can think. This standardizes the file, applies the **same** definitions and thresholds to every company, and flags what needs attention, so the analyst's time goes to the questions for management. It also makes gaps explicit instead of letting a missing quarter quietly distort a trend. The cost estimate for the AI step is about $14.65 per quarter for 275 companies.
*Point to:* README.md model comparison.

**Q3. Why does Python do all the math and Claude only interpret?**
A language model can make arithmetic mistakes, and one wrong number in a board deck undermines every other number in it. Python math is deterministic and can be tested against hand formulas; a prompt can't be tested the same way. So every number is computed and checked in Python first, Claude receives them already formatted ("97.1%"), and the prompt tells it to quote, not calculate. Code then verifies the answer.
*Point to:* CLAUDE.md Rules; `analyze.py` `SYSTEM_PROMPT`.

**Q4. Why build three fake companies instead of one?**
With one company you can't tell whether the flags work or were tuned to produce that company's story. Northwind trips 6 of 9, Alderpeak (healthy) trips none in any quarter, and Fernhollow (distressed) trips 7, with 1 flag it can't evaluate. That proves the logic responds to the data. Each company also has different mess (a title row, a different column order, the Notes tab first), so cleaning isn't tuned to one file either.
*Point to:* `check_companies.py`; OVERNIGHT_REPORT Task 1.

### Design decisions

**Q5. How do you handle missing data? Why not fill it in?**
A blank cell stays blank (NaN), never 0 and never an estimate: an imputed number could reach a board as if it were real. Because any math with NaN gives NaN, a blank quarter automatically spreads to every metric that uses it: QoQ metrics for that quarter and the next, YoY for that quarter and the one 4 later. Flags that depend on a missing value say "cannot evaluate — data missing" instead of pass or fail, and every affected metric and flag is listed as a data gap.
*Point to:* `metrics.data_gaps`; CLAUDE.md "Messy data rules".

**Q6. What's the difference between "data missing" and "not meaningful"?**
YoY growth in the first year of data is blank because there's no earlier year: that's "n/a (no prior period)", not a problem. YoY in Q1 2026 is blank because Q1 2025 exists but wasn't reported: that's "data missing". Infinite values have their own reason too: burn multiple "∞ (ARR shrank)", runway "∞ (not burning)". CLAUDE.md says the two must never look alike, because a reader would draw different conclusions. `data_gaps` tells them apart by asking whether the quarter has enough history for that metric.
*Point to:* `metrics.data_gaps`, `analyze.metric_trend`, `excel_output.cell_value`.

**Q7. Why store ratios as decimals and format as % only at output?**
One representation everywhere means thresholds, comparisons and math never mix 97.1 with 0.971. Formatting happens in one place (`format_value`, or an Excel number format), so the Excel cell still holds 0.971 and works in formulas while displaying 97.1%. The Excel check deliberately tested saving 97.1 instead of 0.971, and it's caught.
*Point to:* `metrics.format_value`, `excel_output.number_format`.

**Q8. Why is NRR annualized, and what's the trade-off?**
The thresholds (NRR 100%, GRR 85%) are annual conventions. One quarter's expansion minus churn is roughly a quarter of the annual effect, so it's multiplied by 4 to compare like with like. The trade-off: one quarter's data is noisier than a true trailing-12-month cohort. So the deck labels it "annualized", and the combo rule looks at a 3-quarter trend rather than a single quarter.
*Point to:* `metrics.nrr`, `metrics.grr`; CLAUDE.md metric definitions.

**Q9. Explain the combo rule. Why "cannot evaluate" instead of False?**
NRR falling while pipeline is rising suggests a retention problem, not a sales problem: sales keeps filling the funnel, but existing customers are leaking, so more pipeline won't fix it. It trips only if NRR fell at **every** step and pipeline rose at **every** step over the last 3 quarters. If any quarter in the window is missing, returning False would tell the board "no retention problem" when we simply don't know. So it returns "cannot evaluate". Fernhollow passes because its pipeline is falling too: that's a sales problem **and** a retention problem.
*Point to:* `metrics.check_combo`; config.yaml comments.

**Q10. Why is "exactly at the threshold" a pass, and why round before comparing?**
The thresholds are limits ("more than 15% over budget"), so hitting 15.0% exactly is within the limit. Computers store decimals in binary, so 3900/3250 − 1 comes out as 0.19999999999999996, not 0.2. A burn exactly at a threshold could trip or pass on invisible noise. `check_threshold` rounds to 6 decimals first. A test proves a real miss of 0.000001 still trips, so rounding doesn't hide real breaches.
*Point to:* `metrics.check_threshold`; LEARNINGS row 1; `tests/test_metrics.py` `test_check_threshold_*`.

**Q11. Why are thresholds in config.yaml and the API key in .env?**
Thresholds are investment judgments, not code: a partner may want runway under 18 months instead of 12, and that shouldn't need a code change. Each line has a comment with the investor reasoning. The API key is a secret: it lives in `.env`, which `.gitignore` keeps out of git, and `python-dotenv` loads it at run time. `.env.example` shows the variable name without a key.
*Point to:* `config.yaml`, `.gitignore`, `analyze.main` (`load_dotenv`).

### What broke (LEARNINGS.md)

**Q12. Tell me about a bug you found. What did it teach you?**
Net new ARR vs budget compared actual **net** new ARR (after churn) with `budget_new_arr`, which is **gross** new sales. So churn counted against the actuals but not the budget, and Q2 2026 showed −17.0% instead of −19.0%. The checks didn't catch it, because the answer key had been written with the same wrong formula. It was caught in review. The fix: the budget side is now budget_arr this quarter minus last quarter, net on both sides. **Lesson:** an answer key only catches errors if it's built independently from the definition, not copied from the code. Every later test follows that rule.
*Point to:* `metrics.budget_net_new_arr`; LEARNINGS row 2.

**Q13. Why does clean.py use `Decimal` to read "$4.03M"?**
In normal float math, 4.03 × 1000 = 4030.0000000000005. `Decimal` does the math the way you would on paper, so "$4.03M" becomes exactly 4030. That matters because the checks require cleaned values to equal the answer key exactly, and because tiny errors can push a value across a threshold.
*Point to:* `clean.parse_number`; LEARNINGS row 1; OVERNIGHT_REPORT Task 4.

**Q14. A cell saying "n/a" was being read as blank. Why was that dangerous, and how did you fix it?**
pandas quietly turns a built-in list of words ("n/a", "NA", "NULL", "None") and every Excel error like `#DIV/0!` into NaN before our code sees them. So a broken formula would have shown on the deck as "data missing" instead of stopping the run to be fixed. A unit test said "n/a" stops, but it only tested `parse_number` on its own, never through pandas. The fix was to read with `keep_default_na=False` so only a truly empty cell is blank, plus `check_no_error_cells` using openpyxl. **Lesson:** test through the real file, not just the helper.
*Point to:* `clean.find_kpi_sheet`, `clean.check_no_error_cells`; LEARNINGS row 6.

**Q15. What happened with a duplicated quarter row?**
Rows were stored in a dictionary keyed by quarter label, and saving a second "Q2 2025" silently overwrote the first. The quarter-order check does catch repeats, but it never saw the duplicate because the dictionary only held one Q2. Now `clean_sheet` stops with "quarter 'Q2 2025' appears twice (also in row 5) - delete one of the rows". Same lesson as Q14: a check can only catch what reaches it.
*Point to:* `clean.clean_sheet`; LEARNINGS row 5.

**Q16. Your Excel read-back check failed on its first run. Why?**
NRR 1.0706921944035346 came back as 1.070692194403535. openpyxl saves 16 significant digits, but a Python float can need 17. The gap is about 1e-16, and Excel itself only keeps 15 digits, so the check compares to 15 significant digits. It still catches real errors like 97.1 saved instead of 0.971. The same investigation found openpyxl silently saves NaN and infinity as an **empty cell**, so `excel_output.py` writes a label that says why instead ("data missing", "∞ (ARR shrank)").
*Point to:* `check_excel_output.same_as_saved`, `excel_output.cell_value`; LEARNINGS row 4.

### The AI part

**Q17. How do you stop Claude from inventing or calculating numbers?**
Four layers:
1. Claude only gets pre-formatted facts.
2. The prompt says to quote numbers exactly and never calculate differences or changes ("went from A to B" instead).
3. `validate_summary` extracts every number from the answer and fails it if any number doesn't appear in the payload. A rounded "97%" or a calculated "11.8 points" gets caught.
4. The saved JSON keeps the payload next to the answer, so a reviewer can trace every claim.

*Point to:* `analyze.find_ungrounded_numbers`, `check_northwind.check_analysis_validation`.

**Q18. What are the known limits of that number check?**
It asks "does this number appear anywhere in the data?", not "is it used correctly". So:
- A calculated number that happens to equal another value passes. "3.3 months faster" passed because −3.3% is a Rule of 40 value.
- It ignores minus signs.
- It can't catch a wrong direction ("improved… down from 20.3 mo" when it went from 20.3 to 20.7, which is worse).
- It can't catch invented attributions ("Management asserts…").
- Questions can come back as raw JSON strings, because the schema only requires strings.

These are logged as open items. That's why human review stays in the loop.
*Point to:* LEARNINGS "Validator gaps found in the blind answers".

**Q19. What happens if Claude's answer fails validation?**
`analyze()` retries **once**, sending back Claude's previous answer and the exact list of problems. If the second answer also fails, it raises `AnalysisError`; nothing unvalidated is ever returned. Both attempts are logged with tokens and timing. One retry is the balance between cost and latency and giving Claude a fair chance to fix a specific problem. Repeated failure points to a prompt or model problem that more retries would just pay for.
*Point to:* `analyze.analyze`, `analyze.build_messages`.

**Q20. How did you improve the prompt?**
Each round: run, find a specific flaw, add one rule, re-run, check the result. v1 → v2 fixed three framing problems:
- The best win (ARR growth YoY 42.8%) was missing.
- A passing flag was framed as a breach.
- NRR's decline was described from the first quarter in the data instead of the peak.

But v2 doubled the output (3,218 → 6,955 tokens, ~$0.043 → ~$0.081 per run) and made details too long for a slide. v3 added a 2-sentence, ~40-word limit that **code enforces**, and banned "significant" for variances under 10%. Output fell to 2,090 tokens, 21.1s, $0.0325.
*Point to:* LEARNINGS "Prompt iterations".

**Q21. Which model did you pick, and how did you decide?**
Sonnet 5 vs Haiku 4.5, 3 runs each on identical input, answers shuffled and scored blind on a 1–5 rubric. The decision rule was written **before** running: Haiku becomes the default only if it averages ≥ 4.0 with a 100% pass rate.
- **Sonnet:** average score 4.0, $0.0533 and 36.0s per run.
- **Haiku:** average score 2.0, $0.0173 and 13.2s per run, and it needed its retry on every run.

So Sonnet stays the default, at about $14.65 per quarter for 275 companies versus $4.77 for Haiku.
*Point to:* README.md; `compare_models.recommend`.

**Q22. What did blind scoring teach you?**
Before scoring, the prediction was that the most polished-sounding answer would score highest. It tied for lowest. The fluent writing hid a number Claude calculated (24.0 − 20.7 = "3.3 months faster") and a false claim ("consistently missing budget", when Q4 2025 beat budget by +16.5%). It had passed every code check. **Lesson:** judge line by line against the data and the rubric, not by how it reads. Code checks catch rule breaks at scale, blind scoring removes bias toward a model, and line-by-line review catches what code can't yet.
*Point to:* LEARNINGS "Reflection".

### Reliability and scale

**Q23. How do you know the numbers are right?**
Four layers:
1. The fake data's answer key is checked to tie out: ARR and cash roll forward.
2. The `check_*.py` scripts run the real workbooks end to end and compare with hand formulas typed like Excel (`3900 / (1850 + 580 - 260 - 510)`), written independently of the code.
3. 212 pytest unit tests check each function, including every edge case, with the hand math in comments.
4. The tests were tested: the code was broken on purpose (in throwaway copies) to confirm the tests notice. 35 of 36 breaks were caught in Task 4 (the missed one can't change any result) and 18 of 18 in Task 5.

*Point to:* `check_companies.py`, `tests/`, OVERNIGHT_REPORT "What failed".

**Q24. In a batch of 275 companies, what happens when one workbook is broken?**
`clean.py` stops that company with a message naming the sheet and cell, e.g. `Sheet 'KPI Tracker', cell G5 (Q2 2025, revenue): Can't read 'TBD' as a number`. `main.py` catches it, prints one ✗ FAILED line, and moves on to the next company. The summary table shows FAILED with the reason, and the run exits with code 1 so a scheduler knows. A bad-input error gets one clean line; an unexpected error (probably a bug) also gets a full traceback.
*Point to:* `main.run_batch`, `main.describe_error`; `check_main.py`.

**Q25. What would you improve or build next?**
Next are the deck (step 4) and the README (step 6). Beyond that, several open definition questions are already logged for a decision:
- **Burn vs budget when a company isn't burning:** burn −200 vs budget −150 comes out +33% and trips.
- **Net new ARR vs budget when budgeted net new ARR is zero or negative:** the sign flips.
- **The combo rule has no minimum size of decline:** a 0.1-point dip counts.
- **`data_gaps` over-reports** a partly blank quarter in the first year.
- **The budget-only row's label isn't checked** against the latest quarter.

For scale: at ~36s per company, 275 companies run one at a time would take about 2.75 hours with AI, so run companies in parallel. Also save the summary table to a file, and warn if companies report different latest quarters.
*Point to:* OVERNIGHT_REPORT "Unresolved" sections.

---

## 6. 10 trace-this-number exercises

**For each number:** name every file and function it passes through, from the workbook cell to the screen, and write out the arithmetic. Answers are in [section 8](#8-answers-to-the-exercises). Try each one first.

**Set up:** run `python metrics.py data/northwind.xlsx` and keep the output open. It helps to open `data/northwind.xlsx` too. In the KPI Tracker tab, row 1 is the header, Q3 2024 is row 2, the blank Q1 2025 is row 4, Q2 2026 is row 9, and the budget-only row is row 10.

**Tools you'll need:** `clean.py` (reading), `metrics.py` (math and flags), `config.yaml` (thresholds), `make_data.py` (where the number was born). Exercise 10 also uses `main.py`.

### Exercise 1: Northwind NRR 97.1% (Q2 2026)
**Where you see it:** `TRIP  NRR (annualized)  97.1% (threshold 100.0%)`
**Your task:** Which 4 input cells does it come from, and which one is stored as text? Which function turns that text into a number? Which function computes NRR, and what's the formula with Northwind's numbers? Why does it trip? Where does 0.97… become "97.1%"?

### Exercise 2: Northwind runway 11.0 mo (Q2 2026)
**Where you see it:** `TRIP  Runway (months)  11.0 mo (threshold 12.0 mo)`
**Your task:** Ending cash is typed as `"$14.3M"`. Trace exactly how `parse_number` turns that text into a number, step by step. Then compute runway. Why divide the burn by 3?

### Exercise 3: Northwind burn vs budget 20.0%
**Where you see it:** `TRIP  Burn vs budget  20.0% (threshold 15.0%)`
**Your task:** Which two cells, and how does the header `"Net Burn (Bud.)"` become `budget_net_burn`? What does Python actually compute for this ratio (not what a calculator shows), and which line of code handles that? Trap: the Notes tab also has 3900, 3250 and "20% over??" in row 7. Does the deck's number come from there? How do you know?

### Exercise 4: Northwind burn multiple 2.35x
**Where you see it:** `TRIP  Burn multiple  2.35x (threshold 2.00x)`
**Your task:** Compute net new ARR first, then the burn multiple. Which two edge-case lines in `burn_multiple` did this quarter **not** hit, and why?

### Exercise 5: Northwind Rule of 40 −12.2%
**Where you see it:** `TRIP  Rule of 40  -12.2% (threshold 40.0%)`
**Your task:** Which two metrics are added? Which quarter's revenue does `.shift(4)` compare Q2 2026 with, and is that quarter blank? Write both parts as decimals, then add them.

### Exercise 6: Northwind net new ARR vs budget −19.0%
**Where you see it:** `PASS  Net new ARR vs budget  -19.0% (threshold -20.0%)`
**Your task:** Compute budgeted net new ARR from two cells in the Budget ARR column. Why does −19.0% **pass** against a −20.0% threshold? Bonus: before the fix in LEARNINGS.md this showed −17.0%. Which number was used as the budget then, and why was that wrong?

### Exercise 7: runway if burn returns to plan, 13.0 mo
**Where you see it:** `Runway at next quarter's budgeted burn: 13.0 mo`
**Your task:** Which row of the workbook does the budget come from, and which function decides that row is budget-only? Which cash figure is used? Why is this number never flagged, even though 13.0 is close to the 12-month threshold?

### Exercise 8: two blanks that look the same
**Where you see it:** in the `metrics.py` printout, `arr_yoy` shows `n/a` for **both** Q4 2024 and Q1 2026. In the Excel Metrics sheet (`python main.py --all --skip-ai`, then open `output/northwind_metrics.xlsx`), Q4 2024 reads "n/a (no prior period)" and Q1 2026 reads "data missing" in a gray cell.
**Your task:** Why is each one blank? Which function decides that only one of them is a data gap, and what are the two conditions it checks? Which functions then pick the words for Excel and for Claude?

### Exercise 9: the combo rule trips for Northwind
**Where you see it:** `TRIP  NRR falling while pipeline rising`
**Your task:** Which 3 quarters are in the window, and where does "3" come from? Give the NRR and pipeline value for each. Show that every step qualifies. Northwind's NRR peak was 108.9% in Q3 2025: why isn't that quarter part of this test?

### Exercise 10: Fernhollow "7 of 9, 1 cannot evaluate"
**Where you see it:** run `python main.py --all --skip-ai`. The summary row says `Fernhollow  7 of 9, 1 cannot evaluate  20 metrics/flags (blank: Q2 2025)`.
**Your task:** Which flag can't be evaluated, and trace why, from the blank row to the words "cannot evaluate" to the summary text. Where does 20 come from? Bonus: burn multiple is one of the 7 trips, but its value is ∞. Why, and what does the Excel file show instead of ∞?

---

## 7. 5 functions to rewrite yourself

Rewriting a function from its description is the fastest way to be able to explain it. These five run from easiest to hardest, and each adds one new skill. The tests already exist, so you'll know right away whether your version is right.

### How to practise safely (same steps for all five)

1. Open the file and find the function. **Keep the `def` line and the docstring; delete only the code under them.**
2. Write your version from the spec below, without peeking at git.
3. Run the test command given. Red = keep going; all green = done.
4. Compare with the original: `git diff metrics.py` (or `clean.py`).
5. Put the original back: `git checkout -- metrics.py`. **This throws away your edits to that file**, which is what you want here. Don't commit practice edits.

If you get stuck for more than 20 minutes, read the original's first line, put it back, and try again tomorrow.

---

### 1. `nrr(df)` in `metrics.py`: easiest

**New skill:** math on whole pandas columns at once.

**Spec:** return annualized NRR for every quarter: 1 + 4 × (expansion − contraction − churn) / starting ARR. Columns: `df["expansion_arr"]`, `df["contraction_arr"]`, `df["churned_arr"]`, `df["starting_arr"]`. Return a decimal (1.05, not 105).

**Hints:**
- A column behaves like one number in a formula: `df["a"] - df["b"]` subtracts row by row, like filling a formula down.
- You don't need a loop, and you don't need to handle blanks: NaN in gives NaN out on its own.

**Traps the tests catch:** forgetting the 4 (not annualized); putting the brackets in the wrong place, so only churn is divided.

**Test:** `python -m pytest -q -k "nrr_annualized"`
**Stretch:** do `grr(df)` too: `python -m pytest -q -k "grr_annualized"`.

---

### 2. `growth(series, quarters_back)` in `metrics.py`

**New skill:** comparing each row with an earlier row (`.shift`).

**Spec:** return growth versus N quarters earlier: value ÷ value N quarters ago − 1. `quarters_back=1` is QoQ, `4` is YoY. The first N quarters have nothing to compare with and must come out NaN, not 0 and not an error.

**Hints:**
- `series.shift(1)` gives a new column where each quarter holds the value from one row up. The first row gets NaN.
- Try it first in a Python prompt: `pd.Series([100, 110, 121]).shift(1)`.

**Traps the tests catch:**
- `.shift(-N)` looks **forward** (wrong direction).
- Forgetting the − 1 gives 1.10 instead of 0.10.
- A blank quarter must spread to the **next** quarter's QoQ (the test `test_growth_blank_quarter_spreads_to_next_quarter`).

**Test:** `python -m pytest -q -k growth`
**Then check the real numbers:** `python check_companies.py` (ARR YoY for all three companies).

---

### 3. `check_threshold(value, threshold, kind)` in `metrics.py`

**New skill:** `if`/`else` logic, and why float noise matters.

**Spec:** return `MISSING`, `TRIP` or `PASS` (constants already defined at the top of the file).
- If `value` is NaN → `MISSING`.
- Otherwise round `value` to 6 decimals.
- `kind == "min"`: trip if the value is **below** the threshold.
- `kind == "max"`: trip if the value is **above** the threshold.
- **Exactly at the threshold passes.**

**Hints:**
- `math.isnan(value)` tests for NaN. Don't use `value == math.nan`: NaN is never equal to anything, not even itself.
- `round(value, 6)`.
- `<` vs `<=` is the whole "exactly at the threshold" rule.

**Traps the tests catch:**
- No rounding: 0.19999999999999996 vs 0.2 misbehaves.
- `<=` instead of `<`: exactly-at trips.
- ∞ must pass a minimum (runway when not burning) and trip a maximum (CAC payback that never pays back). Check yours does this without special code.
- A real miss of 0.000001 must still trip.

**Test:** `python -m pytest -q -k check_threshold` (15 tests)
**Then:** `python check_companies.py` (every flag for all three companies).

---

### 4. `parse_number(value)` in `clean.py`

**New skill:** handling messy text step by step, raising clear errors, and `Decimal`.

**Spec:** turn one cell into a number in $K. In order:
1. Blank (use `is_blank`) → `float("nan")`.
2. `True`/`False` → raise `ValueError` (message must start with `Can't read`).
3. Already a number (`isinstance(value, numbers.Real)`) → `float(value)`.
4. Otherwise it's text. Remove `$` and spaces, uppercase it, and raise `ValueError(f"Can't read {value!r} as a number{NUMBER_HINT}")` unless it matches `NUMBER_TEXT`.
5. If it ends in `M`, remove the M and multiply by 1000. If it ends in `K`, just remove the K.
6. Remove commas and return `float(Decimal(text) * multiplier)`.

**Hints:**
- `text.endswith("M")`, `text[:-1]` (everything except the last character), `text.replace(",", "")`.
- `NUMBER_TEXT.match(text)` returns None if the text doesn't fit.

**Traps the tests catch:**
- **Order matters:** in Python, `True` counts as a number (`isinstance(True, numbers.Real)` is True). So the TRUE/FALSE check must come **before** step 3.
- Using `float(text) * 1000` instead of `Decimal`: `"$4.03M"` becomes 4030.0000000000005.
- A blank cell must be NaN, never 0.
- `"n/a"`, `"(120)"`, `"12.5%"`, `"inf"`, `"1e3"` must all stop.

**Test:** `python -m pytest -q tests/test_clean.py tests/test_bad_inputs.py` (134 tests; the second file reads real workbooks through your function)
**Then:** `python check_companies.py` (cleaning must recover every answer-key value exactly).

---

### 5. `check_combo(metrics, config, quarter)` in `metrics.py`: hardest

**New skill:** taking a window of rows, comparing steps, and treating "can't tell" as its own answer.

**Spec:** return `TRIP` if NRR fell at **every** step **and** pipeline rose at **every** step over the last `config["combo_lookback_quarters"]` quarters, ending at `quarter`. Otherwise `PASS`. Return `MISSING` if:
- there isn't enough history for a full window, or
- any NRR or pipeline value in the window is NaN.

**Hints:**
- `metrics.index.get_loc(quarter)` gives the row position of a quarter label (Q3 2024 = 0).
- `metrics.iloc[start:end]` takes rows `start` up to but **not including** `end`. So for a window of 3 ending at position 7, you want `iloc[5:8]`.
- `.isna().any()` → True if any value is NaN.
- `.diff()` gives this row minus the row above. The first row of the window has nothing above it inside the window, so skip it with `.iloc[1:]`.
- `(steps < 0).all()` → True only if every step is negative.
- Round NRR to 6 decimals **before** `.diff()`, so a tiny float wobble doesn't count as a decline.

**Traps the tests catch:**
- A flat step counts as falling (`<= 0` instead of `< 0`).
- Missing data returns `PASS` instead of `MISSING`.
- A window that's one row off.
- Older quarters outside the window affect the result.
- A quarter other than the latest is evaluated as if it were the latest (`test_check_combo_earlier_quarter`).

**Test:** `python -m pytest -q -k check_combo` (15 tests)
**Then:** `python metrics.py data/northwind.xlsx` should still say `TRIP  NRR falling while pipeline rising`, and `python check_companies.py` should pass (Alderpeak's zigzag NRR must not trip in any quarter).

---

## 8. Answers to the exercises

### Answer 1: NRR 97.1%

1. **Born in** `make_data.py` → `TRUE_DATA`, Q2 2026: starting ARR 25810, expansion 580, contraction 260, churn 510. `TEXT_CELLS` has `("contraction_arr", "Q2 2026"): "K"`, so `make_data_common.to_text` writes contraction as the text `"260K"`.
2. **In the workbook:** KPI Tracker row 9: B9 = 25810 ("Beginning ARR"), D9 = 580 ("expansion arr"), **E9 = "260K"** ("Contraction_ARR"), F9 = 510 (" Churned ARR ").
3. **`clean.py`:** `clean_workbook` → `find_kpi_sheet` → `clean_sheet` → `map_columns` → `name_headers` → `header_name` → `standard_column` → `normalize_header`. `"Beginning ARR"` → `beginning_arr` → alias → `starting_arr`; `" Churned ARR "` → `churned_arr`. Then `parse_row` → `parse_number("260K")` → 260.0.
4. **`metrics.py`:** `compute_metrics` → `nrr(df)`:
   1 + 4 × (580 − 260 − 510) / 25810 = 1 + 4 × (−190) / 25810 = 1 − 760 / 25810 = **0.970554**, stored as a decimal.
5. **Flag:** `evaluate_flags` → `check_threshold(0.970554, 1.00, "min")`. It's below the minimum, so **trip**.
6. **Display:** `format_value("nrr", 0.970554)` → `f"{value:.1%}"` → **"97.1%"**. Claude sees the same text through `analyze.describe_flag`; Excel stores 0.970554 with the `0.0%` format.

**Proven by:** `check_northwind.EXPECTED_LATEST["nrr"]`, typed as `1 + 4 * (580 - 260 - 510) / 25810`.

### Answer 2: runway 11.0 mo

1. **Workbook:** J9 = `"$14.3M"` ("Cash - End of Qtr" → alias → `ending_cash`); I9 = 3900 ("Net Burn").
2. **`clean.parse_number("$14.3M")`:**
   - Not blank, not TRUE/FALSE, not already a number, so it's text.
   - Remove `$` and spaces, uppercase: `"14.3M"`.
   - It matches `NUMBER_TEXT` (digits, decimals, M).
   - It ends with M: multiplier = 1000, text = `"14.3"`.
   - `Decimal("14.3") * 1000` = exactly 14300 → `float` → **14300.0**.
3. **`metrics.runway_months`:** 14300 / (3900 / 3) = 14300 / 1300 = **11.0**. Net burn is **per quarter**, and a quarter has 3 months, so burn ÷ 3 = monthly burn. The `.mask` edge case (not burning → ∞) doesn't apply, because burn is 3900 > 0.
4. **Flag:** `runway_min_months: 12`. 11.0 < 12 → **trip**. `format_value` → "11.0 mo".

### Answer 3: burn vs budget 20.0%

1. **Workbook:** I9 = 3900 ("Net Burn" → `net_burn`), Q9 = 3250 ("Net Burn (Bud.)").
2. **The header:** `normalize_header("Net Burn (Bud.)")`: lowercase `"net burn (bud.)"`; each run of symbols and spaces becomes `_` → `"net_burn_bud_"`; strip the end `_` → `"net_burn_bud"`. `HEADER_ALIASES["net_burn_bud"]` = `"budget_net_burn"`.
3. **`metrics.burn_vs_budget`:** 3900 / 3250 − 1. A calculator says 0.2; **Python gives 0.19999999999999996**, because floats are stored in binary.
4. **`metrics.check_threshold`:** `value = round(value, 6)` turns it into 0.2. Then "max" with `burn_over_budget_max: 0.15`: 0.2 > 0.15 → **trip**. (The rounding line matters for a burn **exactly** 15% over budget, not here.)
5. **The trap:** no. `find_kpi_sheet` loads every tab, but only uses one with a "Quarter" header in its first 10 rows. The Notes tab has none, so its cells never reach the table. Its "3900 / 3250 / 20% over??" is someone's scratch math, left in on purpose (`make_data.py` → `NOTES`) to prove junk tabs are ignored.

### Answer 4: burn multiple 2.35x

1. **`metrics.net_new_arr`:** new C9 1850 + expansion 580 − contraction 260 − churn 510 = **1,660**.
2. **`metrics.burn_multiple`:** `result = 3900 / 1660` = **2.349**.
3. **Edge cases not hit:**
   - `.mask(net_burn <= 0, 0.0)`: burn is 3900, so the company **is** burning.
   - `.mask((net_burn > 0) & (new <= 0), inf)`: net new ARR is 1660 > 0, so ARR didn't shrink.
4. **Flag:** `burn_multiple_max: 2.0`. 2.349 > 2.0 → **trip**. `format_value` → `f"{value:.2f}x"` → "2.35x".

### Answer 5: Rule of 40 −12.2%

1. **`metrics.rule_of_40`** = `growth(df["revenue"], 4)` + `fcf_margin(df)`.
2. **Revenue YoY:** `.shift(4)` moves 4 rows up: Q1 2026, Q4 2025, Q3 2025, **Q2 2025**. Q2 2025 revenue is G5 = 4550. It is **not** blank; the blank quarter is Q1 2025. So 6660 / 4550 − 1 = **0.463736** (46.4%, also shown as `revenue_yoy`).
3. **FCF margin:** −3900 / 6660 = **−0.585586** (−58.6%, shown as `fcf_margin`).
4. **Sum:** 0.463736 + (−0.585586) = **−0.121849** → "−12.2%".
5. **Flag:** `rule_of_40_min: 0.40`. −0.122 < 0.40 → **trip**. Strong growth, but the burn more than cancels it.

### Answer 6: net new ARR vs budget −19.0%

1. **`metrics.budget_net_new_arr`:** `budget_arr − budget_arr.shift(1)` = P9 26800 − P8 24750 = **2,050**.
2. **`metrics.net_new_arr_vs_budget`:** 1660 / 2050 − 1 = **−0.190244** → "−19.0%".
3. **Why it passes:** `net_new_arr_vs_budget_min: -0.20` is a **minimum**, so it trips only **below** −20%. −19.0% is above −20.0% (a smaller miss), so it's a **pass**. Barely: that's why the v2 prompt rule says to call it "passed but close to its threshold".
4. **Bonus:** before the fix, the budget side was `budget_new_arr` = 2000 (O9): 1660 / 2000 − 1 = −17.0%. That compared **net** new ARR (after expansion and churn) with **gross** new sales, so churn counted against the actuals but not the budget. Now both sides are net. Because it looks back one quarter, its gaps follow the QoQ rule (Q1 + Q2 2025).

### Answer 7: runway if burn returns to plan, 13.0 mo

1. **Workbook:** row 10, A10 = `"Q3 2026 (Budget)"`, Q10 = 3300 (the other actual columns are empty).
2. **`clean.py`:** `clean_sheet` → `parse_row` → **`is_budget_only_row`**. `"budget"` is in the label (`BUDGET_LABEL_WORDS`), and no actual column has a value, so it returns True. `clean_sheet` stores it as `next_budget`, not as a quarter. If that row also had revenue filled in, it would stop with an error.
3. **`metrics.runway_at_next_budget`:** cash = `actuals["ending_cash"].iloc[-1]` = the **latest actual** cash, Q2 2026 = 14300. Budgeted burn = 3300. 14300 / (3300 / 3) = 14300 / 1100 = **13.0**.
4. **Never flagged:** it isn't in `FLAG_RULES`. CLAUDE.md says the runway flag uses **current** burn, and this number is context: "if the company got back to plan". It appears as `runway_if_burn_returns_to_plan` in the Claude payload (the prompt forbids calling it a projection) and as an uncolored "(context, not a flag)" row on the Excel Flags sheet.

### Answer 8: two blanks that look the same

1. **Why each is blank:** `arr_yoy = growth(ending_arr, 4)`.
   - **Q4 2024** is the 2nd row. `.shift(4)` points above the top of the table, where nothing exists → NaN.
   - **Q1 2026** points at **Q1 2025**, the blank quarter: its ending ARR is NaN, and NaN math gives NaN.
2. **`metrics.data_gaps`** counts a NaN as a gap if **either**:
   - **has_history:** the quarter's position ≥ `METRIC_LOOKBACK["arr_yoy"]` (4). Q4 2024 is position 1 → no. Q1 2026 is position 6 → **yes**.
   - **incomplete:** the quarter's **own** inputs have a blank. Q4 2024 → no. (That's why Q1 2025 itself is also a gap, even though it's only position 2.)

   So Q4 2024 is not a gap and Q1 2026 is. `arr_yoy` gaps: Q1 2025, Q1 2026.
3. **The words:** `excel_output.cell_value(column, value, is_gap)` → "data missing" (plus gray fill from `style_metric_cell`) or "n/a (no prior period)". For Claude, `analyze.metric_trend` makes the same choice.
4. **Why the printout says "n/a" for both:** `metrics.print_report` passes every value straight to `metrics.format_value`, which doesn't know about gaps and turns any NaN into "n/a". `analyze.metric_trend` and `excel_output.cell_value` check the gap list **before** formatting, which is why the board-facing outputs tell the two apart. A related leftover: a "cannot evaluate" flag's value still reaches Claude as a bare "n/a" through `analyze.describe_flag` (OVERNIGHT_REPORT Task 1, unresolved item 3).

### Answer 9: the combo rule trips

1. **Window:** `combo_lookback_quarters: 3` in config.yaml. In `metrics.check_combo`: `end` = position of Q2 2026 (7) + 1 = 8, so the window is rows 5–7 = **Q4 2025, Q1 2026, Q2 2026**.
2. **NRR** (`metrics.nrr`):
   - Q4 2025: 1 + 4 × (860 − 140 − 290) / 21460 = **1.0801** (108.0%)
   - Q1 2026: 1 + 4 × (710 − 200 − 390) / 23790 = **1.0202** (102.0%)
   - Q2 2026: **0.9706** (97.1%)
3. **Pipeline** (text cells `"$10.1M"`, `"$11.2M"`, `"$12.5M"` in N7–N9 → `parse_number`): **10,100 → 11,200 → 12,500**.
4. **The test:** no value in the window is NaN, so it can be evaluated. NRR, rounded to 6 decimals, then `.diff()`: −0.0600, −0.0496, **all < 0**. Pipeline `.diff()`: +1,100, +1,300, **all > 0**. Both true → **trip**.
5. **Why not Q3 2025's 108.9% peak:** the window is exactly 3 quarters (2 steps) ending at the evaluated quarter. Q3 2025 → Q4 2025 was also a decline, but that step sits outside the window. That's also why the v2 prompt rule says to describe a trend "from the peak, or from the start of the flag's lookback window".

### Answer 10: Fernhollow "7 of 9, 1 cannot evaluate"

1. **The blank:** `make_data_fernhollow.py` → `BLANK_QUARTER = "Q2 2025"`. `write_kpi_sheet` writes only the label. `clean_sheet` keeps it as a row of NaN.
2. **The flag: Rule of 40.** `rule_of_40` = `growth(revenue, 4)` + `fcf_margin`. For Q2 2026, `.shift(4)` lands on **Q2 2025**: blank. So revenue YoY is NaN, and NaN + anything = NaN.
3. **"Cannot evaluate":** `evaluate_flags` → `check_threshold(NaN, 0.40, "min")`. The first line, `if math.isnan(value): return MISSING`, returns `"cannot evaluate — data missing"`, never a pass or a trip.
4. **The summary text:** `main.run_company` collects `result["cannot_evaluate"] = ["Rule of 40"]`. `main.flags_text` → `"7 of 9"` + `", 1 cannot evaluate"`.
5. **Where 20 comes from:** `data_gaps` finds 19 metric columns with a gap (every metric misses Q2 2025; QoQ ones also Q3 2025; YoY ones also Q2 2026). Then `if flag["status"] == MISSING` adds `"flag: Rule of 40": ["Q2 2026"]`: 19 + 1 = **20**. `main.blank_quarters` → Q2 2025. `main.gaps_text` → "20 metrics/flags (blank: Q2 2025)".
6. **Bonus, burn multiple ∞:** net new ARR = 180 + 60 − 150 − 330 = **−240**, while net burn is 1650 > 0. In `burn_multiple`, `.mask((net_burn > 0) & (new <= 0), math.inf)` → ∞ (burning cash while ARR shrank). `check_threshold`: ∞ > 2.0 → trip. Excel can't store ∞, so `excel_output.cell_value` writes **"∞ (ARR shrank)"** from `INFINITE_LABELS`.

**Why this company was designed like this:** Fernhollow's blank is exactly 4 quarters before the latest one, so the "cannot evaluate" path is exercised on real data. Northwind's blank quarter (Q1 2025) is 5 quarters back, so none of its latest-quarter flags are affected.
