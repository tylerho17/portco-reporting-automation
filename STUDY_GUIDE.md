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
| `clean_workbook(path)` | **The entry point.** Finds the KPI tab, checks for error cells, runs `clean_sheet`, and puts `Sheet 'KPI Tracker', ` in front of any error. | Every other file calls this one function to read a workbook. |
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
| `compute_metrics(actuals)` | Calls all of the above and puts the results in one table: a row per quarter, 19 metric columns. `pipeline` is copied in for the combo rule. | The table `python metrics.py` prints. |

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
| `ai_step(skip_ai)` | For now only returns why AI was skipped: "--skip-ai" or "build_deck.py doesn't exist yet". Raises `NotWiredError` if the deck file exists. | No API money is spent while there's no deck to put the commentary in. |
| `deck_step()` | Same for the deck. | |
| `blank_quarters(actuals)` | Quarters with at least one blank input. | Northwind → `["Q1 2025"]`. |
| `run_company(workbook_path, config, skip_ai)` | Clean → metrics → flags → gaps → save Excel → AI step → deck step, printing a ✓ line for each. Returns a result dict for the summary table. | |
| `describe_error(error)` | Prints `✗ FAILED: <type>: <message>`. Bad-input errors (`ValueError`, `OSError`) get one line; anything else also gets a full traceback, because it's probably a bug. | A person fixing a workbook doesn't need a traceback; a developer fixing a bug does. |
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
| `append_learning(stats)` | Adds the result to LEARNINGS.md (once). |
| `command_score(score_args)` | Checks the scores, reveals the key, writes the README and LEARNINGS. |
| `main()` | The `run` / `score` command line. |

---

### `tests/`: unit tests (pytest)

Run with `python -m pytest -q` (212 tests, about 2 seconds). Expected values are **worked out by hand** in comments, not copied from running the code. Tests with `@pytest.mark.parametrize` run the same test on many inputs, each inputs line counting as one test.

| File | Helper functions | What the tests cover |
|---|---|---|
| `test_clean.py` (81) | `write_workbook(path, labels, blank)`: a tiny workbook in pytest's temp folder. | `parse_number` (good text, real numbers, blanks → NaN, 17 kinds of unreadable text stop), `normalize_header` and `standard_column` on every Northwind header, quarter labels and order (Q4 → Q1 rollover, skipped, repeated), and whole workbooks (blank row kept, missing row stops, duplicate row stops). |
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
