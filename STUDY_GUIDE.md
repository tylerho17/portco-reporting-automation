# Study Guide: Board Pack Generator

For you, a finance student learning Python, to understand this project well enough to explain every part of it in an interview.

It matches the code as of 2026-09-17 (branch `polish`, Task 5). **Every build step is done:** 1, 2, 3, 3b, 4 (the deck: `make_template.py`, `build_deck.py`, `charts.py`, `text_fit.py`), 4b, 5 (with the AI step connected to `main.py`) and 6 (the README). One live `main.py --all` run has been made. **Added since:** run manifests and the approval gate (`provenance.py`, `approve.py`), the review status in the deck's footer with the watermark made opt-in (`--draft`), the 4-slide deck, and the web page (`app.py`, started by `run_app.command`). **Not done yet:** the README screenshots, which need you at the screen.

## Contents

1. [How to use this guide](#1-how-to-use-this-guide)
2. [The data flow in plain English](#2-the-data-flow-in-plain-english)
3. [Python words you'll see](#3-python-words-youll-see)
4. [Every file, function by function](#4-every-file-function-by-function)
5. [30 interview questions](#5-30-interview-questions)
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
  python main.py --all --skip-ai             # the batch run: Excel files, and decks with the AI placeholder
  python build_deck.py data/northwind.xlsx   # one deck, using the saved AI analysis (output/northwind_analysis.json)
  python build_deck.py data/northwind.xlsx --draft   # the same, with the DRAFT watermark if nobody has approved it
  python approve.py northwind                # record yourself as the reviewer (writes output/northwind_manifest.json)
  python -m pytest -q                        # 500+ unit tests
  python check_companies.py                  # end-to-end proof for all 3 companies
  python check_deck.py                       # end-to-end proof for the 3 decks
  ```

  **These DO call the API and cost money:** `python main.py` **without** `--skip-ai` (about $0.09 per company), `python analyze.py ...` and `python compare_models.py run`. You don't need them to study: the analyses from the live run are already saved in `output/`.

  **One trap:** `python check_main.py` runs `main.py --all --skip-ai` into `output/`, so afterwards every deck shows "AI summary unavailable" and each manifest's `ai` record says "skipped". Run `python build_deck.py data/<company>.xlsx` to put the saved AI text back on the deck (no API call); the manifest's `ai` record stays "skipped" until the next live run.

  **Or use the web page:** double-click `run_app.command` (or `streamlit run app.py`), drag in `data/northwind.xlsx`, and download the deck. Leave "Include AI commentary" unticked and it's free; ticked, it reuses the saved analysis of the same numbers, also free.
- **Do the exercises in section 6 before you look at section 8.**

---

## 2. The data flow in plain English

### The one-sentence version

A messy Excel file from a portfolio company goes in. Python tidies it, calculates the KPIs, checks them against investor thresholds and lists what's missing. Then it writes an Excel summary, asks Claude to word the commentary (checking that Claude didn't invent or calculate any number), and builds a 4-slide PowerPoint deck.

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
| Build the deck on the fund's PowerPoint template, plus the backup Excel | `build_deck.py` (on `templates/base.pptx` from `make_template.py`) and `excel_output.py` |
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
output/northwind_       numbers → text → Claude →        templates/base.pptx (make_template.py)
metrics.xlsx            check → retry once →             + charts.py (2 matplotlib PNGs)
                        output/northwind_analysis.json ─▶ + text_fit.py (shrink to 12 pt, else stop)
                                                         re-checks the analysis, else "AI summary unavailable"
                                                         → output/northwind_board_pack.pptx

main.py      runs the chain for one workbook or every workbook in data/, always builds the deck,
             prints a summary table and saves output/batch_summary.csv
             and output/northwind_manifest.json (provenance.py: hashes, commit, model, cost)
app.py       the same chain for one workbook dropped onto a web page (run_app.command starts it)
approve.py   a person records their review in the manifest; the next build's footer says "reviewed by"
config.yaml  the flag thresholds            .env  the API key (never committed)
check_*.py and tests/   prove each step gives the right answer
```

### Step by step, with Northwind

**Step 0: make the fake input** (`make_data.py`). The script holds Northwind's true numbers, then messes them up on purpose the way real company files are messy. Headers are inconsistent (`" Churned ARR "` with stray spaces, `"Cash - End of Qtr"`). Some numbers are typed as text (`"$14.3M"`, `"260K"`, `"5,090"`). The Q1 2025 row has a label but no values. There's a budget-only row for next quarter and a junk Notes tab. Because the script keeps the true numbers, the checks can later prove that cleaning got every one of them back.

**Step 1: clean** (`clean.py`). Finds the tab that has a "Quarter" header and ignores the Notes tab. Translates each messy header to one of the 16 standard names from CLAUDE.md. Turns every cell into a plain number in $K ("$14.3M" becomes 14300). Leaves an empty cell empty (NaN), never 0. Splits off the budget-only row and checks that the quarters run in order with none skipped. **If anything is ambiguous, it stops** with a message naming the sheet and cell, e.g. `Sheet 'KPI Tracker', cell G5 (Q2 2025, revenue): Can't read 'TBD' as a number`. The output is a table with one row per quarter.

**Step 2: calculate** (`metrics.py`). Each metric is a small function that does column math on the whole table at once, like writing a formula in row 2 of Excel and filling it down. Ratios stay decimals (0.971, not 97.1). A blank input gives a blank result, so a missing quarter automatically spreads to every metric that uses it.

**Step 3: judge** (`metrics.py` + `config.yaml`). `evaluate_flags` compares the latest quarter with each threshold. Each flag gets one of 3 answers: **trip**, **pass**, or **cannot evaluate**, and "cannot evaluate" always says why: **missing input**, **no prior period** or **not meaningful**. The combo rule looks at the last 3 quarters: NRR falling (by at least 1 point each step) while pipeline is rising. `data_gaps` lists only metrics blank *because an input they use is missing*. That's different from a metric that's blank because there's no earlier period to compare with (YoY growth in the first year), or because the math is undefined (0 ÷ 0).

**Step 4: output.**
- `excel_output.py` writes a 3-sheet workbook: Metrics (red = tripped, gray = data missing), Flags, and Data gaps. It does no math.
- `analyze.py` turns the numbers into display text ("97.1%") and sends it to Claude with writing rules. It checks the answer: the right JSON shape, exactly 3 wins/risks/questions, short enough for a slide, and **every number Claude wrote must appear in the data**. If a check fails, it retries once and tells Claude what was wrong. The saved file keeps the data Claude saw next to what it wrote, so any claim can be traced back.
- `build_deck.py` builds the 4-slide deck on the brand template (`templates/base.pptx`, made once by `make_template.py`):
  1. **Key metrics:** a table of latest quarter, prior quarter, threshold and a red / green / gray status.
  2. **ARR and cash:** two charts from `charts.py`. The blank quarter is a visible gap.
  3. **Risks and flags:** "6 of 9 flags tripped" (counted by Python), each tripped flag vs its threshold, the combo rule, and the Data gaps line.
  4. **AI commentary:** a line under the title, "AI-drafted from computed metrics - review before use", then Claude's headline, and its 3 risks and 3 questions for management side by side. Claude still writes 3 wins, but the deck doesn't show them.

  **Before it uses Claude's text, it checks it again** against numbers rebuilt from today's workbook. If the analysis is missing, failed, is for another quarter or has a number that's no longer in the data, slide 4 says "AI summary unavailable". The other slides are built as normal, because their numbers come from Python. `text_fit.py` measures every piece of text and shrinks it to fit, down to 12 pt; below that, the build stops and names the slide and box.

  **Every slide's footer says where the deck came from and whether a person has reviewed it:** `Fictional data | northwind.xlsx | 2026-09-17 | 9c1b52c | claude-sonnet-5 | AI-drafted | not reviewed`. The last part becomes `AI-drafted | reviewed by Tyler Ho on 2026-09-17` once `approve.py` has recorded a reviewer. There's **no watermark** unless you build with `--draft`, which stamps "DRAFT - NOT REVIEWED" across every slide of a deck nobody has approved.

**Step 5: run the batch** (`main.py`). For each workbook it runs clean → metrics → Excel → AI → deck and prints a ✓ line per step. It catches failures so one broken company doesn't stop the rest. **The deck is always built:**
- **Claude's answer passed:** its text is on the deck, result `OK`.
- **The answer failed twice, or the API call failed:** the deck has the placeholder, result `OK (AI failed)`.
- **`--skip-ai`:** no API call, the placeholder, result `OK (AI skipped)`.

Without `--skip-ai`, it checks for the API key before any company runs. It finishes with a summary table, `output/batch_summary.csv`, and an exit code (0 = every company OK, 1 = a company failed, 2 = wrong arguments). `OK (AI failed)` still counts as OK: every output was built. Each company also gets `output/<company>_manifest.json`: the workbook and config.yaml by SHA-256 hash, the git commit, the model, tokens and cost, and any approval.

**Step 6: a person reviews and approves** (`approve.py`). Someone reads the deck, then runs `python approve.py northwind`. That writes their name and the time into the manifest, next to the hashes of the workbook and config.yaml the deck was built from. It never builds a deck. The next build's footer says "reviewed by ...". If the workbook or a threshold changes afterwards, the hashes no longer match and the footer goes back to "not reviewed" on its own: the reviewer approved numbers that aren't on the deck any more.

**The web page** (`app.py`) runs steps 1 to 5 for one workbook dragged onto a browser page, shows the flags and metrics in the Excel colors, and offers the deck and metrics workbook as downloads. It works in a temporary folder, so it never touches `output/`, and its decks always say "not reviewed".

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
| **python-pptx** | The package that writes PowerPoint files. `Presentation(path)` opens a file; `slide.shapes.add_textbox(...)` adds a box. |
| **slide master / layout** | In PowerPoint, the master holds what every slide shares (colors, fonts, the navy top bar); a layout is one slide design built on it ("Title and Content"). Edit the master once and every slide changes. |
| **placeholder** | A box a layout reserves for something: title, body, footer. `build_deck.py` reads where the layout put its body and footer boxes, so positions live only in the template. |
| **EMU** | PowerPoint's unit of length: 914,400 per inch, 12,700 per point. `Emu(x).pt` converts. |
| **matplotlib figure / axis** | A figure is the whole image; an axis is one chart panel inside it. The ARR chart is one figure with two axes. |
| **PNG / DPI** | The chart image format / dots per inch. 200 DPI keeps charts sharp on a projector. |
| **`lru_cache`** | Remembers a function's answer for the same inputs. `text_fit.py` uses it so the font file loads once, not once per word. |
| **fake client** | A stand-in object with the same method (`messages.parse`) as the real Anthropic client, used in tests so they never call the API. |
| **SHA-256 hash** | A 64-character fingerprint of a file's bytes. The same file always gives the same hash; change one byte and it's completely different. Like a checksum on a wire transfer: it proves the file is the one you think it is, whatever it's called. |
| **manifest** | `output/<company>_manifest.json`: the record of one run (input and config hashes, commit, model, cost) plus the approval. The audit trail behind a deck. |
| **JSON** | A plain-text format for nested data: `{"reviewer": "Tyler Ho"}`. Python's `json` module reads and writes it; it's how the analysis and manifest are saved. |
| **Streamlit** | A package that turns a Python script into a web page (`app.py`). You write `st.checkbox(...)`, it draws a checkbox; no HTML to write. |

---

## 4. Every file, function by function

Order: config, then the pipeline files in the order data flows (clean, metrics, analyze, Excel, the 4 deck files, main), the web page (app), the manifest and approval (provenance, approve), then the data generators, the check scripts, the model comparison, the tests and the small files.

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
- `TRIP = "trip"`, `PASS = "pass"`, `CANNOT_EVALUATE = "cannot evaluate"`: the three flag answers.
- `MISSING_INPUT`, `NO_PRIOR_PERIOD`, `NOT_MEANINGFUL`: the three reasons a value has no number. `REASON_DISPLAY` holds their words ("data missing", "n/a (no prior period)", "n/m (not meaningful)").
- `METRIC_LABELS`, `INPUT_LABELS`: **the one label set** every output uses (printout, Excel, Claude's payload).
- `METRIC_INPUTS`: for every metric, each (input column, quarters back) it uses. The reasons and `data_gaps` come from here. `METRIC_LOOKBACK` is derived from it (the furthest input back).
- `FLAG_THRESHOLDS` / `FLAG_RULES`: one line per flag: (display name, metric column, config.yaml key, "min" or "max"). The display name is taken from `METRIC_LABELS`, so Flags and Metrics use the same words. Adding a flag means adding a line to `FLAG_THRESHOLDS` plus a threshold in config.yaml.
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
| `runway_at_next_budget(actuals, next_budget)` | latest cash / (next quarter's budgeted burn / 3). NaN if there's no budget row or latest cash / budgeted burn is blank (checked first), ∞ if the budget has no burn. Shown as context, never flagged. | 14300 / 1100 = **13.0 mo** |
| `compute_metrics(actuals)` | Calls every metric function above except `runway_at_next_budget` (a single number, which callers compute separately) and puts the results in one table: a row per quarter, 19 metric columns. `pipeline` is copied in for the combo rule. | The table `python metrics.py` prints. |

**Flags and gaps.**

| Function | What it does, in plain English | Example / why it exists |
|---|---|---|
| `input_reason(actuals, metric, position)` | Looks at the metric's inputs (`METRIC_INPUTS`) for one quarter: any blank → missing input; an earlier quarter that doesn't exist → no prior period; else None. A blank wins. | A blank quarter's own YoY is "missing input", not "no prior period". |
| `metric_reasons(actuals, metrics)` | A table shaped like the metrics table: each value's reason, or None. A NaN with every input present is "not meaningful". | Gross margin with revenue 0 and gross profit 0 → not meaningful, not a gap. |
| `not_meaningful_text(actuals, metric, quarter)` | What to show instead of an n/m value: the $K figures for the two budget metrics, else "n/m (not meaningful)". | "n/m: net burn 900 vs budget 0 ($K)". |
| `validate_config(config)` | Stops if `combo_lookback_quarters` isn't a whole number ≥ 2, or `combo_min_nrr_drop` is missing or negative. | With 1 quarter there are no steps, and `all()` of nothing is True, so the combo would always trip. |
| `load_config(path)` | Reads config.yaml into a dictionary, then runs `validate_config`. | `config["nrr_min"]` → 1.0. |
| `check_threshold(value, threshold, kind)` | NaN → CANNOT_EVALUATE. Otherwise rounds to 6 decimals, then "min" trips below the threshold and "max" trips above it. **Exactly at the threshold passes.** | Rounding matters: 3900/3250 − 1 is 0.19999999999999996 in Python. Without rounding, a burn exactly 15% over budget could trip on noise. |
| `check_combo(metrics, reasons, config, quarter)` | Takes the last 3 quarters up to `quarter`. Returns **(status, reason)**. Reasons in the same order as a metric's: a blank NRR/pipeline input in the quarters it has → (cannot evaluate, missing input), even with too little history; then not enough history → no prior period; then the value's own reason. Otherwise trips only if NRR fell by at least `combo_min_nrr_drop` (1 point) at **every** step and pipeline rose at **every** step. | Northwind: NRR 108.0% → 102.0% → 97.1% (−6.0, −5.0 points) while pipeline 10,100 → 11,200 → 12,500 → trip. Missing data is never quietly treated as a pass. |
| `evaluate_flags(metrics, reasons, config, quarter=None)` | Loops over `FLAG_RULES`, runs `check_threshold` for each, then adds the combo rule. Defaults to the latest quarter but accepts any. Returns a list of dicts (flag, metric, quarter, value, threshold, status, reason). | Northwind Q2 2026: 6 trip, 3 pass. |
| `flag_status_text(flag)` | "trip", "pass" or "cannot evaluate — missing input". | |
| `data_gaps(actuals, metrics, flags)` | Every metric value whose reason is **missing input**, plus every flag that can't be evaluated because of a missing input. Returns `{metric or flag: [quarters]}`. | ARR YoY in Q4 2024 is blank but not a gap (no prior period). ARR YoY in Q1 2026 **is** a gap (it uses Q1 2025, which is blank). A partly blank quarter only makes gaps of the metrics that use the blank cell. |
| `format_value(column, value)` | Display text for a real number: $K with commas, "11.0 mo", "2.35x", or a %. ∞ → "∞". | **The only place ratios become %** (plus analyze.py and excel_output.py, which reuse the same rules). |
| `display_value(actuals, metrics, reasons, metric, quarter)` | The formatted number, or the words for its reason. | Used by the printout and Claude's payload, so they say exactly what Excel says. |
| `runway_context_label(runway, has_budget_row)` | For runway at budget: None for a normal number, else "n/a (no budget row)", "data missing" or "∞ (budget not burning)". | |
| `print_report(actuals, next_budget, config)` | Prints the metrics table, runway at budget, flags and data gaps. | What you see with `python metrics.py data/northwind.xlsx`. |

---

### `analyze.py`: Claude writes the commentary, Python checks it

**What it's for:** builds a text "payload" of already-computed facts, sends it to Claude with writing rules, validates the answer, retries once if needed, and saves the result. `main.py` calls it for every company (unless `--skip-ai`), and `build_deck.py` reuses its checks. Run on its own with `python analyze.py data/northwind.xlsx` (**calls the API**).

**Constants worth knowing:** `DEFAULT_MODEL = "claude-sonnet-5"` (the model comparison's pick), `MAX_ATTEMPTS = 2` (first try + one retry), `MAX_HEADLINE_WORDS = 30`, `MAX_DETAIL_WORDS = 45`, `MAX_DETAIL_SENTENCES = 2`. Labels (`METRIC_LABELS`, `INPUT_LABELS`) come from `metrics.py`. `SYSTEM_PROMPT` holds Claude's instructions: role, number rules, what to write, framing rules.

| Function / class | What it does, in plain English | Example / why it exists |
|---|---|---|
| `Point` (class) | The shape of one win or risk: a `title` and a `detail`. | "Retention slipping" / "NRR went from 108.0% to 97.1%…" |
| `BoardSummary` (class) | The shape of the whole answer: headline, wins, risks, questions. | Passed to the API so Claude's reply must be this JSON. |
| `input_trend(actuals, column)` | One raw input (net burn, ending cash) across all quarters, as display text. Blank → "data missing". | `{"Q3 2024": "2,400", ...}` |
| `metric_trend(actuals, metrics, reasons, column)` | One metric across all quarters as text, using `display_value`: the number, "data missing", "n/a (no prior period)" or "n/m ...". | Keeps the three reasons apart for Claude too. |
| `describe_flag(flag, config, actuals, metrics, reasons)` | One flag as text: name, status (TRIPPED / passed / cannot evaluate — reason), value, threshold. The combo rule gets a description instead of a value. | `{"flag": "NRR (annualized)", "status": "TRIPPED", "value": "97.1%", "threshold": "100.0%"}` |
| `build_payload(company, actuals, next_budget, config)` | Runs the metrics, flags and gaps, then packages every fact Claude may use as text: flag counts, flag details, runway if burn returns to plan, all trends, data gaps. | Flag counts are computed here so Claude quotes "6 of 9" instead of counting. |
| `payload_to_text(payload)` | Turns the payload into the exact JSON text sent to Claude. | The number check uses this same text, so "allowed numbers" = exactly what Claude saw. |
| `numbers_in(text)` | Finds every number in a piece of text **with its sign**, ignoring $, %, x, K and commas. A `-`, `−` or `–` counts as a minus only when it isn't right after a letter or digit. | `"-$240K and 97.1%"` → {−240.0, 97.1}; `"2025–2026"` stays {2025, 2026}. So "19.0%" doesn't pass when the data says "-19.0%". |
| `count_sentences(text)` | Counts sentences by splitting after . ! or ? followed by a space. | The decimal point in "97.1%" isn't followed by a space, so it isn't a sentence end. Known limit: "vs. " counts as one. |
| `summary_texts(summary)` | Collects every piece of text Claude wrote into one list. | So the checks can scan everything at once. |
| `find_ungrounded_numbers(summary, payload_text)` | Numbers in Claude's answer that don't appear anywhere in the payload. | "NRR fell 11.8 points" → 11.8 isn't in the data (Claude subtracted) → caught. |
| `validate_summary(summary, payload_text)` | Returns a list of problems: not exactly 3 wins/risks/questions, empty text, headline too long, a detail over 45 words or 2 sentences, numbers not in the data. An empty list = pass. | The code-enforced part of "Claude never does math". |
| `AnalysisError` (class) | The error raised when both attempts fail. Carries `run_info` (tokens, timing, problems) so failures can still be logged. | Used by `compare_models.py` to count failed runs. |
| `build_messages(payload_text, previous)` | The conversation to send. On a retry it adds Claude's last answer and the list of problems. | The retry tells Claude *what* was wrong, not just "try again". |
| `call_claude(client, model, payload_text, previous)` | One API call using structured outputs (`output_format=BoardSummary`). Handles a malformed reply, a refusal or a cut-off answer. Otherwise runs `validate_summary`. Returns the summary, its problems and a log entry. | Timing and token counts are logged for the cost comparison. |
| `make_run_info(model, attempts, passed)` | Adds up tokens and seconds across attempts. | Feeds the step 3b cost table. |
| `analyze(payload_text, model, client)` | **The loop:** up to 2 attempts. Returns the first answer with no problems; if both fail, raises `AnalysisError`. Nothing unvalidated is ever returned. | |
| `save_analysis(path, payload, summary, run_info, error)` | Writes `output/<company>_analysis.json`: the summary (or `null` if there's no validated answer), run info, the `error` (why it failed, or `null`) **and the payload**. | One function, so `analyze.py` and `main.py` write exactly the same file shape. Saving the payload next to the answer makes every claim traceable. |
| `print_commentary(summary)` | Prints headline, wins, risks, questions, with no model name or stats. | Also used for blind scoring in compare_models.py. |
| `print_summary(summary, run_info)` | `print_commentary` plus the run stats. | |
| `main()` | Command line: loads `.env`, cleans the workbook, builds the payload, calls `analyze`, and saves the result with `save_analysis`. A failure is saved too, then the script exits with code 1. | |

---

### `excel_output.py`: the metrics as a spreadsheet (build step 4b)

**What it's for:** writes `output/<company>_metrics.xlsx` with 3 sheets. **It does no math.** It calls the same clean → metrics → flags → gaps chain as everything else, then only writes and formats.

**Constants worth knowing:** `INFINITE_LABELS` (`"∞ (ARR shrank)"`, `"∞ (not burning)"`, `"∞ (never pays back)"`) are the words written for infinity; the words for the three "no number" reasons come from `metrics.reason_text`. `STATUS_COLORS` holds Excel's own "Bad" red, "Good" green, plus gray.

| Function | What it does, in plain English | Example / why it exists |
|---|---|---|
| `number_format(column)` | The Excel display format for a column: `#,##0`, `0.0" mo"`, `0.00"x"` or `0.0%`. | The cell holds 0.971; Excel shows 97.1%. The number stays usable in formulas. |
| `cell_value(actuals, metrics, reasons, column, quarter)` | The number, or a word: the reason's words ("data missing", "n/a (no prior period)", "n/m ..."); ∞ → "∞ (reason)". | Excel can't store NaN or ∞, and openpyxl would silently save both as an empty cell, so "missing" and "infinite" would look the same. |
| `runway_context_value(runway, has_budget_row)` | Same idea for runway at budget (words from `metrics.runway_context_label`): a number, "n/a (no budget row)", "data missing" (the row exists but latest cash or budgeted burn is blank) or "∞ (budget not burning)". | Added in the Task 7 review: a blank input used to say "no budget row". |
| `color_cell(cell, status)` | Fills a cell red/green/gray with matching text color. | |
| `write_header(sheet, headers)` | Bold header row, freezes row 1 and column A. | |
| `set_column_widths(sheet, widths)` | Sets column widths left to right. | |
| `tripped_cells(metrics, reasons, config)` | Runs `evaluate_flags` for **every** quarter and collects each (quarter, metric) that trips. | So a Q1 2026 burn-vs-budget breach is red too, not just the latest quarter. |
| `status_label(flag)` | "Tripped", "Passed" or "Cannot evaluate — <reason>". | |
| `style_metric_cell(cell, column, is_gap, is_tripped)` | Format, right alignment, gray **only** if an input is missing, else red if tripped. | Passed, "no prior period" and "n/m" cells stay uncolored so gray means "chase this data". |
| `write_metrics_sheet(sheet, actuals, metrics, reasons, config)` | One row per quarter, one column per metric, with values, formats and highlights. | |
| `flag_row(flag, actuals, metrics, reasons, config)` | One flag as a row: Flag, Quarter, Value, Threshold, Trips when, Status. The combo rule gets text instead of a value. | "Trips when" says "below threshold" or "above threshold", so −20.0% makes sense. |
| `style_flag_row(sheet, row_number, flag)` | Colors the whole row by status and formats Value and Threshold. | |
| `write_runway_context(sheet, runway, has_budget_row, quarter)` | Adds "Runway at next quarter's budgeted burn (context, not a flag)" below the table, uncolored. | CLAUDE.md: shown as context, not flagged. |
| `write_flags_sheet(sheet, flags, gaps, config, runway_at_budget, has_budget_row)` | Writes every flag row plus the runway context line. | |
| `gap_label(name)` | `"nrr"` → `"NRR (annualized)"`; `"flag: Rule of 40"` → `"Flag: Rule of 40"`. | |
| `write_gaps_sheet(sheet, gaps)` | One row per affected metric or flag, or "None — every metric and flag has the data it needs". | |
| `build_workbook(actuals, next_budget, config)` | Computes metrics, flags and gaps, then writes all 3 sheets into a new workbook. | |
| `output_path(workbook_path, output_dir)` | `data/northwind.xlsx` → `output/northwind_metrics.xlsx`. | The name comes from the input file, so `--all` gets unique names for free. |
| `save_metrics_workbook(workbook_path, config, output_dir)` | Cleans the input, builds the workbook, saves it, returns the path. | **The one function `main.py` calls.** It cleans the workbook itself, and so does `build_deck.save_deck`, so in a `main.py` run each file is read three times. Harmless at this size. |

---

### `make_template.py`: the brand template (build step 4)

**What it's for:** builds `templates/base.pptx`, the PowerPoint template every deck starts from. You run it once (`python make_template.py`); the file is committed, so decks build without running it. Like a firm's PowerPoint template: colors, fonts and footer are set once, not on every slide.

**Why it's built by code:** python-pptx can't create a template from nothing. So it starts from python-pptx's built-in default (4:3, 11 Office layouts), resizes it to 16:9, restyles it and deletes 9 layouts. Some steps edit the file's XML directly, because python-pptx has no function for them.

**Constants worth knowing:**
- Brand: `BRAND_NAME = "Example Capital"` (fictional), `NAVY = "1F2A44"`, `DARK_GRAY`, `MID_GRAY`, `LIGHT_GRAY`, `WHITE`, `FONT = "Arial"`. `build_deck.py` and `charts.py` import these, so the colors live in one place.
- Layout names: `TITLE_LAYOUT = "Title Slide"`, `CONTENT_LAYOUT = "Title and Content"`.
- Geometry, as `(left, top, width, height)`: `TITLE_BOX`, `BODY_BOX` (the content area, ends at 6.75 in), `FOOTER_BOX`, `BRAND_BOX`, `TOP_BAR`, `FOOTER_RULE` (the hairline; `check_deck.py` checks nothing runs past it).

| Function | What it does, in plain English | Example / why it exists |
|---|---|---|
| `set_theme(presentation)` | Rewrites the theme's colors (navy, grays) and both fonts (headings and body) to Arial, in the theme's XML. | Everything that uses "theme colors" turns navy and gray at once. |
| `set_text_style(level_element, size_pt, hex_color, bold, align)` | Sets size, color and bold on one text level of the master. | PowerPoint stores 28 pt as `sz="2800"` (hundredths of a point). |
| `set_master_text_styles(master)` | Titles: navy, bold, 28 pt, left-aligned. Body: dark gray, 20 then 18 pt. | |
| `keep_only_two_layouts(presentation)` | Deletes every layout except Title Slide and Title and Content. | Fewer choices, so a deck can't pick the wrong layout. |
| `remove_unused_placeholders(shapes)` | Removes the date and slide-number boxes. | The run date is in the footer text instead. |
| `place(shape, box)` | Moves and resizes a shape to a box. | |
| `box_for(placeholder_type, cover)` | Which box a placeholder goes to: title, subtitle, body or footer (None = leave it). | The cover slide's title sits lower than a content slide's. |
| `position_placeholders(shapes, cover)` | Moves every placeholder on a master or layout to its 16:9 position. | |
| `set_placeholder_style(placeholder, size_pt, hex_color, bold)` | Gives one placeholder its own text style. | Used for the white cover title. |
| `style_cover_layout(layout)` | Title Slide: navy background, white title, hides the master's bar and footer rule. | The 4-slide deck doesn't use it; it's there for a cover page. |
| `add_master_shape(master, name, box, textbox)` | Adds a rectangle or text box to the slide master, behind everything, by writing its XML. | python-pptx can only add shapes to slides, not to a master. |
| `fill_solid(shape, hex_color)` | Solid fill, no outline. | |
| `add_brand_name(master)` | "Example Capital", navy, bold, 12 pt, bottom right. | |
| `decorate_master(master)` | Adds the navy top bar, the gray hairline above the footer and the brand name. | Every content slide gets them without build_deck.py drawing them. |
| `build_template()` | Runs all of the above in order and returns the template in memory, with no slides. | |
| `save_template(path)` | Builds and saves `templates/base.pptx`. | |

---

### `text_fit.py`: does the text fit its box?

**What it's for:** PowerPoint doesn't tell python-pptx how much room text takes. So this file estimates it, shrinks the text until it fits, and **stops with an error** if it can't fit at 12 pt. A board deck with text running off the slide is worse than no deck.

**How the estimate works:**
- **Width:** measured with a real font file, DejaVu Sans, which ships with matplotlib. DejaVu is wider than Arial, the deck's font, so the estimate errs toward "needs more room".
- **Wrapping:** words go onto a line until the next one doesn't fit, like PowerPoint.
- **Height:** lines × font size × 1.2 (`LINE_SPACING`), plus the space after each paragraph.

**Constants worth knowing:** `MIN_FONT_PT = 12` (the floor), `LINE_SPACING = 1.2`, `MEASURE_SIZE = 100` (fonts load once at 100 pt and are scaled).

| Function / class | What it does, in plain English | Example / why it exists |
|---|---|---|
| `TextDoesNotFitError` (class) | The error raised when text doesn't fit even at 12 pt. Its message names the slide and box. | "Slide 3, Risks and flags: text doesn't fit even at the 12 pt minimum: '...'" |
| `measuring_font(bold)` | Loads DejaVu Sans (regular or bold) once and remembers it (`lru_cache`). | Loading a font file for every word would be slow. |
| `text_width_pt(text, size_pt, bold)` | How wide a piece of text is, in points. | Width at 100 pt × size / 100: width grows in step with size. |
| `count_lines(text, width_pt, size_pt, bold)` | How many lines the text wraps to in a box that wide. A single word wider than the box is split across lines. | |
| `paragraph(text, size, bold, color, space_after)` | Packs one paragraph's settings into a dict. | build_deck.py builds every text box from a list of these. |
| `text_height_pt(paragraphs, width_pt)` | Total height of a list of paragraphs. | |
| `preview(text)` | The first 60 characters, for error messages. | |
| `shrink_to_fit(paragraphs, width_pt, height_pt, where)` | Lowers **every** size by the same step (1 pt at a time) until the text fits. Stops with `TextDoesNotFitError` once the smallest size would go below 12 pt. | A heading stays 4 pt bigger than its body text as both shrink. Known gap (DAY_REPORT Review, finding 3): text that *starts* below 12 pt isn't rejected. |
| `row_heights_pt(rows, column_widths_pt, size_pt, cell_padding_pt)` | Each table row's height: its tallest cell (header row in bold) plus padding. | |
| `fit_table(rows, column_widths_pt, height_pt, where, start_size, cell_padding_pt)` | The biggest font size (from 14 down to 12) at which the whole table fits, plus the row heights. Stops if it can't. | One size for the whole table, so rows don't look mismatched. |

---

### `charts.py`: the two charts on slide 2

**What it's for:** draws the ARR chart and the cash chart with matplotlib and saves them as PNG images. **It does no math**: values arrive computed, and every label goes through `metrics.format_value`.

**Two choices worth explaining:**
- **Two panels, not two y-axes.** Net new ARR is small next to ARR and can go negative, so it gets its own panel and zero line. A chart with two y-axes lets the reader compare heights that aren't comparable.
- **A blank quarter stays visible.** Bars: no bar, and "data missing" written where it would be. Line: the NaN stays in the data, so matplotlib stops the line at the gap instead of joining across it (joining would draw numbers that don't exist).

**Constants worth knowing:** `FONT_SIZE = 12` (same floor as the slides), `DPI = 200`, `HEADROOM = 1.2` (cash axis top = 1.2 × highest cash), `GAP_LABEL` ("data missing", from metrics.py).

| Function | What it does, in plain English | Example / why it exists |
|---|---|---|
| `hex_color(value)` | `"1F2A44"` → `"#1F2A44"`. | PowerPoint colors have no #, matplotlib's need one. |
| `style_axis(axis, quarters, money_column)` | Quiet styling: no top/right border, light grid, "Q2\n2026" labels, y-axis in $K. | |
| `label_gaps(axis, values)` | Writes "data missing" at every blank quarter. | |
| `label_latest(axis, values, column, below)` | Writes the latest value next to its bar or point (below a negative bar). | Only the latest value is labelled, so the chart stays readable. |
| `bar_panel(axis, quarters, values, column)` | Bars for every quarter that has a value, a zero line, the title and labels. | |
| `arr_chart(quarters, ending_arr, net_new_arr, size_inches)` | Figure with two panels: ending ARR (taller) above net new ARR. | |
| `cash_chart(quarters, ending_cash, runway_text, size_inches)` | Ending cash as a line, zero kept on the axis, runway at current and at budgeted burn in the title. | Cash running out means reaching zero, so the axis always shows zero. |
| `save_chart(figure, path)` | Saves the PNG at its own size and frees the memory. | Drawn at the size it has on the slide, so 12 pt in the chart is 12 pt on screen. |

---

### `build_deck.py`: the 4-slide board deck (build step 4)

**What it's for:** `output/<company>_board_pack.pptx`. Run on its own with `python build_deck.py data/northwind.xlsx` (no API call: it uses the saved `output/northwind_analysis.json` if there is one). Add `--no-analysis` for the placeholder, `--analysis PATH` for another file, or `--draft` for the DRAFT - NOT REVIEWED watermark on a deck nobody has approved.

**Rules it follows:**
- **No math and no typed numbers.** Every number comes from metrics.py, config.yaml or the validated analysis. A test reads the code and fails if any text in it contains a digit.
- **Claude's text is checked again before it's used** (`load_analysis`). If it fails, slide 4 says "AI summary unavailable", and the other slides are built as normal.
- **Text must fit** (text_fit.py), and every slide gets a one-line footer: "Fictional data | northwind.xlsx | 2026-09-17 | 9c1b52c | claude-sonnet-5 | AI-drafted | not reviewed". That's the fictional-data note, the source file, the run date, the git commit (with `*` if the code had uncommitted edits), the model ("no AI text" on a placeholder deck) and the review status.
- **The review status comes from the manifest**, judged by `provenance.approval_status` (the only place that decides): "AI-drafted | reviewed by Tyler Ho on 2026-09-17" once a still-valid approval is recorded, otherwise "AI-drafted | not reviewed".
- **No watermark unless asked.** With `--draft` (here or in main.py), every slide of an unreviewed deck also gets a see-through "DRAFT - NOT REVIEWED" drawn on top. `--draft` never stamps an approved deck, because the stamp would be false.
- **The old deck is deleted first**, so a failed build never leaves last run's deck looking current.

**Constants worth knowing:** `PLACEHOLDER_TEXT = "AI summary unavailable"`, `PLACEHOLDER_NOTE` (says the numbers are unaffected), font sizes (`TITLE_SIZE` 28, `HEADLINE_SIZE` 22, `BODY_SIZE` 14, `LIST_SIZE` 16, `TABLE_SIZE` 14), `KPI_COLUMN_SHARES` (table column widths), `CONTEXT_ROWS` (the 3 non-flag rows on slide 1), `SLIDE_BUILDERS` (the 4 slide functions in order), `AI_DRAFTED_LINE` (the line under slide 4's title).

**The call order:** `save_deck` → `clean_workbook` → `collect_deck_data` → `load_analysis` → `approval_status` (from the manifest) → `build_presentation` → for each slide: `new_slide`, that slide's function, `add_footer`, and `add_watermark` only with `--draft` on an unapproved deck → save → `record_deck_status`.

**1. Data and text pieces**

| Function | What it does, in plain English | Example / why it exists |
|---|---|---|
| `collect_deck_data(company, source_name, actuals, next_budget, config)` | Computes metrics, reasons, flags, data gaps and runway at budget once, and returns them in one dict. | Every slide reads from this dict, so no slide computes anything. |
| `value_text(data, metric, quarter)` | One value as slide text: the number, "∞ (ARR shrank)", or the reason words ("data missing", "n/a (no prior period)", "n/m ..."). | Same words as the Excel file. |
| `threshold_text(flag)` | "trips below 100.0%" or "trips above 2.00x". | Says which way a flag trips, so −20.0% makes sense. |
| `flag_count_text(flags)` | "6 of 9 flags tripped", plus ", 1 cannot evaluate" when needed. | Counted by Python, never by Claude. |
| `gaps_lines(gaps)` / `gaps_text(gaps)` | The Data gaps line, grouped by the quarters each metric misses: "Q1 2025 + Q2 2025: ARR growth QoQ, ...". | One bullet per group keeps 19 gaps readable. |
| `points_text(ratio)` | 0.01 → "1.0 pts". | For the combo rule's minimum drop. |
| `combo_text(flag, config)` | The combo result plus its rule, with the settings from config.yaml. | "…trips when NRR falls by at least 1.0 pts and pipeline rises at every step over the last 3 quarters". |
| `runway_lines(data)` | Runway at current burn (the flag) and at next quarter's budgeted burn (context), for the cash chart title. | |

**2. The AI analysis**

| Function | What it does, in plain English | Example / why it exists |
|---|---|---|
| `load_analysis(path, payload)` | Returns (summary, None) only if **all** of these hold: the file exists and is JSON; it has a summary (it passed when made); it's for this company **and** latest quarter; it has the `BoardSummary` shape; and `validate_summary` passes against a payload rebuilt **from today's workbook**. Otherwise (None, why not). | A last-quarter file, or one where someone edited "11.0 mo" to "11.5 mo", never reaches a board. |

**3. Slide helpers**

| Function | What it does, in plain English | Example / why it exists |
|---|---|---|
| `layout_box(layout, placeholder_types)` | Where the template's layout put a placeholder: (left, top, width, height). | Positions live in make_template.py only. |
| `points(length)` | EMU → points. | |
| `write_paragraphs(frame, paragraphs)` | Writes sized paragraphs into a text box, with PowerPoint's autofit switched off. | Autofit off means the saved sizes are the sizes shown, so the fit check means something. |
| `fitted(paragraphs, width, height, where)` | `shrink_to_fit` for a box, minus the box's inner margins. | |
| `add_text_box(slide, name, box, paragraphs, deck)` | Adds a named text box, shrunk to fit (or stops). | Names like "Headline" let check_deck.py find each box. |
| `add_columns(slide, columns, top, height, deck)` | Two boxes side by side. Fits each, then shrinks both by the bigger shrink. | Risks and Questions never show two different font sizes. |
| `set_title(slide, text, deck)` | Writes the slide title, shrunk to fit. | |
| `new_slide(presentation, layout)` | Adds a slide and removes its empty body placeholder. | Each slide places its own boxes in that area. |
| `commit_text()` | The git commit for the footer: "9c1b52c", or "9c1b52c*" when the code had uncommitted edits. | A printed slide names the code that built it. |
| `review_text(approval, with_name)` | "reviewed by Tyler Ho on 2026-09-17", "reviewed on 2026-09-17" (no name), or "not reviewed". | The manifest keeps the time to the second; the footer only has room for the day. |
| `footer_text(deck, run_date, with_name, source_name)` | Joins the footer's parts with " \| ". `source_name` swaps in a shortened file name. | |
| `fits_one_line(text, box)` | True if the footer text, at 12 pt, is no wider than its box. | A reviewer's name can be any length, so its width is measured, not assumed. |
| `shorten_middle(file_name, fits)` / `footer_that_fits(deck, run_date, with_name)` | A file name too long for the footer keeps its start and extension: "Northwin….xlsx". | The web page takes any file name; a long one used to stop the whole deck (LEARNINGS, polish Task 6). |
| `add_footer(slide, deck, run_date)` | The footer on every slide. A long file name is shortened first; if a long reviewer name still doesn't fit, it is left out ("reviewed on DATE") and stays in the manifest. | Anything else that doesn't fit still stops the build. |
| `set_alpha(run, percent)` / `add_watermark(slide, deck)` | Only with `--draft`: "DRAFT - NOT REVIEWED" diagonally across the slide, 25% opaque, on top of everything. | On top, not behind: the table and charts are opaque and would hide it. See-through, so the numbers stay readable. |

**4. The four slides**

| Function | What it does, in plain English | Example / why it exists |
|---|---|---|
| `kpi_rows(data)` | **Slide 1's rows:** Ending ARR (with "vs budget: 2.5%"), ARR growth YoY, gross margin, then all 9 flags in order. The combo row shows "—" and "rule on Risks and flags slide". | All 9 flags, so the statuses add up to slide 3's "of 9". |
| `kpi_header(data)` | Metric, Q2 2026, Q1 2026, Budget or threshold, Status. | |
| `column_widths(total_width)` | Splits the table width by `KPI_COLUMN_SHARES`. | |
| `write_cell(cell, text, size, fill_hex, text_hex, bold)` | One table cell: fill, margins, text. | |
| `fill_table(table, header, rows, size)` | Navy header, striped rows, and each status cell red / green / gray. | Same colors as the Excel file (`excel_output.STATUS_COLORS`). |
| `kpi_slide(slide, deck)` | **Slide 1:** title "Northwind: key metrics — Q2 2026 vs Q1 2026"; fits the table (`fit_table`), then draws it. | The company name is in this title: the slide that used to carry it is gone. |
| `charts_slide(slide, deck)` | **Slide 2:** draws both charts at their slide size, saves the PNGs to `output/charts/`, places them side by side. | |
| `section(heading, lines)` | A bold heading and its bullet lines. | |
| `flags_paragraphs(data)` | **Slide 3, left:** "Tripped flags (6 of 9 flags tripped)", each with value and threshold, a "Cannot evaluate" section if any, and the combo rule. | Fernhollow's Rule of 40 isn't silently absent. |
| `risks_slide(slide, deck)` | **Slide 3:** flags on the left, Data gaps on the right. | Two columns: Fernhollow's text didn't fit in one (LEARNINGS). |
| `commentary_boxes(area)` | **Slide 4's boxes**, top to bottom: the AI-drafted line, the headline, the two columns. | `ai_text_problems` measures these same boxes, so the check can't drift from the slide. |
| `headline_paragraph(text, from_ai)` | The headline: navy for Claude's, gray for the placeholder. | |
| `column_heading(text)` / `points_paragraphs(heading, items)` / `questions_paragraphs(questions)` | A column's bold heading; each risk's title (bold) and detail; the numbered questions. | |
| `commentary_columns(summary)` | Slide 4's two columns: Risks, and Questions for management. | The wins are left out here, so they are neither drawn nor measured. |
| `commentary_slide(slide, deck)` | **Slide 4:** "AI-drafted from computed metrics - review before use", the headline, then risks and questions side by side. Or the gray placeholder and note, without the AI-drafted line. | |

**5. Putting it together**

| Function | What it does, in plain English | Example / why it exists |
|---|---|---|
| `build_presentation(data, summary, run_date, chart_dir, approval, model, draft)` | Opens the template, runs the 4 slide functions in order, adds footers, and the watermark only with `--draft` (`draft=True`) and nobody approved. Sets `deck["where"] = "Slide 3"` first, so a fit error names the slide. | |
| `slide_number(build_slide)` | Which slide a function builds, counting from 1. | Messages say "slide 4" without the number being typed anywhere. |
| `ai_text_problems(summary)` | Problems if the headline, risks or questions don't fit slide 4's boxes even at 12 pt. | analyze.py calls it, so an over-long answer gets the retry. |
| `deck_path(workbook_path, output_dir)` | `data/northwind.xlsx` → `output/northwind_board_pack.pptx`. | |
| `analysis_path(workbook_path, output_dir)` | `data/northwind.xlsx` → `output/northwind_analysis.json`. | main.py and check scripts use it too, so the name is set once. |
| `analysis_details(path)` | The model and prompt version a saved analysis was made with. | The footer names the model whose words are on slide 4. |
| `record_deck_status(workbook_path, output_dir, approval, ai_text)` | After a rebuild, updates the manifest's `deck` part (reviewed or not, AI text or not), if there is a manifest. | So the manifest never says "approved" beside a deck whose footer says "not reviewed". |
| `save_deck(workbook_path, config, analysis_file, run_date, output_dir, draft)` | **The one function `main.py` and `app.py` call.** Deletes the old deck, cleans the workbook, collects the data, loads the analysis (or none), reads the approval from the manifest, builds and saves. Returns (deck path, why the AI text isn't on it, or None). | |
| `main(argv, output_dir)` | Command line: `--analysis`, `--no-analysis`, `--draft`. Prints where the deck was saved, and why the AI text is missing if it is. | |

---

### `memo.py`: the board memo, Word and PDF (final Task 1)

**What it's for:** the same update as the deck, written as a 1 to 2 page memo for board members who read rather than present. `python memo.py data/northwind.xlsx` saves `output/northwind_board_memo.docx` and `.pdf` beside the deck; `main.py` does it for every company.

**What's in it, top to bottom:** title and quarter; the AI headline under "AI-drafted from computed metrics - review before use"; the key metrics table (latest, prior, budget or threshold, status in red / green / gray); runway at next quarter's budgeted burn; "Flags: 6 of 9 flags tripped" with each tripped flag's value and threshold, the flags that can't be evaluated, and the combo rule; data gaps; the AI's 3 questions for management. The footer is the deck's footer, on every page.

**How it's built:** the memo is built once as a list of "blocks" (a title, a heading, a paragraph, a bullet list, a table), then written twice: `write_docx` (python-docx) and `write_pdf` (reportlab). So the Word file and the PDF can't say different things, and `check_memo.py` proves they don't.

**The one rule the deck doesn't have:** every number in the memo must be one the metrics workbook shows. Claude may quote any number in its payload, which includes raw inputs (net burn, ending cash in $K) that the metrics workbook doesn't have. So the memo checks the AI headline and questions against the workbook's numbers too, and says "AI commentary unavailable" if one is missing. An analysis can therefore be on the deck and not in the memo.

**No em dashes:** two labels it shares with the deck and Excel have one (the "Cannot evaluate" status and the "None" data gaps line); the memo shows a colon instead, e.g. "Cannot evaluate: missing input".

| Function | What it does, in plain English | Example / why it exists |
|---|---|---|
| `workbook_texts(data)` | Every text the metrics workbook shows: each metric in each quarter, the quarter labels, flag names and thresholds (as Excel displays them), the combo rule's wording, runway at budget, and the flag count. | The list the AI's numbers are checked against. |
| `workbook_numbers(data)` | The numbers in those texts, read the way analyze.py reads numbers (sign kept, %, x and mo dropped). | Thresholds count as Excel shows them: 15.0, not 0.15 (that bug is in LEARNINGS). |
| `unlisted_numbers(summary, data)` | Numbers in the AI headline and questions that aren't in `workbook_numbers`. | The wins and risks aren't in the memo, so they aren't checked here. |
| `memo_analysis(analysis_file, payload, data)` | First every check the deck makes (`build_deck.load_analysis`), then `unlisted_numbers`. Returns (summary, None) or (None, why not). | Northwind with "ending cash of $14,300K" in a question: fine for the deck, "AI commentary unavailable" in the memo. |
| `no_em_dash(text)` | Swaps a spaced em dash for a colon: the deck's "Cannot evaluate" status becomes "Cannot evaluate: missing input". | The memo has no em dashes. |
| `flag_cells(data, flag)` | One flag's row on the Flags sheet (`excel_output.flag_row`). | The combo rule's words come from here, so its "1 pt" and "3 quarters" are the workbook's own. |
| `runway_context_text(data)` | Runway at next quarter's budgeted burn: "13.0 mo", or why there's no number. | |
| `combo_line(data, flag)` | "NRR falling while pipeline rising: Tripped (trips when NRR falls at least 1 pt and pipeline rises at every step, over the last 3 quarters)". | |
| `kpi_rows(data)` | The key metrics table: Ending ARR, Net new ARR, ARR growth YoY, Gross margin, then every flag with latest, prior, threshold and status. | The same text helpers as slide 1 (`value_text`, `threshold_text`, `status_label`). |
| `flag_lines(data)` | Tripped flags with value and threshold, then "cannot evaluate" flags, then the combo rule. | "Runway at current burn: 11.0 mo (trips below 12.0 mo)". |
| `heading(text)` / `text(words, style, ai)` / `bullets(items, ai)` | Make one block. `ai=True` marks the AI's two slots (Claude's words, or "AI commentary unavailable" in their place). | A test proves every other block is the same with or without the AI text. |
| `ai_blocks(summary, part)` | The headline or the questions under the AI-drafted line, or "AI commentary unavailable" and a note. | |
| `memo_blocks(data, summary)` | Every block of the memo, in order. | |
| `block_texts(blocks)` | Every piece of text in the blocks. | Tests use it to prove both files hold all of it. |
| `memo_footer(data, run_date, model, approval, commit)` | "Fictional data \| northwind.xlsx \| 2026-09-17 \| 2a215a9 \| claude-sonnet-5 \| AI-drafted \| not reviewed". | The same parts as the deck's footer, and the same review status from the manifest. |
| `docx_run` / `docx_paragraph` / `keep_with_next` / `shade_cell` / `docx_cell` / `docx_table` / `docx_block` | Write text, a paragraph, a colored table cell, the table and each block into the Word file. `keep_with_next` stops a heading ending a page. | python-docx has no setting for a cell's color, so `shade_cell` writes the XML itself. |
| `write_docx(blocks, footer, path)` | Saves the Word file: US Letter, 0.7 inch margins, the footer on every page. | |
| `register_pdf_fonts()` | Tells reportlab where DejaVu Sans is. | The PDF's built-in fonts have no "∞", and Fernhollow's burn multiple is "∞ (ARR shrank)". |
| `pdf_color` / `pdf_style` / `pdf_paragraph` / `pdf_table` / `pdf_flowables` | The same blocks in reportlab's terms. `pdf_paragraph` escapes &, < and >, which reportlab would read as markup. | |
| `pdf_sections(blocks)` | Groups the PDF into sections (a heading and what follows), each kept on one page when it fits. | Northwind's first memo left "Questions for management" alone at the foot of page 1. |
| `write_pdf(blocks, footer, path)` | Saves the PDF with the footer drawn on every page; a long footer wraps instead of running off the page. | |
| `memo_paths(workbook_path, output_dir)` | `data/northwind.xlsx` → `output/northwind_board_memo.docx` and `.pdf`. | |
| `save_memo(workbook_path, config, analysis_file, run_date, output_dir)` | **The one function `main.py` calls.** Deletes the old memo files, cleans the workbook, collects the data, checks the analysis, reads the approval, writes both files. Returns the two paths and why the AI text isn't in it, or None. | |
| `memo_approval(approval)` | The approval if it lists "memo" in its documents (approve.py), else None. | Northwind was approved on 2026-09-17, before memos existed: its deck says "reviewed by", its memo says "not reviewed". |
| `record_memo_status(workbook_path, output_dir, files, ai_text)` | After a rebuild, updates the manifest's `memo` part, if there is a manifest. | Same rule as `record_deck_status`. |
| `main(argv, output_dir)` | Command line: `--analysis`, `--no-analysis`. | |

---

### `main.py`: the batch runner (build step 5)

**What it's for:** `python main.py data/northwind.xlsx` runs one company; `python main.py --all` runs every workbook in `data/`; add `--skip-ai` to make no API call, and `--draft` to watermark every deck nobody has approved. One failing company never stops the batch.

**What happens to each company, and what the Result column says:**

| What happened | Deck slide 4 | Result | `output/<company>_analysis.json` |
|---|---|---|---|
| Claude's answer passed validation | Claude's text | `OK` | summary, run info, payload |
| It failed validation twice, **or** the API call failed (connection, rate limit, server) | "AI summary unavailable" | `OK (AI failed)` | summary `null`, the reason in `error` |
| `--skip-ai` | "AI summary unavailable" | `OK (AI skipped)` | not written; an old one is left alone |
| Any other error (bad workbook, or a bug) | no deck | `FAILED: ...` | the old one was already deleted |

**Why "OK (AI failed)" is still OK (CLAUDE.md decision K):** every number on the deck comes from Python, so the deck is still worth building. The exit code stays 0; a warning line under the summary names the companies.

| Function / class | What it does, in plain English | Example / why it exists |
|---|---|---|
| `find_workbooks(data_dir)` | Every `.xlsx` in `data/`, sorted by name, skipping `~$` files. | `~$northwind.xlsx` is the lock file Excel creates while a workbook is open; it would show up as a failing company. |
| `company_name(workbook_path)` | `data/northwind.xlsx` → "Northwind". | Same rule as analyze.py and build_deck.py, so the analysis matches its deck. |
| `shown_path(path)` | Prints a path as `output/...` when it's inside the project, else in full. | Tests save into a temporary folder outside the project. |
| `api_key_problem()` | Loads `.env`, then returns None if `ANTHROPIC_API_KEY` is set, else "ANTHROPIC_API_KEY isn't set: add it to .env, or run with --skip-ai". | The SDK reports a missing key as a plain `TypeError`, which would look like a bug in every company (LEARNINGS). |
| `ai_step(workbook_path, actuals, next_budget, config, output_dir, client)` | Deletes the old analysis JSON, builds the payload, calls `analyze.analyze`, and saves the result **whether it passed or failed** (`save_analysis`). Returns the JSON path if it passed, else None. Catches only `AnalysisError` (failed twice) and `anthropic.AnthropicError` (API failed). | Any other error is a bug and fails the company, so a coding mistake can't hide behind "AI failed". `client` is for tests (a fake client). |
| `deck_step(workbook_path, config, analysis_file, output_dir, draft)` | Calls `build_deck.save_deck` (passing `--draft` on) and prints whether the AI text made it onto the deck. Returns why not, or None. | |
| `ai_record(ai, analysis_file)` | What the manifest says about the AI step: model, prompt version, attempts, tokens, seconds, cost, and passed / failed / skipped. | Nothing is set to 0 when the AI didn't run: a 0 would read like a real cost. |
| `memo_step(workbook_path, config, analysis_file, output_dir)` | Calls `memo.save_memo` with the same analysis as the deck and prints where the memo went, and why the AI text isn't in it if it isn't. | The memo can turn down AI text the deck accepted (see `memo.py`), so it gets its own line. |
| `manifest_step(workbook_path, output_dir, ai, analysis_file, why_unavailable, memo)` | Writes `output/<company>_manifest.json` (`provenance.build_manifest`), carrying over any approval already recorded, plus a `memo` part: its two files and whether it has the AI text. | Whether that approval still counts is decided by `provenance.approval_status`, never assumed. |
| `ai_status(skip_ai, why_unavailable)` | "skipped" with `--skip-ai`; otherwise "ok" only if the AI text is really on the deck, else "failed". | "OK" means the text is on the slide, not just that Claude answered. |
| `blank_quarters(actuals)` | Quarters with at least one blank input. | Northwind → `["Q1 2025"]`. |
| `run_company(workbook_path, config, skip_ai, client, output_dir, draft)` | Clean → metrics → flags → gaps → Excel → AI step (unless `--skip-ai`) → deck → memo → manifest, printing a ✓ line for each. Returns a result dict for the summary table. | |
| `describe_error(error)` | Prints `✗ FAILED: <type>: <message>`. Bad-input errors (`ValueError`, `OSError`) get one line; anything else also gets a full traceback, because it's probably a bug. | A person fixing a workbook doesn't need a traceback; a developer fixing a bug does. Known gap (DAY_REPORT Review, finding 2): a text-fit stop also gets a traceback. |
| `run_batch(workbook_paths, config, skip_ai, client, output_dir, draft)` | Loops over the workbooks with `try`/`except` around each company, records the error and moves on. | The key reliability feature: company 2 breaking doesn't stop company 3. |
| `flags_text(result)` | "6 of 9", or "7 of 9, 1 cannot evaluate". | Without the second part, Fernhollow's "7 of 9" would hide a flag that had no answer. |
| `gaps_text(result)` | "none", or "19 metrics/flags (blank: Q1 2025)". | |
| `result_text(result)` | "OK", "OK (AI skipped)", "OK (AI failed)" or "FAILED: …". | |
| `ai_failed_warning(results)` | "⚠ AI summary unavailable for 1 company (Fernhollow): ..." or None. | Makes an AI failure visible without failing the run. |
| `summary_rows(results)` | One row of text per company; a failed company shows "-" for flags and gaps. | |
| `print_summary(results)` | Prints the table with padded columns, then "3 of 3 companies succeeded". | |
| `csv_row(result)` / `write_summary_csv(results, path)` | Saves the same summary as `output/batch_summary.csv`, with counts as plain numbers so a spreadsheet can sort them. | |
| `quarter_mismatch_warning(results)` | "⚠ Companies end on different quarters (...)" when successful companies' latest quarters differ; None otherwise. Never fails the batch. | Comparing Q1 and Q2 numbers side by side needs care. |
| `parse_args(argv)` | Reads the command line. Exactly one of a file path or `--all`; both or neither → usage error (exit code 2). | |
| `main(argv)` | Parses the arguments, finds the files, checks the API key (unless `--skip-ai`; a missing key exits 1 before any company runs), runs the batch, prints the summary and both warnings, saves the CSV, returns 0 if no company FAILED, else 1. | Exit codes let a scheduler tell from the code alone whether a run worked. |

---

### `app.py`: the web page for non-technical users (Streamlit)

**What it's for:** the same steps as `main.py` for one workbook, but from a web page: drag in an .xlsx, see the flags and metrics, download the deck and the metrics workbook. Start it by double-clicking `run_app.command` (Mac) or with `streamlit run app.py`.

**How Streamlit works, in one paragraph:** Streamlit runs `app.py` from top to bottom every time anything on the page changes (a file dropped in, a box ticked, a button clicked). `st.title`, `st.checkbox`, `st.dataframe` and so on each draw one thing on the page. So the page must not redo the expensive work on every rerun: `st.session_state` (a dictionary that survives reruns) remembers the result for this file and this checkbox choice.

**What doesn't change:** no new math, no new wording. The page shows the same text the deck shows (`build_deck.value_text`, `threshold_text`, `flag_count_text`, `gaps_lines`) in the same colors as the Excel workbook (`excel_output.STATUS_COLORS`, `tripped_cells`).

**What's different from `main.py`:** it builds in a temporary folder and never writes to `output/`, so it can't overwrite a command-line deck, manifest or approval. It writes no manifest, so its deck's footer always says "AI-drafted | not reviewed", and it never passes `--draft`, so there's no watermark. Approving is a command-line step (`approve.py`) on the `output/` decks.

| Function | What it does, in plain English | Example / why it exists |
|---|---|---|
| `ai_checkbox_label()` | The checkbox text, with the typical cost. | "Include AI commentary (typically about $0.09 and 70 seconds per workbook)". The number is `TYPICAL_AI_COST_USD`, copied from README's Cost table: a label, not a calculation. |
| `save_upload(file_name, data, folder)` | Writes the uploaded bytes to a file, keeping only the file's own name. | `clean_workbook` reads a file path, not bytes. "../x.xlsx" becomes "x.xlsx", so an odd name can't write outside the folder. |
| `is_excel_workbook(path)` | True if the file is a zip with `xl/workbook.xml` inside. | A .pptx or .docx renamed .xlsx is a zip too; checking only "is it a zip" let one through to a cryptic pandas error (LEARNINGS, polish Task 6). |
| `status_css(status)` | A status's Excel colors as a style for the on-screen table. | trip → "background-color: #FFC7CE; color: #9C0006" (light red). |
| `metrics_table(data)` / `metrics_colors(data)` | The metrics as text, one row per metric and one column per quarter; and a same-shaped table of styles: gray = data missing, red = the flag tripped that quarter, else none. | Same rules as the Excel Metrics sheet. Rows are metrics (not quarters) so 8 quarters fit across a screen. |
| `flag_row(data, flag)` / `flags_table(data)` / `flags_colors(data)` | The latest quarter's flags as Flag, Value, Threshold, Status; each whole row in its status color. `combo_rule_text(config)` describes the combo rule. | Same as the Excel Flags sheet. |
| `reusable_analysis(workbook_path, saved_dir, config)` | The saved `output/<company>_analysis.json` if Claude saw **exactly** today's facts (the whole payload matches) and it still passes the deck's checks; else None. | Ticking the AI box for a workbook `main.py` already ran costs nothing. An edited workbook never gets an old analysis. |
| `ai_commentary(...)` | Box not ticked → no analysis. Ticked → reuse a saved one, else check for a key, else ask Claude (`main.ai_step`). Returns the file for the deck and a note for the page. | The deck is always built: without AI text, slide 4 shows "AI summary unavailable" (CLAUDE.md decision K). |
| `build_in_folder(...)` | Every step for one upload, in a temporary folder: check it's a real .xlsx → clean → metrics → Excel → AI → deck. Returns everything the page shows, files as bytes. | |
| `build_outputs(file_name, file_bytes, include_ai, saved_dir, client)` | Runs `build_in_folder` and **never raises**: a bad workbook returns clean.py's message, a non-Excel file returns `NOT_A_WORKBOOK`, and a bug returns "Something unexpected went wrong ..." (its traceback goes to the Terminal window only). | "Input errors show as plain messages, never a traceback." |
| `styled(table, colors)` | Puts the colors on the table for `st.dataframe`. | pandas' `Styler.apply`. |
| `show_downloads(result)` / `show_result(result)` | Draws the page: the error, or the heading, the two download buttons, the AI note, the flags, the data gaps, the metrics. | |
| `main()` | The page: title, checkbox, file drop. Builds once per file and checkbox choice. | Runs only when Streamlit runs the file, so the tests can import `app.py` without drawing anything. |

---

### `provenance.py` and `approve.py`: where a deck came from, and who signed it off

**What they're for:** a printed slide should be traceable to the exact inputs behind it, and a person, not the code, should decide when a deck is fit for a board.
- `provenance.py` writes and reads `output/<company>_manifest.json`, and decides whether an approval still counts. `main.py` writes a manifest for every company it runs.
- `approve.py` is the human step: `python approve.py northwind` (your name from `git config user.name`) or `--reviewer "A Name"`. It writes the name and the time into the manifest, next to the workbook's and config.yaml's hashes. **It never builds a deck**: producing a deck and vouching for it stay two separate acts.

**The finance analogy:** a manager's sign-off on a reconciliation, stapled to the exact version of the file they checked. If someone edits the file afterwards, the sign-off doesn't carry over to the new version.

**What the deck shows:** the footer ends "AI-drafted | not reviewed" until someone approves, then "AI-drafted | reviewed by Tyler Ho on 2026-09-17" on the next build. There's no watermark by default; `--draft` adds "DRAFT - NOT REVIEWED" to unapproved decks only.

**`provenance.py`**

| Function | What it does, in plain English | Example / why it exists |
|---|---|---|
| `file_sha256(path)` | The file's SHA-256 fingerprint, read a chunk at a time. | Two files both called northwind.xlsx can hold different numbers; only the hash tells them apart. |
| `git_output(arguments, folder)` | Runs a read-only git command; None if git can't answer. | A copy of the project without git still runs. |
| `git_commit(folder)` | The commit, and whether there were uncommitted edits. | "9c1b52c" alone wouldn't describe code that had been edited since. |
| `timestamp()` | Now, to the second: "2026-09-17T14:03:11". | |
| `manifest_path(workbook_path, output_dir)` | `data/northwind.xlsx` → `output/northwind_manifest.json`. | Beside the deck and the metrics workbook. |
| `read_manifest(path)` | The saved manifest, or None if it's missing or unreadable. | A broken manifest doesn't stop a run; the deck just isn't approved. |
| `approval_status(manifest, input_hash, config_hash)` | (the approval, None) if a reviewer is recorded **and** today's workbook and config.yaml hash the same as when they approved; else (None, why not). | **The only judge.** main.py, build_deck.py and approve.py all ask it, so they can't disagree. |
| `deck_status(approval)` | The manifest's `deck.status`: "approved by Tyler Ho on 2026-09-17T22:33:47", or `NOT_REVIEWED` (the same words the `--draft` watermark uses; the footer says "not reviewed"). | |
| `build_manifest(...)` | Everything about one run in one dict: input and config hashes, commit, AI record, deck file, AI text or not, status, approval. | |
| `save_manifest(path, manifest)` | Writes it as indented JSON. | Meant to be opened and read by a person. |

**`approve.py`**

| Function | What it does, in plain English | Example / why it exists |
|---|---|---|
| `reviewer_name(given, folder)` | The name given, else git's `user.name`; stops if there's neither. | An approval must have a person's name on it. |
| `reviewed_documents(manifest)` | What the approval covers: `["deck", "memo"]` if that run built a memo, else `["deck"]`. | The memo's footer says "reviewed by" only if "memo" is in this list. An approval from before memos existed has no list, so it never vouches for a memo nobody read. |
| `approve(company, reviewer, ...)` | Writes the approval into the manifest, with the documents it covers. Stops if there's no manifest, or if the workbook or config.yaml changed since that run. | Approving then would put a name against numbers the reviewer never saw. |
| `main(argv, ...)` | Command line. Prints "Approved Northwind by ..." and "Rebuild the deck so its footer says reviewed: python build_deck.py data/northwind.xlsx", or one "Not approved: ..." line. | |

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
| `excel_number(cell)` | A cell back to a metric value: number → float, a reason label → NaN, an **exact** ∞ label → ∞. |
| `expected_format(column)` | The number format each kind of metric must have. |
| `check_file(path, answer_key)` | File name and sheet order. |
| `check_value_cell(cell, column, value, reason)` | One cell holds the right number, or the exact label for its reason or its ∞, with the right format. |
| `tripped_by_quarter(metrics, reasons, config)` | Every (quarter, metric) that trips, in every quarter. |
| `check_metrics_sheet(sheet, metrics, reasons, gaps, config)` | Headers, quarters, every value, format and fill. |
| `check_latest_against_hand_formulas(sheet, company, metrics)` | Latest-quarter cells equal the hand formulas. |
| `check_story_highlights(sheet, company, metrics)` | Latest-quarter red cells are exactly the flags the story says trip. |
| `check_flag_row(...)` / `check_status_and_color(row, name, status)` | One Flags row's values, threshold, status text and row color. |
| `check_flags_sheet(...)` | Every flag row, the combo row, and the runway context line. |
| `check_gaps_sheet(sheet, gaps)` | The Data gaps sheet lists exactly `data_gaps()`. |
| `check_company(company, config)` / `main()` | Runs it all for each company. |

#### `check_deck.py` (step 4 proof)

Builds each company's deck (with its saved analysis if there is one), **opens the saved .pptx**, and checks it against the saved metrics workbook. It compares with Excel as Excel displays it, not with the deck's own formatting code, so a formatting bug can't pass its own check. It saves the decks into `output/`, so they keep the AI text.

| Function | What it does |
|---|---|
| `excel_display(cell)` | What Excel shows in a cell, using the cell's own number format (`0.0%` → "97.1%"). |
| `number_tokens(text)` | Every number in a text with its sign and unit: "-19.0%", "11.0 mo". |
| `read_metrics_workbook(path)` | The Metrics sheet as {(quarter, label): shown text}, and the Flags sheet as {flag: row}. Skips the runway context row (it has no Status). |
| `allowed_numbers(table, flags)` | Every number shown anywhere in the metrics workbook. |
| `shape(slide, name)` / `slide_text(slide)` / `table_rows(slide)` | Find a box by name (exactly one must exist) / all text on a slide / slide 1's table as text. |
| `check_titles(slides, company)` | 4 slides, with titles typed by hand in the check. |
| `expected_flag_count(company)` | "6 of 9 flags tripped", counted from the story in check_companies.py. |
| `check_ai_slide(slides, summary, name)` | Slide 4 shows the "AI-drafted from computed metrics - review before use" line, the JSON's headline, every risk and question, and no win anywhere on the deck, **or** exactly "AI summary unavailable" without the AI-drafted line. Never a mix, and no AI text on slides 1 to 3. |
| `check_kpi_numbers(slide, table, flags, name)` | Every number on slide 1 is in the metrics workbook. The footer date is skipped here (checked in `check_footers`). |
| `check_kpi_rows(slide, table, flags, name)` | Row by row: latest and prior cells equal the Excel cells, thresholds and statuses match the Flags sheet, status cell colors are right. |
| `check_charts(slide, company)` | Two pictures; redraws the charts and checks the blank quarter has no bar and the cash line has a NaN there. |
| `check_risks_slide(slide, company, table_flags, gap_labels)` | Slide 3: the flag count, every tripped flag, the combo result, and every "data missing" metric (or "None"). |
| `expected_review(workbook, output_dir)` | What the footer must end with ("AI-drafted \| not reviewed" or "... reviewed by NAME on DATE"), worked out from the manifest here, **not** with build_deck's code. |
| `check_footers(slides, source_name, name, review)` | Every footer: the fictional-data note, file name, today's date, ends with the expected review status, and fits on one line. |
| `watermark_count(slides)` | How many slides carry the watermark. Without `--draft` it must be 0. |
| `check_draft_option(config, folder)` | Builds all 3 decks with `--draft` in a temporary folder: Alderpeak and Fernhollow get the watermark on 4 of 4 slides; Northwind, which is approved, gets none. |
| `frame_paragraphs(frame, where)` | Reads a saved text box back into text_fit.py's paragraph form; stops on any font below 12 pt. |
| `check_text_fits(...)` / `check_table_fits(frame, where)` | Re-measures the saved text and table cells: they need no more room than they have. |
| `check_no_overflow(presentation, name)` | Every shape inside the slide and above the footer line; every text fits. |
| `check_overflow_check_catches_overflow(path)` | Breaks a saved deck on purpose (a very long headline, then an 11 pt font): the overflow check must fail both times. |
| `check_company(company, config, output_dir)` | Runs all the checks for one company. |
| `tampered_analysis(folder, change)` / `check_bad_analysis_gets_placeholder(config, folder)` | Copies Northwind's analysis with one change ("11.0 mo" → "11.5 mo", or the wrong quarter): the deck must show the placeholder. |

#### `check_memo.py` (memo proof)

Builds each company's memo (with its saved analysis if there is one), **opens the saved Word file and PDF**, and checks every number in them against the saved metrics workbook, read with check_deck.py's own Excel-reading functions, not memo.py's.

| Function | What it does |
|---|---|
| `flag_counts(flags)` | The Flags sheet's rows counted by status, so "6 of 9 flags tripped" counts as in the workbook. |
| `runway_context(path)` | The runway-at-budget cell below the flag table, as Excel shows it. |
| `workbook_allowed(path)` | Every number the metrics workbook shows: check_deck's `allowed_numbers`, plus the flag counts and runway at budget. |
| `docx_body(path)` / `docx_footer(path)` / `pdf_pages(path)` | Read the saved files back: paragraphs and table cells, the Word footer, each PDF page's text. |
| `flat(text)` / `without_footer(page, footer)` | One space between words (a PDF wraps lines where it likes); a PDF page less its footer. |
| `sentence_tokens(text)` | check_deck's `number_tokens`, but "Q2 2026, compared with" reads as 2026, not "2026,". |
| `check_numbers(texts, allowed, where)` | Every number in the texts is in the metrics workbook. |
| `check_kpi_table(tables, table, flags, name)` | Row by row: latest and prior cells equal the Excel cells; thresholds and statuses match the Flags sheet; every flag has a row. |
| `check_flags_and_gaps(paragraphs, company, gap_labels)` | The flag count from the story, every tripped flag, every data gap (or None). |
| `check_ai_text(paragraphs, summary, name)` | The JSON's headline and questions under the AI-drafted line, no wins or risks; or "AI commentary unavailable" twice and no AI-drafted line. |
| `expected_memo_review(workbook, output_dir)` | "reviewed by NAME on DATE" only if a still-valid approval lists the memo, else "not reviewed"; worked out from the manifest here, not with memo.py's code. |
| `expected_footer(workbook, output_dir, model)` / `check_footer(...)` | The footer worked out here from git, the analysis and the manifest; it must be the Word footer and on every PDF page. |
| `check_same_text_and_no_em_dash(paragraphs, tables, pages, name)` | Every piece of the Word text is in the PDF; neither has an em dash. |
| `memo_numbers_check(...)` / `check_company(company, config, output_dir)` | Runs all the checks for one company. |
| `check_bad_analyses_are_unavailable(config, folder)` | An invented number, and a quoted ending cash the deck accepts: both give "AI commentary unavailable" with every computed number still there. |
| `check_number_check_catches_a_planted_number(memo, config, folder)` | Changes NRR's 97.1% to 44.4% in a copy of the saved memo: the number check must fail. |

#### `check_main.py` (batch runner proof)

Every run uses `--skip-ai`. The `main.py` process also gets no API key and an API address where nothing listens, so even a bug couldn't reach Claude.

| Function | What it does |
|---|---|
| `run_main(*args)` | Runs `python main.py ...` as a real separate process, with no key and a dead API address, and captures its output. |
| `summary_table(stdout)` | Reads the printed summary table back into rows. |
| `story_flags_text(company)` / `story_gaps_text(company)` | The expected table cells, built from each company's story, **not** from main.py. |
| `write_broken_workbook(folder)` | A workbook with most columns missing, in a temp folder. |
| `run_quietly(function, *args, **kwargs)` | Calls a function and captures what it prints. |
| `analysis_file_state(workbook)` | (size, modified time) of a company's analysis JSON, or None. Used to prove `--skip-ai` didn't touch it. |
| `headline_on_deck(path)` | The text in slide 4's Headline box. |
| `check_finds_the_three_companies()` | `--all` finds exactly the 3 workbooks. |
| `check_batch_run()` | Exit 0, the AI skip message 3 times, a fresh Excel file and deck per company, "OK (AI skipped)" rows matching each story, no analysis JSON written, replaced or deleted. |
| `check_summary_csv(started)` | `output/batch_summary.csv` was written by this run and matches each story. |
| `check_failure_does_not_stop_batch(folder, config)` | A broken workbook and a missing file between good companies: the good ones still succeed, and no traceback is printed for input errors. |
| `check_code_bug_does_not_stop_batch(config)` | Temporarily swaps in a `compute_metrics` that fails once (`buggy_first_call`): the error is recorded **with** a traceback and the batch continues. The real function is always put back. |
| `check_exit_codes(broken_path)` | 1 when a company fails; 2 for no arguments or both a file and `--all`. |
| `refuse(what)` | Makes a stand-in function that stops the check if it's ever called. |
| `check_skip_ai_never_calls_claude(folder, config)` | Replaces `analyze()`, the key check and the Anthropic client with `refuse` stand-ins, runs all 3 companies with `--skip-ai` into a temp folder, and checks each deck's headline is the placeholder and no JSON was saved. |
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
| `assign_letters(records)` | Shuffles the runs and labels them A–F (up to A–Z; more than 26 runs stops). |
| `save_blind_files(lettered)` | `answer_A.json` … holds commentary only; `key_DO_NOT_OPEN.json` holds models and stats. |
| `print_blind_answers(lettered)` | Prints the answers by letter with the rubric. |
| `command_run(force)` | Runs S, H, S, H, S, H, then saves and prints the blind set. Refuses to overwrite an existing set without `--force`. |
| `parse_scores(score_args, runs)` | `["A=4", "B=3"]` → `{"A": 4, "B": 3}`. Every passed answer needs a whole-number score from 1 to 5, given once. |
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

Run with `python -m pytest -q` (575 tests, about a minute). Expected values are **worked out by hand** in comments, not copied from running the code. Tests with `@pytest.mark.parametrize` run the same test on many inputs, each inputs line counting as one test.

| File | Helper functions | What the tests cover |
|---|---|---|
| `test_clean.py` (81) | `write_workbook(path, labels, blank)`: a tiny workbook in pytest's temp folder. | `parse_number` (good text, real numbers, blanks → NaN, 17 kinds of unreadable text stop), `normalize_header` and `standard_column` on most of the Northwind headers, quarter labels and order (Q4 → Q1 rollover, skipped, repeated), and whole workbooks (blank row kept, missing row stops, duplicate row stops). |
| `test_metrics.py` (163) | `table(**columns)`: a small table with only the needed columns. `values(series)`: compare with NaN allowed. `burn_table`, `cac_table`: tables for one metric. `full_actuals(blank, blank_cells)`: 8 realistic quarters with optional blanks. `combo_metrics`, `reasons_for`, `combo`, `flags_for`, `gaps_for`, `reasons_of`: shortcuts for combo, flag, gap and reason tests. `TEST_CONFIG`: thresholds typed into the test file, so editing config.yaml never breaks a test. | Every metric against hand math; every CLAUDE.md edge case (∞, 0, −0.0); a missing input never becomes 0 or ∞; `check_threshold` exactly at the threshold, float noise, real misses, NaN, ∞; `check_combo` trip/pass/cannot evaluate with its reason and the 1-point minimum; config validation; the three reasons (every input blanked one at a time must say missing input); not-meaningful budgets; `data_gaps` following the QoQ/YoY rules; the printout's words. |
| `test_bad_inputs.py` (63) | `good_table()`: a valid table. `set_cell`, `drop_column`, `add_column`: break one thing. `write_workbook(path, rows, empty_columns_left)`: Notes tab, title, empty row, table from row 3. `error_from`, `assert_stops_with`: run `clean_workbook` and check how the error message starts. `reorder_quarters(labels)`: quarter rows in a given order. | Broken workbooks stop with the sheet name and Excel address: missing columns, two headers with one meaning, unknown headers, quarters out of order, a budget row with actuals, a budget row for the wrong quarter or with no quarter, two KPI tabs, footnote rows, unreadable text, Excel error cells. `test_good_workbook_cleans` proves the starting workbook is valid, so each failure comes from the one thing that was broken. |
| `test_analyze.py` (46) | `summary_saying(text)`: an answer with one piece of text. `northwind`: the real Northwind payload, built once. | Minus signs in the number check; the payload's words for each reason and its flag names; `save_analysis` output that `build_deck.load_analysis` accepts (passed) or rejects (failed). |
| `test_excel_output.py` (9) | `two_quarters`, `budget_row`, `runway_context_cell`, `fill`, `metrics_cell`, `reason_workbook`. | Runway-at-budget labels (a blank wins over ∞), the words and gray fill for each reason, the Flags status text, the Data gaps sheet. |
| `test_compare_models.py` (6) | none | Blind letters never drop a run; a letter scored twice stops. |
| `test_make_template.py` (9) | `placeholder_types(layout)`, `theme(presentation)`. | 16:9, exactly the 2 layouts, no slides, title/body/footer placeholders inside the slide in the right order, navy/gray theme and Arial, the saved file opens again. |
| `test_text_fit.py` (13) | none | Wider text measures wider (bold wider still), wrapping, a word wider than the box, height, shrinking all sizes together, never below 12 pt, failing loudly with the box's name, table fitting. |
| `test_charts.py` (5) | `texts(axis)`, `bar_positions(axis)`. | No bar for a blank quarter and "data missing" written there; latest values labelled; the cash line keeps the NaN (so it breaks); runway text in the title and zero on the axis; figure drawn at slide size. |
| `test_build_deck.py` (53) | `flag(...)`, `three_quarters(blank)`, `deck_data(company, blank)`: a tiny 3-quarter company. `summary_dict()`, `write_analysis(...)`, `payload`: analysis files. `build(tmp_path, summary)`, `shape`, `all_text`, `status_fills`, `run_sizes`: build and read a deck. | Flag count and threshold wording; data gaps grouped by quarter; **every way `load_analysis` must reject a file** (missing, not JSON, failed, wrong shape, other quarter, other company, a number not in today's data, 2 questions); 4 slides in order; placeholder vs AI text on slide 4, the AI-drafted line, no wins on the deck; footer on every slide; status colors; "data missing" in the table; slide 3 contents and flag count; matching column font sizes; 2 chart pictures; text too long fails loudly; **no digit typed in any text in build_deck.py, charts.py or memo.py**; no watermark by default, 4 of 4 with `--draft`, none on an approved deck even with `--draft`; both footer review wordings, the one-line fit and the long-name fallback. |
| `test_memo.py` (35) | `three_quarters(blank)`, `memo_data(blank)`: the same tiny 3-quarter company as test_build_deck.py. `summary_dict(headline, question)`, `summary_from`, `write_analysis`, `payload`: analyses. `all_text`, `docx_text`: read the memo back. | Title, quarter and sections in order; the key metrics table's values, thresholds and statuses worked out by hand (NRR −300.0%, runway 36.0 mo); tripped flags with value and threshold; the combo rule in the Flags sheet's words; data gaps or None; **no em dash**; the AI headline and questions but no wins or risks, and questions as bullets (not "1.", "2.", "3.", which aren't in the workbook); every way the AI text is refused, including a number in the payload the workbook doesn't show; thresholds counted as Excel displays them; the footer; the Word file and the PDF hold every piece of text, the PDF is 1 or 2 pages with the footer on each, draws "∞", keeps a heading with its text and wraps a long footer; old files deleted before a build; the manifest's memo part; "reviewed by" only when the approval lists the memo. |
| `test_main.py` (32) | `ok(...)`, `failed(...)`: result dicts. `FakeClient`: stands in for the Anthropic client, returns a fixed answer (or raises) and counts calls. `no_real_client`: runs before every test and makes creating a real client fail the test. `summary`, `run_northwind`, `headline_on_deck`, `saved_analysis`. | The CSV, both warnings and the result texts; a passing answer lands on the deck (1 call); an answer with an invented number is called exactly twice, then the placeholder and "OK (AI failed)"; an API error is "OK (AI failed)", not FAILED; a bug in the AI step fails the company and leaves no old analysis; `--skip-ai` never calls Claude or looks for a key; a missing key stops the run before any company; the manifest; no watermark by default and "not reviewed" in the footer, `--draft` from the command line stamps every slide, a still-valid approval is named on a re-run; the memo is built beside the deck (AI text, skipped, AI failed, and AI text the deck takes but the memo refuses) and recorded in the manifest. |
| `test_provenance.py` (13) | `manifest_for(...)`: a manifest for a tiny workbook in a temp folder. | The hash is the standard SHA-256 and changes only when the bytes do; the commit (or "unknown" outside git); a missing or broken manifest reads as None; a manifest records every input and output; no approval → not reviewed; an approval holds for today's files and is void once the workbook or config.yaml changes. |
| `test_approve.py` (10) | `company`: a workbook, config and manifest in a temp folder. `approve_testco(...)`. | The reviewer and time are recorded and nothing else in the manifest changes; no manifest, a changed workbook, changed thresholds or no name each stop; the command line says to rebuild so the footer says reviewed (never "watermark": that needs `--draft`), or prints one line when it refuses. |
| `test_app.py` (19) | `FakeClient`, `summary(headline)`, `no_real_client` (as in test_main.py). `build_northwind`: the Northwind workbook's bytes through `build_outputs`. `save_northwind_analysis`: a saved analysis of today's numbers. `headline_in(deck_bytes)`. `render_northwind`, `render_bad_file`: draw the results with Streamlit's `AppTest`. | The checkbox names the cost; an upload can't escape its folder; the Excel colors; clean.py's message word for word, a non-Excel file and a bug each give a plain message with no traceback; Northwind gives a 4-slide deck and a 3-sheet workbook, 6 of 9 flags in red and green rows, "data missing" gray and a tripped NRR red; a saved analysis of the same numbers is reused with **no** call, one of other numbers is not; no saved analysis → 1 call; no key or an API error still builds the placeholder deck; the page and the results draw with no error; `run_app.command` is executable and starts app.py. |
| `test_docs.py` (15) | `python_files_named(text)`, `study_guide_tables()`, `defined_in(files, name)`, `files_named`, `functions_named`, `watermark_lines(text)`. | README keeps the model comparison markers, and rewriting that block leaves the rest alone; every `.py` file named in README, CLAUDE.md, this guide and LOOM_SCRIPT.md exists; **every function in this guide's tables exists** in the file its heading names; INTERVIEW_PREP.md's files, functions and group order; README, CLAUDE.md, this guide and LOOM_SCRIPT.md never describe the watermark without `--draft`, all name `--draft`, app.py and run_app.command, and count 4 slides. |

---

### Small files

| File | What it is |
|---|---|
| `requirements.txt` | The packages: pandas (tables), openpyxl (Excel), python-pptx (the deck) and matplotlib (the charts, plus the font text_fit.py measures with), anthropic (Claude API), python-dotenv (.env), pyyaml (config.yaml), pydantic (answer shape), pytest (tests), streamlit (the web page, app.py). Pillow, which text_fit.py imports, comes in with matplotlib and isn't listed. |
| `pytest.ini` | Tells pytest to look only in `tests/`, and lets tests `import clean` from the project folder. |
| `.env` / `.env.example` | `.env` holds `ANTHROPIC_API_KEY` and is never committed. `.env.example` shows the variable name with no key. |
| `run_app.command` | Double-click it on a Mac to start the web page (app.py). The first time, it creates `.venv` and installs `requirements.txt`; then it starts Streamlit and opens the browser. The `.command` ending is what makes Finder run it in Terminal. |
| `.streamlit/config.toml` | Streamlit's settings for app.py: headless (no first-run email question), no usage statistics, **no tracebacks on the page**, and a minimal toolbar. |
| `.gitignore` | Keeps `.env`, `.venv/`, `__pycache__/`, `output/` and `.DS_Store` out of git. |
| `CLAUDE.md` | The project spec: goal, rules, metric definitions, edge cases, build order. |
| `LEARNINGS.md` | Everything that broke and how it was fixed, plus the prompt iterations and model comparison. **Interview gold.** |
| `OVERNIGHT_REPORT.md` | What each overnight task built, the decisions made, what failed, what's unresolved (a historical record). |
| `DAY_REPORT.md` | The same for today's tasks (the deck, main.py wiring, the live run, docs), plus a review of everything committed today. |
| `README.md` | What it does, how to run it, data flow, design decisions, the model comparison, cost, screenshots to capture, next steps. |
| `LOOM_SCRIPT.md` | The 2-minute demo video script, with timestamps and recording prep. |
| `POLISH_REPORT.md` | The polish tasks (review status in the footer, the 4-slide deck, the web page, interview prep, these docs): what each built, the decisions made, what failed, what's unresolved. |
| `INTERVIEW_PREP.md` | Every interview question in the order it tends to be asked, with 30–60 second answers and where to point. |
| `templates/base.pptx` | The brand template, built by `make_template.py` and committed. |
| `output/` | Generated files (git-ignored): `*_board_pack.pptx`, `*_metrics.xlsx`, `*_analysis.json`, `*_manifest.json` (with any approval), `charts/*.png`, `batch_summary.csv`, `compare/`, `day_logs/` (the saved live run). |

---

## 5. 30 interview questions

Short answers you can say out loud in 30–60 seconds. **"Point to"** says where to look in the code or docs if they ask for detail. Practise saying the answers in your own words; don't memorize them.

> **For interview practice, use INTERVIEW_PREP.md.** It has these 30 in the order they'd be asked, plus 9 more (direction checks and their cost, provenance, scale, working method), with facts that changed since this section was written (cost, test count, the direction check) brought up to date.

### The project

**Q1. Walk me through the project in one minute.**
A messy portfolio-company KPI workbook goes in; a board pack comes out. `clean.py` reads the messy Excel: odd headers, "$14.3M" typed as text, a blank quarter. It produces a standard table in $K. `metrics.py` calculates NRR, GRR, burn multiple, runway, Rule of 40, CAC payback and budget variances, then checks 9 flags against thresholds in `config.yaml` and lists any data gaps. `excel_output.py` writes a highlighted Excel summary, and `analyze.py` has Claude write the headline, wins, risks and questions for management. Code then checks that Claude used only numbers from the data. `build_deck.py` puts it all on a 4-slide deck: key metrics table, ARR and cash charts, risks and flags with data gaps, and the AI commentary (headline, risks and questions, marked as AI-drafted). `main.py` runs it all for one company or a whole folder, and still builds the deck if the AI step fails.
*Point to:* section 2 of this guide.

**Q2. Why would a PE fund want this?**
Portfolio companies report KPIs in their own formats, so analysts spend time re-typing and reconciling before they can think. This standardizes the file, applies the **same** definitions and thresholds to every company, and flags what needs attention, so the analyst's time goes to the questions for management. It also makes gaps explicit instead of letting a missing quarter quietly distort a trend. The AI step costs about $0.05 per company: $12.81 per quarter for 275 companies, from the live run of all three (the model comparison estimated $14.65 from Northwind alone).
*Point to:* README.md Cost.

**Q3. Why does Python do all the math and Claude only interpret?**
A language model can make arithmetic mistakes, and one wrong number in a board deck undermines every other number in it. Python math is deterministic and can be tested against hand formulas; a prompt can't be tested the same way. So every number is computed and checked in Python first, Claude receives them already formatted ("97.1%"), and the prompt tells it to quote, not calculate. Code then verifies the answer.
*Point to:* CLAUDE.md Rules; `analyze.py` `SYSTEM_PROMPT`.

**Q4. Why build three fake companies instead of one?**
With one company you can't tell whether the flags work or were tuned to produce that company's story. Northwind trips 6 of 9, Alderpeak (healthy) trips none in any quarter, and Fernhollow (distressed) trips 7, with 1 flag it can't evaluate. That proves the logic responds to the data. Each company also has different mess (a title row, a different column order, the Notes tab first), so cleaning isn't tuned to one file either.
*Point to:* `check_companies.py`; OVERNIGHT_REPORT Task 1.

### Design decisions

**Q5. How do you handle missing data? Why not fill it in?**
A blank cell stays blank (NaN), never 0 and never an estimate: an imputed number could reach a board as if it were real. Because any math with NaN gives NaN, a blank quarter automatically spreads to every metric that uses it: QoQ metrics for that quarter and the next, YoY for that quarter and the one 4 later. Flags that depend on a missing value say "cannot evaluate — missing input" instead of pass or fail, and every affected metric and flag is listed as a data gap. Edge-case rules (∞, 0) only apply when every input is present, so a blank cell can never become a red flag.
*Point to:* `metrics.data_gaps`; CLAUDE.md "Messy data rules".

**Q6. What's the difference between "data missing" and "not meaningful"?**
YoY growth in the first year of data is blank because there's no earlier year: that's "n/a (no prior period)", not a problem. YoY in Q1 2026 is blank because Q1 2025 exists but wasn't reported: that's "data missing". A third case is "not meaningful": every input is there but the math is undefined, e.g. 0 ÷ 0, or burn vs a budget of 0 (shown with the $K figures instead). Infinite values have their own words too: burn multiple "∞ (ARR shrank)", runway "∞ (not burning)". CLAUDE.md says these must never look alike, because a reader would draw different conclusions. Each metric declares its inputs (`METRIC_INPUTS`), and `metric_reasons` checks them: a blank input → missing input; an earlier quarter that doesn't exist → no prior period; otherwise NaN → not meaningful. Only missing input is a data gap.
*Point to:* `metrics.metric_reasons`, `metrics.data_gaps`, `metrics.display_value`, `excel_output.cell_value`.

**Q7. Why store ratios as decimals and format as % only at output?**
One representation everywhere means thresholds, comparisons and math never mix 97.1 with 0.971. Formatting happens in one place (`format_value`, or an Excel number format), so the Excel cell still holds 0.971 and works in formulas while displaying 97.1%. The Excel check deliberately tested saving 97.1 instead of 0.971, and it's caught.
*Point to:* `metrics.format_value`, `excel_output.number_format`.

**Q8. Why is NRR annualized, and what's the trade-off?**
The thresholds (NRR 100%, GRR 85%) are annual conventions. One quarter's expansion minus churn is roughly a quarter of the annual effect, so it's multiplied by 4 to compare like with like. The trade-off: one quarter's data is noisier than a true trailing-12-month cohort. So the deck labels it "annualized", and the combo rule looks at a 3-quarter trend rather than a single quarter.
*Point to:* `metrics.nrr`, `metrics.grr`; CLAUDE.md metric definitions.

**Q9. Explain the combo rule. Why "cannot evaluate" instead of False?**
NRR falling while pipeline is rising suggests a retention problem, not a sales problem: sales keeps filling the funnel, but existing customers are leaking, so more pipeline won't fix it. It trips only if NRR fell by at least 1 point (`combo_min_nrr_drop`) at **every** step and pipeline rose at **every** step over the last 3 quarters, so a 0.1-point wobble isn't called a retention problem. If any quarter in the window is missing, returning False would tell the board "no retention problem" when we simply don't know. So it returns "cannot evaluate". Fernhollow passes because its pipeline is falling too: that's a sales problem **and** a retention problem.
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
- A calculated number that happens to equal another value passes. The first planned test, "NRR fell 11 points", would have passed because runway is 11.0 mo, so the tests use numbers proven to be absent. In the blind runs, "3.3 months faster" passed because −3.3% is a Rule of 40 value. (A second number problem, ignoring minus signs, **is fixed**: "19.0%" now fails when the data says −19.0%.)
- It can't catch a wrong direction ("improved… down from 20.3 mo" when it went from 20.3 to 20.7, which is worse). The live run found more: Alderpeak's "persistent decline" in an NRR that goes up and down, and Fernhollow's passing combo flag called a win when it passes only because pipeline is falling too.
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
3. 415 pytest unit tests check each function, including every edge case, with the hand math in comments. `check_deck.py` also opens each saved deck and checks every number on slide 1 against the metrics workbook.
4. The tests were tested: the code was broken on purpose (in throwaway copies) to confirm the tests notice. 35 of 36 breaks were caught in Task 4 (the missed one can't change any result) and 18 of 18 in Task 5.

*Point to:* `check_companies.py`, `tests/`, OVERNIGHT_REPORT "What failed".

**Q24. In a batch of 275 companies, what happens when one workbook is broken?**
`clean.py` stops that company with a message naming the sheet and cell, e.g. `Sheet 'KPI Tracker', cell G5 (Q2 2025, revenue): Can't read 'TBD' as a number`. `main.py` catches it, prints one ✗ FAILED line, and moves on to the next company. The summary table shows FAILED with the reason, and the run exits with code 1 so a scheduler knows. A bad-input error gets one clean line; an unexpected error (probably a bug) also gets a full traceback.
*Point to:* `main.run_batch`, `main.describe_error`; `check_main.py`.

**Q25. What would you improve or build next?**
Every build step is done, so next is making it safer and easier to run at scale:
- **Fit Claude's text to the slide:** today a valid answer a few words longer than the live run's fails the whole company. I'd check fit inside the validator, so it uses the retry, and fall back to the placeholder if it still doesn't fit.
- **Check trend words:** that "rose" or "fell" matches the numbers, because the live run found claims that passed every check but were misleading. Until then a person reads each deck.
- **A SharePoint or Power Automate trigger,** so a company's upload starts the run.
- **A portfolio rollup** across all companies.
- **Run companies in parallel:** about 33 s of API time each means about 2.5 hours for 275 companies one at a time.
*Point to:* README Next steps; DAY_REPORT.md Review.

### The deck

**Q26. Why does the deck check Claude's analysis again, when it already passed validation?**
Because the deck may be built later than the analysis, from a workbook that has changed, or from a file someone edited. `load_analysis` accepts the saved JSON only if it's for the same company and latest quarter, has the right shape, and `validate_summary` still passes against a payload **rebuilt from today's workbook**. So every number in Claude's text is still in today's data. If any check fails, slide 4 says "AI summary unavailable" and the terminal prints why. `check_deck.py` proves it by changing "11.0 mo" to "11.5 mo" in a copy, and by labelling a copy for the previous quarter: both get the placeholder.
*Point to:* `build_deck.load_analysis`; `check_deck.check_bad_analysis_gets_placeholder`.

**Q27. What happens to the deck if the AI step fails? Why build it at all?**
The deck is still built. Every number on it comes from Python, so the metrics table, charts, flags and data gaps are valid whatever Claude did. Only slide 4 changes: it says "AI summary unavailable", with a note that the other numbers are unaffected. The batch result reads `OK (AI failed)`, a warning names the company, and the reason is saved in its analysis JSON. "AI failed" means only two things: the answer failed validation twice, or the API call itself failed. Any other error is treated as a bug and fails the company, so a coding mistake can't hide behind "AI failed". A missing API key stops the run before any company starts.
*Point to:* `main.ai_step`, `main.ai_status`; `tests/test_main.py` (fake client).

**Q28. How do you make sure text doesn't run off a slide, when PowerPoint doesn't tell Python how big text is?**
`text_fit.py` estimates it. It measures each word's width with a real font file (DejaVu Sans, slightly wider than Arial, so it errs toward "needs more room"), wraps words into lines, and adds up line heights. If the text is too tall, every font size in the box drops by 1 pt together until it fits. If it doesn't fit at 12 pt, the build stops with the slide and box named, because a board deck with text off the page is worse than no deck. PowerPoint's own autofit is switched off, so the saved sizes are what's shown. `check_deck.py` re-measures the saved file, and was itself proven by breaking a deck on purpose. **The honest limit:** it's an estimate, not PowerPoint's own layout, and slide 4's Risks column has little spare room for Claude's text today: Northwind's and Fernhollow's risks fit only at the 12 pt floor (DAY_REPORT Review, finding 1; POLISH_REPORT Task 2).
*Point to:* `text_fit.shrink_to_fit`, `check_deck.check_no_overflow`.

**Q29. How does the deck show a quarter the company never sent?**
It never fills it in. On slide 1, a value that needs the blank quarter says "data missing" (Northwind's Q1 2026 ARR growth YoY compares with the blank Q1 2025), which is different from "n/a (no prior period)". On slide 2, the ARR chart draws no bar for Q1 2025 and writes "data missing" there. The cash line keeps the blank as NaN, so matplotlib breaks the line instead of drawing a straight join that would invent the numbers in between. On slide 3, a Data gaps column lists every affected metric and flag, grouped by the quarters they miss. `check_deck.py` checks there's no bar, the line breaks, and every "data missing" metric in the Excel file appears on slide 3.
*Point to:* `charts.bar_panel`, `charts.cash_chart`, `build_deck.gaps_lines`.

**Q30. How do you know no number on the deck was typed in by hand or mis-formatted?**
Two ways. First, a unit test reads the source code of `build_deck.py` and `charts.py` and fails if any piece of text in them (docstrings aside) contains a digit, so even "of 9" or "12 months" can't be typed in. Every number comes from metrics.py, config.yaml or the validated analysis, formatted by `metrics.format_value`. Second, `check_deck.py` opens the saved deck, collects every number on slide 1, and checks each one appears in the saved metrics workbook **as Excel displays it**, using Excel's own number formats rather than the deck's formatting code. It also checks each row's cells, thresholds, statuses and colors. Breaking a cell on purpose (a wrong number, a right number in the wrong column, a wrong status) was caught each time.
*Point to:* `tests/test_build_deck.py::test_no_digit_in_any_text_written_in_the_code`, `check_deck.check_kpi_numbers`, `check_deck.check_kpi_rows`.

---

## 6. 10 trace-this-number exercises

**For each number:** name every file and function it passes through, from the workbook cell to the screen, and write out the arithmetic. Answers are in [section 8](#8-answers-to-the-exercises). Try each one first.

**Set up:** run `python metrics.py data/northwind.xlsx` and keep the output open. It helps to open `data/northwind.xlsx` too. In the KPI Tracker tab, row 1 is the header, Q3 2024 is row 2, the blank Q1 2025 is row 4, Q2 2026 is row 9, and the budget-only row is row 10.

**Tools you'll need:** `clean.py` (reading), `metrics.py` (math and flags), `config.yaml` (thresholds), `make_data.py` (where the number was born). Exercise 10 also uses `main.py`.

### Exercise 1: Northwind NRR 97.1% (Q2 2026)
**Where you see it:** `TRIP  NRR (annualized)  97.1% (threshold 100.0%)`
**Your task:** Which 4 input cells does it come from, and which one is stored as text? Which function turns that text into a number? Which function computes NRR, and what's the formula with Northwind's numbers? Why does it trip? Where does 0.97… become "97.1%"?

### Exercise 2: Northwind runway 11.0 mo (Q2 2026)
**Where you see it:** `TRIP  Runway at current burn  11.0 mo (threshold 12.0 mo)`
**Your task:** Ending cash is typed as `"$14.3M"`. Trace exactly how `parse_number` turns that text into a number, step by step. Then compute runway. Why divide the burn by 3?

### Exercise 3: Northwind burn vs budget 20.0%
**Where you see it:** `TRIP  Net burn vs budget  20.0% (threshold 15.0%)`
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
**Where you see it:** in the `metrics.py` printout, the ARR growth YoY row reads "n/a (no prior period)" for Q4 2024 and "data missing" for Q1 2026. The Excel Metrics sheet (`python main.py --all --skip-ai`, then open `output/northwind_metrics.xlsx`) says the same, with Q1 2026 in a gray cell.
**Your task:** Why is each one blank? Which function decides that only one of them is a data gap, and what does it check? Which function then picks the words for the printout, Excel and Claude?

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

**Spec:** return `CANNOT_EVALUATE`, `TRIP` or `PASS` (constants already defined at the top of the file).
- If `value` is NaN → `CANNOT_EVALUATE`.
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

**Test:** `python -m pytest -q tests/test_clean.py tests/test_bad_inputs.py` (144 tests; the second file reads real workbooks through your function)
**Then:** `python check_companies.py` (cleaning must recover every answer-key value exactly).

---

### 5. `check_combo(metrics, reasons, config, quarter)` in `metrics.py`: hardest

**New skill:** taking a window of rows, comparing steps, and treating "can't tell" as its own answer.

**Spec:** return a pair `(status, reason)`. `(TRIP, None)` if NRR fell by at least `config["combo_min_nrr_drop"]` at **every** step **and** pipeline rose at **every** step over the last `config["combo_lookback_quarters"]` quarters, ending at `quarter`. Otherwise `(PASS, None)`. Start with `validate_config(config)`. Return `CANNOT_EVALUATE` with a reason, checked in this order:
1. any NRR or pipeline cell in the window (or in the quarters that exist, if the history is short) is `MISSING_INPUT` → `MISSING_INPUT` (a blank wins),
2. there isn't enough history for a full window → `NO_PRIOR_PERIOD`,
3. any other NRR or pipeline cell in the window has a reason → the first reason found.

**Hints:**
- `metrics.index.get_loc(quarter)` gives the row position of a quarter label (Q3 2024 = 0).
- `metrics.iloc[start:end]` takes rows `start` up to but **not including** `end`. So for a window of 3 ending at position 7, you want `iloc[5:8]`. With a short history, `max(end - size, 0)` keeps `start` from going negative.
- `.isna().any()` → True if any value is NaN.
- `.diff()` gives this row minus the row above. The first row of the window has nothing above it inside the window, so skip it with `.iloc[1:]`.
- `(steps <= -min_drop).all()` → True only if every step fell at least the minimum.
- Round **after** `.diff()` (`.round(6)`), so 1.07 − 1.08 = −0.010000000000000009 counts as exactly 1 point and a tiny float wobble doesn't count as a decline.
- `isinstance(r, str)` tells a reason from None in the reasons table.

**Traps the tests catch:**
- A flat step, or a 0.1-point dip, counts as falling.
- Exactly 1.0 point doesn't count (rounded before subtracting instead of after).
- Missing data returns `PASS` instead of `CANNOT_EVALUATE`, or the wrong reason.
- A window that's one row off.
- Older quarters outside the window affect the result.
- A quarter other than the latest is evaluated as if it were the latest (`test_check_combo_earlier_quarter`).

**Test:** `python -m pytest -q -k check_combo` (20 tests)
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
2. **`metrics.input_reason`** (called by `metric_reasons`) checks `METRIC_INPUTS["arr_yoy"]`: starting ARR and the 4 ARR flows, this quarter **and** 4 quarters back.
   - **Any of those cells blank → missing input.** Q1 2026 looks back to Q1 2025, which is blank → **missing input**. (Q1 2025 itself: its own cells are blank → also missing input.)
   - **Else, a needed quarter doesn't exist → no prior period.** Q4 2024 is position 1; 4 back would be position −3 → **no prior period**.

   `data_gaps` lists only missing input, so `arr_yoy` gaps: Q1 2025, Q1 2026.
3. **The words:** `metrics.reason_text` picks them from the reason: "data missing" or "n/a (no prior period)". `metrics.display_value` uses it for the printout and Claude's payload (`analyze.metric_trend`), and `excel_output.cell_value` for Excel, where `style_metric_cell` adds gray only for missing input.
4. **Before the hardening pass** the printout said "n/a" for both, because it formatted every NaN without asking why (LEARNINGS row J).

### Answer 9: the combo rule trips

1. **Window:** `combo_lookback_quarters: 3` in config.yaml. In `metrics.check_combo`: `end` = position of Q2 2026 (7) + 1 = 8, so the window is rows 5–7 = **Q4 2025, Q1 2026, Q2 2026**.
2. **NRR** (`metrics.nrr`):
   - Q4 2025: 1 + 4 × (860 − 140 − 290) / 21460 = **1.0801** (108.0%)
   - Q1 2026: 1 + 4 × (710 − 200 − 390) / 23790 = **1.0202** (102.0%)
   - Q2 2026: **0.9706** (97.1%)
3. **Pipeline** (text cells `"$10.1M"`, `"$11.2M"`, `"$12.5M"` in N7–N9 → `parse_number`): **10,100 → 11,200 → 12,500**.
4. **The test:** no value in the window has a reason, so it can be evaluated. NRR `.diff()`, rounded to 6 decimals: −0.0600, −0.0496, **all ≤ −0.01** (`combo_min_nrr_drop`: at least 1 point). Pipeline `.diff()`: +1,100, +1,300, **all > 0**. Both true → **trip**.
5. **Why not Q3 2025's 108.9% peak:** the window is exactly 3 quarters (2 steps) ending at the evaluated quarter. Q3 2025 → Q4 2025 was also a decline, but that step sits outside the window. That's also why the v2 prompt rule says to describe a trend "from the peak, or from the start of the flag's lookback window".

### Answer 10: Fernhollow "7 of 9, 1 cannot evaluate"

1. **The blank:** `make_data_fernhollow.py` → `BLANK_QUARTER = "Q2 2025"`. `write_kpi_sheet` writes only the label. `clean_sheet` keeps it as a row of NaN.
2. **The flag: Rule of 40.** `rule_of_40` = `growth(revenue, 4)` + `fcf_margin`. For Q2 2026, `.shift(4)` lands on **Q2 2025**: blank. So revenue YoY is NaN, and NaN + anything = NaN.
3. **"Cannot evaluate":** `evaluate_flags` → `check_threshold(NaN, 0.40, "min")`. The first line, `if math.isnan(value): return CANNOT_EVALUATE`, never returns a pass or a trip. The flag's reason comes from `metric_reasons`: Rule of 40 uses revenue 4 quarters back, which is blank → **missing input**.
4. **The summary text:** `main.run_company` collects `result["cannot_evaluate"] = ["Rule of 40"]`. `main.flags_text` → `"7 of 9"` + `", 1 cannot evaluate"`.
5. **Where 20 comes from:** `data_gaps` finds 19 metric columns with a gap (every metric misses Q2 2025; QoQ ones also Q3 2025; YoY ones also Q2 2026). Then a flag that can't be evaluated **because of a missing input** is added: `"flag: Rule of 40": ["Q2 2026"]`: 19 + 1 = **20**. `main.blank_quarters` → Q2 2025. `main.gaps_text` → "20 metrics/flags (blank: Q2 2025)".
6. **Bonus, burn multiple ∞:** net new ARR = 180 + 60 − 150 − 330 = **−240**, while net burn is 1650 > 0. In `burn_multiple`, `.mask(present & (net_burn > 0) & (new <= 0), math.inf)` → ∞ (burning cash while ARR shrank; `present` makes sure no input is blank). `check_threshold`: ∞ > 2.0 → trip. Excel can't store ∞, so `excel_output.cell_value` writes **"∞ (ARR shrank)"** from `INFINITE_LABELS`.

**Why this company was designed like this:** Fernhollow's blank is exactly 4 quarters before the latest one, so the "cannot evaluate" path is exercised on real data. Northwind's blank quarter (Q1 2025) is 5 quarters back, so none of its latest-quarter flags are affected.
