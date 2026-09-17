# Day Report

## Task 1: Hardening (decisions A, B, C, I, J)

### Where this started

Most of this task was already built and committed in `5cd93dc` ("Hardening pass A-O"). That commit covered A-O, and this task is the A, B, C, I, J part of it. So today I:

1. Re-ran everything: 338 tests and all 4 check scripts passed.
2. Went through each decision again, looking for any place the code still breaks its own rule.
3. Found and fixed 2 places where a blank input could still look like a real answer (commit `df6de63`).
4. Wrote this report.

Now: **344 tests pass**, and `check_northwind.py`, `check_companies.py`, `check_excel_output.py` and `check_main.py` all print "All checks passed". I didn't call the Anthropic API or edit config.yaml or CLAUDE.md today.

### What was built

**A. Each metric lists its own inputs; every value with no number has one reason**
- `METRIC_INPUTS` in metrics.py lists, for each metric, every input column it uses and how many quarters back it looks. For example, `rule_of_40` uses revenue and net burn from this quarter, plus revenue from 4 quarters earlier.
- `input_reason` walks that list for one quarter:
  - a blank input → **missing input**
  - otherwise, an input quarter before the first quarter in the workbook → **no prior period**
- `metric_reasons` builds a table shaped like the metrics table, holding each value's reason. If every input is present but there's still no number, the reason is **not meaningful**.
- Flags carry a `reason` (None unless the flag can't be evaluated). `data_gaps` lists a metric or flag only when its reason is missing input.
- Effect: a partly blank quarter no longer turns unrelated metrics into data gaps, and 0 ÷ 0 is no longer called a gap.

**B. CAC payback with a blank input can't be evaluated; it never trips**
- The edge-case rules (∞ or 0) now apply only when every input is present (`all_present`).
- The same bug existed in runway (blank cash while not burning showed ∞ and passed) and burn multiple (blank ARR flows while not burning showed 0). Both are fixed the same way.

**C. Budget comparisons against a budget of 0 or less are "not meaningful"**
- `burn_vs_budget` is NaN when `budget_net_burn` ≤ 0.
- `net_new_arr_vs_budget` is NaN when budgeted net new ARR (this quarter's `budget_arr` minus last quarter's) ≤ 0.
- Since every input is present, the reason is **not meaningful**:
  - the flag shows "cannot evaluate — not meaningful" and never trips
  - it isn't a data gap
- `not_meaningful_text` shows the $K figures instead, e.g. "n/m: net burn 900 vs budget 0 ($K)". The printout, the Excel file and Claude's payload all use this same function.

**I. One label set**
- `METRIC_LABELS` and `INPUT_LABELS` now live in metrics.py; before, they were in analyze.py.
- Flag names are built from `METRIC_LABELS`, so the Flags sheet and the Metrics sheet can't drift apart.
- excel_output.py and analyze.py both import the labels from metrics.py.

**J. The printout tells "data missing" apart from "n/a (no prior period)"**
- `display_value` picks the words from the reason: "data missing", "n/a (no prior period)" or "n/m ...".
- The printout, the Excel file and the payload all call it, so they say the same thing.

### Today's two fixes (found by re-checking A and B)

| Bug | Example | Fix |
|---|---|---|
| **Runway at next quarter's budgeted burn** showed "∞ (budget not burning)" when the latest cash cell was **blank** and budgeted burn was 0 or less. | Missing cash looked like "never runs out". | `runway_at_next_budget` checks for a blank cash or budgeted-burn cell first. It now says "data missing". |
| **Combo rule with a short history**: if the workbook had fewer quarters than the window (3) and NRR or pipeline was blank, the flag said "no prior period". | That flag was left off Data gaps, which breaks CLAUDE.md's rule that a blank input wins. | `check_combo` checks the quarters that do exist for a blank first. The order now matches the metrics: missing input, then no prior period, then any other reason. |

- I wrote the tests first: 5 new test cases failed on the old code and pass now.
- No real workbook hits either case (all three have both values and 8 quarters), so the check scripts didn't change.
- Both are logged in LEARNINGS.md, and STUDY_GUIDE.md (function table and exercise 5) now describes the new order.

### Tests and checks whose expected results changed (and why)

All of these changes were made in `5cd93dc`. Today's commit only added tests; it didn't change any existing expected result.

| Where | Before | After | Why |
|---|---|---|---|
| Flag names in `check_northwind.py`, `check_companies.py`, test files | "Burn vs budget", "Runway (months)", "CAC payback (months)" | "Net burn vs budget", "Runway at current burn", "CAC payback" | Decision I: flag names are now the Metrics sheet labels. The check scripts still type the names by hand, so a wrong label can't pass its own check. |
| Fernhollow's Rule of 40 flag | status `MISSING` | status `CANNOT_EVALUATE`, reason `missing input` | Decision A: "can't evaluate" is one status with a reason attached, not its own status. Same meaning as before, now with the reason. `check_companies.py` also asserts that every can't-evaluate flag in the 3 stories is missing input. |
| `check_threshold(NaN)` tests | `MISSING` | `CANNOT_EVALUATE` | Same rename. |
| `check_combo` tests | returned a status, e.g. `MISSING` | return `(status, reason)`, e.g. `(CANNOT_EVALUATE, NO_PRIOR_PERIOD)` | Decision A: the combo carries a reason too. The 2-quarter history test changed from "missing" to "no prior period", because a short history isn't missing data. |
| `evaluate_flags` / `data_gaps` tests | `evaluate_flags(metrics, config)` | `evaluate_flags(metrics, reasons, config)` | Flags now need the reasons table to fill in `reason`. |
| `test_burn_multiple_missing_input_is_nan`, CAC payback and runway missing-input tests | only checked blanks while burning | also check blanks while **not** burning, and blank S&M with no new ARR | Decision B: these cases used to come out 0 or ∞. |
| `check_excel_output.py` value cells | any NaN was "data missing" if it was a gap, otherwise "n/a (no prior period)"; any text starting with "∞" was accepted | the exact words for the cell's reason, and the exact "∞ (...)" label per metric | Decisions A and J: 3 reasons instead of 2, and "n/m ..." must never be read as "no prior period". |
| `check_excel_output.py` gray cells | gray where a value was missing | gray only where the reason is missing input | Decision A: "no prior period" and "not meaningful" aren't problems to chase. |

**What didn't change:**
- the flag statuses in all three stories (Northwind 6 trip / 3 pass, Alderpeak 0 trip, Fernhollow 7 trip)
- Fernhollow's 20 data gaps
- every metric value

The hardening changed how missing values are labelled, not the math.

### Decisions I made that you didn't specify

1. **The exact words for each reason**: "data missing", "n/a (no prior period)", "n/m (not meaningful)". For the budget cases in C: "n/m: net burn X vs budget Y ($K)" and "n/m: net new ARR X vs budget Y ($K)".
2. **Which reason wins when more than one applies**: missing input, then no prior period, then not meaningful. Missing input first comes from CLAUDE.md. Putting no prior period before not meaningful was my choice, so a first-year quarter reads "n/a" rather than "n/m". Today I made the combo rule follow the same order.
3. **"Not meaningful" is the leftover reason**: any value with every input present but no number. That includes an infinity outside the three allowed edge cases (e.g. growth from a zero base). It isn't a hand-written list, so a new 0 ÷ 0 case can't slip through as a gap.
4. **A flag's status and its reason are separate fields**: the status stays trip / pass / cannot evaluate, and `reason` is filled only for cannot evaluate. Text outputs join them as "cannot evaluate — missing input".
5. **Excel colors**: on the Metrics sheet, only missing-input cells are gray, and "n/a" and "n/m" cells have no fill. On the Flags sheet, every "cannot evaluate" row is gray, whatever the reason.
6. **For the flag renames, the metric labels won**: e.g. "Net burn vs budget", not "Burn vs budget". The flag names changed, not the Metrics sheet headers.
7. **A raw input shown in Claude's payload** (net burn, ending cash) reads "data missing" when blank. A raw input has no "no prior period" or "n/m" case.
8. **Runway at next quarter's budgeted burn keeps its own labels**: "n/a (no budget row)", "data missing" and "∞ (budget not burning)". It's one context number, not a per-quarter metric. Today I made a blank input win over the ∞ label.
9. **Combo with a short history and a not-meaningful NRR** reads "no prior period" (follows decision 2). There's a test pinning this.

### What failed and how it was fixed

- The two bugs above: a blank input looked like ∞ in runway-at-budget, and like "no prior period" in the combo rule. I wrote failing tests, fixed the code, then re-ran all 344 tests and 4 checks.
- `5cd93dc` logged its own failures in LEARNINGS.md, rows A-J:
  - B: blank S&M tripped CAC payback
  - C: a budget of −150 made +33% "over budget"
  - I: labels differed between sheets
  - J: "n/a" was used for both data missing and no prior period
- Two of my shell commands were blocked by the sandbox, one using a loop variable and one using `source .venv/bin/activate`. I re-ran them calling `.venv/bin/python` directly. No effect on the code.

### Unresolved

1. **ARR vs budget with a `budget_arr` of 0 or less isn't covered by C.**
   - A budget of 0 already comes out as plain "n/m (not meaningful)", because ÷ 0 is infinite. It doesn't show the $K figures like the two C metrics do.
   - A negative budgeted ending ARR would compute a meaningless %. That can't happen in real data, so I left it.
   - Tell me if you want it handled like C.
2. **The printout shows a bare "∞"**, while the Excel file shows "∞ (ARR shrank)" / "∞ (not burning)" / "∞ (never pays back)". J was only about missing vs no prior period, so I didn't change it. It's a one-line fix if you want every output to match.
3. **`5cd93dc` also edited config.yaml and CLAUDE.md.** That was for decisions D-O (the new `combo_min_nrr_drop` setting and the rule text), not this task. Today's rules forbid editing them, and I didn't touch either file. I'm noting it so you can review those edits separately.
4. **build_deck.py isn't built yet**, so the "Data gaps" line on the Risks/Flags slide is still untested. It should take its list from `data_gaps`, which already has only missing input.

## Task 2: The deck (make_template.py, build_deck.py, check_deck.py)

### Where this started

Nothing existed for this task: `templates/` was empty and nothing was committed. I built it in 4 commits:

| Commit | What |
|---|---|
| `a06ba68` | make_template.py and templates/base.pptx |
| `7927f38` | text_fit.py, charts.py, build_deck.py, and the deck step connected in main.py |
| `edda2be` | check_deck.py |
| `f690350` | side-by-side columns share one font size |

Now: **397 tests pass** (344 before, 53 new), and all 5 check scripts print "All checks passed": check_companies, check_deck, check_excel_output, check_main, check_northwind. I didn't call the Anthropic API or edit config.yaml or CLAUDE.md.

**To see the deck:** open `output/northwind_board_pack.pptx`. If you run `check_main.py` or `main.py --all --skip-ai` again, it overwrites that file with the placeholder version. Run `python build_deck.py data/northwind.xlsx` to get the AI text back (see Unresolved 2).

### What was built

**make_template.py → templates/base.pptx** (committed, so decks build without running it first)
- A fictional, neutral brand, "Example Capital":
  - colors: navy `1F2A44`, dark gray text, mid-gray rules and footer, light gray table stripes
  - font: Arial
  - size: 16:9 (13.333 × 7.5 in)
- Exactly two layouts:
  - **Title Slide**: navy background, white title
  - **Title and Content**: navy title, a content area, a footer strip
- Every content slide also gets a thin navy bar at the top, a gray line above the footer, and the brand name at the bottom right.
- How it's made: python-pptx can't create a template from nothing, so it starts from python-pptx's built-in default (4:3, 11 layouts). It resizes it, restyles it (theme colors, fonts, title and body text styles) and removes the other 9 layouts.
- build_deck.py reads the content area and the footer position from the layout, so moving them in make_template.py moves them on every deck.

**text_fit.py: making text fit**
- **Measuring:** PowerPoint doesn't tell python-pptx how much room text takes, so text_fit.py measures it with a real font file. It uses DejaVu Sans, which ships with matplotlib. DejaVu is wider than Arial, so the estimate errs toward "needs more room".
- **Wrapping:** words go onto a line until the next word doesn't fit.
- **Height:** number of lines × font size × 1.2, plus the space after each paragraph.
- **`shrink_to_fit`:** lowers every size in a box by the same step (a heading stays bigger than its body text) until the text fits. If it still doesn't fit at 12 pt, it stops with `TextDoesNotFitError`, naming the slide and the box, e.g. "Slide 4, Risks and flags: text doesn't fit even at the 12 pt minimum".
- **`fit_table`:** does the same for the key metrics table.

**charts.py: the two charts on slide 3**
- **ARR chart:**
  - ending ARR bars on top, net new ARR bars below
  - two panels with their own scales, never one chart with two y-axes: net new ARR is small next to ARR and can go negative
- **Cash chart:**
  - ending cash as a line, with zero kept on the axis
  - the title gives runway at current burn and at next quarter's budgeted burn
- **A blank quarter:**
  - bars: no bar is drawn, and the gap says "data missing"
  - line: the blank stays in the data as NaN, so matplotlib stops the line at the gap instead of joining across it
- **Size:** drawn at the size they appear on the slide, so 12 pt chart text is 12 pt on screen.

**build_deck.py → output/<company>_board_pack.pptx**
1. **Summary:**
   - title "Northwind: Q2 2026 board update"
   - the AI headline
   - "6 of 9 flags tripped"
   - wins and risks in two columns, each point a bold title plus its detail
2. **Key metrics:**
   - columns: Metric | Q2 2026 | Q1 2026 | Budget or threshold | Status
   - 3 context rows, then all 9 flags
   - status cells: red = tripped, green = passed, gray = cannot evaluate. These are the Excel file's colors.
3. **ARR and cash:** the two charts side by side.
4. **Risks and flags:**
   - left column: tripped flags (value and threshold), flags that can't be evaluated, the combo rule
   - right column: the Data gaps line
5. **Questions for management:** the AI's 3 questions, numbered.
- Every slide has the footer: "Fictional data, generated for demonstration | Source: northwind.xlsx | Run date: 2026-09-17".
- **The AI analysis is checked again before it's used.** `load_analysis` accepts the JSON only if every one of these holds:
  - the file exists and is JSON
  - it passed validation when it was made
  - it is for this company and this latest quarter
  - it has the BoardSummary shape
  - `analyze.validate_summary` still passes against a payload rebuilt from today's workbook. So every number in the text is still in today's data.
- If any check fails, slides 1 and 5 say "AI summary unavailable", and the terminal prints why. Everything else is built as normal.
- `output/northwind_analysis.json` passes all of these checks.
- **No number typed by hand:** a test reads build_deck.py and charts.py and fails if any piece of text in the code (docstrings excepted) contains a digit. Every number a reader sees comes from metrics.py, config.yaml or the analysis, formatted by `metrics.format_value`.
- **Run:** `python build_deck.py data/northwind.xlsx`. It uses `output/northwind_analysis.json` if it exists. Add `--analysis PATH` to use another file, or `--no-analysis` for the placeholder.

**check_deck.py: the end-to-end proof.** For each company it builds the deck, opens the saved file and checks:
1. **Slides:** 5 of them, titles typed by hand in the check.
2. **Slide 1 and slide 5:**
   - Northwind: the headline equals the JSON headline, and every win, risk and question appears.
   - Alderpeak and Fernhollow (no analysis file): exactly "AI summary unavailable" on both slides.
   - The flag count matches each company's story from check_companies.py.
3. **Slide 2:**
   - Every number on it appears in `output/<company>_metrics.xlsx`, in the Metrics sheet or among the Flags sheet thresholds.
   - The workbook values are shown in Excel's own number formats, so this doesn't reuse the deck's formatting code.
   - Row by row, the latest and prior cells equal that metric's Excel cells.
   - Statuses and thresholds match the Flags sheet, and each status cell has the right color.
   - I broke cells on purpose: a wrong number, a right number in the wrong column and a wrong status were all caught.
4. **Slide 3:** two pictures; the blank quarter has no bar, and the cash line isn't joined across it.
5. **Slide 4:**
   - every tripped flag
   - the combo result
   - every metric with "data missing" in the Excel file, or "None" for Alderpeak
6. **Nothing overflows:**
   - every text box and table cell is measured again from the saved file
   - every font is at least 12 pt
   - every shape is inside the slide and above the footer line
   - proven by breaking a deck on purpose (a very long headline, then an 11 pt font)
7. **Footer** on every slide, with today's date.
8. **Bad analyses** get the placeholder: a copy of Northwind's analysis with one number changed (11.0 → 11.5 mo), and one labelled for Q1 2026.

**main.py:** the deck step is now connected (see decision 14).

### Decisions I made that you didn't specify

1. **Brand:**
   - name "Example Capital" (clearly made up)
   - colors: navy `1F2A44` plus three grays
   - font: Arial
   - details: top navy bar, footer line, brand name at the bottom right
   - The Title Slide layout exists in the template, but the 5-slide deck doesn't use it.
2. **Template details:**
   - the date and slide-number placeholders are removed, since the run date is in the footer text instead
   - templates/base.pptx is committed to git
3. **Two extra files:** `text_fit.py` and `charts.py`, so build_deck.py stays readable (it's still ~560 lines). They aren't in CLAUDE.md's Architecture list, and I wasn't allowed to edit CLAUDE.md.
4. **How fit is estimated:**
   - DejaVu Sans widths (more cautious than Arial)
   - line height = 1.2 × font size
   - boxes keep their proportions as they shrink
   - the two columns on slides 1 and 4 always share one font size
   - even the footer respects the 12 pt minimum
   - PowerPoint's own autofit is switched off, so the saved sizes are the sizes shown
   - Starting sizes (pt): title 28, headline 22, headings 18, lists 16, wins/risks and table 14, footer 12. Today's Northwind deck needed wins/risks at 13 pt and the headline at 21 pt.
5. **Slide 1 without AI:**
   - the headline box reads exactly "AI summary unavailable" (so a check can match it)
   - a gray note replaces both columns: "...written by Claude... Every number on the other slides is computed in Python and is unaffected."
   - The flag count still shows, because it comes from Python.
   - The reason the analysis was rejected goes to the terminal, not the slide.
6. **A saved analysis must match today's data, not only have passed when it was made.** It must be for the same company and latest quarter, and pass `validate_summary` against a freshly built payload. A JSON from last quarter or a hand-edited one never reaches the board.
7. **Slide 2 rows:**
   - 3 context rows: Ending ARR ($K), ARR growth YoY, Gross margin
   - then all 9 flags in FLAG_RULES order, including the combo, so the statuses add up to slide 1's "of 9"
   - **Budget column:** Ending ARR shows "vs budget: 2.5%" rather than the $K budget figure, because a $K budget isn't in the metrics table and would break "every number on slide 2 appears in the metrics table". The other vs-budget comparisons are flag rows with thresholds.
   - **Threshold wording:** "trips below 100.0%" / "trips above 2.00x"
   - **Combo row:** "—" for values and "rule on Risks and flags slide" as its threshold (its settings, 1 pt and 3 quarters, aren't in the metrics table)
   - **Non-flag rows:** status "—", no color
   - **Missing values:** a value with no number shows the same words as Excel: "data missing", "n/a (no prior period)", "∞ (ARR shrank)"
8. **Charts:**
   - only the latest value is labelled, as the chart guidance recommends
   - blank quarters are labelled "data missing"
   - the cash chart title includes runway at next quarter's budgeted burn (context, not a flag)
   - PNGs are kept in `output/charts/` for README screenshots
9. **Slide 4:**
   - The tripped heading says "Tripped flags (7 of 9 flags tripped, 1 cannot evaluate)".
   - I added a "Cannot evaluate" section (you didn't ask for it) so Fernhollow's Rule of 40 isn't silently absent.
   - The combo line spells out its rule with the config values: "trips when NRR falls by at least 1.0 pts and pipeline rises at every step over the last 3 quarters".
   - Data gaps are grouped by the quarters they miss, one bullet per group: "Q1 2025 + Q2 2025: ARR growth QoQ, ...". Flag gaps read "Flag: Rule of 40".
10. **Slide 5:** numbered questions, or the placeholder plus the same note as slide 1.
11. **Footer:** the file name only (not the folder path) and an ISO date (2026-09-17).
12. **File name:** `output/northwind_board_pack.pptx` (lowercase, like `northwind_metrics.xlsx`).
13. **Old decks are deleted first:** the old deck is deleted before a build starts, so a failed build never leaves last run's deck looking current.
14. **main.py:**
    - **Why I touched it:** creating build_deck.py tripped main.py's own "exists but not wired" guard. Every company failed, and check_main.py would have stopped the day runner.
    - **What I changed:** I connected only the deck step. With `--skip-ai`, each company gets a placeholder deck (decision L), and the result reads "OK (AI skipped)".
    - **What I left for Task 3:** the AI step. Without `--skip-ai`, a run still fails loudly ("AI commentary isn't connected to main.py yet - run with --skip-ai").
    - check_main.py and one test in tests/test_main.py were updated to match.
15. **check_deck.py compares slide 2 with the saved metrics workbook** as Excel displays it, not with the deck's own formatting functions. The footer's run date is left out of the slide-2 number check and checked on its own.
16. **Pillow:** text_fit.py imports Pillow (`PIL`) directly. It was already installed as a dependency of matplotlib and python-pptx, so I didn't add it to requirements.txt. Tell me if you want it listed.

### What failed and how I fixed it

All 5 are logged in LEARNINGS.md:
1. **main.py guard:** creating build_deck.py made `main.py --all --skip-ai` fail all 3 companies with NotWiredError. This was expected (the step 5 guard). Fixed by connecting the deck step (decision 14).
2. **Fernhollow's Risks and flags slide didn't fit at 12 pt** after I split data gaps into bullets. The build stopped with the slide and box named. Fixed with two columns. This was the fail-loudly rule doing its job.
3. **Two bugs in my own check:**
   - check_deck.py counted the footer's run date ("09", "17") as invented numbers on slide 2
   - it read the runway context row of the Flags sheet as a 10th flag
   - Both fixed.
4. **Chart labels overlapped** lines, axes and the slide edge. Found by looking at the PNGs, not by a check. Fixed with padding, label placement and a white background behind labels.
5. **A new test proved nothing at first:** "side-by-side columns use the same font sizes" passed on the old code, because nothing needed to shrink. Proved by swapping the old behavior back in, then made stricter.

A few of my shell commands were blocked by the sandbox (heredocs containing braces, `git stash`). I wrote small patch scripts in /tmp and ran them with `.venv/bin/python` instead. To look at the slides, I used a rough matplotlib preview script in /tmp. It isn't part of the project.

### Unresolved

1. **I couldn't open the decks in PowerPoint or Keynote.** There's no slide renderer here.
   - python-pptx reopens every file, and the fit checks pass.
   - But the template is built by editing slide XML (theme, text styles, shapes on the slide master), and only real PowerPoint can confirm it opens cleanly and looks as intended.
   - The fit is also an estimate, not PowerPoint's own text layout. It errs toward "needs more room", but the text could still differ slightly if PowerPoint doesn't have Arial.
   - **Please open `output/northwind_board_pack.pptx` and `output/fernhollow_board_pack.pptx`.**
2. **main.py overwrites the Northwind deck with the placeholder version.** Any `main.py --all --skip-ai` run does this, including the one inside check_main.py. Run `python build_deck.py data/northwind.xlsx` to rebuild it with the AI text. Task 3 wires AI into main.py.
3. **CLAUDE.md's Architecture section is out of date:**
   - it still says build_deck.py is "not built yet"
   - it doesn't list make_template.py, text_fit.py, charts.py or check_deck.py
   - I wasn't allowed to edit it (Task 5 can).
4. **Northwind's saved analysis quotes old flag names** ("burn vs budget", from before Task 1 renamed flags). Its numbers still pass today's validation, so it's used. Fresh wording needs an API run (Task 4).
5. **Chart text isn't size-checked.** Text inside the PNGs is set to 12 pt in charts.py, but check_deck.py can't measure text inside an image.
6. **Slide 2 shows budget variances, not budget $K figures** (decision 7). If you want the $K budget on the slide, the "every number is in the metrics table" rule would need to include the input columns.
7. **STUDY_GUIDE.md isn't updated** for the new files. That's Task 6.

## Task 3: Wiring main.py (AI step + deck)

### Where this started

Nothing for this task existed yet. Task 2 had connected only the deck step, so a run without `--skip-ai` still stopped with "AI commentary isn't connected". I built it in 2 commits:

| Commit | What |
|---|---|
| `f5e5bc7` | `analyze.save_analysis`: one function writes the analysis JSON for both analyze.py and main.py |
| `d09bc80` | main.py runs the AI step and always builds the deck; tests; check_main.py |

Now: **408 tests pass** (397 before, 11 new), and all 5 check scripts print "All checks passed". I didn't call the Anthropic API or edit config.yaml or CLAUDE.md.

**⚠ Read "What failed" item 1 first:** a break test of mine deleted `output/northwind_analysis.json`. I restored Claude's summary word for word, but the file's token counts are lost.

### What was built

**main.py: what happens to each company**
1. Clean, metrics, flags, data gaps, Excel file: unchanged.
2. **AI step** (skipped with `--skip-ai`):
   - deletes the old `output/<company>_analysis.json`
   - builds the payload and calls `analyze.analyze` (first try + one retry, as before)
   - saves `output/<company>_analysis.json` **whether it passed or failed**
3. **Deck:** always built.

| What happened | Deck slides 1 and 5 | Result column | Analysis JSON |
|---|---|---|---|
| AI passed validation | AI text | `OK` | summary, run_info, payload |
| AI failed validation after the retry | "AI summary unavailable" | `OK (AI failed)` | summary `null`, run_info (both attempts and their problems), `error`, payload |
| API error (connection, rate limit, server...) | "AI summary unavailable" | `OK (AI failed)` | summary `null`, `error`, payload |
| `--skip-ai` | "AI summary unavailable" | `OK (AI skipped)` | not written; an existing one is left alone |
| Any other error (a bug) | not built | `FAILED: ...` with a traceback | the old one is already deleted |

- **The terminal says why:** e.g. "✗ AI commentary: failed validation after the retry (output/northwind_analysis.json):" followed by each problem. A pass prints attempts, tokens and seconds.
- **Before any company runs** (not with `--skip-ai`): main.py loads .env and checks for `ANTHROPIC_API_KEY`. If it's missing, main.py prints "ANTHROPIC_API_KEY isn't set: add it to .env, or run with --skip-ai" and exits 1.
- **Under the summary table:** "⚠ AI summary unavailable for 1 company (Fernhollow): their decks show the placeholder; the reasons are in output/<company>_analysis.json".
- **output/batch_summary.csv:** already existed (step 5). Same columns, one row per company. The Result column now reads `OK`, `OK (AI failed)`, `OK (AI skipped)` or `FAILED: ...`.

**analyze.py:** `save_analysis(path, payload, summary, run_info, error)`. analyze.py's own command line uses it too, so both write the same file shape. The only change there is a new `error` field.

**tests/test_main.py (9 new tests).** They run the real Northwind workbook through `run_batch` into a temporary folder, with a **fake Claude client** that returns a fixed answer and counts calls. They prove:
- a passing answer lands on the deck and in the JSON, with 1 call
- an answer with an invented number (555.5%) is called exactly twice (first try + retry), then the deck shows the placeholder, the result reads "OK (AI failed)", and the JSON keeps the problem. An old good analysis in the folder is replaced, not used.
- an API error gives "OK (AI failed)", not a failed company
- a bug inside the AI step fails the company and doesn't leave the old analysis behind
- `--skip-ai` never calls `analyze()`, never looks for a key, and leaves an existing analysis JSON untouched
- a missing key stops `main.py --all` before any company runs
- the result texts, the CSV and the AI-failed warning line

A guard in every test makes creating a real Anthropic client fail the test, so no test can reach the API.

**tests/test_analyze.py (2 new tests):** a saved passing analysis loads in `build_deck.load_analysis`, and a saved failure gives the placeholder.

**check_main.py:**
- Still uses `--skip-ai` for every run.
- The `main.py --all --skip-ai` process now runs with no API key and with `ANTHROPIC_BASE_URL` pointed at a dead local port. Even a bug that called Claude couldn't reach it.
- It also checks that the run didn't write, replace or delete any company's analysis JSON.
- **New:** `check_skip_ai_never_calls_claude` runs all 3 companies into a temporary folder, with `analyze()`, the key check and the Anthropic client each replaced by a function that stops the check if called. It then checks each deck's slide 1 headline reads exactly "AI summary unavailable" and no JSON was saved.
- I proved both parts catch a regression by making main.py ignore `--skip-ai`: the process check failed, and so did the in-process check ("--skip-ai called analyze()").
- **Removed:** "without --skip-ai the run fails loudly" and "if build_deck.py were missing, the deck is skipped". Both behaviors are gone (decisions 1 and 9).

### Decisions I made that you didn't specify

1. **"AI failed" also covers API errors**, not only validation failures. That means any `anthropic.AnthropicError`: connection, rate limit, overload, server error. The numbers on the deck are still valid, which is the reason behind decision K. Any other exception is treated as a bug: the company fails with a traceback, so a coding error can't hide behind "AI failed".
2. **A missing API key stops the whole run up front** (exit 1) instead of giving every company "OK (AI failed)". It's a setup problem, not an AI answer problem, and the SDK reports it as a plain `TypeError` (see What failed, item 2).
3. **"OK (AI failed)" still exits 0.** Exit code 1 stays reserved for a company with no outputs. The warning line and the Result column make the failure visible. Tell me if you'd rather a failed AI step made the run exit 1, e.g. for a scheduled job.
4. **The analysis JSON is saved on failure too**, with `summary: null` and the reason. You can see why Claude's answer was rejected, and `build_deck.py` run later on its own gives the placeholder too. The problems were already in `run_info`; I added an `error` field with the one-line reason.
5. **The old analysis JSON is deleted before the AI step starts.** That matches Task 2's rule for decks: a crash never leaves last run's file looking current. The cost is that a crash loses an analysis that was paid for, which is exactly what happened to me today (What failed, item 1). Tell me if you'd rather keep the old file until the new one is saved.
6. **`--skip-ai` leaves an existing analysis JSON alone** and still builds the placeholder deck, as the task says. `python build_deck.py data/northwind.xlsx` can still put the saved AI text on a deck.
7. **"OK" means the AI text is actually on the deck.** The deck checks the analysis again (Task 2 decision 6). If a passing analysis were ever rejected there, the result reads "OK (AI failed)" and the terminal prints the deck's reason. This shouldn't happen, since both checks run on the same numbers seconds apart.
8. **The CSV columns didn't change.** The Result column carries the AI status. There's no column for why the AI failed: that's in the terminal and the JSON.
9. **Removed the "not connected yet" guard and the "build_deck.py missing" skip** (`NotWiredError`, `BUILD_DECK_PATH`, "OK (AI + deck skipped)"). main.py imports build_deck.py at the top, so if that file were missing main.py couldn't even start: the skip could never happen.
10. **`run_batch` and `run_company` gained `client` and `output_dir` arguments,** used only by tests and checks: a fake Claude client and a temporary folder. The command line always uses the real client and `output/`.
11. **No `--model` option on main.py.** It uses analyze.py's default, claude-sonnet-5, the recommendation from the model comparison.
12. **Per-company terminal line on a pass:** attempts, input/output tokens and seconds. No cost in dollars: Task 4 records that.
13. **Paths inside the project print as `output/...`; others print in full** (`shown_path`). A temporary folder isn't inside the project, and the old `relative_to` call would have crashed on it.
14. **The AI-on paths are proven in pytest with a fake client, not in check_main.py.** The task says checks must use `--skip-ai`.

### What failed and how I fixed it

All 3 are logged in LEARNINGS.md.

1. **I deleted the real `output/northwind_analysis.json`.**
   - **How:** to prove check_main.py catches a `--skip-ai` regression, I temporarily made main.py ignore `--skip-ai` and ran the check. That run used the real output/ folder. The AI step deleted the old analysis first (decision 5), then stopped on the missing key. No API call was made. main.py was restored right away (confirmed identical).
   - **Noticed when:** check_deck.py failed with "No such file". output/ isn't in git, and the v1/v2 files are older prompt versions that fail today's length limits.
   - **Restored:**
     - Claude's summary: recovered word for word from a Task 2 session transcript that had printed it. It passes today's validation.
     - Payload: rebuilt from today's workbook.
     - `run_info` (tokens, seconds): lost, so it's `null`.
     - A `restored_note` field in the file says all of this.
   - **Result:** check_deck.py passes again. Slide 1 shows the same headline ("Runway has fallen to 11.0 mo against a 12.0 mo threshold...").
   - **Lesson:** deliberate breaks go into a temporary folder. My second break test did.
2. **A missing API key isn't an Anthropic error.** I tried it against a dead local address before writing the code. The SDK raises a plain `TypeError`, and only when the request is sent, so every company would have shown as a code bug. Fixed by checking the key before the batch (decision 2).
3. **Two of my new tests were wrong at first.**
   - The "no old analysis left behind" test passed even with the delete line removed, because there was no old file. It now writes one first; proved by removing the line.
   - A test looked for "987654.3", but the validator prints it as "987654". Changed the number to 555.5.

Also: the sandbox blocked a few shell commands (a heredoc with braces, a `for` loop, `sed`, and `env -u`). I used the Edit tool, one command per script, or small scripts in /tmp run with `.venv/bin/python`. No effect on the code.

### Unresolved

1. **`output/northwind_analysis.json` is a restoration** (What failed, item 1). The AI text is genuine and validated, but its token counts are gone and its payload was rebuilt. Task 4's live run replaces it.
2. **The real Claude call through main.py has never run.** Everything AI-on is proven with a fake client that returns the same shape as `client.messages.parse`. Task 4's live run is the first real test. If a real response differs from the fake (e.g. a new stop reason), the existing checks in `analyze.call_claude` still apply.
3. **Running the checks overwrites real outputs** (this was already true before this task):
   - check_main.py's last `main.py` run is the broken-workbook test, so `output/batch_summary.csv` ends up holding one "Broken" row
   - every deck in output/ gets the placeholder
   - Run `python main.py --all` (or `--skip-ai`) yourself afterwards for the real CSV. Moving check outputs to a temporary folder would need an `--output-dir` option on main.py; tell me if you want it.
4. **Companies run one after another.** At ~36 s per Sonnet call (model comparison), 275 companies would take about 2.75 hours. Running them in parallel is a later step.
5. **Exit code on "OK (AI failed)"** (decision 3) and **delete-first** (decision 5) are judgment calls worth a look.

## Task 4: One live run (`main.py --all` with the real API)

### Where this started

Nothing for this task existed yet. Task 3's checks had all passed (408 tests, 5 check scripts). **One live run was made: all 3 companies passed, so there was no second run.** I changed no code: this run found no bugs outside the API step to fix. I didn't edit config.yaml or CLAUDE.md.

### Result

```
Company     Flags tripped              Data gaps                          Result
Alderpeak   0 of 9                     none                               OK
Fernhollow  7 of 9, 1 cannot evaluate  20 metrics/flags (blank: Q2 2025)  OK
Northwind   6 of 9                     19 metrics/flags (blank: Q1 2025)  OK
3 of 3 companies succeeded      (exit code 0)
```

**What I checked myself, not just the terminal output:**
- **Each `output/<company>_analysis.json`:** `run_info.passed` is true, `error` is null, 1 attempt, stop reason `end_turn`, model claude-sonnet-5.
- **I re-ran `validate_summary`** on each saved summary against its saved payload: 0 problems for all 3.
- **Each deck has 5 slides.** Slide 1 shows Claude's headline, not the placeholder.
- **check_deck.py passes:** for all 3 companies it says "AI text from the JSON". It's the first time Alderpeak and Fernhollow have had real analyses. Both tampered Northwind analyses still get the placeholder.
- **The flag counts match each company's story:** Alderpeak 0, Fernhollow 7 plus 1 cannot evaluate, Northwind 6.

**Headlines:**
- **Alderpeak:** "All nine flags passed in Q2 2026, with ARR reaching $18,450K and 108.0 mo of runway, though YoY growth eased to 47.5%."
- **Fernhollow:** "Fernhollow's growth has stalled and cash runway has shrunk to 6.0 mo, with 7 of 9 flags tripped, including burn and retention metrics."
- **Northwind:** "Northwind's ARR growth remains strong at 42.8% YoY, but burn multiple 2.35x, runway 11.0 mo, and NRR 97.1% signal deteriorating unit economics."

### Tokens and cost (also in LEARNINGS.md, "Live runs")

| Company | Attempts | Input tokens | Output tokens | Seconds | Cost |
|---|---|---|---|---|---|
| Alderpeak | 1 | 5,335 | 2,368 | 22.5 | $0.0343 |
| Fernhollow | 1 | 5,956 | 4,692 | 42.6 | $0.0588 |
| Northwind | 1 | 5,915 | 3,472 | 33.1 | $0.0466 |
| **Total** | 3 | **17,206** | **10,532** | 98.1 | **$0.1397** |

- **Total: 27,738 tokens, $0.14.** On average that's $0.0466 per company, or **$12.81 per quarter for 275 companies**. The README's model comparison said $14.65.
- **Prices:** Sonnet 5 at $2 input / $10 output per million tokens, from `compare_models.PRICES` (as of 2026-06-24).
- **Time:** about 33 s of API time per company, or about 2.5 hours for 275 companies run one after another.

### Decisions I made that you didn't specify

1. **I saved a copy of the restored Northwind analysis before the run** (`output/northwind_analysis_restored_task3.json`). The run deletes the old analysis first (Task 3 decision 5). If the live run had failed, you'd still have the only validated Northwind text. The run replaced it with a real one, so the copy can be deleted.
2. **No second run.** The task allows one more run after a fix, and there was nothing to fix. A second run only to get "better" text would cost money and prove nothing new about the code.
3. **Cost uses the price table already in compare_models.py,** so the numbers are comparable with the README's comparison. I didn't check a newer pricing page.
4. **Cost for 275 companies = the average of the 3 companies × 275.** A real portfolio's mix of healthy and distressed companies changes this: the distressed company used about 2x the healthy one's output tokens.
5. **I kept a copy of this run's terminal output and CSV** (`output/day_logs/task4_live_run.txt`, `output/day_logs/task4_batch_summary.csv`). check_main.py runs after this task and overwrites `output/batch_summary.csv` and the decks (see Unresolved 1).
6. **I read all 3 summaries against their payloads line by line,** not just the validator's pass. That's the model comparison's lesson. Findings are under What failed, item 2.
7. **I didn't change the prompt or validator for those findings.** The task says to fix only non-API bugs, and checking a prompt change needs more paid runs.

### What failed and how I fixed it

Both are logged in LEARNINGS.md.

1. **My first run command was blocked before it started.**
   - **What happened:** I added `; echo "EXIT CODE: $?"` after the run, and the permission rules only allow commands that start with `.venv/bin/python`.
   - **Proof nothing ran:** before trying again, I checked the timestamps. `output/northwind_analysis.json` and `batch_summary.csv` were unchanged, so no API call had been made.
   - **Fix:** ran `.venv/bin/python main.py --all` on its own. That was run 1. The exit code (0) came from the background task's status.
2. **Claude's text passed validation but has 3 false or misleading claims** (not fixed):
   - **Alderpeak, win 1:** YoY growth of 47.5% is "well above prior-period levels". **False:** it fell every quarter (53.9% → 51.5% → 49.5% → 47.5%), and Claude's own risk 1 says so.
   - **Alderpeak, risk 3:** NRR has a "persistent gentle decline" from 110.0% to 109.8%. **False:** it goes up and down (110.2%, 109.5%, 110.0%, 109.6%, 109.8%, 109.7%, 109.8%) and rose last quarter. That's the story, which says NRR never falls at every step.
   - **Fernhollow, win 3:** the combo flag passing is presented as good news. **Misleading:** it passes because pipeline is falling too (4,200 → 2,500 $K), so the company has a sales problem and a retention problem. Pipeline isn't mentioned anywhere in Fernhollow's text.
   - **Smaller:** Alderpeak calls a 0.4-point Rule of 40 dip (43.1% → 42.7%) a "slide" and asks how to "arrest" it. That's the same overstatement the v1 → v2 prompt rule was meant to stop.
   - **Why the checks missed them:** the validator checks the shape, lengths and that every number appears in the payload. It can't check direction or whether something is good news. This is the "wrong direction" gap already logged in the model comparison.
   - **Prompt-rule slips (known from the v3 notes, still happening):**
     - **"Describe a trend from its peak or lookback window, not the first quarter":** Alderpeak and Fernhollow break it several times, e.g. "NRR fell from 98.0% in Q3 2024" and "burn multiple improved from 0.70x (Q3 2024)". Northwind's answer follows it.
     - **"When a flag passed, say so":** Northwind question 3 asks what is lengthening CAC payback (18.9 → 20.7 mo) without saying the flag passed. Its win 3 does say it.
     - **Fixed since v3:** Northwind's pipeline now has $K.

### Unresolved

1. **Running the check scripts after this task replaces the AI decks with placeholders.** check_main.py runs `main.py --all --skip-ai` into output/ (Task 3, Unresolved 3), and it also overwrites `batch_summary.csv`. The analysis JSONs survive, because check_main.py checks that they aren't touched.
   - **To get the AI decks back without an API call:** `python build_deck.py data/alderpeak.xlsx`, and the same for fernhollow and northwind.
   - **To get the CSV back:** copy `output/day_logs/task4_batch_summary.csv`.
   - **Proper fix:** an `--output-dir` option so checks write to a temporary folder. It needs your go-ahead.
2. **Direction and "is this really good news" claims aren't validated** (What failed, item 2). Options, cheapest first:
   - **(a) Prompt rules:** "only call a trend rising or falling if every step in the window moves that way" and "a passing flag is not a win unless its value is good in itself". Needs paid runs to confirm.
   - **(b) Code check:** a validator check that finds "rose/fell/declined/improved … from X to Y" and compares X and Y. Direction words next to two numbers can be tested without the API.
   - **(c) Human review:** keep a person reading each claim before a deck goes to a board. It's the only thing that caught these today.
3. **Only 1 run per company.** 3 first-attempt passes don't prove the retry rate at 275 companies. The model comparison's Sonnet runs didn't need retries either, so the retry path has only run for real with Haiku.
4. **`output/northwind_analysis_restored_task3.json`** can be deleted now that a real analysis exists (output/ isn't in git).

## Task 5: Docs (README.md, CLAUDE.md Architecture)

### Where this started

Nothing for this task existed yet: README.md was the one-line stub plus the model comparison block. No API calls were made. I didn't edit config.yaml or any pipeline code.

### What I built

1. **`tests/test_docs.py`** (4 tests, written first):
   - The README keeps exactly one model comparison block, start marker before end marker.
   - Rewriting that block with `compare_models.write_readme_section` (on a copy in a temporary folder) leaves everything above and below it unchanged. So re-running the model comparison can't wipe the new README.
   - Every `.py` file named in the README, and in CLAUDE.md's Architecture section, exists. I proved the check catches a made-up `ghost.py`.
2. **README.md, rewritten:**
   - **What it does:** 3 sentences.
   - **How to run it:** setup, the 3 main.py commands, a sample summary table, the output files, the other commands, and the 5 checks + pytest.
   - **Data flow:** a diagram from the workbook to the deck.
   - **Key design decisions:** 11 of them, each with its reason.
   - **Model comparison:** the existing block, unchanged. I checked it's identical character for character.
   - **Cost:** per company, and for 275 companies per quarter.
   - **Screenshots:** 11 placeholders.
   - **Next steps.**
3. **CLAUDE.md, Architecture section only:**
   - build_deck.py is no longer "not built yet".
   - Added make_template.py, charts.py, text_fit.py, check_deck.py, README.md, LEARNINGS.md and DAY_REPORT.md.
   - main.py's line now shows its full chain.

**Checks:** 412 tests pass (408 + 4 new). All 5 check scripts print "All checks passed".

### Decisions I made that you didn't specify

1. **6 of the 8 Next steps are already built, so the README marks them done.** They are the combo minimum drop, lookback guard, budget row check, two-KPI-tab check, same-quarter warning and sign-aware number check. Each one is in the code with tests: LEARNINGS rows D, E, F, G, N and O, from the hardening pass. Listing them as future work would say the code lacks things it has. The README shows them as `[x]`, each with the function that does it. Only the Power Automate or SharePoint trigger and the portfolio rollup are `[ ]`. **Tell me if you meant something different by any of the 6.**
2. **I added 3 open items to Next steps from earlier reports.** They are: checking the direction of trend claims (Task 4), an `--output-dir` option (Tasks 3 and 4), and running companies in parallel (Task 3). They're real gaps a reader should know about. Delete them if you want only your list.
3. **I didn't add a new stories section to CLAUDE.md.** "Company stories" already has the Alderpeak and Fernhollow stories (added in the hardening pass). I checked them against today's data (Fernhollow NRR 77.9%, GRR 74.7%, runway 6.0 mo, Rule of 40 "cannot evaluate") and against Task 4's flag counts. They match, so I left them word for word.
4. **Where the cost numbers come from:** Task 4's live run. $0.0466 per company is the average of the 3 real costs. $12.81 for 275 companies is that average × 275. The range is $9.45 (all like Alderpeak) to $16.18 (all like Fernhollow). Prices come from `compare_models.PRICES` (as of 2026-06-24). I worked these out with a Python one-liner, not by hand.
5. **I suggest a budget of $13–15 per quarter**, between the live run ($12.81) and the model comparison ($14.65), because a retry doubles one company's cost.
6. **Screenshot placeholders are text lines (📸), not image links.** A link to an image that doesn't exist shows as a broken icon on GitHub. Each line names the file, slide and what should be visible. I added two that show the flags aren't hard-coded: Alderpeak slide 1 ("0 of 9") and Fernhollow slide 4.
7. **The README tells you to run `build_deck.py` before taking deck screenshots.** check_main.py leaves placeholder decks behind. I ran it for all 3 companies at the end of this task, so the decks in output/ have the AI text right now (I checked Northwind's slide 1). The next check run will replace it again.
8. **I left the Python version (3.14) and "400+ tests" in the README.** An exact test count goes out of date with every new test.
9. **One README heading (Model comparison) is written by compare_models.py.** I placed the block after the design decisions, and the table of contents links to it.
10. **CLAUDE.md, outside the sections I was allowed to edit:** the heading "Step 4 decisions (build_deck.py, not built yet)" and "Should trip 4+ flags" in the Northwind story (it trips 6) are now out of date. I didn't change them.

### What failed and how I fixed it

All 3 are logged in LEARNINGS.md.

1. **Your Next steps list included 6 finished items** (decision 1). I caught this by checking each item in the code before writing.
2. **My first README draft blamed `check_deck.py` as well as `check_main.py` for placeholder decks.**
   - **Caught:** when I ran the checks myself. Only check_main.py does it.
   - **Fixed:** corrected the note. I also proved that `build_deck.py` restores the AI text.
3. **Three shell commands were refused before running.**
   - **Which:** two `for` loops using `$f`, and one ending in `awk`.
   - **Workaround:** ran one command per script instead. Nothing ran, and there was no effect on the code.

### Unresolved

1. **Screenshots aren't taken.** Placeholders only: capturing them needs you (PowerPoint and Excel open on screen).
2. **CLAUDE.md, outside this task's sections, has 2 lines that are out of date** (decision 10).
3. **Cost and time come from a single live run** of 3 companies (Task 4, Unresolved 3). The 275-company numbers are a projection.
4. **The README's Next steps say "a person still reads each deck before it goes to a board"** because direction claims aren't validated (Task 4, Unresolved 2). That stays true until the direction check exists.

## Review: everything committed today (`git diff main..HEAD`)

### What I reviewed and how

- **Scope:** 17 commits, `df6de63` to `9bd1951`. 24 files, +3,320 / −113 lines.
- **Read in full:** build_deck.py, main.py, charts.py, text_fit.py, make_template.py.
- **Read the diffs of:** analyze.py (`save_analysis`), metrics.py (the two blank-input fixes), check_main.py, tests/test_excel_output.py.
- **Skimmed:** check_deck.py, and every new test file's test names.
- **Ran again:** 412 tests pass. All 5 check scripts print "All checks passed".
- **Probes:** 4 small scripts in /tmp that try one thing each. None calls the API, and none writes to output/ (they use temporary folders). They aren't part of the project.

### Findings, most serious first

**1. A valid AI answer that's a little longer makes the whole company FAIL, with no deck. (Medium-high: your decision)**
- **What happens:** the validator allows each win and risk detail up to 45 words. Slide 1 has room for less than that. When the text doesn't fit at 12 pt, text_fit.py stops the build (as designed). But in main.py that stop happens inside the deck step, so:
  - the result is `FAILED: TextDoesNotFitError: Slide 1, Risks: ...`, with a traceback
  - no deck is saved (the old one was already deleted)
  - the analysis JSON is saved as **passed**
- **Proof (fake Claude client, temporary folder):** I took today's real Northwind answer and added "this quarter" to each of the 6 win and risk details. The longest detail was 43 words, still under the 45-word limit, and main.py reported Northwind as FAILED.
- **How close today's real answers are** (the same extra words added to every detail until slide 1 breaks):

  | Company | Wins/risks font today | Breaks with |
  |---|---|---|
  | Northwind | 12 pt (the floor) | +1 word per detail |
  | Fernhollow | 12 pt (the floor) | +3 words per detail |
  | Alderpeak | 13 pt | +4 words per detail |

  So the next live run could fail a company just because Claude wrote a bit more. Task 2's report said Northwind needed 13 pt. That was the old restored text. The live-run text is longer and already needs 12 pt.
- **Why no check caught it:**
  - the tests use short fake answers
  - check_deck.py uses the saved answers, which happen to fit
  - the 45-word limit was set in step 3, before the slides existed
- **Why this conflicts with the rules:** decision K says the deck is built whenever the numbers are valid. Here both the numbers and the AI text are valid, and there's no deck.
- **Options:**
  - **(a) Placeholder fallback:** if the AI text doesn't fit slide 1 or 5, build the deck with "AI summary unavailable". The result reads `OK (AI failed)`, with the reason "AI text doesn't fit slide 1".
  - **(b) Fit check in the validator:** `validate_summary` also measures the text against slide 1's boxes. A too-long answer then uses the existing retry, with "too long for the slide" as the problem. It costs nothing unless it happens.
  - **(c) Tighter limits or more room:** a lower word limit in the prompt (needs paid runs to confirm), or more room for the columns (e.g. a shorter headline box).
- **My recommendation:** (b) plus (a). The retry fixes most cases, and the placeholder guarantees the deck.
- **Not fixed:** each option changes what the AI step does or what a result means, so it's your call.

**2. A text-fit stop prints a traceback as if it were a code bug. (Low)**
- `TextDoesNotFitError` isn't in main.py's `INPUT_ERRORS`, so it gets the full traceback that's meant for bugs.
- The message already names the slide and the box, so the traceback adds nothing.
- It goes with finding 1: fixing that removes most of the cases.

**3. `shrink_to_fit` accepts text that already starts below 12 pt. (Low, can't happen today)**
- **Proof:** `shrink_to_fit([paragraph('short', 10)], ...)` returns 10 pt without an error.
- **Why it matters:** `add_columns` shrinks both columns by the bigger column's shrink. If the two columns started with different smallest sizes, one could be pushed below 12 pt, and nothing would stop it.
- **Why it can't happen today:** on both slides that use columns (1 and 4), the two columns use the same sizes. check_deck.py also checks every saved font is at least 12 pt.
- **Fix, if you want it:** one line that stops when the smallest size is already below `MIN_FONT_PT`.

**4. A hand-edited analysis JSON can crash the deck step instead of giving a reason. (Low)**
- **Examples:** a `payload` that isn't an object, or a file that isn't UTF-8 text. `load_analysis` raises an error instead of returning "why not".
- **Effect:** in main.py that company fails.
- **How likely:** only a hand-edited file can do this. `save_analysis` always writes a valid shape.

**5. The template script depends on python-pptx internals, and no package version is pinned. (Low)**
- **Internals:** make_template.py uses python-pptx's private attributes: `theme_part._blob` and `shapes._spTree`.
- **No pins:** requirements.txt pins no versions, and Pillow isn't listed (Task 2 decision 16).
- **Risk:** a python-pptx upgrade could break re-running make_template.py.
- **Limit of the risk:** templates/base.pptx is committed, so decks keep building either way.

**6. Each workbook is read 3 times per company. (Low, efficiency)**
- `run_company`, `save_metrics_workbook` and `save_deck` each clean it again and recompute the metrics. The payload is also built twice: once for Claude, once to re-check the answer.
- **Cost:** well under a second per company, next to about 33 s of API time.
- **Verdict:** not worth changing now.

**7. Chart PNGs aren't deleted before a build. (Very low)**
- Decks and analyses are deleted first; the images in `output/charts/` aren't.
- A failed build can leave last run's PNGs. They're only used for README screenshots.

**8. Docs out of date**
- **check_deck.py's docstring** said Alderpeak and Fernhollow have no analysis. They do since Task 4. **Fixed** (docstring only).
- **STUDY_GUIDE.md** still described main.py before the deck existed (`NotWiredError`, "AI + deck skipped"). **Updated in this task.**
- **CLAUDE.md** still has the 2 out-of-date lines from Task 5 decision 10. Not edited (this task doesn't allow it).

### What I checked and found right

- **No typed numbers:** a test reads build_deck.py and charts.py and fails on any text in the code (docstrings aside) that contains a digit. Every number comes through `metrics.format_value`, so ratios become % only at output.
- **One label set:** the deck uses `METRIC_LABELS` and `INPUT_LABELS` from metrics.py, and Excel's status and gap labels. No names are typed a second time.
- **A stale or edited analysis can't reach a slide:** `load_analysis` re-checks it against a payload rebuilt from today's workbook. check_deck.py proves it with a changed number and a wrong quarter.
- **The AI failure paths are sorted correctly in main.py:**
  - failed validation twice, or an API error → deck with the placeholder, `OK (AI failed)`
  - any other error → the company fails with a traceback
  - a missing key → the run stops before any company starts
- **`--skip-ai` can't reach Claude:** check_main.py runs it with no key and a dead API address, and also swaps in stand-ins that stop the check if `analyze()`, the key check or the client is ever called.
- **Blank inputs win:** the two metrics.py fixes put missing input first in both runway-at-budget and the combo rule, the same order as `metric_reasons`.
- **Delete first:** the old deck and the old analysis are deleted before a rebuild, so a crash never leaves last run's file looking current.
- **No secrets or real data:** `.env` isn't tracked (only `.env.example`), and every company and number is fictional.
