# Board Pack Generator

## Goal
Messy portfolio-company KPI workbook (.xlsx) → 4-slide board update deck (.pptx) + AI summary.
Demo project for a PE AI automation role. Must be clean, explainable, and reliable.

## About me
- Finance background, learning Python. Explain every file after writing it, in plain English.
- Work in small steps. Use plan mode before writing code. Stop after each build step so I can review.

## Environment
- macOS, Python venv at .venv (activate: `source .venv/bin/activate`)
- Packages in requirements.txt. Don't add packages without telling me why.

## Architecture
- data/                  fake input workbooks (3 made-up software companies)
- make_data.py           Northwind's answer key (TRUE_DATA) and mess settings -> data/northwind.xlsx
- make_data_alderpeak.py, make_data_fernhollow.py   the healthy and distressed companies, same shape
- make_data_common.py    shared checking and writing code for the three make_data scripts
- clean.py               normalize headers, parse "$1.2M"-style text, handle blanks; stops (with sheet/row/cell) on anything it can't read with certainty
- metrics.py             every metric, why a value has no number (reasons), flags, data gaps, and the ONE label set (METRIC_LABELS, INPUT_LABELS)
- analyze.py             Claude API → JSON (headline, 3 wins, 3 risks, 3 mgmt questions); validate schema + sign-aware number check + slide-fit check + direction check (a direction word its numbers contradict, a "persistent" trend that isn't), retry once on failure
- compare_models.py      step 3b: blind Sonnet vs Haiku comparison -> README table
- excel_output.py        step 4b: output/<company>_metrics.xlsx (Metrics, Flags, Data gaps sheets)
- make_template.py       builds templates/base.pptx: fictional "Example Capital" brand, navy/gray, 16:9, title + content layouts
- build_deck.py          step 4: output/<company>_board_pack.pptx on templates/base.pptx; slides: Key metrics, ARR and cash charts, Risks and flags, AI commentary (headline + 3 risks + 3 questions under "AI-drafted from computed metrics - review before use"; wins aren't shown). Re-checks the analysis JSON; "AI summary unavailable" if missing, failed, or too long for slide 4 (ai_text_problems) - a company is never left without a deck. Footer carries the git commit and model; every slide is watermarked "DRAFT - NOT REVIEWED" until approve.py records a reviewer
- charts.py              the two matplotlib charts for slide 2 (ARR + net new ARR, ending cash with runway); a blank quarter is a visible gap
- text_fit.py            measures slide text, shrinks it to a 12 pt floor, then stops naming the slide and box
- provenance.py          run manifests: input/config SHA-256 hashes, git commit, model + prompt version, tokens, cost; and whether a human approved this deck (approval_status is the only judge)
- approve.py             `python approve.py northwind [--reviewer NAME]`: records reviewer + time in the manifest, so rebuilding drops the DRAFT watermark. Never builds a deck
- main.py                CLI: `python main.py data/northwind.xlsx` or `--all` [--skip-ai]; clean -> metrics -> Excel -> AI JSON -> deck per company; summary table + output/batch_summary.csv
- app.py                 Streamlit web page: drag in an xlsx, see flags + metrics table in the Excel colors, download deck + metrics workbook; "Include AI commentary" checkbox labelled with the typical cost (reuses a saved analysis of the same numbers for free); errors in plain words, never a traceback; builds in a temp folder, never touches output/
- run_app.command        double-click on a Mac: sets up .venv on first run, starts app.py, opens the browser (.streamlit/config.toml: headless, no tracebacks on the page)
- config.yaml            flag thresholds with investor reasoning in comments (NRR < 100%, burn > 15% over budget, runway < 12 mo)
- tests/ + pytest.ini    pytest unit tests (`python -m pytest -q`), expected values worked out by hand; tests/test_docs.py keeps README/CLAUDE.md file names real
- check_northwind.py, check_companies.py, check_excel_output.py, check_deck.py, check_main.py   end-to-end proofs against each company's answer key (no API calls)
- README.md              what it does, how to run it, data flow, design decisions, model comparison (compare_models.py rewrites the block between its marker comments), cost, screenshots, next steps
- STUDY_GUIDE.md         data flow, glossary, interview questions, exercises
- LEARNINGS.md           what broke and what it taught; model comparison and live-run costs
- DAY_REPORT.md          per-task report of the day's build (decisions, failures, unresolved); OVERNIGHT_REPORT.md is a historical record
- output/                generated files (git-ignored)

## Rules
- Python computes every number. Claude only interprets computed metrics — never does math.
- All ratios are stored as decimals; formatting to % happens only at output.
- API key lives in .env (never committed). Load with python-dotenv.
- Fictional data only. No real company names or numbers.
- Keep functions small and commented. I need to explain every line.
- Log anything that breaks in LEARNINGS.md.
- One label set: metric and input display names live only in metrics.py. Flag names ARE the metric labels, so every output (printout, Excel, Claude's payload, deck) uses the same words.

## Input columns (per quarter, $K)
starting_arr, new_arr, expansion_arr, contraction_arr, churned_arr, revenue, gross_profit, net_burn, ending_cash, sm_spend, new_customers, headcount, pipeline, budget_new_arr, budget_arr, budget_net_burn
- This list supersedes the column list in Build order step 1. NRR and gross margin are computed in metrics.py, not stored as inputs.
- After the actual quarters, the workbook has one budget-only row for next quarter (e.g. "Q3 2026 (Budget)") with only the 3 budget columns filled. It feeds "runway at next quarter's budgeted burn" and is a forecast, not a blank quarter.

## Metric definitions (put these in metrics.py docstrings later)
- NRR (annualized) = 1 + 4 * (expansion - contraction - churn) / starting_arr
- GRR (annualized) = 1 - 4 * (contraction + churn) / starting_arr
  - Annualized so thresholds keep their usual annual meaning; uses one quarter's data only. Deck labels them "annualized".
- Net new ARR = new + expansion - contraction - churn
- Ending ARR = starting_arr + net new ARR
- Net new ARR vs budget = net new ARR / (budget_arr this quarter - budget_arr last quarter) - 1
  - Both sides are net of expansion and churn. Needs the prior quarter, so it follows the QoQ gap rule.
- ARR vs budget = ending ARR / budget_arr - 1 (budget_arr is budgeted *ending* ARR)
- Burn vs budget = net_burn / budget_net_burn - 1
  - Not meaningful when budget_net_burn <= 0: the flag doesn't trip, and outputs show the $K figures instead ("n/m: net burn 1,650 vs budget 0 ($K)").
- Net new ARR vs budget is also not meaningful when budgeted net new ARR <= 0 (same handling).
- Burn multiple = net_burn / net new ARR
- Rule of 40 = YoY revenue growth % + FCF margin, where FCF margin = -net_burn / revenue
- CAC payback (months) = sm_spend / (new_arr * gross margin) * 12
- Runway (months) = ending_cash / (net_burn / 3), at current burn and at next quarter's budgeted burn
  - The runway flag uses current burn. Runway at next quarter's budgeted burn is shown as context, not flagged.
- Flags are evaluated on the latest actual quarter (functions accept any quarter, for step 4b).
- Edge cases apply ONLY when every input is present (a blank input is always "missing input", never ∞ or 0):
  - Burn multiple when net new ARR <= 0 and net_burn > 0 -> infinite, trips ("ARR shrank")
  - Burn multiple when net_burn <= 0 -> 0, passes ("not burning")
  - Runway when net_burn <= 0 -> infinite, passes ("not burning")
  - CAC payback when new_arr * gross margin <= 0 -> infinite, trips
  - Any other infinity (e.g. growth from a zero base) is "not meaningful"
- Why a value has no number: exactly one of three reasons (metrics.py `metric_reasons`), and they must never look alike:
  - missing input: an input the metric uses (METRIC_INPUTS) is blank -> "data missing" (a data gap)
  - no prior period: it needs an earlier quarter the workbook doesn't have -> "n/a (no prior period)"
  - not meaningful: every input is there but the math is undefined (0 / 0, the budget rules above) -> "n/m ..."
  - A blank input wins: a blank quarter's own YoY is "missing input", not "no prior period".
- Flags: trip / pass / cannot evaluate. "Cannot evaluate" always carries its reason ("cannot evaluate — missing input").
- Combo rule (retention problem) = NRR falling AND pipeline rising over the last `combo_lookback_quarters` quarters (config.yaml)
  - Window = the last N quarters, including the latest (N=3 means 2 quarter-over-quarter comparisons). N must be at least 2, or config loading stops.
  - Falling = NRR fell by at least `combo_min_nrr_drop` (0.01 = 1 point) in every comparison in the window: a 0.1-point dip is noise
  - Rising = pipeline increased in every quarter-over-quarter comparison in the window
  - Not enough history -> "cannot evaluate — no prior period"; a blank NRR or pipeline input -> "cannot evaluate — missing input"; never False

## Messy data rules
- Parse "$1.2M" / "850K" text to numbers; normalize header names
- Blank quarter stays blank and is flagged "data missing" on the deck. Never impute numbers that could reach a board.
- Missing values spread to every metric that uses them. Any metric whose inputs include the blank quarter shows "data missing":
  - QoQ (and net new ARR vs budget): the blank quarter and the next quarter
  - YoY: the blank quarter and the quarter 4 quarters later
- Only missing input counts: a metric is a data gap only when one of ITS inputs is blank. A partly blank quarter doesn't make unrelated metrics gaps.
- Any flag or combo rule that depends on a missing value returns "cannot evaluate — missing input" instead of pass/fail
- The Risks/Flags slide gets a "Data gaps" line listing every metric and flag affected (missing input only; "no prior period" and "not meaningful" are not gaps)
- Stop, don't guess (clean.py):
  - The budget-only row must be labelled with the quarter right after the last actual quarter
  - Two or more tabs with a "Quarter" header -> stop and name them
  - Footnote rows under the table -> stop (a guessed "note" could hide a real quarter)

## Company stories (each make_data script must produce its story)

### Northwind story (make_data.py)
8 quarters. ARR grows strongly. NRR slides 108% -> 97% over last 3 quarters. Pipeline keeps rising. Burn ~20% over budget. Runway ends ~11 months. Should trip 4+ flags.

### Alderpeak story (make_data_alderpeak.py): healthy, no flags in any quarter
ARR grows ~47% a year; NRR holds ~110%, GRR ~95%. Burn shrinks every quarter and stays under budget, so runway is long. Rule of 40 stays above 40% once a year of history exists. NRR moves up and down (never falls at every step) while pipeline rises, so the combo passes. No blank quarter (zero data gaps). Mess differs from Northwind: other header spellings, the Notes tab comes BEFORE the KPI tab, budget row "Q3 2026 Plan".

### Fernhollow story (make_data_fernhollow.py): distressed, 7 of 9 flags trip
ARR stalls then shrinks (net new ARR negative from Q1 2026, so burn multiple is ∞ "ARR shrank"). Churn and contraction climb: NRR ~78%, GRR ~75%. Burn rises ~20% over budget; runway 6 months. New sales dry up, so CAC payback is over 10 years. Pipeline FALLS too, so the combo passes (a sales AND retention problem). Q2 2025 is blank, exactly 4 quarters before Q2 2026, so Rule of 40 is "cannot evaluate — missing input" in the latest quarter. Mess: a title line and empty row above the header, budget columns next to their actuals, budget row "Q3 2026 - Bud".

The three companies together prove flags aren't hard-coded.

## Build order
1. make_data.py: generate fake messy workbook for "Northwind Software" (8 quarters: ARR, new ARR, churned ARR, NRR, gross margin, burn, cash, headcount, pipeline + budget columns; inconsistent headers, "$1.2M" text, one blank quarter, junk notes tab)
2. clean.py + metrics.py — verify outputs against manual Excel math
3. analyze.py
4. build_deck.py + charts
5. main.py batch mode + 2 more fake companies
6. README with before/after screenshots

## Step 4 decisions (build_deck.py, not built yet)
- K. If the AI summary fails after its retry, still build the deck: the numbers are valid. The AI commentary slide says "AI summary unavailable" and the batch result reads "OK (AI failed)".
- L. `--skip-ai` builds the deck without AI text, using the same "AI summary unavailable" slide.

## Added build steps
- 3b. Model comparison: run analyze.py on the same metrics with claude-sonnet-5 and claude-haiku-4-5. Log cost, latency, schema pass/fail, and my 1-5 quality score to a table in README with a one-line recommendation.
- 4b. Excel output: write output/<company>_metrics.xlsx with computed metrics and flagged cells highlighted.
- Feedback round after first full deck: log "reviewer said X -> changed Y" in LEARNINGS.md.
