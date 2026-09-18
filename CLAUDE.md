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
- eval/                  the evaluation set: eval/make_eval_data.py writes 12 edge-case companies to eval/data/ (exactly at every threshold, blank first/second-to-last/last quarter, zero revenue, negative budget, 2 quarters of history, a repeated and a missing row that must stop), each with a hand-typed answer key; eval/run_eval.py scores clean, metrics, flags and gaps against them and names every mismatch
- clean.py               normalize headers, parse "$1.2M"-style text, handle blanks; stops (with sheet/row/cell) on anything it can't read with certainty
- mapping.py             a header clean.py doesn't know: proposes the standard column it means (name similarity + its values), with a confidence and reason; nothing runs until a person confirms (`python mapping.py <workbook> --confirm`, or the web page's Review mapping); saved to mappings/<company>.yaml, whose hash goes in the manifest
- config_schema.py       the rules config.yaml must follow (type, range, example, meaning per setting), checked on every load; every problem at once, each naming the key and a line to copy
- cache.py               clean_workbook, compute_metrics and metric_reasons remember their answer by SHA-256 of their inputs, so a run reads each workbook once (benchmark.py measures it)
- metrics.py             every metric, why a value has no number (reasons), flags, data gaps, and the ONE label set (METRIC_LABELS, INPUT_LABELS)
- analyze.py             Claude API → JSON (headline, 3 wins, 3 risks, 3 mgmt questions); validate schema + sign-aware number check + slide-fit check + direction check (a direction word its numbers contradict, a "persistent" trend that isn't), retry once on failure
- compare_models.py      step 3b: blind Sonnet vs Haiku comparison -> README table
- excel_output.py        step 4b: output/<company>_metrics.xlsx (Metrics, Flags, Data gaps sheets)
- theme.py               the only palette, font (Arial, then Helvetica, then DejaVu Sans) and type sizes; the template, deck, charts, memo and web page all take theirs from it (a test fails if any other code file types a color). TEXT_PAIRS: every text/background pair, checked for WCAG AA contrast
- make_template.py       builds templates/base.pptx: fictional "Example Capital" brand, navy/gray, 16:9, title + content layouts
- build_deck.py          step 4: output/<company>_board_pack.pptx on templates/base.pptx; slides: Key metrics, ARR and cash charts, Risks and flags, AI commentary (headline + 3 risks + 3 questions under "AI-drafted from computed metrics - review before use"; wins aren't shown). Re-checks the analysis JSON; "AI summary unavailable" if missing, failed, or too long for slide 4 (ai_text_problems) - a company is never left without a deck. Footer carries the git commit, model and review status ("AI-drafted | not reviewed", or "AI-drafted | reviewed by NAME on DATE" once approve.py records a reviewer). The "DRAFT - NOT REVIEWED" watermark is opt-in: --draft (build_deck.py, main.py) stamps every slide of a deck nobody has approved. --appendix (build_deck.py, main.py) adds one optional slide after the four: every metric for every quarter, with a key
- memo.py                the board memo, 1 to 2 pages, Word + PDF (output/<company>_board_memo.docx/.pdf): headline, what changed since the last run, key metrics, flags, data gaps, questions, the deck's footer. Uses the AI text only if the deck would AND every number in it is one the metrics workbook shows; else "AI commentary unavailable". "Reviewed" only if the approval lists the memo
- charts.py              the two matplotlib charts for slide 2 (ARR + net new ARR, ending cash with runway); a blank quarter is a visible gap; one axis style for both, no label on another, a shrinking quarter amber and hatched (never color alone)
- diff_runs.py           what changed since the last run with different numbers (saved in the manifest): flags that flipped, metrics that moved more than diff_min_points / diff_min_relative, data gaps opened or closed; in the memo, on the company page and in main.py's printout
- rollup.py              `python rollup.py`: output/portfolio_rollup.pptx and .xlsx across every company: ranked by flags tripped with each one's worst flag (WORST_FIRST order), companies by status, runway by company; computed metrics only, no AI text
- export.py              `python export.py <workbook>` or --all: <company>_metrics.csv (one row per quarter and metric), _flags.csv, _export.json (with source hashes, never NaN) and _email.html (slide 1's table, laid out to paste into Outlook); the metrics workbook's numbers to the digit
- text_fit.py            measures slide text, shrinks it to a 12 pt floor, then stops naming the slide and box
- provenance.py          run manifests: input/config SHA-256 hashes, git commit, model + prompt version, tokens, cost; and whether a human approved this deck (approval_status is the only judge)
- approve.py             `python approve.py northwind [--reviewer NAME]`: records reviewer + time in the manifest, so the rebuilt deck's footer says "reviewed by" (and --draft no longer stamps it). Never builds a deck
- main.py                CLI: `python main.py data/northwind.xlsx` or `--all` [--skip-ai] [--draft] [--appendix] [--resume] [--max-cost] [--timeout] [--workers]; also --version, --list-companies, --help (exit codes 0/1/2/130, EXIT_CODES); clean -> metrics -> what changed -> Excel -> AI JSON -> deck -> memo -> manifest per company, reusing a saved analysis of exactly the same numbers for free; summary table + output/batch_summary.csv + output/batch_manifest.json
- resilience.py          what keeps a long batch going: each company built in a private folder and moved into output/ only on success; --resume, --max-cost, --timeout, --workers; rate-limit retries with growing waits
- run_log.py             output/logs/run_<timestamp>.jsonl: one line per step per company (step, seconds, result, error) then one for the whole company; from main.py and the web page
- app.py                 Streamlit web page, two pages. Portfolio: every company in data/ (latest quarter, flags, gaps, last run, deck status), search, per-row Generate and Download (deck, memo, Excel), Generate all with progress, Download rollup, Add a company (with Review mapping for unknown headers), Recent runs. Company: flags, what changed, data gaps, metrics in the status colors, both charts, AI commentary, Generate, downloads, Export, Approve. AI box off by default and labelled with the typical cost; errors in plain words, never a traceback; writes to output/ exactly as main.py does
- portfolio.py           the work behind the page's buttons: rows, search, Generate / Generate all (main.run_company), Add a company (only if clean.py reads it), mapping proposals and confirmations, Approve (approve.approve), what changed
- run_app.command        double-click on a Mac: sets up .venv on first run, starts app.py, opens the browser (.streamlit/config.toml: headless, no tracebacks on the page, theme.py's colors)
- demo_reset.py          before a demo (DEMO.md): deletes every file the tool builds, keeps every saved AI analysis (checked by hash), builds last quarter's run then today's as Generate all does with the AI box unticked; never calls the API; exit 1 with what to fix
- golden.py              golden files: an approved text copy of every deck, memo and metrics workbook in tests/golden/ (rebuilt by pytest from tests/golden/analysis/, fixed date and commit); `python golden.py` shows differences, --update accepts them
- benchmark.py           seconds per company run and workbook reads, in a temp folder, no API call
- config.yaml            flag thresholds with investor reasoning in comments (NRR < 100%, burn > 15% over budget, runway < 12 mo)
- tests/ + pytest.ini    pytest unit tests (`python -m pytest -q`), expected values worked out by hand; tests/test_docs.py keeps the docs' file and function names real, and README/CLAUDE.md/STUDY_GUIDE/LOOM_SCRIPT on the 4-slide deck, the opt-in (--draft) watermark and the web page
- check_northwind.py, check_companies.py, check_excel_output.py, check_deck.py, check_main.py   end-to-end proofs against each company's answer key (no API calls)
- check_memo.py, check_rollup.py, check_diff.py, check_export.py   the same for the memo (every number is in the metrics workbook, Word and PDF), the rollup, what changed (last quarter's run then today's) and the exports (Outlook-safe email)
- mappings/              confirmed column mappings, one <company>.yaml each, committed like config.yaml (empty today: the three companies' headers are all known)
- DEMO.md                a 5 minute walkthrough of the web page for a non-technical viewer; run demo_reset.py first
- README.md              what it does, how to run it, data flow, design decisions, model comparison (compare_models.py rewrites the block between its marker comments), cost, screenshots, next steps
- STUDY_GUIDE.md         data flow, glossary, interview questions, exercises
- INTERVIEW_PREP.md      every interview question in the order asked, 30-60 s answers with where to point; how to answer when I can't recall (tests/test_docs.py checks the files and functions it names)
- LEARNINGS.md           what broke and what it taught; model comparison and live-run costs
- DAY_REPORT.md          per-task report of the day's build (decisions, failures, unresolved); OVERNIGHT_REPORT.md is a historical record
- POLISH_REPORT.md, FINAL_REPORT.md   the same for the polish run and the final run (the memo onward): what each task built, decisions, what failed, planted-bug results, unresolved
- output/                generated files (git-ignored): per company the deck, memo (.docx/.pdf), metrics workbook, analysis JSON, manifest and exports; portfolio_rollup.pptx/.xlsx, batch_summary.csv, batch_manifest.json, logs/

## Rules
- Python computes every number. Claude only interprets computed metrics. It never does math.
- All ratios are stored as decimals; formatting to % happens only at output.
- API key lives in .env (never committed). Load with python-dotenv.
- Fictional data only. No real company names or numbers.
- Keep functions small and commented. I need to explain every line.
- Log anything that breaks in LEARNINGS.md.
- One label set: metric and input display names live only in metrics.py. Flag names ARE the metric labels, so every output (printout, Excel, Claude's payload, deck) uses the same words.
- No em dashes in any .md file or user-facing text: use a comma, colon or full stop. tests/test_docs.py fails on one in any doc, any string in the project's .py files, the AI system prompt or the metric labels.

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
- Flags: trip / pass / cannot evaluate. "Cannot evaluate" always carries its reason ("cannot evaluate: missing input").
- Combo rule (retention problem) = NRR falling AND pipeline rising over the last `combo_lookback_quarters` quarters (config.yaml)
  - Window = the last N quarters, including the latest (N=3 means 2 quarter-over-quarter comparisons). N must be at least 2, or config loading stops.
  - Falling = NRR fell by at least `combo_min_nrr_drop` (0.01 = 1 point) in every comparison in the window: a 0.1-point dip is noise
  - Rising = pipeline increased in every quarter-over-quarter comparison in the window
  - Not enough history -> "cannot evaluate: no prior period"; a blank NRR or pipeline input -> "cannot evaluate: missing input"; never False

## Messy data rules
- Parse "$1.2M" / "850K" text to numbers; normalize header names
- Blank quarter stays blank and is flagged "data missing" on the deck. Never impute numbers that could reach a board.
- Missing values spread to every metric that uses them. Any metric whose inputs include the blank quarter shows "data missing":
  - QoQ (and net new ARR vs budget): the blank quarter and the next quarter
  - YoY: the blank quarter and the quarter 4 quarters later
- Only missing input counts: a metric is a data gap only when one of ITS inputs is blank. A partly blank quarter doesn't make unrelated metrics gaps.
- Any flag or combo rule that depends on a missing value returns "cannot evaluate: missing input" instead of pass/fail
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
ARR stalls then shrinks (net new ARR negative from Q1 2026, so burn multiple is ∞ "ARR shrank"). Churn and contraction climb: NRR ~78%, GRR ~75%. Burn rises ~20% over budget; runway 6 months. New sales dry up, so CAC payback is over 10 years. Pipeline FALLS too, so the combo passes (a sales AND retention problem). Q2 2025 is blank, exactly 4 quarters before Q2 2026, so Rule of 40 is "cannot evaluate: missing input" in the latest quarter. Mess: a title line and empty row above the header, budget columns next to their actuals, budget row "Q3 2026 - Bud".

The three companies together prove flags aren't hard-coded.

## Build order
1. make_data.py: generate fake messy workbook for "Northwind Software" (8 quarters: ARR, new ARR, churned ARR, NRR, gross margin, burn, cash, headcount, pipeline + budget columns; inconsistent headers, "$1.2M" text, one blank quarter, junk notes tab)
2. clean.py + metrics.py: verify outputs against manual Excel math
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
