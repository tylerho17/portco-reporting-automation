# Board Pack Generator

## Goal
Messy portfolio-company KPI workbook (.xlsx) → 5-slide board update deck (.pptx) + AI summary.
Demo project for a PE AI automation role. Must be clean, explainable, and reliable.

## About me
- Finance background, learning Python. Explain every file after writing it, in plain English.
- Work in small steps. Use plan mode before writing code. Stop after each build step so I can review.

## Environment
- macOS, Python venv at .venv (activate: `source .venv/bin/activate`)
- Packages in requirements.txt. Don't add packages without telling me why.

## Architecture
- data/          fake input workbooks (3 made-up software companies)
- clean.py       normalize headers, parse "$1.2M"-style text, handle blanks
- metrics.py     QoQ, YoY, variance vs budget, runway, Rule of 40, burn multiple, CAC payback, threshold flags
- analyze.py     Claude API → JSON (headline, 3 wins, 3 risks, 3 mgmt questions); validate schema, retry once on failure
- build_deck.py  python-pptx using templates/base.pptx; slides: Summary, KPI table, Charts, Risks/Flags, Questions
- main.py        CLI: `python main.py data/northwind.xlsx` or `--all` for batch
- config.yaml    flag thresholds with investor reasoning in comments (NRR < 100%, burn > 15% over budget, runway < 12 mo)
- output/        generated decks (git-ignored)

## Rules
- Python computes every number. Claude only interprets computed metrics — never does math.
- All ratios are stored as decimals; formatting to % happens only at output.
- API key lives in .env (never committed). Load with python-dotenv.
- Fictional data only. No real company names or numbers.
- Keep functions small and commented. I need to explain every line.
- Log anything that breaks in LEARNINGS.md.

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
- Burn multiple = net_burn / net new ARR
- Rule of 40 = YoY revenue growth % + FCF margin, where FCF margin = -net_burn / revenue
- CAC payback (months) = sm_spend / (new_arr * gross margin) * 12
- Runway (months) = ending_cash / (net_burn / 3), at current burn and at next quarter's budgeted burn
  - The runway flag uses current burn. Runway at next quarter's budgeted burn is shown as context, not flagged.
- Flags are evaluated on the latest actual quarter (functions accept any quarter, for step 4b).
- Edge cases ("not meaningful" must never look like "data missing"):
  - Burn multiple when net new ARR <= 0 and net_burn > 0 -> infinite, trips ("ARR shrank")
  - Burn multiple when net_burn <= 0 -> 0, passes ("not burning")
  - Runway when net_burn <= 0 -> infinite, passes ("not burning")
  - CAC payback when new_arr * gross margin <= 0 -> infinite, trips
- Combo rule (retention problem) = NRR falling AND pipeline rising over the last `combo_lookback_quarters` quarters (config.yaml)
  - Window = the last N quarters, including the latest (N=3 means 2 quarter-over-quarter comparisons)
  - Falling = NRR decreased in every quarter-over-quarter comparison in the window
  - Rising = pipeline increased in every quarter-over-quarter comparison in the window
  - Any quarter in the window missing -> "cannot evaluate — data missing", never False

## Messy data rules
- Parse "$1.2M" / "850K" text to numbers; normalize header names
- Blank quarter stays blank and is flagged "data missing" on the deck. Never impute numbers that could reach a board.
- Missing values spread to every metric that uses them. Any metric whose inputs include the blank quarter shows "data missing":
  - QoQ (and net new ARR vs budget): the blank quarter and the next quarter
  - YoY: the blank quarter and the quarter 4 quarters later
- Any flag or combo rule that depends on a missing value returns "cannot evaluate — data missing" instead of pass/fail
- The Risks/Flags slide gets a "Data gaps" line listing every metric and flag affected

## Northwind story (make_data.py must produce this)
8 quarters. ARR grows strongly. NRR slides 108% -> 97% over last 3 quarters. Pipeline keeps rising. Burn ~20% over budget. Runway ends ~11 months. Should trip 4+ flags.
Later: one healthy company (no flags) and one distressed company, to prove flags aren't hard-coded.

## Build order
1. make_data.py: generate fake messy workbook for "Northwind Software" (8 quarters: ARR, new ARR, churned ARR, NRR, gross margin, burn, cash, headcount, pipeline + budget columns; inconsistent headers, "$1.2M" text, one blank quarter, junk notes tab)
2. clean.py + metrics.py — verify outputs against manual Excel math
3. analyze.py
4. build_deck.py + charts
5. main.py batch mode + 2 more fake companies
6. README with before/after screenshots

## Added build steps
- 3b. Model comparison: run analyze.py on the same metrics with claude-sonnet-5 and claude-haiku-4-5. Log cost, latency, schema pass/fail, and my 1-5 quality score to a table in README with a one-line recommendation.
- 4b. Excel output: write output/<company>_metrics.xlsx with computed metrics and flagged cells highlighted.
- Feedback round after first full deck: log "reviewer said X -> changed Y" in LEARNINGS.md.
