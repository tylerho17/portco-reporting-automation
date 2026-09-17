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
- API key lives in .env (never committed). Load with python-dotenv.
- Fictional data only. No real company names or numbers.
- Keep functions small and commented. I need to explain every line.
- Log anything that breaks in LEARNINGS.md.

## Build order
1. make_data.py: generate fake messy workbook for "Northwind Software" (8 quarters: ARR, new ARR, churned ARR, NRR, gross margin, burn, cash, headcount, pipeline + budget columns; inconsistent headers, "$1.2M" text, one blank quarter, junk notes tab)
2. clean.py + metrics.py — verify outputs against manual Excel math
3. analyze.py
4. build_deck.py + charts
5. main.py batch mode + 2 more fake companies
6. README with before/after screenshots
