# Loom script: Board Pack Generator (2 minutes)

About 285 spoken words, at roughly 150 words a minute. **Say** is what you say; **Show** is what's on screen. Practise with a timer: if a section runs long, cut words, not the timestamps.

The script quotes only numbers Python computes (6 of 9 flags, NRR 97.1%, runway 11.0 mo, and so on). Those are the same on every run. Claude's wording changes from run to run, so no AI sentence is quoted word for word.

---

## Before you record

1. **Put the AI text back on the decks (no API call):** `python demo_reset.py` (it must end "Ready for the demo."), or for one deck `python build_deck.py data/northwind.xlsx`. It's needed if `check_main.py` has run since, because that script leaves "AI summary unavailable" decks behind. Open slide 4 and check you see the headline, not the placeholder.
   - **Build without `--draft`**, so there's no watermark over the numbers. The footer does the job instead: check it ends "AI-drafted | reviewed by Tyler Ho on ...". `demo_reset.py` clears every approval (a demo starts "not reviewed"), so after it run `python approve.py northwind`, then `python build_deck.py data/northwind.xlsx`.
2. **Choose how to show the run** (0:35):
   - **Option A (free, no risk; recommended for now):** don't run the AI again. Show the terminal output of today's real live run of `python main.py --all`, saved in `output/day_logs/task4_live_run.txt` (e.g. open it in the VS Code terminal with `cat`). The narration doesn't say "live", so it's accurate either way.
   - **Option B (live, about $0.09 and 70 s for one company):** run `python main.py data/northwind.xlsx` on camera and trim the wait in Loom.
   - **Why A for now:** today's Northwind risks only just fit slide 4 (at the 12 pt floor). A slightly longer answer gets one retry and then the "AI summary unavailable" placeholder (DAY_REPORT.md, Review, finding 1; POLISH_REPORT.md, Task 2). A live run could hit that on camera.
3. **Open these windows in this order:**
   1. `data/northwind.xlsx` in Excel, on the KPI Tracker tab
   2. VS Code with `CLAUDE.md` open and the Claude Code panel beside it
   3. a terminal in the project folder, with the final line of `python -m pytest -q` already showing
   4. `output/northwind_board_pack.pptx` in PowerPoint, on slide 1
   5. the web page: double-click `run_app.command` (it starts `app.py` and opens the browser), then click Northwind in the portfolio table, AI box unticked, so its flags and metrics are already showing
4. **Don't show Alderpeak's or Fernhollow's slide 4:** Claude's text there has claims that passed validation but are misleading (LEARNINGS.md, Task 4 live run). The walk-through uses Northwind only.

---

## 0:00 – 0:15 The problem (~35 words)

**Show:** PowerPoint on slide 1, or your face.

**Say:**
> "A PE fund with 275 portfolio companies gets a KPI spreadsheet from each one every quarter, each in its own format. Someone re-keys them, checks the metrics and writes the board update by hand. I automated that."

## 0:15 – 0:35 The messy input, and how I built it (~55 words)

**Show, 0:15 – 0:22:** Excel, `data/northwind.xlsx`. Point at " Churned ARR " and "Cash - End of Qtr", the "$14.3M" typed as text, and the empty Q1 2025 row.

**Say:**
> "Here's a fictional company's workbook: odd headers, money typed as text, and a quarter that's blank."

**Show, 0:22 – 0:35:** VS Code. Scroll `CLAUDE.md` to **Rules** and **Metric definitions**, then the Claude Code panel, then the terminal's test line (`841 passed`).

**Say:**
> "I build it in VS Code with Claude Code. CLAUDE.md holds the spec: every metric definition and the rules. I work in small steps: plan, tests first, then code. Each step ends with over 500 tests passing."

## 0:35 – 1:20 Run it, and walk the Northwind deck (~105 words)

**Show, 0:35 – 0:45:** the terminal: the ✓ lines, then the summary table.

**Say:**
> "One command runs the batch. Same code, three stories: healthy Alderpeak trips 0 of 9 flags, distressed Fernhollow 7, and Northwind 6."

**Show, 0:45 – 0:57:** slide 1, Key metrics. Point at NRR 97.1%, runway 11.0 mo, then "data missing" in the Q1 2026 column of the ARR growth YoY row (it compares with the blank Q1 2025).

**Say:**
> "Every number here is from Python. NRR is 97.1%, below 100%, and runway is 11 months against a 12-month floor. Where a number needs the blank quarter, it says data missing. It never guesses."

**Show, 0:57 – 1:04:** slide 2, the charts. Point at the gap at Q1 2025.

**Say:**
> "The blank quarter stays a visible gap in the charts."

**Show, 1:04 – 1:12:** slide 3 ("6 of 9 flags tripped" at the top left, the Combo rule under the flags, Data gaps on the right).

**Say:**
> "Each tripped flag against its threshold, plus a combo rule: NRR falling while pipeline rises points to retention, not sales."

**Show, 1:12 – 1:20:** slide 4, AI commentary. Point at the gray line under the title, then the end of the footer ("AI-drafted | reviewed by ...").

**Say:**
> "Only the last slide is Claude's, marked AI-drafted. Every footer says whether a person has reviewed the deck."

## 1:20 – 1:45 Design choices (~60 words)

**Show:** README.md, **Key design decisions**, then scroll to the **Model comparison** table.

**Say:**
> "Four choices. Python does all the math; Claude only interprets. Claude's answer is validated in code: the right shape, and every number it writes must appear in the data, sign included. If it fails twice, the deck still builds with a placeholder. And I picked the model with a blind test: Haiku scored 2 of 5, so Sonnet stays."

## 1:45 – 2:00 The web page, and next steps (~30 words)

**Show, 1:45 – 1:52:** the web page's portfolio table (three companies, their flags and deck status), then Northwind's page with its flags and download buttons.

**Say:**
> "Colleagues who don't use a terminal get a web page: every company in one table, click Generate, download the deck."

**Show, 1:52 – 2:00:** README.md, **Next steps**.

**Say:**
> "Next: a SharePoint trigger, and a portfolio rollup across all companies."

---

## Where each number in the script comes from

| Said on camera | Where it comes from | Computed by |
|---|---|---|
| 275 portfolio companies | The demo's premise (README cost section) | Not computed |
| 0 of 9, 7 of 9, 6 of 9 flags | Summary table from `main.py` | `metrics.evaluate_flags` |
| NRR 97.1%, below 100% | Slide 1, NRR row; threshold from `config.yaml` (`nrr_min`) | `metrics.nrr` |
| Runway 11 months vs 12-month floor | Slide 1, Runway row; `config.yaml` (`runway_min_months`) | `metrics.runway_months` |
| Q1 2025 blank | Northwind's blank quarter | `make_data.py` (`BLANK_QUARTER`) |
| Over 500 tests | `python -m pytest -q` (1252 today) | pytest counts them |
| 12 of 12 edge cases | `python eval/run_eval.py` | each eval company's answer key in `eval/make_eval_data.py` |
| Portfolio rollup: Fernhollow 7, Northwind 6, Alderpeak 0; worst flag runway | `python rollup.py` or the web page's Download rollup (`output/portfolio_rollup.pptx`) | `rollup.py` from `metrics.evaluate_flags`; proved by `python check_rollup.py` |
| Northwind Q1 to Q2: 5 flags flipped, 8 metrics moved, 1 data gap closed | The memo's "What changed since the last run", the company page's card, or `python diff_runs.py data/northwind.xlsx` (needs a run a quarter earlier: `python demo_reset.py` makes one) | `diff_runs.py` from the saved results in the manifest; proved by `python check_diff.py` |
| Metrics and flags as CSV and JSON, and a table for email | `python export.py data/northwind.xlsx`, or the company page's Export | `export.py` from the same numbers as the metrics workbook; proved by `python check_export.py` |
| Each deck, memo and workbook matches its approved copy | `python golden.py` ("9 of 9 match their goldens") | the approved text copies in `tests/golden/` |
| Haiku scored 2 of 5 | README model comparison (needed 4.0) | `compare_models.py` |

## Timing check

| Section | Seconds | Words (approx.) |
|---|---|---|
| Problem | 15 | 35 |
| Input and workflow | 20 | 55 |
| Run and deck | 45 | 105 |
| Design choices | 25 | 60 |
| Web page and next steps | 15 | 30 |
| **Total** | **120** | **~285** |
