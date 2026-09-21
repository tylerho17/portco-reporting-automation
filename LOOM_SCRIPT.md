# Loom script: Board Pack Generator (2 minutes)

About 293 spoken words at 150 words a minute. **Show** is what's on screen; **Say** is what you say. The whole recording happens on the web page, the downloaded deck and memo, and two docs: no terminal on camera. Practise with a timer: if a section runs long, cut words, not the timestamps.

The script quotes only numbers Python computes, plus the model scores and the counts in the docs (table at the end). Claude's wording can change between runs, so no AI sentence is quoted word for word.

---

## Before you record

1. **Reset the output folder, with no API call:** `python demo_reset.py`. It must end "Ready for the demo." It rebuilds every deck and memo from the saved analysis JSONs (the AI text is reused, nothing is sent to Claude), and builds last quarter's run first, so the memo has "what changed" to show.
   - One warning is expected: Fernhollow has no saved analysis of today's numbers, so its deck says "AI summary unavailable". The recording opens Northwind only, so it never shows.
   - If it ends "Not ready for the demo", fix each line marked ✗ (each says how) and run it again.
2. **Commit first.** `git status` must be clean, or every footer shows the commit with a `*` ("built from code that wasn't committed"). If you commit, run `python demo_reset.py` again.
3. **Don't approve Northwind before recording.** The footer should read "AI-drafted | not reviewed" on camera: that is the point of the approval gate. `demo_reset.py` clears approvals, so this is the state it leaves.
4. **No watermark:** don't build with `--draft`. It stamps a diagonal DRAFT across every slide, over the numbers you point at. The footer carries the review status instead.
5. **Open the web page:** double-click `run_app.command` in Finder (it sets up and starts `app.py`, then opens the browser). Leave its Terminal window open but behind the browser. Browser full screen, zoom 100%, portfolio page at the top, not scrolled.
6. **Leave the AI box unticked.** Ticked, Generate could call Claude (about $0.09 and 70 s). Unticked, it reuses Northwind's saved analysis for free, in about a second. Click Generate on Northwind once before recording to see the green line that says the files are ready, then go back to the portfolio page.
7. **Clear old downloads.** Delete any older `northwind_board_pack.pptx` and `northwind_board_memo.pdf` from Downloads, so the files you open on camera are the ones you just downloaded. PowerPoint and a PDF viewer open in the background.
8. **Open these tabs, in this order, for 1:03 onward:**
   1. README.md at **Key design decisions** (decision 1 at the top)
   2. README.md at **Model comparison**, with the Results table in view
   3. LEARNINGS.md at "v3 → v4: false direction and good-news claims (Task 7)"
   4. README.md at **Next steps**, "Still to do"
9. **Notifications off** (Focus mode), mouse cursor large enough to follow.

---

## 0:00 to 0:11 The problem (27 words)

**Show:** the portfolio page, not scrolled.

**Say:**
> "Every quarter, each portfolio company sends a KPI spreadsheet, each in its own format. Someone re-keys it and writes the board update by hand. I automated that."

## 0:11 to 0:29 The portfolio page (44 words)

**Show:** the table. Run the cursor along the column headers (Latest quarter, Flags tripped, Data gaps), then down **Flags tripped** row by row, then Northwind's **Data gaps** cell ("20 metrics/flags (blank: Q1 2025)").

**Say:**
> "This is the portfolio page. Three fictional companies, one row each: latest quarter, flags tripped, data gaps. Same code, three stories: Alderpeak trips 0 of 9 flags, Fernhollow 7, Northwind 6. Northwind left a quarter blank, and the tool says so instead of guessing."

## 0:29 to 0:39 Pick a company and generate (24 words)

**Show:** click **Northwind** in the first column. On its page, click the navy **Generate** button (AI box unticked) and let the green line appear.

**Say:**
> "I click Northwind, then Generate. It cleans the messy workbook, computes every metric, checks nine thresholds and builds the deck and memo in seconds."

## 0:39 to 1:03 The deck and the memo (56 words)

**Show, 0:39 to 0:53:** click **Download deck**, open it in PowerPoint. Slide 1: point at NRR 97.1% and runway 11.0 mo with their thresholds. Slide 2: the gap at Q1 2025. Slide 3: "6 of 9 flags tripped". Slide 4: the gray "AI-drafted from computed metrics - review before use" line.

**Say:**
> "The deck: four slides. Key metrics: NRR 97.1%, below 100; runway 11 months against a 12 month floor. Charts with the blank quarter left as a gap. Every tripped flag against its threshold. Then Claude's commentary, marked AI-drafted."

**Show, 0:53 to 1:03:** back in the browser, click **Download memo (PDF)** and open it. Point at **What changed since the last run** and its five "Flags that flipped" lines.

**Say:**
> "The memo says the same in two pages, plus what changed since last quarter's run: five flags flipped."

## 1:03 to 1:43 Design decisions (102 words)

**Show, 1:03 to 1:20:** README.md, **Key design decisions**: decision 1, then decision 2.

**Say:**
> "Five decisions. One: Python computes every number; Claude only interprets. Two: its answer is checked in code: the right shape, it fits the slide, and every number must appear in the data, sign included."

**Show, 1:20 to 1:27:** README.md, **Model comparison**, the Results table (Avg score column).

**Say:**
> "Three: I picked the model by blind test. Haiku scored 2 of 5, Sonnet 4, so Sonnet stays."

**Show, 1:27 to 1:35:** LEARNINGS.md, the "v3 → v4" entry: point at "persistent gentle decline".

**Say:**
> "Four: passing checks isn't the same as right. Reading the output, I caught three false claims, like a trend that moved both ways called persistent, and wrote rules against them."

**Show, 1:35 to 1:43:** PowerPoint, slide 4's footer ("... | claude-sonnet-5 | AI-drafted | not reviewed"), then the browser: scroll Northwind's page to the **Approve** box.

**Say:**
> "Five: every deck records its input hashes, code version and model, and says not reviewed until a person approves it."

## 1:43 to 2:00 Next steps (40 words)

**Show:** README.md, **Next steps**, "Still to do".

**Say:**
> "Next: any file format in, with each number cited to its source and a person confirming it. A SharePoint trigger, so a dropped workbook starts the run. And the portfolio rollup, built for three companies today, run across all 275."

---

## Where every number said on camera comes from

| Said on camera | Where it's shown | Computed or recorded by |
|---|---|---|
| Three fictional companies | The portfolio page's rows, one per workbook in `data/` | `make_data.py`, `make_data_alderpeak.py`, `make_data_fernhollow.py` |
| 0 of 9, 7 of 9, 6 of 9 flags | Portfolio page, **Flags tripped**; each deck's slide 3 (Fernhollow's also says "1 cannot evaluate") | `metrics.evaluate_flags`, thresholds in `config.yaml`; proved by `python check_companies.py` |
| A quarter blank (Q1 2025) | Portfolio page, **Data gaps**; slide 3's Data gaps line; the gap in slide 2's charts | `make_data.py` (`BLANK_QUARTER`); spread by `metrics.metric_reasons` |
| Nine thresholds | The nine flags on slide 3 and in the memo | `metrics.evaluate_flags` with `config.yaml` |
| In seconds | The green line after Generate | `python benchmark.py`: 1.11 s for Northwind with no AI call (FINAL_REPORT.md, Task 17) |
| Four slides | The deck | `build_deck.py` (the appendix slide is off unless `--appendix`) |
| NRR 97.1%, below 100 | Slide 1, NRR (annualized) row; threshold `nrr_min` in `config.yaml` | `metrics.nrr` |
| Runway 11 months, 12 month floor | Slide 1, Runway at current burn row; `runway_min_months` in `config.yaml` | `metrics.runway_months` |
| Two pages | `northwind_board_memo.pdf` has 2 pages | `memo.py` |
| Five flags flipped | The memo's "What changed since the last run": NRR, burn multiple, runway, Rule of 40 and the combo rule | `diff_runs.py`, against last quarter's run that `demo_reset.py` builds; proved by `python check_diff.py` |
| Five decisions | The order of this script (README decisions 1, 2, 10, 11 and 12, plus LEARNINGS.md) | Not computed |
| Haiku 2 of 5, Sonnet 4 | README, **Model comparison**, Avg score (Haiku needed 4.0 to replace Sonnet) | `compare_models.py`, scored blind by hand |
| Three false claims | LEARNINGS.md, "v3 → v4: false direction and good-news claims (Task 7)": growth "well above prior-period levels" while it fell, an up-and-down NRR called a "persistent gentle decline", a flag that passed for a bad reason called a win | Found by reading the Task 4 live run; the direction part is now also checked in code (`analyze.claim_problems`) |
| 275 companies | The demo's premise (README, Cost and Model comparison) | Not computed |

## Honesty notes for the next steps

- **Any file format in, with citations and a confirm:** not built. What exists is the confirm step for workbooks: `mapping.py` proposes what an unknown header means, and nothing runs until a person confirms it. Today a number's source is named only when `clean.py` stops (sheet, row and cell).
- **SharePoint trigger:** not built (README, Next steps).
- **Portfolio rollup:** built and proved for three companies (`rollup.py`, `python check_rollup.py`). A batch of 275 has not been run; the time and cost for it are estimates (README, Cost).

## Timing check

| Section | Seconds | Words |
|---|---|---|
| The problem | 11 | 27 |
| The portfolio page | 18 | 44 |
| Pick a company and generate | 10 | 24 |
| The deck and the memo | 24 | 56 |
| Design decisions | 40 | 102 |
| Next steps | 17 | 40 |
| **Total** | **120** | **293** |

Numbers read aloud ("97.1%", "275") take longer than their word count, which is why the total sits under 300.
