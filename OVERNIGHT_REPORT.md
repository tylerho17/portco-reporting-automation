# Overnight Report

## Task 1 — Step 5 data: a healthy and a distressed company

**Result:** done. Both new workbooks are in `data/`. `check_northwind.py` and the new `check_companies.py` both print "All checks passed". Four commits, from `638498d` to `24c1dd4`.

### What I built

| File | What it is, in plain English |
|---|---|
| `make_data_common.py` (new) | The shared "workbook factory". The code that checks an answer key and writes a messy workbook used to live inside `make_data.py`. It moved here so all three companies use the same code. `check_true_data` confirms ARR and cash roll forward, money values are multiples of 10, and the mess settings point at real columns and quarters. `write_kpi_sheet` writes the table. `save_workbook` runs the check and then saves the file. |
| `make_data.py` (changed) | Now holds **only Northwind's** numbers, headers, text cells and notes, and hands them to `save_workbook`. I regenerated `northwind.xlsx` and compared it cell by cell with the previous file: **identical**. |
| `make_data_alderpeak.py` (new) | **Alderpeak Software, the healthy company.** ARR grows 47.5% YoY and NRR holds around 110%. Burn shrinks every quarter and stays under budget, so runway is 108 months. Rule of 40 is 42.7%. **0 of 9 flags trip in Q2 2026, and no flag trips in any earlier quarter either.** |
| `make_data_fernhollow.py` (new) | **Fernhollow Systems, the distressed company.** ARR shrinks, NRR is 77.9%, GRR is 74.7% and runway is 6.0 months. Burn is 19.6% over budget and CAC payback is 134 months. **7 of 9 flags trip.** Q2 2025 is blank, so Rule of 40 is "cannot evaluate — data missing", and the combo rule passes (see decisions). |
| `data/alderpeak.xlsx`, `data/fernhollow.xlsx` | The generated messy workbooks. |
| `check_companies.py` (new) | Runs the same four checks on **all three** companies: (1) cleaning recovers the answer key exactly, (2) key Q2 2026 metrics equal hand formulas typed from the answer key, (3) every flag status matches the story, (4) data gaps are exactly the ones the CLAUDE.md rules predict, no more and no fewer. It also checks that Alderpeak trips nothing in any quarter. |

**Flag results, Q2 2026:**

| Flag | Northwind | Alderpeak (healthy) | Fernhollow (distressed) |
|---|---|---|---|
| NRR (annualized) | trip | pass (109.8%) | trip (77.9%) |
| GRR (annualized) | pass | pass (96.0%) | trip (74.7%) |
| Burn multiple | trip | pass (0.12x) | trip (∞, ARR shrank) |
| Burn vs budget | trip | pass (−9.1%) | trip (+19.6%) |
| Runway | trip | pass (108.0 mo) | trip (6.0 mo) |
| CAC payback | pass | pass (17.7 mo) | trip (134.4 mo) |
| Net new ARR vs budget | pass | pass (+5.8%) | trip (−150.0%) |
| Rule of 40 | trip | pass (42.7%) | **cannot evaluate — data missing** |
| NRR falling while pipeline rising | trip | pass | pass |
| **Total** | 6 trip, 3 pass | 0 trip, 9 pass | 7 trip, 1 pass, 1 cannot evaluate |

### Decisions you didn't specify

1. **Company names:** "Alderpeak Software" (healthy) and "Fernhollow Systems" (distressed), picked to sound invented. Files are `make_data_alderpeak.py` / `data/alderpeak.xlsx` and `make_data_fernhollow.py` / `data/fernhollow.xlsx`.
2. **Shared code instead of copies:** I moved the writing and checking functions into `make_data_common.py` rather than pasting ~100 lines into each new file. Each company file is now just data plus a short `main()`. The functions take one `company` dictionary instead of reading global variables.
3. **Same 8 quarters for everyone (Q3 2024 – Q2 2026):** a fund's board pack covers the same quarter for every company, so the batch run in step 5 lines up.
4. **Alderpeak has no blank quarter.** A well-run company with clean reporting should show "Data gaps: none", and this tests the zero-gaps path, which Northwind can't. The rest of the mess is still there: odd headers, text numbers, a budget-only row and a notes tab. To add a blank quarter, change `BLANK_QUARTER = None` to a quarter label. Its flags would still pass, as long as the blank isn't Q2 2025 or one of the last 3 quarters.
5. **Fernhollow's blank quarter is Q2 2025, exactly 4 quarters before the latest.** That puts the YoY gap on Q2 2026, so a latest-quarter flag (Rule of 40) returns "cannot evaluate — data missing". Northwind's blank quarter never reaches its latest flags, so this was the one CLAUDE.md rule the real data didn't exercise yet.
6. **Fernhollow's pipeline falls, so the combo rule passes.** The combo means "retention problem while sales keep filling the funnel". Fernhollow has a sales problem too, so a pass is the correct diagnosis, and it proves the combo isn't just "NRR is bad". The trade-off: 7 of 9 trip rather than 8. To make it trip, make pipeline rise over Q4 2025–Q2 2026.
7. **Alderpeak's NRR zigzags on purpose** (109.8% → 109.7% → 109.8% over the last 3 quarters). It falls in one step and rises in the next, which tests that the combo needs a decline at *every* step.
8. **Alderpeak is clean in every quarter, not just the latest.** CLAUDE.md only requires the latest quarter, but step 4b will show flags for any quarter, and a "healthy" company tripping Rule of 40 in 2025 would be confusing. `check_companies.py` enforces it.
9. **New kinds of mess, all of which `clean.py` already handles:** Alderpeak has new header spellings ("S & M spend", "contraction-arr", "revenue " with a trailing space) and puts the Notes tab *before* the KPI tab. Fernhollow has a title line and an empty row above the header, a different column order (budget columns next to their actuals), and budget rows labelled "Q3 2026 Plan" / "Q3 2026 - Bud". **I made no changes to `clean.py` or `metrics.py`.** I deliberately avoided header names that would need new `HEADER_ALIASES`, so the checks prove existing behavior instead of new code.
10. **Budgeted burn is always positive.** Alderpeak's burn falls to $200K but never goes to zero or negative, because of unresolved item 1 below.
11. **`check_companies.py` also covers Northwind**, reusing `EXPECTED_LATEST`, `EXPECTED_FLAGS` and `EXPECTED_RUNWAY_AT_BUDGET` from `check_northwind.py`. The expected values live in one place, and "all three companies" is visible in one run. `check_northwind.py` itself is unchanged; it also still tests the analysis validator.
12. **Hand formulas for the new companies** follow the lesson in LEARNINGS.md: each one is typed from the answer-key numbers (e.g. `200 / (1220 + 580 - 70 - 100)`), not copied from `metrics.py`. Where a metric should be "data missing" (Fernhollow's Rule of 40 and ARR YoY), the expected value is `NaN`, and the check requires the result to be NaN too.

### What failed and how I fixed it

1. **First healthy design would have tripped flags in history.** In my first draft, Alderpeak's NRR slid a little every quarter from Q3 2024 to Q4 2025 (110.0% → 108.4%) while pipeline rose, which trips the combo in earlier quarters. Rule of 40 was also 38.6% in Q3 2025. Fix: expansion now moves up and down, and early burn is lower (Rule of 40 is now 42.4%–43.1%).
2. **Second draft would have tripped the combo in the latest quarter.** NRR went 109.82% → 109.68% → 109.51%, a drop of only 0.3 points, but the rule trips on *any* decline. Fix: Q2 2026 expansion went from $570K to $580K, so NRR ends at 109.75%. This led to unresolved item 2 below.
3. **Checks passed on the first run, which proves nothing by itself.** So I broke things on purpose, in memory, without editing any files. `check_companies.py` caught all 7 breaks: wrong expected flag, NaN expected but a number found, a number expected but NaN found, wrong YoY gap quarter, missing flag gap, a cleaned value off by 1, and a healthy company tripping in an earlier quarter. `check_true_data` caught all 4: broken ARR chain, broken cash chain, a value not divisible by 10, and a text cell on the blank quarter.
4. **Minor tooling issue:** the shell sandbox blocked multi-line inline scripts, so I wrote scratch scripts to `/tmp` instead. No effect on the project.

### Unresolved: needs your call

1. **"Burn vs budget" gives the wrong answer when a company isn't burning.** With net burn −200 and budget −150 (generating *more* cash than planned), `net_burn / budget_net_burn − 1` = +33%, which **trips** "burn over budget". A budget of 0 gives ∞, which also trips. CLAUDE.md defines edge cases for burn multiple and runway but not this one. I didn't change the formula; that's a definition decision. A possible rule: if budgeted burn ≤ 0, compare in $K, or mark it "not meaningful".
2. **The combo rule has no minimum size of decline.** A 0.1-point NRR dip per quarter at 110% NRR counts as "NRR falling" and trips "retention problem". You may want a minimum drop (e.g. 1 point per step) in `config.yaml`. I didn't edit config.yaml, per the rules.
3. **"Cannot evaluate" flags display their value as "n/a".** In both `metrics.py`'s printout and the analyze payload, Fernhollow's Rule of 40 value is "n/a" next to the status "cannot evaluate — data missing". The status is right, but "n/a" on its own reads like "not applicable", and CLAUDE.md says missing must never look like not meaningful. Best fixed when building the deck (step 4). I didn't touch analyze.py.
4. **Negative percentages will be hard to read on a slide.** Fernhollow's net new ARR vs budget is −150.0% (−$240K actual vs +$480K budget). The math is correct, but the deck may want to show the $K figures alongside it.
5. **Not done here, by design:** no analysis was run for the new companies (no API calls), and CLAUDE.md wasn't updated with their stories. The stories are in each `make_data_*.py` docstring. You may want a "Healthy / distressed company story" section in CLAUDE.md like the Northwind one.
6. **Overlap:** Northwind's cleaning, flag and gap checks now run in both `check_northwind.py` and `check_companies.py`. It's harmless, but if you'd rather have one home, the four Northwind checks could be deleted from `check_northwind.py`.

---

## Task 2 — Step 4b Excel output

**Result:** done. `python excel_output.py data/northwind.xlsx` saves `output/northwind_metrics.xlsx`. The new `check_excel_output.py` builds the file for all three companies, reads each one back from disk and prints "All checks passed". `check_companies.py` and `check_northwind.py` still pass. Two commits: `4942059` (writer) and `85abfd5` (check + LEARNINGS entry). I made no changes to `clean.py`, `metrics.py`, `analyze.py`, `config.yaml` or `CLAUDE.md`.

### What I built

| File | What it is, in plain English |
|---|---|
| `excel_output.py` (new) | **Turns the computed numbers into a spreadsheet a partner can open.** It does no math: it calls the same `clean_workbook` → `compute_metrics` → `evaluate_flags` → `data_gaps` chain as everything else, then only writes and formats. `number_format` picks how a column displays (`0.0%`, `#,##0`, `0.0" mo"`, `0.00"x"`). `cell_value` decides what goes in a cell: the number itself, or a word when there's no usable number. `write_metrics_sheet`, `write_flags_sheet` and `write_gaps_sheet` build one sheet each. `build_workbook` puts them together, and `save_metrics_workbook(path, config)` is the one function `main.py` will call in Task 3. |
| `check_excel_output.py` (new) | **Proves the spreadsheet says exactly what the metrics table says.** For each company it saves the file, **reloads it from disk** and checks: (1) the file name and sheet order; (2) every Metrics cell against the metrics table, its number format, and its fill (red only where a flag trips that quarter, gray only where data is missing, nothing else); (3) latest-quarter cells against the **hand formulas** already in `check_companies.py`; (4) every Flags row's value, threshold, status text and row color against each company's **story** (the expected statuses in `check_companies.py`, not recomputed); (5) the Data gaps rows against `data_gaps()`, no more and no fewer. |
| `LEARNINGS.md` (changed) | One new row: openpyxl's number precision (see "What failed"). |

**What each file shows:**

| Sheet | Northwind | Alderpeak (healthy) | Fernhollow (distressed) |
|---|---|---|---|
| Metrics | 8 rows; Q1 2025 all gray; red cells wherever a flag tripped | 8 rows, no red or gray anywhere | 8 rows; Q2 2025 all gray; lots of red |
| Flags (Q2 2026) | 6 red, 3 green; runway at budget 13.0 mo | 9 green; runway at budget 144.0 mo | 7 red, 1 gray (Rule of 40), 1 green; runway at budget 7.6 mo |
| Data gaps | 19 rows | "None — every metric and flag has the data it needs" | 20 rows (incl. "Flag: Rule of 40") |

### Decisions you didn't specify

1. **File name comes from the input file name:** `data/northwind.xlsx` → `output/northwind_metrics.xlsx`. That needs no company-name lookup, and `--all` in Task 3 gets unique names for free.
2. **The Metrics sheet holds the 19 computed metrics only** (the same columns as `compute_metrics`), not raw inputs like net burn or cash. That keeps it "one row per quarter of metrics" and makes the read-back check a 1-to-1 comparison.
3. **Column headers reuse `METRIC_LABELS` from analyze.py** (imported, not edited), so Excel, the Claude payload and later the deck use the same names. Side effect: `excel_output.py` imports the `anthropic` package even though it never calls the API (see unresolved item 5).
4. **No value = a word, never an empty cell.** Excel can't store NaN or infinity, and openpyxl silently saves both as a blank. So each case gets its own label: `data missing` (a gap), `n/a (no prior period)` (e.g. YoY in the first 4 quarters), `∞ (ARR shrank)` (burn multiple), `∞ (not burning)` (runway), `∞ (never pays back)` (CAC payback). This follows "not meaningful must never look like data missing". One nice consequence: a formula like `=D9-D8` on a "data missing" cell shows `#VALUE!` instead of quietly treating it as 0.
5. **The Metrics sheet highlights flagged cells in every quarter.** CLAUDE.md step 4b says "flagged cells highlighted", and `evaluate_flags` already accepts any quarter. A cell is red if its flag trips in *that* quarter and gray if it's a data gap. Passed cells are **not** green: 150 green cells would drown out the red. The combo rule has no single cell, so it appears only on the Flags sheet.
6. **The Flags sheet has 2 columns beyond the 4 you listed:** `Quarter` (so a printout is self-explanatory) and `Trips when` ("below threshold" / "above threshold"). Without it, a threshold of −20.0% doesn't say which direction is bad. Columns: Flag, Quarter, Value, Threshold, Trips when, Status.
7. **The whole flag row is colored**, not just the Status cell, with darker text in the same color (Excel's standard "Good/Bad" light red `FFC7CE` and light green `C6EFCE`; gray `D9D9D9`).
8. **Status wording:** "Tripped", "Passed", "Cannot evaluate — data missing". The payload to Claude uses "TRIPPED" / "passed"; for a spreadsheet I used sentence case.
9. **Combo rule row:** Value = "see NRR and Pipeline on Metrics sheet", Threshold = "last 3 quarters" (read from config), Trips when = "NRR falls and pipeline rises at every step".
10. **Value and Threshold stay real numbers** with the metric's format, so a threshold of 0.15 displays as 15.0% but can still be used in formulas.
11. **Runway at next quarter's budgeted burn** is on the Flags sheet, one empty row below the table, labeled "(context, not a flag)" and uncolored. CLAUDE.md says to show it as context without flagging it.
12. **Data gaps sheet:** one row per affected metric or flag, with quarters joined as "Q1 2025, Q2 2025", in the same order as `data_gaps()`. Flag gaps read "Flag: Rule of 40".
13. **$K cells use `#,##0` with no `$` sign**, because the header already says "($K)", matching metrics.py's printout.
14. **Small usability touches:** bold headers, frozen header row and quarter column, fixed column widths, and right-aligned text labels in number columns so they line up.
15. **The check writes the real files into `output/`** (git-ignored) instead of a temp folder, so after any check run the latest files are there to open.
16. **The check writes its expected labels, formats and colors out literally** instead of importing them from `excel_output.py`. If a constant there is wrong, the check can't agree with it by accident.
17. **The CLI takes one file.** Batch mode is Task 3's `main.py`, which can loop over `save_metrics_workbook`.

### What failed and how I fixed it

1. **The read-back check failed on its first run.** NRR `1.0706921944035346` came back as `1.070692194403535`. openpyxl saves numbers with 16 significant digits (`"%.16g"`), but a Python float can need 17. The difference is about 1e-16, and Excel itself only keeps 15 digits. **Fix:** the check compares to 15 significant digits (`same_as_saved`). I didn't change the writer, because nothing is lost at any precision Excel can show. Logged in LEARNINGS.md, together with openpyxl silently writing NaN or infinity as a blank cell.
2. **Checks passing proves nothing by itself, so I broke the writer on purpose** (in memory, without editing files) and reran the check. **13 of 13 caught:** ratios saved as 97.1 instead of 0.971, values rounded to 3 decimals, `%` format missing, "data missing" written as "n/a", "no prior period" written as "data missing", ∞ written as "n/a", tripped cells not red, passed rows gray, a changed status label, one data gap dropped, a wrong threshold, runway at budget off by 1, and sheets in the wrong order. The first break is the one the 15-digit tolerance needed to keep catching, and it does.
3. **Minor tooling:** the shell sandbox blocked a `for` loop and a `sed` pipe, so I ran the three companies as separate commands and read full dumps instead. No effect on the project. `pytest` reports "no tests ran" (exit code 5) because `tests/` doesn't exist until Task 4; the overnight runner only runs pytest when that folder exists.

### Unresolved: needs your call

1. **Flag names and metric headers don't quite match.** The Flags sheet uses `FLAG_RULES` names ("Burn vs budget", "Runway (months)", "CAC payback (months)"). The Metrics sheet uses `METRIC_LABELS` ("Net burn vs budget", "Runway at current burn", "CAC payback"). Both come from existing code I wasn't asked to change. One set of names before the deck is built would be cleaner.
2. **Burn multiple = 0 when not burning shows as "0.00x"** with no "not burning" note, while runway in the same situation shows "∞ (not burning)". It's a real number, so I kept it numeric. An Excel cell comment could explain it. No company in the data hits this case.
3. **The Task 1 "burn vs budget" edge case carries into Excel.** A company generating more cash than budgeted would get a red cell. It's a definition decision in metrics.py, not an output one.
4. **I haven't opened the files in Excel.** Everything was verified by reading the saved file back with openpyxl: values, formats, fills. Column widths and how the colors look should get a quick visual check, which also gives step 6 a before/after screenshot.
5. **`excel_output.py` depends on analyze.py for labels**, which pulls in the `anthropic` package. It's harmless (it's in requirements.txt and no API call happens), but if you want the Excel step to run without any AI code, `METRIC_LABELS` could move to metrics.py. That means editing analyze.py's import line (not the prompt), so I left it.
6. **The Flags sheet covers the latest quarter only**, as specified. Earlier quarters' trips show only as red cells on the Metrics sheet; the combo rule's history isn't shown anywhere.

---

## Task 3 — Batch runner (main.py, no AI yet)

**Result:** done. `python main.py --all --skip-ai` runs all three companies and exits 0. The new `check_main.py` prints "All checks passed", and so do `check_northwind.py`, `check_companies.py` and `check_excel_output.py`. Two commits: `4f1a989` (main.py) and `0deef52` (check). I made no changes to `clean.py`, `metrics.py`, `excel_output.py`, `analyze.py`, `config.yaml` or `CLAUDE.md`, and made no API calls.

### What I built

| File | What it is, in plain English |
|---|---|
| `main.py` (new) | **The one command that runs the whole pipeline.** `python main.py data/northwind.xlsx` runs one company; `python main.py --all` runs every workbook in `data/`. For each company it cleans the workbook, computes metrics, flags and gaps, saves the Excel file, then reaches the AI and deck steps and prints why they were skipped. `find_workbooks` lists the input files. `run_company` runs the steps for one company and prints a ✓ line per step. `run_batch` loops over the companies. A failure is caught, printed as `✗ FAILED: <reason>`, and the loop moves on. `print_summary` prints the final table. `main` returns exit code 0 if everything worked and 1 if any company failed. |
| `check_main.py` (new) | **Proves the runner works and survives bad input.** (1) It runs `main.py --all --skip-ai` as a real command and checks that it exits 0, finds exactly the 3 workbooks, saves a fresh Excel file for each, prints both skip messages 3 times, and prints a summary table that matches each company's **story** from `check_companies.py`, not main.py's own counting. (2) It puts a broken workbook, a missing file and a simulated code bug between good companies and checks that the good companies still succeed. (3) It checks exit codes: 1 when a company fails, 2 for wrong arguments. (4) It checks that the run fails loudly if `build_deck.py` exists but isn't wired in. (5) It checks that `--all` ignores Excel lock files. |

**Output of `python main.py --all --skip-ai`, summary part:**

```
Company     Flags tripped              Data gaps                          Result
----------  -------------------------  ---------------------------------  ----------------------
Alderpeak   0 of 9                     none                               OK (AI + deck skipped)
Fernhollow  7 of 9, 1 cannot evaluate  20 metrics/flags (blank: Q2 2025)  OK (AI + deck skipped)
Northwind   6 of 9                     19 metrics/flags (blank: Q1 2025)  OK (AI + deck skipped)

3 of 3 companies succeeded
```

Above the table, each company gets a block like `✓ Cleaned: 8 quarters (Q3 2024 to Q2 2026), budget row: Q3 2026 (Budget)`, the names of the tripped flags, `✓ Excel: output/northwind_metrics.xlsx`, `- AI commentary: skipped (--skip-ai)` and `- Deck: skipped (build_deck.py doesn't exist yet - build step 4)`.

### Decisions you didn't specify

1. **How I read "AI and deck steps are skipped if build_deck.py does not exist":** while `build_deck.py` is missing, **both** steps are skipped even without `--skip-ai`. The commentary has nowhere to go until there's a deck, so spending API money on it would be waste. As a result, main.py can't call the API tonight no matter which flags you pass. For now `--skip-ai` only changes the message ("skipped (--skip-ai)" instead of "skipped (build_deck.py doesn't exist yet)"). Once the deck exists, it will skip the Claude call.
2. **AI and deck are not wired up, and main.py fails loudly if `build_deck.py` shows up.** I didn't write code that calls `analyze()`, because it couldn't be tested tonight without the API, and I don't know build_deck's functions yet. Once `build_deck.py` exists, every company fails with `NotWiredError: build_deck.py exists but main.py doesn't call it yet`. This happens even with `--skip-ai`, because the deck step isn't wired either. That way nobody gets an "OK" with no deck. **Heads-up for step 4:** creating `build_deck.py` will make `check_main.py` fail until main.py calls it. That's intended, but you'll need to update `ai_step`, `deck_step`, `result_text` and `EXPECTED_OK` in the check together.
3. **Company name = file name in title case** (`northwind.xlsx` → "Northwind"), the same rule `analyze.py` uses. The workbooks don't have a reliable company-name cell.
4. **"Flags tripped" column:** "6 of 9" (tripped out of all flags, including the combo rule). When some flags can't be evaluated, it says so: "7 of 9, 1 cannot evaluate". Otherwise Fernhollow's "7 of 9" would hide that a 9th flag had no answer. The per-company block above the table lists the tripped flag names.
5. **"Data gaps" column:** the number of metrics and flags affected, the same count as the Excel "Data gaps" sheet, plus the blank quarter(s) that caused them: "19 metrics/flags (blank: Q1 2025)", or "none". A blank quarter here means any quarter with at least one blank input.
6. **"Result" column:** `OK (AI + deck skipped)` while there is no deck, `OK` once there is, and `FAILED: <error type>: <message>` otherwise. A failed company shows "-" for flags and gaps, because there's nothing reliable to count.
7. **Exit codes:** 0 = every company OK; 1 = at least one failed (the batch still finishes first); 2 = wrong arguments (argparse's standard). A scheduler or the overnight script can tell from the code alone whether a run worked.
8. **Arguments:** exactly one of a file path or `--all`. Both or neither gives "give one workbook path, or --all (not both)". Only one path is accepted. For several, use `--all`.
9. **A missing file is a company failure, not an argument error.** `main.py data/nope.xlsx` goes through the normal path and shows `FAILED: FileNotFoundError ...` in the table, so the batch path and the single-file path behave the same.
10. **Bad input vs. code bugs:** `ValueError` (clean.py's messages, and pandas' "not an Excel file") and `OSError` (missing or unreadable file) print one clear line. Any other error is probably a bug, so its full traceback is printed as well. The batch continues either way.
11. **`--all` reads `data/*.xlsx` only, sorted by name,** and skips files starting with `~$`. Excel creates those lock files while a workbook is open, and they would otherwise show up as a failing company. `.xls`, `.csv` etc. are ignored.
12. **Nothing is written except the Excel file.** I didn't add a log or summary file; the table goes to the terminal only.
13. **The check runs `main.py` as a separate process** for the batch, exit-code and argument tests, which is how you'll actually use it. The failure-isolation tests call `run_batch` directly with temporary files, so no broken workbook is ever put in `data/`. The simulated code bug and the fake `build_deck.py` are swapped in only in memory or in a temp folder, then put back. The real project files are never touched.
14. **Expected table values in the check come from the stories**, not from main.py: flag counts are counted from `expected_flags`, and gap counts come from `expected_gaps()` in `check_companies.py`.

### What failed and how I fixed it

1. **Nothing in the project code failed.** The check passed on its first run, and a first pass proves nothing by itself. So I broke `main.py` on purpose **14 ways**, each in a throwaway copy of the project in a temp folder, and ran `check_main.py` against each copy. **14 of 14 caught:** the batch stopping at the first failure, exit code always 0, "cannot evaluate" dropped from the flags text, gap count off by one, lock files not skipped, the deck silently skipped when `build_deck.py` exists, the Excel file not saved, tracebacks printed for input errors, no traceback for real bugs, both a file and `--all` accepted, "cannot evaluate" counted as tripped, `--skip-ai` ignored, the blank quarter not shown, and a failed company shown as OK.
2. **Wording, fixed before committing:** the first version printed "skipped: --skip-ai" and "Flags (Q2 2026): 0 of 9", which were hard to read. They now say "skipped (--skip-ai)" and "Flags tripped (Q2 2026): 0 of 9".
3. **Minor tooling:** the shell sandbox blocked `echo $?`, a heredoc and chained commands, so I made the edits with the file editor and wrote scratch scripts to `/tmp`. No effect on the project.

### Unresolved: needs your call

1. **When `build_deck.py` exists, main.py needs a real AI + deck step** (see decision 2). Two questions for step 4. Should a company that fails AI validation after the retry still get a deck with the commentary slide marked "AI summary unavailable", or should it FAIL? And should `--skip-ai` build a deck without the Summary slide, or skip the deck too?
2. **Does `--skip-ai` mean anything tonight?** Since AI is skipped anyway without a deck, the flag only changes the message. If you meant "without `--skip-ai`, call Claude now even without a deck" (saving `output/<company>_analysis.json`), that's a small change to `ai_step`. It wasn't allowed tonight because it calls the API.
3. **The summary table only goes to the terminal.** For 275 companies you'd probably want it saved too, e.g. `output/batch_summary.xlsx` or a CSV with one row per company. I didn't add it because you didn't ask for it.
4. **Runs are one company at a time.** That's fine for 3 companies and for the Python steps. With the Claude call added, 275 companies × ~36 s (the Sonnet timing in README) is about 2.75 hours in a row. Running several at once can wait until the AI step exists.
5. **The runner doesn't cross-check companies.** For example, it doesn't warn if one company's latest quarter is Q1 2026 while the others are Q2 2026. A board pack for a fund probably wants every company on the same quarter. It's easy to add as a warning line if you want it.
