# Overnight Report

## Task 1: Step 5 data: a healthy and a distressed company

**Result:** done. Both new workbooks are in `data/`. `check_northwind.py` and the new `check_companies.py` both print "All checks passed". Four commits, from `638498d` to `24c1dd4`.

### What I built

| File | What it is, in plain English |
|---|---|
| `make_data_common.py` (new) | The shared "workbook factory". The code that checks an answer key and writes a messy workbook used to live inside `make_data.py`. It moved here so all three companies use the same code. `check_true_data` confirms ARR and cash roll forward, money values are multiples of 10, and the mess settings point at real columns and quarters. `write_kpi_sheet` writes the table. `save_workbook` runs the check and then saves the file. |
| `make_data.py` (changed) | Now holds **only Northwind's** numbers, headers, text cells and notes, and hands them to `save_workbook`. I regenerated `northwind.xlsx` and compared it cell by cell with the previous file: **identical**. |
| `make_data_alderpeak.py` (new) | **Alderpeak Software, the healthy company.** ARR grows 47.5% YoY and NRR holds around 110%. Burn shrinks every quarter and stays under budget, so runway is 108 months. Rule of 40 is 42.7%. **0 of 9 flags trip in Q2 2026, and no flag trips in any earlier quarter either.** |
| `make_data_fernhollow.py` (new) | **Fernhollow Systems, the distressed company.** ARR shrinks, NRR is 77.9%, GRR is 74.7% and runway is 6.0 months. Burn is 19.6% over budget and CAC payback is 134 months. **7 of 9 flags trip.** Q2 2025 is blank, so Rule of 40 is "cannot evaluate: data missing", and the combo rule passes (see decisions). |
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
| Rule of 40 | trip | pass (42.7%) | **cannot evaluate: data missing** |
| NRR falling while pipeline rising | trip | pass | pass |
| **Total** | 6 trip, 3 pass | 0 trip, 9 pass | 7 trip, 1 pass, 1 cannot evaluate |

### Decisions you didn't specify

1. **Company names:** "Alderpeak Software" (healthy) and "Fernhollow Systems" (distressed), picked to sound invented. Files are `make_data_alderpeak.py` / `data/alderpeak.xlsx` and `make_data_fernhollow.py` / `data/fernhollow.xlsx`.
2. **Shared code instead of copies:** I moved the writing and checking functions into `make_data_common.py` rather than pasting ~100 lines into each new file. Each company file is now just data plus a short `main()`. The functions take one `company` dictionary instead of reading global variables.
3. **Same 8 quarters for everyone (Q3 2024 – Q2 2026):** a fund's board pack covers the same quarter for every company, so the batch run in step 5 lines up.
4. **Alderpeak has no blank quarter.** A well-run company with clean reporting should show "Data gaps: none", and this tests the zero-gaps path, which Northwind can't. The rest of the mess is still there: odd headers, text numbers, a budget-only row and a notes tab. To add a blank quarter, change `BLANK_QUARTER = None` to a quarter label. Its flags would still pass, as long as the blank isn't Q2 2025 or one of the last 3 quarters.
5. **Fernhollow's blank quarter is Q2 2025, exactly 4 quarters before the latest.** That puts the YoY gap on Q2 2026, so a latest-quarter flag (Rule of 40) returns "cannot evaluate: data missing". Northwind's blank quarter never reaches its latest flags, so this was the one CLAUDE.md rule the real data didn't exercise yet.
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
3. **"Cannot evaluate" flags display their value as "n/a".** In both `metrics.py`'s printout and the analyze payload, Fernhollow's Rule of 40 value is "n/a" next to the status "cannot evaluate: data missing". The status is right, but "n/a" on its own reads like "not applicable", and CLAUDE.md says missing must never look like not meaningful. Best fixed when building the deck (step 4). I didn't touch analyze.py.
4. **Negative percentages will be hard to read on a slide.** Fernhollow's net new ARR vs budget is −150.0% (−$240K actual vs +$480K budget). The math is correct, but the deck may want to show the $K figures alongside it.
5. **Not done here, by design:** no analysis was run for the new companies (no API calls), and CLAUDE.md wasn't updated with their stories. The stories are in each `make_data_*.py` docstring. You may want a "Healthy / distressed company story" section in CLAUDE.md like the Northwind one.
6. **Overlap:** Northwind's cleaning, flag and gap checks now run in both `check_northwind.py` and `check_companies.py`. It's harmless, but if you'd rather have one home, the four Northwind checks could be deleted from `check_northwind.py`.

---

## Task 2: Step 4b Excel output

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
| Data gaps | 19 rows | "None: every metric and flag has the data it needs" | 20 rows (incl. "Flag: Rule of 40") |

### Decisions you didn't specify

1. **File name comes from the input file name:** `data/northwind.xlsx` → `output/northwind_metrics.xlsx`. That needs no company-name lookup, and `--all` in Task 3 gets unique names for free.
2. **The Metrics sheet holds the 19 computed metrics only** (the same columns as `compute_metrics`), not raw inputs like net burn or cash. That keeps it "one row per quarter of metrics" and makes the read-back check a 1-to-1 comparison.
3. **Column headers reuse `METRIC_LABELS` from analyze.py** (imported, not edited), so Excel, the Claude payload and later the deck use the same names. Side effect: `excel_output.py` imports the `anthropic` package even though it never calls the API (see unresolved item 5).
4. **No value = a word, never an empty cell.** Excel can't store NaN or infinity, and openpyxl silently saves both as a blank. So each case gets its own label: `data missing` (a gap), `n/a (no prior period)` (e.g. YoY in the first 4 quarters), `∞ (ARR shrank)` (burn multiple), `∞ (not burning)` (runway), `∞ (never pays back)` (CAC payback). This follows "not meaningful must never look like data missing". One nice consequence: a formula like `=D9-D8` on a "data missing" cell shows `#VALUE!` instead of quietly treating it as 0.
5. **The Metrics sheet highlights flagged cells in every quarter.** CLAUDE.md step 4b says "flagged cells highlighted", and `evaluate_flags` already accepts any quarter. A cell is red if its flag trips in *that* quarter and gray if it's a data gap. Passed cells are **not** green: 150 green cells would drown out the red. The combo rule has no single cell, so it appears only on the Flags sheet.
6. **The Flags sheet has 2 columns beyond the 4 you listed:** `Quarter` (so a printout is self-explanatory) and `Trips when` ("below threshold" / "above threshold"). Without it, a threshold of −20.0% doesn't say which direction is bad. Columns: Flag, Quarter, Value, Threshold, Trips when, Status.
7. **The whole flag row is colored**, not just the Status cell, with darker text in the same color (Excel's standard "Good/Bad" light red `FFC7CE` and light green `C6EFCE`; gray `D9D9D9`).
8. **Status wording:** "Tripped", "Passed", "Cannot evaluate: data missing". The payload to Claude uses "TRIPPED" / "passed"; for a spreadsheet I used sentence case.
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

## Task 3: Batch runner (main.py, no AI yet)

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

---

## Task 4: Unit tests (tests/ with pytest)

**Result:** done. `python -m pytest -q` runs **147 tests, all passing, in about 1 second** (69 for clean.py, 78 for metrics.py). I found **one real bug** in clean.py (a repeated quarter row silently overwrote the earlier one), fixed it and logged it in LEARNINGS.md. After the fix, `check_northwind.py`, `check_companies.py`, `check_excel_output.py` and `check_main.py` all still print "All checks passed". Two commits: `0c0ab7f` (pytest setup, clean.py tests, bug fix) and `8993d7d` (metrics.py tests). metrics.py, config.yaml, CLAUDE.md and analyze.py are unchanged. No build_deck.py, no API calls.

### What I built

| File | What it is, in plain English |
|---|---|
| `pytest.ini` (new) | **Two settings for pytest.** `testpaths = tests` makes pytest look only in `tests/`. `pythonpath = .` lets the tests `import clean` and `import metrics` from the project folder. Without it, plain `pytest` can't find them. |
| `tests/test_clean.py` (new, 69 tests) | **Tests for clean.py.** `parse_number`: "$12.5M", "120K", "5,090", lowercase units, stray spaces, negatives, numbers that are already numbers (always returned as a plain float), blank cells becoming NaN (never 0), and unreadable text ("n/a", "(120)", "12.5%") stopping with an error. `normalize_header` and `standard_column`: every messy header from the Northwind workbook, plus unknown headers ("EBITDA", "Quarter") stopping. Quarter labels: good and bad formats. Quarter order: the Q4 → Q1 year rollover, skipped, repeated, reversed, a whole year skipped. Three tests write a tiny workbook to a temp folder. A blank quarter row is kept as all-NaN, a missing row stops, and a duplicate row stops (the bug below). |
| `tests/test_metrics.py` (new, 78 tests) | **Tests for metrics.py.** **Every metric function** is checked against hand math written in the comment beside it. **Edge cases:** burn multiple with zero or negative net new ARR (∞), zero or negative burn (0, including 0 burn with 0 net new ARR, which would otherwise be 0/0), runway when not burning (∞), CAC payback with no new ARR or zero/negative margin (∞). A missing input stays NaN and never turns into 0 or ∞. **`check_threshold`:** exactly at the threshold passes for both min and max. The next possible float above or below a threshold passes, so float noise is ignored. A real Rule of 40 of exactly 40% (Python computes 0.39999999999999997) passes. A real miss of 0.000001 still trips. NaN → cannot evaluate. ∞ passes a minimum (runway) and trips a maximum (burn multiple, CAC). **`check_combo`:** trips, several passes (NRR flat or rising in one step, pipeline flat or falling), noise-level NRR changes counted as flat, older quarters outside the window ignored. **Cannot evaluate:** a missing NRR at each window position, missing pipeline, missing data where the rest would have passed, not enough history. **`data_gaps`:** no blanks → no gaps. A blank quarter matches the CLAUDE.md rules exactly: its own quarter, QoQ = blank + next, YoY = blank + 4 later. Also: a blank first quarter, a blank recent quarter that makes flags "cannot evaluate", a blank year-ago quarter that hits the Rule of 40 flag, and **one blank cell spreading only to the metrics that use it**. |

### Decisions you didn't specify

1. **Expected values are hand-worked from the CLAUDE.md definitions and written as formulas in comments** (e.g. `# 600 / (400 * 750 / 1000) * 12 = 24 months`). None are copied from running the code. That's the LEARNINGS lesson from the net new ARR vs budget mistake: an answer key copied from the code can't catch the code.
2. **Tests use small made-up tables, not the three company workbooks.** The `check_*.py` scripts already cover the real companies end to end. Unit tests isolate one function each, so a failure points at one line of math. Each table has only the columns that function reads, with round numbers you can check in your head.
3. **Thresholds are written into the test file (`TEST_CONFIG`) instead of read from config.yaml**, so tuning a threshold later never breaks a test. One separate test reads config.yaml (without changing it) and checks that every flag rule's threshold key exists, which catches a typo there.
4. **"Exactly at the threshold" is tested two ways:** with `math.nextafter` (the smallest possible float step above or below the threshold) and with real inputs that produce noise (revenue 600 → 1020, burn 306: exactly 40%, computed as 0.39999999999999997). I also checked the other side: 0.150001 against a 0.15 maximum still trips, so the rounding doesn't hide real misses.
5. **I tested `evaluate_flags` briefly (2 tests)** even though you didn't list it, because `data_gaps` depends on its output. Most `data_gaps` tests run the real chain (`compute_metrics` → `evaluate_flags` → `data_gaps`). One test passes hand-built flags to check the "flag: <name>" format on its own.
6. **I did not write tests for behavior I think is questionable** (see Unresolved 2–5). A test would lock that behavior in as correct. The report lists those cases instead.
7. **One test turns warnings into errors** (`runway_at_next_budget` with zero budgeted burn). Without that, numpy quietly divides by zero, returns ∞ with only a warning, and hides whether the zero-burn case is handled on purpose.
8. **Three clean.py tests build a real workbook in a temp folder** (blank row kept, missing row, duplicate row), because the duplicate bug only exists in `clean_workbook`, not in the helper functions. Nothing is written to `data/`. Task 5 (bad inputs) can reuse the `write_workbook` helper.
9. **Test names say what should happen** (`test_burn_multiple_missing_input_is_nan`), and edge-case tables carry a "why" note that pytest prints if the test fails.

### What failed and how I fixed it

1. **Real bug in clean.py: a repeated quarter silently lost data.** With rows Q1, Q2, Q2, Q3 2025, the second Q2 row replaced the first with no error. Rows are stored in a dictionary keyed by the quarter label, so the second save overwrote the first. The dictionary then held only one Q2, so `check_quarters_in_order` (which does catch repeats) never saw it. **Fix:** 3 lines in `clean_workbook` that stop with "Quarter 'Q2 2025' appears twice - check the workbook". None of the three workbooks has a duplicate, so no output changed. The four check scripts still pass. Logged in LEARNINGS.md.
2. **One of my own test comments was wrong.** I wrote that `21.85 * 1000` gives 21849.999… in plain Python. I checked, and it gives exactly 21850.0. I searched for a value where plain float math really is off and switched the test to "$4.03M" (4.03 × 1000 = 4030.0000000000005), so the test proves why clean.py uses `Decimal`.
3. **All 147 tests passed on the first run, which proves nothing by itself, so I broke the code on purpose** in a throwaway copy (36 breaks, each run against the tests; the real files were never touched). **First round: 32 of 36 caught.** Of the other 4:
   - One was a typo in my break script. Once fixed, it was caught.
   - **Two were real holes in the tests, now fixed.** (a) CAC payback checked with `< 0` instead of `<= 0` went unnoticed. New ARR of 0 with a negative gross margin gives `0 × negative = −0.0`, and dividing by that gives **−∞** instead of ∞. I added that case. (b) Runway at budget checked with `< 0` instead of `<= 0` went unnoticed because numpy returned ∞ anyway. Hence decision 7.
   - One can't be caught: burn multiple checked with `new < 0` instead of `new <= 0` gives the same answer, because burn ÷ 0 is already ∞ in pandas.
   
   **Final: 35 of 36 caught**, and the one left can't change any result. Breaks included: no rounding in `check_threshold`, trip at exactly the threshold, NRR/GRR not annualized, FCF margin sign flipped, net new ARR vs budget using gross `budget_new_arr`, combo ignoring missing pipeline, flat NRR counted as falling, data gaps ignoring blank quarters, gaps counting "no prior period" as missing, a lookback entry deleted, no `Decimal`, no Q4 → Q1 rollover, header aliases ignored, and the duplicate-quarter check removed.
4. **Minor tooling:** the shell sandbox blocked running the `pytest` program directly and output pipes, so I used `python -m pytest -q` (the same command `output/overnight.sh` uses) and wrote scratch scripts to `/tmp`. No effect on the project.

### Unresolved: needs your call

1. **A partly blank quarter among the first 4 quarters over-reports data gaps.** I checked it: if only `pipeline` is blank in Q3 2024, `data_gaps` also lists arr_qoq, arr_yoy, revenue_qoq, revenue_yoy, rule_of_40 and net_new_arr_vs_budget as "data missing" for Q3 2024. Their real reason is "no prior period", and their own inputs are fine. `data_gaps` treats a quarter with *any* blank input as incomplete for *every* metric. That breaks the CLAUDE.md rule that "not meaningful" must never look like "data missing", though it errs toward saying too much, not hiding a gap. None of the three companies hits it: their blank quarters are fully blank, and a fully blank quarter's gaps are correct. A proper fix needs a list of which inputs each metric uses, which is a design change, so I left it. I didn't add a test for the current behavior.
2. **`parse_number` accepts some text it probably shouldn't.** Python's `Decimal` reads "inf" / "Infinity" as ∞, "nan" as blank, and "1e3" as 1000. A TRUE cell counts as a number, so it becomes 1.0. A typo cell reading "inf" would put ∞ on a board slide instead of stopping. None of the workbooks has these. The fix is small (reject non-finite results and booleans), but it's behavior you didn't ask me to change, and Task 5 (bad inputs) is the natural place for it.
3. **Net new ARR vs budget breaks when budgeted net new ARR is zero or negative** (budget_arr flat or falling). Dividing by a negative flips the sign: actual −100 vs budget −200 comes out as −50% (a miss) when it actually beat plan. Dividing by zero gives ∞. The lowest budgeted net new ARR in the three companies is 470 (Fernhollow), so nothing is affected today. It's a definition question, like the "burn vs budget" edge case from Task 1.
4. **`combo_lookback_quarters` below 2 makes the combo rule always trip.** I checked with 1: there are no quarter-over-quarter steps, and "fell in every step" is then true by default. config.yaml says 3, so it doesn't happen today. A guard that stops on a value below 2 would prevent it. I didn't add one because the task said not to change behavior.
5. **Not burning + blank ARR or cash inputs:** burn multiple shows 0 and runway shows ∞ (matching CLAUDE.md's "not burning" rules), even when net new ARR or ending cash for that quarter is blank. The answer really doesn't depend on the blank value, so I think it's correct, but it's a case where a metric with a blank input doesn't say "data missing". I didn't test it either way.
6. **Overlap with Task 5:** my missing-row and duplicate-row workbook tests touch "quarters out of order", which Task 5 also lists. Task 5 can keep them or move them into its bad-input test file.

---

## Task 5: Bad inputs (broken workbooks must stop with a clear message)

**Result:** done. `python -m pytest -q` runs **212 tests, all passing, in about 2 seconds**: 53 new in `tests/test_bad_inputs.py`, 81 in `test_clean.py` (12 added), and 78 in `test_metrics.py` (unchanged). I found **one real bug**: "n/a"-style words and Excel error cells were silently read as blank. It's fixed and logged in LEARNINGS.md. `check_northwind.py`, `check_companies.py`, `check_excel_output.py` and `check_main.py` all still print "All checks passed", so the three real workbooks read exactly as before. Two commits: `194efc4` (clearer clean.py errors) and `81ff917` (bad-input tests + bug fix). Only clean.py, the two clean test files and LEARNINGS.md changed. metrics.py, main.py, config.yaml, CLAUDE.md and analyze.py are unchanged. No build_deck.py, no API calls.

### What I built

**Every error from the KPI tab now starts with the sheet name, then an Excel address**, so you can go straight to the cell. Before → after:

| Broken input | Before | After |
|---|---|---|
| Missing column | `Workbook is missing columns: ['pipeline']` | `Sheet 'KPI Tracker', row 3 (header) is missing columns: pipeline` |
| Two headers, one meaning | `Two headers both mean 'starting_arr' - check the workbook` | `Sheet 'KPI Tracker', row 3 (header): columns B ('Starting ARR') and R ('Beginning ARR') both mean 'starting_arr' - keep one and delete the other` |
| Unknown header | `Unknown column header 'EBITDA' - add it to HEADER_ALIASES` | `Sheet 'KPI Tracker', cell R3 (header): Unknown column header 'EBITDA' - if it means one of the 16 input columns, add it to HEADER_ALIASES in clean.py; otherwise delete the column` |
| Quarters out of order | `After Q1 2025 expected Q2 2025, found Q3 2025` | `Sheet 'KPI Tracker', row 5: After Q1 2025 expected Q2 2025, found Q3 2025 - quarters must run oldest to newest with none skipped or repeated (a quarter with no data still needs its own row)` |
| Budget row with actuals | `Budget row 'Q1 2026 (Budget)' also has actual values in: ['revenue']` | `Sheet 'KPI Tracker', row 8 ('Q1 2026 (Budget)') is labelled as a budget-only row but also has actual values in G8 (revenue) - a budget-only row may fill only budget_new_arr, budget_arr, budget_net_burn` |
| Unreadable text | `Q2 2025, revenue: Can't read 'TBD' as a number` | `Sheet 'KPI Tracker', cell G5 (Q2 2025, revenue): Can't read 'TBD' as a number (expected e.g. 1250, '$1.2M' or '850K'; leave the cell empty if there's no data)` |
| Repeated quarter | `Quarter 'Q2 2025' appears twice - check the workbook` | `Sheet 'KPI Tracker', row 6: quarter 'Q2 2025' appears twice (also in row 5) - delete one of the rows` |
| Second budget row | `Workbook has more than one budget-only row` | `Sheet 'KPI Tracker', row 9 ('Q2 2026 (Plan)') is a second budget-only row (the first is 'Q1 2026 (Budget)') - keep only next quarter's budget` |

main.py already prints input errors as one line, e.g. `✗ FAILED: ValueError: Sheet 'KPI Tracker', cell G5 (Q2 2025, revenue): Can't read 'n/a' as a number (...)`, so the batch table picks these up with no change to main.py.

| File | What changed, in plain English |
|---|---|
| `clean.py` | **New small helpers:** `excel_row` and `excel_column` turn pandas' positions (counted from 0) into Excel's row 7 / column F. `at_row` adds "row 7: " to quarter-order errors. `header_name` treats "Quarter" like any other header, so two Quarter columns are caught by the same duplicate check. `name_headers` stops on unknown headers and repeated meanings. `check_no_headerless_values` and `check_unlabelled_row_is_empty` stop on numbers that would otherwise be skipped without a word. `check_no_error_cells` stops on #DIV/0!, #REF! and similar (the bug below). `clean_sheet` is the old body of `clean_workbook`. `clean_workbook` now finds the tab and adds `Sheet '<name>', ` in front of any error from that tab, all in one place. **`parse_number` is stricter:** number text must match one shape (`NUMBER_TEXT`: optional minus, digits with commas only between groups of 3, optional decimals, optional K/M), and TRUE/FALSE cells stop. |
| `tests/test_bad_inputs.py` (new, 53 tests) | **Builds small broken workbooks in pytest's temp folder** (never `data/`). Each workbook looks like a real one: a Notes tab first, then "KPI Tracker" with a title row, an empty row, the header in row 3, Q1–Q4 2025 in rows 4–7 and a budget row in row 8. That way the addresses in the messages are really tested. A column map at the top of the file lets you check every expected address by eye. Helpers: `good_table`, `set_cell`, `drop_column`, `add_column`, `write_workbook`, `assert_stops_with`. Test groups: **missing column** (one, several, no Quarter column), **duplicate meaning** (alias vs. standard name, exact repeat, budget alias, two Quarter columns), **unknown header** (3 names, how-to-fix wording, values under a blank header, table shifted right to column C), **quarter order** (swapped, skipped, newest first, wrong year rollover, repeated, unreadable labels, values with no label, header with no rows), **budget row** (one actual, several actuals listed, allowed columns named, second budget row), **unreadable text** (16 bad values, 4 Excel error values, a date, a bad budget-row cell, how-to-fix wording). Plus `test_good_workbook_cleans`, which proves the starting workbook is valid, so each failure comes from the one thing broken. |
| `tests/test_clean.py` | Added direct `parse_number` cases: "inf", "Infinity", "nan", "1e3", "1.250,5", "1,2,5", ".5M", "-", True and False stop; "$-1.2M" and "1,250,000" still read. |
| `LEARNINGS.md` | One row for the bug below. |

### Decisions you didn't specify

1. **Addresses use Excel's own terms:** "cell G5" for a single value, "row 5" for a row problem, "column R" / "cell R3 (header)" for a header, always after `Sheet '<name>'`. Cell errors also give the quarter and standard column name, e.g. `(Q2 2025, revenue)`, so the message makes sense without opening the file.
2. **Tests check how the message starts, not the full text.** The start (sheet, place, problem) is fixed. The how-to-fix wording at the end has 4 separate tests of its own, so it can be reworded without breaking 50 tests.
3. **Expected addresses are typed out by hand** from the layout in the test file's docstring, not copied from the code (same rule as Task 4).
4. **I made `parse_number` stricter, closing Task 4's unresolved item 2.** Text now needs a known shape before it's read. Newly rejected: "inf"/"Infinity" (was ∞), "nan" (was blank), "1e3" (was 1000), TRUE/FALSE (was 1/0), "1.250,5" (was silently 1.2505), "1,2,5" (was 125) and a typographic minus "−1.2M". **Also rejected now: ".5M" and "12."**, which are harmless but unusual. I chose to stop rather than guess. None of the three workbooks uses them.
5. **"n/a", "NA", "NULL", "None" in a number cell now stop instead of counting as blank** (see the bug below). They probably mean "no data", but CLAUDE.md says blank means an empty cell and never to guess. The message tells the person to clear the cell.
6. **Two new stops that the task didn't list, both in the "unknown header" family:** a column with values but **no header** (previously skipped silently), and a row with values but **no quarter label** (previously skipped silently, losing that quarter's numbers). Side effect: a stray comment typed to the right of the table or under it now stops the run too. None of the three workbooks has one.
7. **Also added:** two "Quarter" columns stop (before, the second one silently won), and a header with no quarter rows under it stops (before, clean.py returned an empty table without complaint and the problem surfaced later, in another file).
8. **Missing columns are listed together, in the standard order, as plain names** (`starting_arr, budget_arr`, not a Python list), so one fix round is enough.
9. **The "no 'Quarter' header" error lists the tabs it checked** (`tabs: ['Notes', 'KPI Tracker']`). It names the file rather than a sheet, because no sheet qualified.
10. **An Excel error anywhere in the KPI tab stops the run**, even outside the table. A broken formula in that tab is worth fixing whatever it feeds.
11. **Kept the Task 4 duplicate-row and missing-row tests where they are** in `test_clean.py`. Task 5 covers the same cases with row numbers, so both stay (they're quick).
12. **openpyxl was already in requirements.txt,** so no new package. clean.py now imports it for column letters and the error-cell scan.

### What failed and how I fixed it

1. **Real bug: pandas silently blanked "n/a"-style text and Excel error cells.** My first run of the new tests had 4 failures: "n/a", "nan" and "#DIV/0!" in a revenue cell did **not** stop. I checked in a scratch script. `pd.read_excel` turns a built-in list of words ("n/a", "NA", "NULL", "None", "nan"…) and every error cell into NaN before clean.py sees them. A "#DIV/0!" from a broken formula would have reached the deck as "data missing". Task 4's test said "n/a" stops, but it called `parse_number` directly, so it never went through pandas. **Fix:** read with `keep_default_na=False, na_values=[""]` (only a truly empty cell is blank; I checked that empty cells still come back as NaN). Plus `check_no_error_cells`, because error cells get blanked either way and only openpyxl still sees them. All four check scripts still pass. Logged in LEARNINGS.md.
2. **Checked the address math before writing it:** I wrote a workbook with the table starting at C4 and confirmed pandas keeps the empty rows and columns above and left of it, so Excel row = pandas index + 1 and column = position + 1. The test with the table shifted to column C locks this in.
3. **All tests passing proves nothing by itself, so I broke clean.py on purpose 18 ways,** each in a throwaway copy in a temp folder, and ran the full test suite against each. **18 of 18 caught:** pandas' default NA words back on, error-cell scan removed, row number off by one, column letter off by one, TRUE counted as a number, number-shape check removed, commas allowed anywhere, headerless and unlabelled values skipped again, duplicate-header check removed, sheet name dropped, row dropped from order errors, budget row ignoring actuals, second budget row allowed, duplicate quarter allowed, empty table allowed, missing columns listed in the wrong order, and a second Quarter column allowed. My first version of that last break was too blunt (it broke 41 tests for an unrelated reason), so I replaced it with the realistic one.
4. **Minor tooling:** the shell sandbox blocked a heredoc and a `for` loop, so I wrote scratch scripts to `/tmp/t5` and ran the check scripts one at a time. No effect on the project.

### Unresolved: needs your call

1. **A formula cell with no saved result reads as blank.** Excel always saves formula results, so this only happens with files written by scripts (e.g. openpyxl) and never opened in Excel. The quarter would show "data missing", not a wrong number. Catching it means opening the file a second time with formulas on. I left it.
2. **The budget-only row's label isn't checked.** "Q4 2026 (Budget)" after a Q2 2026 actual, or a budget row placed above the actuals, is accepted. Runway at next quarter's budgeted burn would then use the wrong quarter's budget without a word. The fix is small (the budget label must be the quarter after the last actual). You didn't ask for it, and it changes what counts as a valid workbook.
3. **If two tabs both have a "Quarter" header, the first tab wins silently** (e.g. a copied "KPI Tracker (old)" tab placed first). Stopping with "two tabs look like KPI tabs: ..." would be safer. Not added, for the same reason.
4. **Row numbers depend on pandas keeping the empty rows and columns at the top and left.** I verified this for openpyxl-written files. I didn't have a real Excel-saved file with unusual layout info to test. If a real workbook ever shows an address that's off, that's where to look. The error-cell message uses openpyxl's own address, so it's always right.
5. **Any text in the label column under the table stops the run**, e.g. a "Source: finance team" note row: `row 14: Can't read quarter label 'Source: finance team'`. That was true before Task 5 too; only the message is new. Real workbooks often have footnotes there. If you want to allow them, the rule would need to say which rows may be ignored, and that is a design decision.

---

## Task 6: Study guide (STUDY_GUIDE.md, no code changes)

**Result:** done. `STUDY_GUIDE.md` (about 930 lines) has the data flow in plain English, every file explained function by function, 25 interview questions with short answers, 10 trace-this-number exercises with answers at the bottom, and 5 functions to rewrite yourself, easiest first. **No code changed:** no `.py` file, `config.yaml` or `CLAUDE.md` was touched, `build_deck.py` doesn't exist, and nothing called the Anthropic API. `python -m pytest -q` still reports 212 passed, and `check_northwind.py` and `check_companies.py` still print "All checks passed". Five commits: `60db889` (data flow + file guide), `b9baf3b` (interview questions), `e2673f9` (exercises + answers), `77ad97b` (functions to rewrite), `ad0aa67` (fact-check fixes).

### What I built

| Section of STUDY_GUIDE.md | What it is, in plain English |
|---|---|
| 1. How to use this guide | Which commands are free to run (clean, metrics, main `--skip-ai`, pytest, checks) and which **cost money** (`analyze.py`, `compare_models.py run`). |
| 2. The data flow | One sentence, a table matching each step to what an analyst does by hand, a diagram of every file and what it passes on, then each step walked through with Northwind's numbers. Ends with the "Python computes, Claude interprets" rule and two reasons to give in an interview. |
| 3. Python words you'll see | 22 terms (DataFrame, NaN, `.shift`, `.mask`, `try`/`except`, `Decimal`…), each explained with an Excel comparison where one exists. |
| 4. Every file, function by function | A table per file: each function, what it does, and an example or why it exists. The metrics table shows each formula with **Northwind's Q2 2026 numbers worked through**. clean.py also gets the order the functions call each other. Covers config.yaml, clean, metrics, analyze, excel_output, main, the 4 make_data files, the 4 check scripts, compare_models, the 3 test files and the small files. |
| 5. 25 interview questions | Grouped as: the project (4), design decisions (7), what broke (5), the AI part (6), reliability and scale (3). Every LEARNINGS.md row, the prompt iterations, the model comparison and the validator gaps are covered. Each answer ends with **"Point to:"**, the file, function or LEARNINGS row to open if they dig deeper. |
| 6. 10 trace-this-number exercises | Each names a line you can see by running a free command and asks you to trace it from the workbook cell to the screen: NRR 97.1%, runway 11.0 mo, burn vs budget 20.0% (with the Notes-tab trap), burn multiple 2.35x, Rule of 40 −12.2%, net new ARR vs budget −19.0%, runway at budget 13.0 mo, "data missing" vs "n/a (no prior period)", the combo rule, and Fernhollow's "7 of 9, 1 cannot evaluate". |
| 7. 5 functions to rewrite yourself | `nrr` → `growth` → `check_threshold` → `parse_number` → `check_combo`. Each has the skill it teaches, a spec, hints, the traps the existing tests will catch, the exact pytest command, and a follow-up check. Starts with a safe routine: delete the body, rewrite it, run the tests, `git diff`, then `git checkout -- <file>`. |
| 8. Answers | Numbered chains for all 10: the `make_data.py` value, the workbook cell (e.g. E9 `"260K"`, J9 `"$14.3M"`), each clean.py and metrics.py function, the arithmetic, the flag decision and the display format. |

### Decisions you didn't specify

1. **Section order:** exercises (6), then rewrites (7), then answers (8). You asked for the answers "at the bottom", so they're the very last section and you can't see them by accident while doing the exercises.
2. **Two extra sections:** "How to use this guide" and a short Python glossary. Both are aimed at a finance student reading real pandas code for the first time.
3. **How far "function by function" goes:** every function in the pipeline files, `make_data_common.py`, the check scripts and `compare_models.py` gets its own row. The three company data files are described as constants plus `main()`, because that's all they contain. **Test files are described by their helper functions and test groups, not all 212 tests one by one.** The test names already say what they test, and a 212-row table would bury the helpers you actually need to understand.
4. **Explanations use Northwind's real numbers** wherever a formula appears, so each function has a concrete example you can check in the `metrics.py` printout.
5. **Exercise choice:** 9 from Northwind and 1 from Fernhollow. Fernhollow's is the only real data that reaches "cannot evaluate" and ∞ in the latest quarter, which Northwind can't show. Every exercise starts from a line printed by a free command.
6. **Every number and cell address was checked by running the code,** not copied from docs. I read the workbook cells with openpyxl and recomputed each formula in a scratch script in `/tmp` (not in the project, not committed).
7. **Interview answers are short (30–60 seconds spoken)** with a "Point to" line, and the guide tells you to practise them in your own words rather than memorize them. Q2 ("Why would a PE fund want this?") is my framing of the business case. It only quotes numbers that are in the project (the $14.65 per quarter for 275 companies).
8. **Rewrite picks:** 5 functions that are central to interview talking points and already have tests, so you get instant feedback. Each adds one skill: column math, `.shift`, if/else and float noise, text parsing and errors, windows and "can't tell". I left out `data_gaps` (too many ideas at once for practice) and the analyze.py functions (their tests are mixed into `check_northwind.py`, not pytest).
9. **The practice routine uses `git checkout -- <file>`** to restore the original, with a bold warning that it discards edits to that file.
10. **An independent fact-check before the report.** A separate read-only agent compared the whole guide with the code and the docs; see "What failed" item 3. It had no permission to edit or call the API.
11. **No LEARNINGS.md entry.** Nothing in the project broke. The errors below were mistakes in my draft of the guide, caught before you saw it.
12. **I didn't rerun `check_excel_output.py` or `check_main.py`.** No code changed, and both rewrite the files in `output/`. pytest, `check_northwind.py` and `check_companies.py` were run and pass.
13. **The guide is pinned to commit `aa8838b`** in its first lines, so it's clear which version of the code it describes.

### What failed and how I fixed it

1. **Tooling:** the shell sandbox blocked `source .venv/bin/activate`, a `PYTHONPATH=…` prefix and a `for` loop. I ran `.venv/bin/python` directly, and the scratch script adds the project folder to its import path itself. No effect on the project.
2. **A wrong count in my first draft:** I wrote that 20+ kinds of unreadable text are tested. `pytest --collect-only -k unreadable_text_stops` shows **17**. Fixed before the first commit.
3. **The fact-check found 9 wording errors (no wrong numbers), all fixed in `ad0aa67`:**
   - `compute_metrics` doesn't call `runway_at_next_budget`.
   - `NotWiredError` doesn't print a traceback.
   - `ai_step` still skips with `--skip-ai` even if `build_deck.py` exists.
   - The header tests cover most Northwind headers, not every one ("Revenue" and "Net Burn" aren't listed).
   - The Notes tab **is** loaded, just never used.
   - The "n/a for both" explanation named the wrong function and the wrong report item.
   - One trace skipped `header_name`.
   - The "Sheet '…'" prefix isn't added to the "no Quarter header" error.
   - `append_learning` itself doesn't enforce "once"; `command_score` does.

   It confirmed that every cell address, calculation, test count, cost, token count and LEARNINGS row reference is correct, and that all internal links work.

### Unresolved: needs your call

1. **The guide will go out of date** when step 4 adds `build_deck.py` and main.py starts calling analyze.py. Section 2, main.py's table in section 4, and Q1/Q25 will need updating. Step 6 (README) is a natural time to refresh it.
2. **Found while tracing, not changed (no code changes allowed):**
   - **main.py reads each workbook twice.** `run_company` calls `clean_workbook`, and `save_metrics_workbook` calls it again. It's harmless at this size and a one-line change when the deck step is wired in.
   - **The number check can't catch a flipped sign.** `numbers_in` ignores minus signs (its docstring says so), so an answer saying net new ARR vs budget was "19.0%" when the data says "−19.0%" would pass. LEARNINGS.md logs direction errors in general, but not this case.
   - **`python metrics.py` prints "n/a" for both a data gap and "no prior period"** (Exercise 8). Only that debugging printout is affected; Excel and the Claude payload tell them apart. The exercise uses it as a teaching point.
3. **Check the interview answers against your own view,** especially Q2 (business case) and Q25 (what's next), before you use them. They're a starting point written from the repo, not from your conversations with the fund.

---

## Task 7: Review of everything committed tonight (`git diff main..HEAD`)

**Result:** done. I reviewed all 25 commits (from `aaaa606` to `14bf7f1`, 25 files). I found and fixed **3 clear bugs**, each with a test that fails on the old code and passes on the new one. Everything else is listed below for you to decide. `python -m pytest -q` now runs **223 tests, all passing** (212 before, +11). `check_northwind.py`, `check_companies.py`, `check_excel_output.py` and `check_main.py` all print "All checks passed". Two commits: `1fd5db0` (Excel runway label) and `1d53668` (compare_models.py). No refactoring. CLAUDE.md, config.yaml, the analyze.py prompt and metrics.py are unchanged. No build_deck.py and no API calls.

### Review

The findings are grouped as you asked. **"Still open" means I didn't change it.** Items already in an earlier task's "Unresolved" list are only named here, with a pointer, unless the review found something new about them.

#### Inconsistent with CLAUDE.md

1. **"Not meaningful" still shows as "data missing" in 3 cases, even with no blank cells.** All three come from `metrics.data_gaps`, which existed before tonight. Tonight's Excel output and main.py summary pass them on. I confirmed each one with a scratch script:
   - **0 ÷ 0 counts as a gap.** A complete quarter with revenue 0 (a pre-revenue company) makes gross margin and CAC payback "data missing". So does zero net new ARR against a flat `budget_arr` (net new ARR vs budget). main.py then prints `7 metrics/flags (blank: none)`, which contradicts itself.
   - **Flags with too little history count as gaps.** `data_gaps` lists every "cannot evaluate" flag, whatever the reason. A company with fewer than 5 quarters gets "Flag: Rule of 40" in Data gaps, even though the real reason is "no prior year". With fewer than 3 quarters, the combo rule gets the same treatment. Metric cells in the same situation correctly say "n/a (no prior period)", so the flag and its own metric disagree.
   - **A partly blank quarter among the first 4** (already Task 4, Unresolved 1). New finding: the test meant to guard this (`test_data_gaps_one_blank_cell_spreads_only_to_metrics_that_use_it`) blanks Q4 2025, which has a full year of history, so it can't catch the problem.
2. **CAC payback can trip a flag on a blank cell.** If S&M spend is blank and new ARR is 0, `cac_payback_months` gives ∞ and **trips**, and no data gap is listed. The CLAUDE.md rule is "any flag that depends on a missing value returns cannot evaluate". This is the same family as Task 4, Unresolved 5 (runway and burn multiple), but those cases *pass*, and this one puts a red flag on the board slide. It's in metrics.py (from before tonight), so I didn't change it.
3. **The budget-only row isn't checked to be "next quarter"** (Task 5, Unresolved 2). "Q1 2026 (Budget)" after Q2 2026 actuals is accepted and used for runway at budget without a word.
4. **CLAUDE.md's Architecture list is out of date.** It doesn't mention `excel_output.py`, `make_data_common.py`, `make_data_alderpeak.py`, `make_data_fernhollow.py`, the four `check_*.py` scripts, `compare_models.py`, `tests/` or `pytest.ini`. The healthy and distressed company stories aren't in CLAUDE.md either (Task 1, Unresolved 5). I'm not allowed to edit CLAUDE.md.
5. **`python main.py data/northwind.xlsx` doesn't do what CLAUDE.md describes yet.** Steps 4b and 5 were built before step 4. Without `build_deck.py`, main.py skips both the AI summary and the deck, even without `--skip-ai` (Task 3, decision 1). That's correct for tonight, but the README and demo shouldn't promise a deck until step 4 is done.

#### Risky

6. **compare_models.py understates cost when a run fails.** This doesn't affect tonight's README numbers, because all 6 real runs passed.
   - If the retry hits an API error, the first attempt's tokens are thrown away and the run is logged at $0.
   - If a reply doesn't match the JSON shape, analyze.py logs `input_tokens: None`, and `make_run_info` counts it as 0.
   - Either way, average cost per run and the 275-company projection come out low. Fixing it means changing analyze.py's error path (not the prompt).
7. **compare_models.py saves nothing until all 6 runs finish.** A crash or Ctrl-C during run 5 loses 4 paid results.
8. **Prices are hard-coded** (Sonnet 5 $2/$10, Haiku 4.5 $1/$5 per million tokens, "as of 2026-06-24"), and the README prints them with no decimals, so a price like $0.80 would show as "$1". I couldn't check current prices tonight (no web access, no API). Check them before quoting "$14.65 for 275 companies" in an interview.
9. **Commit `aaaa606` is titled "Add pytest" but also contains all of step 3b:** `compare_models.py`, the README comparison table, the LEARNINGS entry, the analyze.py change and the new check_northwind checks. Anyone reading `git log` would miss step 3b. I didn't rewrite history, since the branch may already be shared. If you squash-merge, a better PR title fixes it.
10. **main.py reads each workbook twice** (Task 6, Unresolved 2): once in `run_company`, and again in `save_metrics_workbook`. It's harmless now. If a file changed between the two reads, the Excel file and the summary line could disagree.
11. **`make_data_common.py` checks the answer key with `assert`**, which `python -O` switches off. It's low risk, because `check_companies.py` would still catch a bad workbook.

#### Unclear

12. **`check_excel_output.py`'s docstring overstates how independent it is.** It says expected labels and formats are written out "instead of imported", but it imports the column headers (`METRIC_LABELS`) and the lists that decide number formats (`DOLLAR_COLUMNS`, `MONTH_COLUMNS`) from the same places excel_output.py uses. A wrong entry there would pass both. Also, ∞ cells are only checked with `startswith("∞")`, so a wrong reason (e.g. "∞ (ARR shrank)" on runway) would pass.
13. **Northwind's ARR vs budget is never checked against a hand formula.** `EXPECTED_LATEST` in `check_northwind.py` has no `arr_vs_budget` entry, but the other two companies have one. It would be one line: `"arr_vs_budget": 27470 / 26800 - 1`. That's a missing check, not a bug, so I didn't add it.
14. **Two small compare_models.py edges.** Running `score` a second time rewrites the README but not LEARNINGS.md, so the two can disagree. A README with the start marker but no end marker crashes with `IndexError`.
15. **Names still differ across outputs** (Task 2, Unresolved 1): the Flags sheet says "Burn vs budget" and the Metrics sheet says "Net burn vs budget". `python metrics.py` and the analyze payload also still print "n/a" for both a gap and "no prior period" (Task 1, Unresolved 3). Both are worth settling before the deck.

#### Checked and fine

- **The three company stories match their docstrings and CLAUDE.md.** Northwind: NRR 108.0% → 102.0% → 97.1%, pipeline rising, burn 20.0% over budget, runway 11.0 months, 6 flags trip. Alderpeak trips no flag in any quarter. Fernhollow: 7 trip, 1 cannot evaluate, 1 pass.
- **Nothing writes into `data/`.** Tests use pytest's temp folder and `check_main.py` uses a temp directory. The only project writes go to `output/`, which is git-ignored.
- **No API call happens on import.** `load_dotenv` and the client are only created inside `analyze()` / `command_run`.
- **Ratios are stored as decimals everywhere.** Formatting to % happens only in the printouts, the Excel number formats and the payload.
- **All data is fictional.**

### What I built (the 3 bug fixes)

| File | What changed, in plain English |
|---|---|
| `excel_output.py` | **Bug 1: the Excel file said "no budget row" when a number was just blank.** Runway at next quarter's budgeted burn is NaN in three cases: there is no budget row, the budget row's burn cell is blank, or the latest quarter's cash is blank. `runway_context_value` only looked at the NaN, so all three said "n/a (no budget row)". It now also receives `has_budget_row`: no row → "n/a (no budget row)"; a row but NaN → **"data missing"**. `write_runway_context` and `write_flags_sheet` pass that one extra value through, and `build_workbook` supplies it (`next_budget is not None`). |
| `tests/test_excel_output.py` (new, 5 tests) | Builds a tiny 2-quarter table and reads the runway context cell from the Flags sheet. The 5 cases are a normal number (1200 / (300 / 3) = 12.0), no budget row, a budget of 0 burn (∞), a blank budgeted burn and a blank latest cash. The last 2 fail on the old code. |
| `compare_models.py` | **Bug 2: runs past the 10th silently disappeared.** Blind letters came from `"ABCDEFGHIJ"`, and `zip` stops at the shorter list without an error. With `RUNS_PER_MODEL = 6`, 12 paid runs would become 10 in the blind set and the stats. Letters are now A–Z, and more than 26 runs stops with a message. **Bug 3: a letter could be scored twice.** `score A=4 B=3 A=2` quietly kept 2. It now stops with "answer A was already scored - give each letter once". |
| `tests/test_compare_models.py` (new, 6 tests) | Checks that 6, 11 and 26 runs each keep every run exactly once, that 27 runs stop, and that a repeated letter stops (including `A=4 a=2`). 5 of the 6 fail on the old code. The 6-run case passes on both, which is expected. |
| `LEARNINGS.md` | Two rows: the runway label, and the two compare_models.py holes. |
| `STUDY_GUIDE.md` | 5 table rows updated so the function signatures and descriptions match the fixed code. |

### Decisions you didn't specify

1. **What counts as a "clear bug" I could fix:** wrong output with no design question behind it, in code committed tonight. A fix had to be a few lines, with the right answer obvious from CLAUDE.md or from what the code already claims to do. Everything that needs a definition choice, or touches code from before tonight (metrics.py's gap and edge-case rules), is listed instead. The same goes for anything needing analyze.py changes (the cost accounting).
2. **The compare_models.py fixes are latent.** Neither changed tonight's 6-run results. I still fixed them because both lose data silently, which is the same class of problem as the duplicate-quarter bug in LEARNINGS.md.
3. **"Data missing" for the runway context cell, not a new label.** It's the exact word the Metrics sheet already uses for a gap, and CLAUDE.md wants one consistent "data missing".
4. **The review covered the diff, not just the code.** I also checked the reports, README, LEARNINGS, STUDY_GUIDE and commit messages, because "unclear" findings (e.g. item 9) live there.
5. **I used a separate read-only agent** for the test files, check scripts, data generators and compare_models.py, while I read clean.py, metrics.py, excel_output.py and main.py. It had no permission to edit files or call the API. I re-checked each of its findings against the code before listing it. I dropped one of its items (the wording of a made-up Claude answer inside a test, which is test data, not a claim).
6. **New pytest files instead of adding to `check_northwind.py`.** The unit tests for `parse_scores` so far live in `check_northwind.py`, but tonight's convention is pytest under `tests/`, and pytest runs automatically in the overnight script.
7. **I didn't update the "212 tests" counts in STUDY_GUIDE.md.** The guide is pinned to commit `aa8838b`. I only fixed the rows that describe functions I changed, so no row describes code that no longer exists.

### What failed and how I fixed it

1. **Proving each new test catches its bug, without editing committed files.** The shell sandbox blocked `git stash`, `git archive` and copying the project to `/tmp`. One run I meant to do "without the fix" actually ran with it, because the stash had been refused; I noticed before relying on it. **Fix:** a scratch script in `/tmp/t7` loads the committed version of the file (`git show HEAD:<file>`) under the module's name and runs the new tests against it. Result: 2 of 5 Excel tests and 5 of 6 compare_models tests fail on the old code, all pass on the new code. The project files were never touched.
2. **A heredoc was blocked** (`Parser skipped input`), so I wrote the scratch scripts with the file editor instead. No effect on the project.
3. **The review agent couldn't run scratch scripts** (blocked by permissions), so it checked its edge cases by reading the code. I re-ran the ones I relied on (runway label, 0 ÷ 0 gaps, flag gaps with short history) myself.

### Unresolved: needs your call

1. **Decide how `data_gaps` should tell "missing" from "not meaningful"** (Review items 1–2). One fix covers all of them: a list of which inputs each metric uses, and flags that say *why* they can't be evaluated (no history / missing input / undefined). It's a design change to metrics.py, and the deck's "Data gaps" line depends on it, so it should come before step 4.
2. **Should the budget row's label be checked against the last actual quarter?** (Review item 3; also Task 5.)
3. **Should compare_models.py count failed attempts' tokens?** (Review items 6–7.) Only matters if you rerun the comparison, e.g. for a new prompt version.
4. **Verify the model prices** before quoting costs (Review item 8).
5. **Update CLAUDE.md's Architecture section** with tonight's files (Review item 4). I wasn't allowed to.
6. **Cheap test additions I left out, because they aren't bug fixes:** Northwind `arr_vs_budget` hand formula (item 13), exact ∞ labels in `check_excel_output.py` (item 12), and a partly-blank-first-quarter test (item 1, once you decide the rule).
