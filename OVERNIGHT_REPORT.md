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
