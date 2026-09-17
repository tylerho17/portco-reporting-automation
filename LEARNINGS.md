# Learnings

Log what broke, why, and how it was fixed.

| Date | What broke | Why | Fix |
|------|-----------|-----|-----|
| 2026-09-17 | Threshold checks at risk of wrong results: Python computes 3900/3250 − 1 as 0.19999999999999996, not 0.2. A burn exactly 15% over budget could trip a "max 0.15" flag. | Computers store decimals in binary, so most decimal fractions are a tiny bit off. Comparing a value with its threshold then fails on noise. | `check_threshold` rounds to 6 decimals before comparing, and the combo rule rounds NRR before `.diff()`. `clean.py` uses `Decimal` so "$21.85M" parses to exactly 21850. Tested: 15.0% over budget → pass. |
| 2026-09-17 | Net new ARR vs budget compared the wrong things: actual net new ARR (after expansion and churn) against `budget_new_arr` (gross new sales only). Q2 2026 showed −17.0% instead of −19.0%. | The definition mixed a net figure with a gross one, so churn counted against actuals but not the budget. Caught in review, not by the checks, because the answer key used the same wrong formula. | Budget side is now `budget_arr` this quarter − last quarter (net on both sides). It looks back 1 quarter, so gaps follow the QoQ rule (Q1 + Q2 2025). Updated metrics.py, CLAUDE.md and the answer key. Lesson: an answer key only catches errors if it's built from the definition independently, not copied from the code. |
| 2026-09-17 | The planned test for the "Claude never does math" check ("NRR fell 11 points") would have passed validation, even though 11 is a subtraction. | The number check only asks "does this number appear anywhere in the payload?" 11 appears as runway (11.0 mo), so a calculated 11 slips through by coincidence. | Tests now use numbers verified to be absent from the payload (11.8 = 108.9 − 97.1, and 97 rounded from 97.1). Known limit: the check catches invented or rounded numbers, not a calculation that happens to equal another value in the data. |
| 2026-09-17 | The Excel read-back check failed on its first run: NRR 1.0706921944035346 came back from output/northwind_metrics.xlsx as 1.070692194403535. | openpyxl saves numbers with 16 significant digits (`"%.16g"`), but a Python float can need 17 to round-trip exactly. The gap is ~1e-16, invisible at any display format, and Excel itself only keeps 15 digits. openpyxl also silently saves NaN and infinity as an empty cell, which would make "data missing" and "∞" look identical. | The check compares to 15 significant digits (`same_as_saved`), which still catches 97.1 saved instead of 0.971 or a value rounded to 0.971 (both proven by deliberate breaks). excel_output.py never hands NaN or infinity to openpyxl; it writes a label that says why ("data missing", "n/a (no prior period)", "∞ (ARR shrank)"). |
| 2026-09-17 | clean.py silently lost data when a quarter label appeared twice: with rows Q1, Q2, Q2, Q3 2025, the second Q2 row replaced the first with no error. Found while writing unit tests (Task 4), not by any check. | Rows are stored in a dictionary keyed by label, and saving under a label that already exists overwrites it. The dictionary keeps only one Q2, so the quarter-order check (which does catch repeats) never saw the duplicate. | `clean_workbook` now stops with "Quarter 'Q2 2025' appears twice - check the workbook". Test: `test_workbook_duplicate_quarter_row_stops`. Lesson: a check can only catch what reaches it; test the whole path, not only the helper. |
| 2026-09-17 | A cell reading "n/a", "NA", "NULL" or "None", or holding an Excel error like #DIV/0! or #REF!, was silently read as an empty cell and would have shown on the deck as "data missing". The Task 4 unit test said "n/a" stops, but it only tested `parse_number` on its own. Found by the Task 5 broken-workbook tests. | `pd.read_excel` turns a built-in list of words ("n/a", "NA", "NULL", "nan"...) into NaN before clean.py ever sees them, and reads every Excel error cell as NaN too. `parse_number` never got the chance to stop. | clean.py reads with `keep_default_na=False, na_values=[""]`, so only a truly empty cell is blank and the words reach `parse_number`, which stops. `check_no_error_cells` scans the KPI tab with openpyxl (which still sees error cells) and stops with the cell address. Tests: `test_unreadable_number_names_the_cell`, `test_excel_error_cell_stops`. Same lesson as the duplicate-quarter bug: test through the real file, not only the helper. |
| 2026-09-17 | The Excel Flags sheet showed "n/a (no budget row)" for runway at next quarter's budgeted burn when the budget row **did** exist but its burn cell was blank, or when the latest quarter's cash was blank. Missing data looked like "not applicable". Found in the Task 7 review, not by any check (all three workbooks have both values). | `runway_at_next_budget` returns NaN in all three cases, and `runway_context_value` only looked at the NaN, so it couldn't tell "no row" from "blank cell". | `runway_context_value` now also gets `has_budget_row`: no row → "n/a (no budget row)", row but NaN → "data missing". Test: `tests/test_excel_output.py::test_runway_context_label` (the 2 blank cases fail on the old code). Lesson: when one value (NaN) has more than one cause, the label needs the cause passed in, not guessed from the value. |

## Prompt iterations (analyze.py, Northwind, claude-sonnet-5)

- **2026-09-17, v1 → v2: missing top win**
  - **Flaw:** ARR growth YoY (42.8%) wasn't among the wins; "CAC payback within threshold" was listed instead.
  - **Rule added:** wins lead with the strongest growth or scale metric; "within threshold" is a win only if nothing stronger exists.
  - **Result:** fixed. ARR growth YoY 42.8% is win #1 and also in the headline.
- **2026-09-17, v1 → v2: passing flag read as a breach**
  - **Flaw:** net new ARR vs budget (−19.0%, passed) appeared inside a risk as "(threshold −20.0%)".
  - **Rule added:** say explicitly when a flag passed, quoting value vs threshold, never framed as a breach. Your example "within 1 point" was reworded because it asked Claude to calculate a gap.
  - **Result:** fixed. It now reads "passed but only barely (−19.0% vs −20.0% threshold) - watch".
- **2026-09-17, v1 → v2: trend told from the wrong start**
  - **Flaw:** NRR described as "dropped from 106.0% in Q3 2024", hiding the 108.9% peak.
  - **Rule added:** describe trends from the peak or the start of the flag's lookback window, not the first quarter in the data.
  - **Result:** fixed. NRR now reads "from 108.0% in Q4 2025 to 97.1% in Q2 2026".
- **2026-09-17, v2 side effects (not fixed yet)**
  - **Longer:** output tokens went from 3,218 to 6,955, run time from 29.7s to 61.8s, and cost from ~$0.043 to ~$0.081.
  - **Too wordy:** risk details now run to 3+ clauses, too long for a slide.
  - **One overstatement:** "−1.5% in Q1 2026, a second straight quarter of significant budget miss". A −1.5% miss isn't significant, and Q4 2025 beat budget (+16.5%).
  - **One mischaracterization:** 13.0 mo runway at budgeted burn is called a "projected improvement".
- **2026-09-17, v2 → v3: details too long for a slide**
  - **Flaw:** v2 risk details ran 65–70 words, and output doubled to 6,955 tokens.
  - **Rule added:** each detail is at most 2 sentences and ~40 words. `validate_summary` enforces it (45-word buffer), so a long answer fails and retries.
  - **Result:** fixed. All details are within limits on attempt 1; output fell to 2,090 tokens, 21.1s, $0.0325 (v2: 6,955 tokens, 61.8s, $0.0808). Re-validating the saved v2 answer under the new limits fails all 3 risks, so the check works.
- **2026-09-17, v2 → v3: overstated variance**
  - **Flaw:** a −1.5% quarter was called "a second straight quarter of significant budget miss".
  - **Rule added:** never call a quarter a significant miss or beat unless its value is quoted and more than 10% off budget.
  - **Result:** fixed. No "significant/large/sharp" language appears.
- **2026-09-17, v2 → v3: runway at budget called a projection**
  - **Flaw:** 13.0 mo runway at budgeted burn was called a "projected improvement".
  - **Rule added:** describe it only as "runway if burn returns to plan", never as a projection or improvement. The payload key was also renamed `runway_if_burn_returns_to_plan`.
  - **Result:** fixed. It now reads "runway if burn returns to plan is 13.0 mo".
- **2026-09-17, v3 remaining issues (not fixed yet)**
  - **Rule 3 slip:** question 1 describes burn vs budget "from 0.0% in Q3 2024", the first quarter in the data.
  - **Passing flag not labeled:** question 3 treats net new ARR vs budget (−19.0%, passed) as a problem to "address" without saying it passed.
  - **Missing units:** risk 2 quotes pipeline "from 9,200 to 12,500" without $K.

## Model comparison (step 3b)

- **2026-09-17, Sonnet 5 vs Haiku 4.5 on Northwind (v3 prompt, 3 runs each, blind)**
  - **Sonnet:** passed 3/3, avg score 4.0, $0.0533/run, 36.0s
  - **Haiku:** passed 3/3, avg score 2.0, $0.0173/run, 13.2s
  - **Result:** Keep claude-sonnet-5 as the default ($14.65 per quarter for 275 companies): Haiku averaged 2.0 (needs ≥ 4.0).
  - **Haiku needed its retry on every run** (avg attempts 2.0). The 100% pass rate hides that none of its first answers passed.
  - **Validator gaps found in the blind answers (not fixed yet):**
    - **Coincidental number:** "3.3 months faster than the 24.0-month threshold" is a subtraction, but passed because −3.3% (Rule of 40, Q4 2025) is in the payload. A second real case of the known limit.
    - **Malformed questions:** raw JSON strings (`{"title": ..., "detail": ...}`) were accepted, because the schema only requires strings.
    - **Wrong direction:** "improved… down from 20.3 mo" (actually 20.3 → 20.7, worse). Direction errors pass every current check.
    - **Invented claim:** "Management asserts…" appeared, but nothing checks attributions.
  - **Reflection:**
    - **Prediction before scoring:** the most polished-sounding answer would score highest.
    - **Result:** it tied for lowest (A, Haiku, score 2). Fluent writing hid two problems:
      - A number Claude calculated: "3.3 months faster than the 24.0-month threshold" (24.0 − 20.7).
      - A false claim: net new ARR was "consistently missing budget", but Q4 2025 beat budget by +16.5% and the flag passed at −19.0% vs −20.0%.
    - **Lesson:** judge against the rubric line by line, not by how it reads. This is why validation runs in code and scoring is blind. Neither is enough alone: A passed every code check, and only checking each claim against the data caught it. Code catches rule breaks at scale, blind scoring removes bias toward a model, and line-by-line review catches what code can't yet.
