# Learnings

Log what broke, why, and how it was fixed.

| Date | What broke | Why | Fix |
|------|-----------|-----|-----|
| 2026-09-17 | Threshold checks at risk of wrong results: Python computes 3900/3250 − 1 as 0.19999999999999996, not 0.2. A burn exactly 15% over budget could trip a "max 0.15" flag. | Computers store decimals in binary, so most decimal fractions are a tiny bit off. Comparing a value with its threshold then fails on noise. | `check_threshold` rounds to 6 decimals before comparing, and the combo rule rounds NRR before `.diff()`. `clean.py` uses `Decimal` so "$21.85M" parses to exactly 21850. Tested: 15.0% over budget → pass. |
| 2026-09-17 | Net new ARR vs budget compared the wrong things: actual net new ARR (after expansion and churn) against `budget_new_arr` (gross new sales only). Q2 2026 showed −17.0% instead of −19.0%. | The definition mixed a net figure with a gross one, so churn counted against actuals but not the budget. Caught in review, not by the checks, because the answer key used the same wrong formula. | Budget side is now `budget_arr` this quarter − last quarter (net on both sides). It looks back 1 quarter, so gaps follow the QoQ rule (Q1 + Q2 2025). Updated metrics.py, CLAUDE.md and the answer key. Lesson: an answer key only catches errors if it's built from the definition independently, not copied from the code. |
| 2026-09-17 | The planned test for the "Claude never does math" check ("NRR fell 11 points") would have passed validation, even though 11 is a subtraction. | The number check only asks "does this number appear anywhere in the payload?" 11 appears as runway (11.0 mo), so a calculated 11 slips through by coincidence. | Tests now use numbers verified to be absent from the payload (11.8 = 108.9 − 97.1, and 97 rounded from 97.1). Known limit: the check catches invented or rounded numbers, not a calculation that happens to equal another value in the data. |

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
