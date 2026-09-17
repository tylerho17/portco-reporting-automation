# Learnings

Log what broke, why, and how it was fixed.

| Date | What broke | Why | Fix |
|------|-----------|-----|-----|
| 2026-09-17 | Threshold checks at risk of wrong results: Python computes 3900/3250 − 1 as 0.19999999999999996, not 0.2. A burn exactly 15% over budget could trip a "max 0.15" flag. | Computers store decimals in binary, so most decimal fractions are a tiny bit off. Comparing a value with its threshold then fails on noise. | `check_threshold` rounds to 6 decimals before comparing, and the combo rule rounds NRR before `.diff()`. `clean.py` uses `Decimal` so "$21.85M" parses to exactly 21850. Tested: 15.0% over budget → pass. |
| 2026-09-17 | Net new ARR vs budget compared the wrong things: actual net new ARR (after expansion and churn) against `budget_new_arr` (gross new sales only). Q2 2026 showed −17.0% instead of −19.0%. | The definition mixed a net figure with a gross one, so churn counted against actuals but not the budget. Caught in review, not by the checks, because the answer key used the same wrong formula. | Budget side is now `budget_arr` this quarter − last quarter (net on both sides). It looks back 1 quarter, so gaps follow the QoQ rule (Q1 + Q2 2025). Updated metrics.py, CLAUDE.md and the answer key. Lesson: an answer key only catches errors if it's built from the definition independently, not copied from the code. |
