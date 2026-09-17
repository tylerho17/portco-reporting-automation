# Board Pack Generator

Messy portfolio-company KPI workbook (.xlsx) → board update deck (.pptx) + AI summary. Full README comes in build step 6.

<!-- model-comparison:start -->
## Model comparison (build step 3b)

**Setup:** Northwind (fictional data), analyze.py v3 prompt, identical input for every run, 3 runs per model, alternating order. Answers were shuffled and scored blind before the model key was opened. Each model ran at its default settings (Sonnet 5 uses adaptive thinking by default; Haiku 4.5 does not), which is how they would be deployed. Runs that failed validation after analyze.py's built-in retry were not rerun: they count in the pass rate, cost and time, but aren't scored.

**Scoring rubric**

| Score | Meaning |
|---|---|
| 5 | Board-ready, no edits |
| 4 | One or two word-level fixes |
| 3 | Framing or emphasis wrong somewhere |
| 2 | Misleading or missing a key point |
| 1 | Wrong or unusable |

**Results**

| Model | Pass rate | Avg attempts | Avg cost / run | Avg seconds | Avg score | Cost for 275 companies / quarter |
|---|---|---|---|---|---|---|
| claude-sonnet-5 | 3/3 (100%) | 1.0 | $0.0533 | 36.0 | 4.0 | $14.65 |
| claude-haiku-4-5 | 3/3 (100%) | 2.0 | $0.0173 | 13.2 | 2.0 | $4.77 |

Blind scores by answer: A claude-haiku-4-5 = 2, B claude-haiku-4-5 = 2, C claude-haiku-4-5 = 2, D claude-sonnet-5 = 4, E claude-sonnet-5 = 4, F claude-sonnet-5 = 4.

**Recommendation rule (set before running):** if Haiku averages ≥ 4.0 with a 100% pass rate, Haiku becomes the batch default with Sonnet for flagged or board-critical companies. Otherwise Sonnet stays the default.

**Recommendation:** Keep claude-sonnet-5 as the default ($14.65 per quarter for 275 companies): Haiku averaged 2.0 (needs ≥ 4.0).

**Caveats:** 3 runs per model on one company, so averages are indicative, not conclusive. Prices as of 2026-06-24 (per million input/output tokens: Sonnet 5 $2/$10, Haiku 4.5 $1/$5). Projected cost = average cost per run × 275, assuming other companies' data is similar in size to Northwind's.
<!-- model-comparison:end -->
