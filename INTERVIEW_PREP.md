# Interview prep

Every question you are likely to get about this project, in the order an interviewer tends to ask
them: first what it is, then why it is built this way, then the AI, then what went wrong, then
whether it would survive real use, then how you work.

- **Each answer is 30–60 seconds out loud.** Read it, then say it in your own words. Don't memorize.
- **"Point to"** is where to look (or what to open on screen) if they ask for detail.
- Questions 1–30 come from STUDY_GUIDE.md section 5, re-ordered into these groups. The old number
  is in brackets, e.g. "(guide Q5)". A few facts in them have changed since the guide was written
  (the cost, the test count, the direction check); they are updated here.
- The last section is how to answer honestly when you can't recall a detail. Read it twice.

## Contents

1. [The project](#the-project): Q1–Q4
2. [Design decisions](#design-decisions): Q5–Q16
3. [The AI layer](#the-ai-layer): Q17–Q25
4. [What broke](#what-broke): Q26–Q31
5. [Scale and risk](#scale-and-risk): Q32–Q36
6. [Working method](#working-method): Q37–Q39
7. [Deep dives on the final run](#deep-dives-on-the-final-run): Q40–Q72, the follow-ups an
   interviewer asks after the first answer (column mapping, the eval set, batch resilience, cost
   ceilings, golden files, the approval gate, the run diff, the rollup, exports, what breaks at 275
   companies)
8. [When you can't recall a detail](#when-you-cant-recall-a-detail)

---

## The project

**Q1. Walk me through the project in one minute.** (guide Q1)
A messy portfolio-company KPI workbook goes in; a board pack comes out. `clean.py` reads the messy
Excel: odd headers, "$14.3M" typed as text, a blank quarter. It produces a standard table in $K.
`metrics.py` calculates NRR, GRR, burn multiple, runway, Rule of 40, CAC payback and budget
variances, then checks 9 flags against thresholds in `config.yaml` and lists any data gaps.
`excel_output.py` writes a highlighted Excel summary, and `analyze.py` has Claude write the
headline, risks and questions for management. Code then checks that Claude used only numbers from
the data, pointed them the right way, and fits the slide. `build_deck.py` puts it all on a 4-slide
deck: key metrics table, ARR and cash charts, risks and flags with data gaps, and the AI commentary
(marked as AI-drafted). `main.py` runs it for one company or a whole folder, and still builds the
deck if the AI step fails. There's also a web page (`app.py`): every company in one table with its flags and deck status,
Generate and download buttons, a page per company, and Approve.
*Point to:* STUDY_GUIDE.md section 2; README.md "Data flow".

**Q2. Why would a PE fund want this?** (guide Q2)
Portfolio companies report KPIs in their own formats, so analysts spend time re-typing and
reconciling before they can think. This standardizes the file, applies the **same** definitions and
thresholds to every company, and flags what needs attention, so the analyst's time goes to the
questions for management. It also makes gaps explicit instead of letting a missing quarter quietly
distort a trend. The AI step costs about $0.09 per company, about $25 per quarter for 275
companies, from the latest live run of all three.
*Point to:* README.md "Cost".

**Q3. Why build three fake companies instead of one?** (guide Q4)
With one company you can't tell whether the flags work or were tuned to produce that company's
story. Northwind trips 6 of 9, Alderpeak (healthy) trips none in any quarter, and Fernhollow
(distressed) trips 7, with 1 flag it can't evaluate. That proves the logic responds to the data.
Each company also has different mess (a title row, a different column order, the Notes tab first),
so cleaning isn't tuned to one file either.
*Point to:* `check_companies.py`; OVERNIGHT_REPORT.md Task 1.

**Q4. Why fake data at all?**
Two reasons. Real portfolio data is confidential, and a demo that leaks a company's numbers is a
serious problem, so CLAUDE.md makes fictional data a rule. And fake data lets me write the answer
first: each `make_data` script holds the true numbers (the answer key) and then messes them up on
purpose. So every check can compare what the pipeline recovered with what I know is true. With a
real file I would only know what the file says, not what it should say.
*Point to:* `make_data.py` (`TRUE_DATA`); `make_data_common.py`; CLAUDE.md Rules.

---

## Design decisions

**Q5. Why does Python compute every number and Claude only interpret?** (guide Q3, expanded)
A language model can make arithmetic mistakes, and one wrong number in a board deck undermines every
other number in it. Python math is deterministic: the same inputs give the same answer every time,
and it can be tested against formulas I worked out by hand. A prompt can't be tested the same way.
So every number is computed and checked in Python first; Claude receives them already formatted
("97.1%"), and the prompt tells it to quote, never calculate. Then code checks the answer. It also
splits the job cleanly: if a number is wrong, it's a Python bug I can find and test; if the wording
is wrong, it's the AI layer. In an audit, "the model did the math" is not an answer you want to give.
*Point to:* CLAUDE.md Rules; `analyze.SYSTEM_PROMPT`; `analyze.find_ungrounded_numbers`.

**Q6. How do you handle missing data? Why not fill it in?** (guide Q5)
A blank cell stays blank (NaN), never 0 and never an estimate: an imputed number could reach a
board as if it were real. Because any math with NaN gives NaN, a blank quarter automatically spreads
to every metric that uses it: QoQ metrics for that quarter and the next, YoY for that quarter and the
one 4 later. Flags that depend on a missing value say "cannot evaluate: missing input" instead of
pass or fail, and every affected metric and flag is listed as a data gap. Edge-case rules (∞, 0)
only apply when every input is present, so a blank cell can never become a red flag.
*Point to:* `metrics.data_gaps`; CLAUDE.md "Messy data rules".

**Q7. What's the difference between "data missing" and "not meaningful"?** (guide Q6)
YoY growth in the first year of data is blank because there's no earlier year: that's "n/a (no
prior period)", not a problem. YoY in Q1 2026 is blank because Q1 2025 exists but wasn't reported:
that's "data missing". A third case is "not meaningful": every input is there but the math is
undefined, e.g. 0 ÷ 0, or burn vs a budget of 0 (shown with the $K figures instead). Infinite values
have their own words too: burn multiple "∞ (ARR shrank)", runway "∞ (not burning)". They must never
look alike, because a reader would draw different conclusions. Each metric declares its inputs
(`METRIC_INPUTS`), and `metric_reasons` checks them: a blank input → missing input; an earlier
quarter that doesn't exist → no prior period; otherwise → not meaningful. Only missing input is a
data gap.
*Point to:* `metrics.metric_reasons`, `metrics.data_gaps`, `metrics.display_value`, `excel_output.cell_value`.

**Q8. Why store ratios as decimals and format as % only at output?** (guide Q7)
One representation everywhere means thresholds, comparisons and math never mix 97.1 with 0.971.
Formatting happens in one place (`format_value`, or an Excel number format), so the Excel cell still
holds 0.971 and works in formulas while displaying 97.1%. The Excel check deliberately tested saving
97.1 instead of 0.971, and it's caught.
*Point to:* `metrics.format_value`, `excel_output.number_format`.

**Q9. Why is NRR annualized, and what's the trade-off?** (guide Q8)
The thresholds (NRR 100%, GRR 85%) are annual conventions. One quarter's expansion minus churn is
roughly a quarter of the annual effect, so it's multiplied by 4 to compare like with like. The
trade-off: one quarter's data is noisier than a true trailing-12-month cohort. So the deck labels it
"annualized", and the combo rule looks at a 3-quarter trend rather than a single quarter.
*Point to:* `metrics.nrr`, `metrics.grr`; CLAUDE.md metric definitions.

**Q10. Explain the combo rule. Why "cannot evaluate" instead of False?** (guide Q9)
NRR falling while pipeline is rising suggests a retention problem, not a sales problem: sales keeps
filling the funnel, but existing customers are leaking, so more pipeline won't fix it. It trips only
if NRR fell by at least 1 point (`combo_min_nrr_drop`) at **every** step and pipeline rose at
**every** step over the last 3 quarters, so a 0.1-point wobble isn't called a retention problem. If
any quarter in the window is missing, returning False would tell the board "no retention problem"
when we simply don't know. So it returns "cannot evaluate". Fernhollow passes because its pipeline is
falling too: that's a sales problem **and** a retention problem.
*Point to:* `metrics.check_combo`; config.yaml comments.

**Q11. Why is "exactly at the threshold" a pass, and why round before comparing?** (guide Q10)
The thresholds are limits ("more than 15% over budget"), so hitting 15.0% exactly is within the
limit. Computers store decimals in binary, so 3900/3250 − 1 comes out as 0.19999999999999996, not
0.2. A burn exactly at a threshold could trip or pass on invisible noise. `check_threshold` rounds
to 6 decimals first. A test proves a real miss of 0.000001 still trips, so rounding doesn't hide
real breaches.
*Point to:* `metrics.check_threshold`; LEARNINGS.md first row; `tests/test_metrics.py`.

**Q12. Why are thresholds in config.yaml and the API key in .env?** (guide Q11)
Thresholds are investment judgments, not code: a partner may want runway under 18 months instead of
12, and that shouldn't need a code change. Each line has a comment with the investor reasoning. The
API key is a secret: it lives in `.env`, which `.gitignore` keeps out of git, and `python-dotenv`
loads it at run time. `.env.example` shows the variable name without a key.
*Point to:* `config.yaml`, `analyze.main` (`load_dotenv`).

**Q13. What happens to the deck if the AI step fails? Why build it at all?** (guide Q27)
The deck is still built. Every number on it comes from Python, so the metrics table, charts, flags
and data gaps are valid whatever Claude did. Only slide 4 changes: it says "AI summary unavailable",
with a note that the other numbers are unaffected. The batch result reads `OK (AI failed)`, a warning
names the company, and the reason is saved in its analysis JSON. "AI failed" means only two things:
the answer failed validation twice, or the API call itself failed. Any other error is treated as a
bug and fails the company, so a coding mistake can't hide behind "AI failed". A missing API key
stops the run before any company starts.
*Point to:* `main.ai_step`, `main.ai_status`; `tests/test_main.py` (fake client).

**Q14. How does the deck show a quarter the company never sent?** (guide Q29)
It never fills it in. On slide 1, a value that needs the blank quarter says "data missing"
(Northwind's Q1 2026 ARR growth YoY compares with the blank Q1 2025), which is different from "n/a
(no prior period)". On slide 2, the ARR chart draws no bar for Q1 2025 and writes "data missing"
there. The cash line keeps the blank as NaN, so matplotlib breaks the line instead of drawing a
straight join that would invent the numbers in between. On slide 3, a Data gaps column lists every
affected metric and flag. `check_deck.py` checks there's no bar, the line breaks, and every "data
missing" metric in the Excel file appears on slide 3.
*Point to:* `charts.bar_panel`, `charts.cash_chart`, `build_deck.gaps_lines`.

**Q15. How do you know no number on the deck was typed in by hand or mis-formatted?** (guide Q30)
Two ways. First, a unit test reads the source code of `build_deck.py` and `charts.py` and fails if
any piece of text in them (docstrings aside) contains a digit, so even "of 9" or "12 months" can't
be typed in. Every number comes from metrics.py, config.yaml or the validated analysis. Second,
`check_deck.py` opens the saved deck, collects every number on slide 1, and checks each one appears
in the saved metrics workbook **as Excel displays it**, using Excel's own number formats rather than
the deck's formatting code. Breaking a cell on purpose (a wrong number, a right number in the wrong
column, a wrong status) was caught each time.
*Point to:* `tests/test_build_deck.py::test_no_digit_in_any_text_written_in_the_code`, `check_deck.check_kpi_numbers`, `check_deck.check_kpi_rows`.

**Q16. How do you make sure text doesn't run off a slide, when PowerPoint doesn't tell Python how big text is?** (guide Q28)
`text_fit.py` estimates it. It measures each word's width with a real font file (slightly wider than
Arial, so it errs toward "needs more room"), wraps words into lines, and adds up line heights. If the
text is too tall, every font size in the box drops by 1 pt together until it fits. If it doesn't fit
at 12 pt, the build stops with the slide and box named, because a board deck with text off the page
is worse than no deck. For Claude's text the same measurement runs inside validation, so an answer
that's too long uses the retry instead of stopping the build. **The honest limit:** it's an estimate,
not PowerPoint's own layout, and slide 4's Risks column has little spare room: Northwind's and
Fernhollow's risks fit only at the 12 pt floor.
*Point to:* `text_fit.shrink_to_fit`, `check_deck.check_no_overflow`, `build_deck.ai_text_problems`; POLISH_REPORT.md Task 2.

---

## The AI layer

**Q17. How do you stop Claude from inventing or calculating numbers?** (guide Q17)
Four layers:
1. Claude only gets pre-formatted facts.
2. The prompt says to quote numbers exactly and never calculate differences or changes ("went from A
   to B" instead).
3. `validate_summary` extracts every number from the answer, sign included, and fails it if any
   number doesn't appear in the payload. A rounded "97%" or a calculated "11.8 points" gets caught.
4. The saved JSON keeps the payload next to the answer, so a reviewer can trace every claim.

*Point to:* `analyze.find_ungrounded_numbers`, `check_northwind.check_analysis_validation`.

**Q18. What are the known limits of that number check?** (guide Q18, updated)
It asks "does this number appear anywhere in the data?", not "is it used correctly". So a calculated
number that happens to equal another value passes: in the blind runs, "3.3 months faster" passed
because −3.3% is a Rule of 40 value. It can't catch invented attributions ("Management asserts…"),
or a claim that two figures moved "over the same period" when they didn't. Wrong direction used to
be on this list; it's now partly covered by the direction check (Q21), but a claim with no numbers
in it still can't be checked in code. That's why a person reads every deck.
*Point to:* LEARNINGS.md "Validator gaps found in the blind answers"; README.md "Known limitations".

**Q19. What happens if Claude's answer fails validation?** (guide Q19)
`analyze()` retries **once**, sending back Claude's previous answer and the exact list of problems.
If the second answer also fails, it raises `AnalysisError`; nothing unvalidated is ever returned.
Both attempts are logged with tokens and timing. One retry is the balance between cost and latency
and giving Claude a fair chance to fix a specific problem. Repeated failure points to a prompt or
model problem that more retries would just pay for.
*Point to:* `analyze.analyze`, `analyze.build_messages`.

**Q20. How did you improve the prompt?** (guide Q20)
Each round: run, find a specific flaw, add one rule, re-run, check the result. v1 → v2 fixed three
framing problems: the best win (ARR growth YoY 42.8%) was missing, a passing flag was framed as a
breach, and NRR's decline was described from the first quarter in the data instead of the peak. But
v2 doubled the output (3,218 → 6,955 tokens) and made details too long for a slide. v3 added a
2-sentence, ~40-word limit that **code enforces**, and banned "significant" for variances under
10%. v4 added the direction rules (Q21). The prompt is now `PROMPT_VERSION` v4, and every saved
analysis records which version wrote it.
*Point to:* LEARNINGS.md "Prompt iterations"; `analyze.PROMPT_VERSION`.

**Q21. Tell me about the direction-claim failures. How did you find them, and how do you catch them now?** (new)
The first live run passed every code check, and I still read all three summaries line by line
against the data. Three claims were false or misleading. Alderpeak's growth was "well above
prior-period levels", but it fell every quarter, 53.9% down to 47.5%. Alderpeak's NRR had a
"persistent gentle decline", but it goes up and down and rose last quarter. And Fernhollow's passing
combo flag was called a win, when it only passes because pipeline is falling too. The fix has two
parts. Two prompt rules: call a trend rising or falling only if every step moves that way, and a
passing flag is a win only if its value is good in itself. And a two-layer code check. Layer 1: the
direction word has to agree with its own two numbers ("fell from 97.1% to 108.9%" fails). Layer 2: a
trend called persistent or consistent is checked step by step against the quarterly series in the
payload. Honestly: layer 2 catches 1 of the 3; the other 2 have no numbers to check, so they rest on
the prompt rules, and the next live run had all 3 fixed.
*Point to:* `analyze.claim_problems`, `analyze.contradiction_problem`, `analyze.persistence_problem`; `tests/test_analyze.py::test_claims_only_the_prompt_rules_can_catch`; LEARNINGS.md Task 4 live-run row and Task 7 rows.

**Q22. Did the direction check ever get it wrong?**
Yes, twice, and both times it failed a true claim. "NRR went from 108.0% to 97.1% while pipeline
rose": the check took "rose" from anywhere in the sentence, but "rose" was about pipeline. The fix:
the direction word has to sit between the previous number and the pair it describes. One of these
false positives happened in a live run and cost Alderpeak a retry, about 3 cents; the retry passed,
so no company failed. The lesson: a check on Claude's words needs its false positives tested as
hard as its catches. Both sentences are now test fixtures. I also deliberately left "improved" and
"deteriorated" out, because an improvement is a smaller burn multiple but a bigger NRR, so the word
alone doesn't tell you the direction.
*Point to:* `tests/test_analyze.py::test_a_direction_word_after_the_numbers_belongs_to_something_else`; LEARNINGS.md Task 7 row "the new direction check failed two claims that were true".

**Q23. What did the direction rules cost, and is that worth it?** (new)
Real numbers from two live runs of all three companies. Before the direction rules: about 3,500
output tokens and $0.047 per company, $12.81 per quarter for 275 companies. After: about 7,500
output tokens and $0.091 per company, $25.06 per quarter. So roughly 1.1 million extra output tokens
and about $12 more per quarter, about $50 a year, or 4–5 cents per company. Part of that was one
retry from a false positive I've since fixed, but the two companies that passed first time still
roughly doubled: checking every step of a trend before writing is thinking Claude has to pay for.
The trade is clearly right. A false "persistent decline" in a board deck costs credibility for every
other number on it, and 4 cents is far less than the analyst time to catch it by hand. The bigger
cost is time, not money: about 70 seconds per company instead of 33, so a 275-company batch run one
at a time goes from 2.5 to 5 hours. That's a reason to run companies in parallel, not to drop the rules.
*Point to:* README.md "Cost"; LEARNINGS.md "Live runs" (Task 4 and Task 7 tables).

**Q24. Which model did you pick, and how did you decide?** (guide Q21, expanded)
Sonnet 5 against Haiku 4.5, 3 runs each on identical Northwind input. The answers were shuffled,
given letters, and scored blind on a 1–5 rubric before I opened the key. The decision rule was
written **before** running: Haiku becomes the default only if it averages at least 4.0 with a 100%
pass rate. Sonnet averaged 4.0, at $0.053 and 36 s per run. Haiku averaged 2.0, at $0.017 and 13 s,
and it needed its retry on every run, so its 100% pass rate hid that none of its first answers
passed. So Sonnet stays the default. Writing the rule first matters: otherwise it's easy to find a
reason afterwards to pick the one you already liked. The caveat: 3 runs on one company is
indicative, not conclusive.
*Point to:* README.md "Model comparison"; `compare_models.recommend`.

**Q25. What did blind scoring teach you?** (guide Q22)
Before scoring, I predicted the most polished-sounding answer would score highest. It tied for
lowest. The fluent writing hid a number Claude calculated (24.0 − 20.7 = "3.3 months faster") and a
false claim ("consistently missing budget", when Q4 2025 beat budget by +16.5%). It had passed every
code check. **Lesson:** judge line by line against the data and the rubric, not by how it reads.
Code checks catch rule breaks at scale, blind scoring removes bias toward a model, and line-by-line
review catches what code can't yet.
*Point to:* LEARNINGS.md "Reflection".

---

## What broke

**Q26. Why is the analysis checked again when the deck is built, when it already passed?** (guide Q26, expanded)
Because the deck can be built later than the analysis, from a workbook that has changed since, or
from a JSON file someone edited by hand. `load_analysis` accepts a saved analysis only if it's for
the same company and latest quarter, has the right shape, and `validate_summary` still passes
against a payload **rebuilt from today's workbook**. So every number in Claude's text is in today's
data, and the text still fits today's slide. If any check fails, slide 4 says "AI summary
unavailable" and the terminal says why. `check_deck.py` proves it by changing "11.0 mo" to "11.5 mo"
in a copy, and by labelling a copy for the previous quarter: both get the placeholder. It's the same
idea as not trusting last month's reconciliation for this month's numbers.
*Point to:* `build_deck.load_analysis`; `check_deck.check_bad_analysis_gets_placeholder`.

**Q27. Tell me about a bug you found. What did it teach you?** (guide Q12)
Net new ARR vs budget compared actual **net** new ARR (after churn) with `budget_new_arr`, which is
**gross** new sales. So churn counted against the actuals but not the budget, and Q2 2026 showed
−17.0% instead of −19.0%. The checks didn't catch it, because the answer key had been written with
the same wrong formula. It was caught in review. The fix: the budget side is now budget_arr this
quarter minus last quarter, net on both sides. **Lesson:** an answer key only catches errors if it's
built independently from the definition, not copied from the code.
*Point to:* `metrics.budget_net_new_arr`; LEARNINGS.md second row.

**Q28. Why does clean.py use `Decimal` to read "$4.03M"?** (guide Q13)
In normal float math, 4.03 × 1000 = 4030.0000000000005. `Decimal` does the math the way you would on
paper, so "$4.03M" becomes exactly 4030. That matters because the checks require cleaned values to
equal the answer key exactly, and because tiny errors can push a value across a threshold.
*Point to:* `clean.parse_number`; LEARNINGS.md first row.

**Q29. A cell saying "n/a" was being read as blank. Why was that dangerous, and how did you fix it?** (guide Q14)
pandas quietly turns a built-in list of words ("n/a", "NA", "NULL", "None") and every Excel error like
`#DIV/0!` into blanks before our code sees them. So a broken formula would have shown on the deck as
"data missing" instead of stopping the run to be fixed. A unit test said "n/a" stops, but it only
tested `parse_number` on its own, never through pandas. The fix was to read with
`keep_default_na=False` so only a truly empty cell is blank, plus `check_no_error_cells` using
openpyxl. **Lesson:** test through the real file, not just the helper.
*Point to:* `clean.find_kpi_sheet`, `clean.check_no_error_cells`; LEARNINGS.md row on "n/a".

**Q30. What happened with a duplicated quarter row?** (guide Q15)
Rows were stored in a dictionary keyed by quarter label, and saving a second "Q2 2025" silently
overwrote the first. The quarter-order check does catch repeats, but it never saw the duplicate
because the dictionary only held one Q2. Now `clean_sheet` stops and names both rows. Same lesson as
Q29: a check can only catch what reaches it.
*Point to:* `clean.clean_sheet`; LEARNINGS.md duplicate-quarter row.

**Q31. Your Excel read-back check failed on its first run. Why?** (guide Q16)
NRR 1.0706921944035346 came back as 1.070692194403535. openpyxl saves 16 significant digits, but a
Python float can need 17. The gap is about 1e-16, and Excel itself only keeps 15 digits, so the
check compares to 15 significant digits. It still catches real errors like 97.1 saved instead of
0.971. The same investigation found openpyxl silently saves NaN and infinity as an **empty cell**, so
`excel_output.py` writes a label that says why instead ("data missing", "∞ (ARR shrank)").
*Point to:* `check_excel_output.same_as_saved`, `excel_output.cell_value`; LEARNINGS.md read-back row.

---

## Scale and risk

**Q32. How do you know the numbers are right?** (guide Q23, updated)
Four layers:
1. The fake data's answer key is checked to tie out: ARR and cash roll forward.
2. The `check_*.py` scripts run the real workbooks end to end and compare with hand formulas typed
   like Excel (`3900 / (1850 + 580 - 260 - 510)`), written independently of the code.
3. Over 1,300 pytest unit tests check each function, including every edge case, with the hand math in
   comments. `check_deck.py` also opens each saved deck and checks every number on slide 1 against
   the metrics workbook.
4. The tests were tested: the code was broken on purpose (in throwaway copies) to confirm the tests
   notice. 35 of 36 breaks were caught in one round (the missed one can't change any result) and 18
   of 18 in the next.

*Point to:* `check_companies.py`, `tests/`, OVERNIGHT_REPORT.md "What failed".

**Q33. In a batch of 275 companies, what happens when one workbook is broken?** (guide Q24)
`clean.py` stops that company with a message naming the sheet and cell, e.g. `Sheet 'KPI Tracker',
cell G5 (Q2 2025, revenue): Can't read 'TBD' as a number`. `main.py` catches it, prints one ✗ FAILED
line, and moves on to the next company. The summary table shows FAILED with the reason, a CSV copy
is saved, and the run exits with code 1 so a scheduler knows. A bad-input error gets one clean line;
an unexpected error (probably a bug) also gets a full traceback.
*Point to:* `main.run_batch`, `main.describe_error`; `check_main.py`.

**Q34. What are provenance and the approval gate for?** (new)
Two questions an auditor or a partner will ask about any deck: "where did this come from?" and "who
checked it?". Provenance answers the first. Every run writes a manifest per company: the workbook and
config.yaml by SHA-256 hash, the git commit, the model and prompt version, tokens, cost and whether
validation passed. A hash matters because two files with the same name can hold different numbers.
The footer repeats the commit and model, so even a printed slide is traceable. The approval gate
answers the second. `approve.py` records a reviewer's name and the time, and the footer then says
"reviewed by NAME on DATE"; otherwise it says "not reviewed", and `--draft` stamps a DRAFT watermark.
The approval is tied to the hashes, so if the data or a threshold changes, the deck goes back to
not reviewed by itself: a stale approval is worse than none. And `approve.py` never builds a deck, so
making a deck and vouching for it stay two separate acts.
*Point to:* `provenance.build_manifest`, `provenance.approval_status`, `approve.approve`, `build_deck.record_deck_status`; `tests/test_approve.py`.

**Q34b. A new company's workbook says "Opening ARR" and "Plan Burn". What happens?** (new, Task 5)
clean.py doesn't know those headers, so the run stops, but not with a bare error. `mapping.py`
proposes what each header means: "Opening ARR" becomes starting ARR from its words ("opening" is a
synonym for starting), and "Plan Burn" becomes budgeted net burn because of its words and because it
has a value in the budget-only row, which only budget columns may. Each proposal carries a
confidence and a reason, and the page shows the first few values under the header. A person
confirms or changes each one; only then is it saved, to `mappings/<company>.yaml`, and next
quarter's workbook runs without asking. Nothing is used on confidence alone, even at 99%: a wrong
column would put the wrong number on a board slide with nothing to show it. The mapping file is
hashed into the manifest, so changing it voids an approval. The proof: all three companies' workbooks
with their headers renamed get every proposal right and identical metrics after confirming. It's
heuristics, not a model, and I'd add a Claude call next, for headers like "Cash Burn" whose name
fits two columns.
*Point to:* `mapping.propose`, `clean.clean_workbook`, `app.review_mapping`; `tests/test_mapping.py`.

**Q34c. Three demo companies can't cover every edge case. How do you know it handles the odd ones?** (new, Task 6)
An evaluation set: 12 more fictional companies, each built for one case, each with its own answer
key typed by hand. A company sitting exactly on every threshold (NRR 100%, runway 12 months, Rule
of 40 at 40: all pass, because exactly at the line passes). A blank quarter first, second to last
and last. A pre-launch company with zero revenue, where 0 divided by 0 must say "not meaningful"
rather than a number. A plan with negative budgeted burn. Two quarters of history, where YoY must
say "no prior period" and never count as missing data. And two workbooks that must stop: a row
pasted twice and a row missing. `python eval/run_eval.py` runs clean, metrics, flags and gaps for
all 12 and prints a scorecard that names any mismatch, like "Tidewell, Flags: Rule of 40: expected
pass, got trip". To prove it works I planted 24 bugs, one at a time, in throwaway copies, and each
failed the company built for it. It also found a hole in itself: at first only the flag metrics
were value-checked, so a wrong revenue YoY slipped through. The healthy company now checks all 20.
*Point to:* `eval/make_eval_data.py`, `eval/run_eval.py`; `tests/test_eval.py::test_every_company_matches_its_answer_key`.

**Q34d. A 275-company batch runs overnight. What stops it going wrong?** (new, Task 7)
Five switches. `--workers 4` runs four companies side by side, because the time is all waiting for
Claude. A rate limit waits and tries again: what the API asks for, else 5, 10, 20, 40 seconds, then
gives up and that deck gets the placeholder, like any AI failure. `--timeout` gives up on a company
that hangs and moves on. `--max-cost` starts no more companies once the spend reaches the ceiling.
And `--resume` skips any company whose outputs were built from exactly today's workbook and
thresholds, checked by hash, so a rerun after a crash only redoes what's missing. The part I care
most about: each company is built in a private folder and moved into output only if it succeeds, so
a failure or a timeout leaves last quarter's deck untouched, never a new deck beside an old memo.
Every skip, timeout and stop is written into that company's manifest, the summary table and a batch
manifest. I proved it with fake clients (one that hangs, one that's rate limited, one that counts how
many calls are in flight) and by planting bugs in throwaway copies.
*Point to:* `main.run_batch`, `resilience.RateLimitRetry`, `resilience.resume_problem`, `resilience.commit_stage`; `tests/test_batch.py::test_a_company_past_its_timeout_is_given_up_and_the_next_one_still_runs`.

**Q34e. Your checks prove the numbers. How do you know the deck still looks right?** (new, Task 8)
Golden files. I read each company's deck, memo and metrics workbook once, approved them, and saved
a text copy: every word with its size and color, where each box sits, the table fills, the memo's
page breaks, every Excel cell's format. Every test run rebuilds them and fails with the exact
changed lines. Text, not the files, because a PowerPoint file's bytes change on every save. To make
the test repeatable I fixed the two things that change by themselves, the date and the commit in the
footer, and the AI text comes from saved analyses, so no API call. I planted 17 bugs a reader would
notice but that change no number, like charts swapped sides or an Excel header row no longer
frozen. The goldens caught all 17; my number checks caught 2. The trade-off: an intended change
fails too, so you rerun with `--update` and read the diff before committing it.
*Point to:* `golden.dump_deck`, `golden.compare`, `tests/golden/northwind_deck.txt`; `tests/test_golden.py::test_the_output_matches_its_approved_golden`.

**Q34f. A partner wants one page on the whole portfolio. What does the rollup show, and how did you pick the "worst" flag?** (new, Task 9)
One deck and one workbook across every company. Slide 1 ranks the companies by flags tripped, with
each one's worst flag and its value against the threshold: Fernhollow first with 7, worst is runway
at 6 months. Slide 2 counts companies by flag status and by review status, so you also see how many
decks nobody has signed off. Slide 3 is runway by company, red where the flag trips, with the
12-month line. For "worst" I didn't score how far past the threshold each flag is, because months,
percentages and multiples can't be compared, so any score would be a made-up number. It's a fixed
order an investor reads in, written in the code with the reason beside each: runway first, because
cash running out leaves months to act; the combo rule last, because it explains an NRR drop that's
already higher up. Same rules as every deck: Python computes, no AI text on it, text must fit, and
the numbers are worked out from the workbooks when it's built, so it's never stale.
*Point to:* `rollup.WORST_FIRST`, `rollup.rank_key`, `check_rollup.py`; `tests/test_rollup.py::test_the_portfolio_is_ranked_by_flags_tripped`.

**Q34g. A board member asks "what's different from last quarter?" How does the tool answer?** (new, Task 10)
Every run saves its results in the manifest: each metric's value and the words the deck showed, each
flag's status, each data gap. The next run compares and lists three things: flags that flipped
("Runway at current burn: Tripped (was Passed)"), metrics that moved more than a set amount, and data
gaps that opened or closed. For Northwind, Q1 to Q2: five flags flipped, eight metrics moved, one gap
closed. It's in the memo after the headline and on the company page. Two choices I'd defend. First,
"moved" is points for percentages and percent-of-old-value for everything else, because one number
can't mean both: 5% of a 97% NRR is under 5 points. Second, it compares with the last run whose
numbers were different, not literally the last run: otherwise approving and rebuilding would make the
memo say "nothing changed" against a run ten seconds old. Proof: `check_diff.py` rebuilds each company's
workbook as it stood a quarter ago, runs both, and checks every line against lists I worked out by
hand from the answer keys.
*Point to:* `diff_runs.baseline`, `diff_runs.move_text`, `check_diff.py`; `tests/test_diff_runs.py::test_a_rebuild_with_the_same_results_keeps_the_comparison_it_already_had`.

**Q34h. The data team wants your numbers in their warehouse, and a partner wants the table in an email. How?** (new, Task 11)
`export.py` writes the same numbers four ways: a metrics CSV with one row per quarter and metric
(the shape a database or Power BI reads without reshaping), a flags CSV, a JSON with both plus the
hashes of the inputs, and an HTML email. Three choices I'd defend. A missing number is an empty
cell or null with a status saying why, never 0 and never NaN, because a tool that reads 0 for a
blank quarter gets the portfolio quietly wrong. The values are the metrics workbook's to the digit:
the test comparing them failed first, because openpyxl stores 16 significant digits and Python has
17, so I export what the workbook stores. And the email is built for Outlook, which draws email with
Word's engine: inline styles only, a font on every cell or it falls back to Times New Roman, fills
repeated as bgcolor. `check_export.py` reads every file back from disk and compares each of the
160 metric values per company with the workbook, and checks the email against each of those rules.
*Point to:* `export.as_stored`, `export.email_html`, `check_export.outlook_problems`; `tests/test_export.py::test_a_value_has_the_16_significant_digits_the_workbook_stores`.

**Q34i. How do you make sure a demo doesn't go wrong in front of someone?** (new, Task 12)
Two things. A script, DEMO.md: five minutes, the exact clicks and what to say, what to do if
something goes wrong. And a reset, `python demo_reset.py`, run before every demo. Practice leaves
things behind: an approval, decks built without the AI text, exports. The reset deletes only the
files the tool builds, by the names the code gives them, and keeps the saved AI analyses, which are
the one thing that costs money to remake: it copies them first and checks their hashes after. It
rebuilds through the page's own Generate all button with a client that refuses any API call, and
builds last quarter first so "what changed" has something to show. Then it checks the state and says
"Ready for the demo." or what to fix. The script can't drift from the page either: a test walks its
clicks through the real page after a reset, and fails if a button it names is renamed.
*Point to:* `demo_reset.remove_built`, `demo_reset.readiness`, `demo_reset.NoApiClient`; `tests/test_demo_reset.py::test_demo_md_walkthrough_runs_on_the_page_after_a_reset`.

**Q34j. A partner edits a threshold and makes a typo. What happens?** (new, Task 13)
It stops before anything is built, and says exactly what to fix. config.yaml is checked against a
schema every time it is read: every setting has a kind, a range and an example. The dangerous typo is
a percent typed as a whole number: `rule_of_40_min: 40` is 4,000%, so every company would trip and
nobody would know why. The message says "rule_of_40_min must be a decimal from -1 to 1 (got 40). 40% is
written 0.4. Example: rule_of_40_min: 0.40". A misspelt key gets "Did you mean nrr_min?", because
otherwise it would be ignored and the old threshold would quietly stay. Every problem is listed at
once, and `main.py` prints them and exits 1; the web page shows them instead of the portfolio. I
proved the tests catch what they claim by breaking the checker 30 ways in a temporary copy.
*Point to:* `config_schema.SETTINGS`, `config_schema.config_problems`; `tests/test_config_schema.py::test_a_percent_typed_as_a_whole_number_gets_the_decimal_it_meant`.

**Q34k. How would an ops person run this on a schedule, and know if it worked?** (new, Task 14)
By the exit code. `python main.py --all` ends with 0 when every company was built, 1 when something
needs fixing (a bad config.yaml, no API key, a company that failed or timed out), 2 when the command
itself is wrong, and 130 when someone pressed Ctrl+C. A scheduler only reads that number, so it's
documented in one place in the code, `EXIT_CODES`; `--help` prints it and README has the same table,
and a test fails if the two ever differ. Before a run, `--list-companies` shows every company with its
flags and whether its deck is reviewed, in the web page's own words, and writes nothing. `--version`
names the commit, which is the same one every deck footer carries, so a deck can be matched to the
code that built it. I proved the tests catch what they claim by planting 28 bugs in a copy; three got
through at first, and each now has a test.
*Point to:* `main.EXIT_CODES`, `main.list_companies`, `main.check_combination`; `tests/test_cli.py::test_readme_documents_every_exit_code_in_mains_words`.

**Q34l. Last night's run took 40 minutes and Fernhollow failed. How do you find out why?** (new, Task 15)
From the run log. Every run writes `output/logs/run_<timestamp>.jsonl`: one line per step per company
(clean, metrics, what changed, Excel, AI, deck, memo, manifest) with how many seconds it took, its
result and the error, then a "whole company" line with the outcome. So I'd see which step failed and
the exact error, and which step the time went on (it's nearly always the AI call). Each line is a
complete JSON object written the moment its step ends, so even a run that crashed leaves everything
up to the crash; a single JSON file would be unreadable if cut off. Companies built side by side
share the file, with a lock so their lines never mix. The web page's Recent runs card reads the same
files, so someone who never opens a terminal sees it too, and a log that can't be written never stops
a board pack. The first real run caught a bug the tests had missed (the company's line said it
started when it finished), so there's now a test for it, and I planted 43 bugs in a copy to prove
the tests catch what they claim.
*Point to:* `run_log.CompanyLog.step`, `main.log_whole_company`, `run_log.recent_runs`; `tests/test_run_log.py::test_a_failing_step_is_logged_with_its_error_and_no_later_step_runs`.

**Q34m. An analyst runs this at 11pm and it says FAILED. What do they see?** (new, Task 16)
A sentence that says what's wrong, where, and what to do next, like "Sheet 'KPI Tracker', row 3
(header) is missing columns: pipeline - add a column headed with each name". I audited every message
the core files raise against those three things. Most already named the sheet and cell; what was
missing was usually the "what next", and in a few places the reader got the library's words instead
of mine: "ValueError:" in front of every failure, pandas' "you must specify an engine manually" for a
file that isn't a workbook, "AuthenticationError: Error code: 401" for a bad key, a YAML parser
error for a typo in a mappings file. Now each says what to do. The one place I kept the Python name
is a real bug, because the person fixing it needs it; it's labelled "probably a bug in this tool
rather than the workbook" so the analyst doesn't go hunting in their spreadsheet. The tests pin
each message's "what to do next", and I undid each fix one at a time in a copy to prove the tests
notice.
*Point to:* `run_log.error_text`, `analyze.api_error_text`, `clean.check_is_workbook`; `tests/test_error_messages.py::test_api_errors_say_what_happened_and_what_to_do`.

**Q34n. You added a cache. How do you know it can't put an old number on a board deck?** (new, Task 17)
First, why: each output was built to stand on its own, so one run read the same workbook five times
and worked out the same metrics up to eight. I measured before touching anything: 1.1 to 1.4 seconds
per company. The cache sits inside the three functions that do the work, so nothing that calls them
changed. It can't serve a stale number for three reasons. The key is a hash of the file's contents
and of its column mapping, never the file name, so an edited workbook is a new key. Errors are never
kept, so a problem is found again every time. And every answer is copied on the way in and out, so
one step changing its table can't change the next step's. The proof is two layers: the goldens (every
deck, memo and workbook, word for word) still match, and tests for each of those rules, each of which
I checked by planting that exact bug in a copy of the project and watching the test fail. One test
passed a bug the first time: it only changed an answer the cache had never handed out, so I made it
go three rounds. Result: 1 read instead of 5, runs 23 to 26% faster.
*Point to:* `cache.ResultCache.get`, `clean.clean_workbook`, `metrics.table_key`; `tests/test_cache.py::test_an_edited_workbook_is_read_again`; `python benchmark.py`.

**Q34o. Can everyone on the board actually read your deck?** (new, Task 18)
I checked every text color against every background it sits on, using WCAG's contrast standard:
AA needs 4.5 : 1 for text. Most of the palette passes easily (body text is about 10 : 1), but two
didn't. The "Passed" green on its light green cell was 4.25 : 1, so I darkened it a shade to 5.0 : 1.
And the template's cover page had a footer with no color of its own, so it inherited a dark gray on
navy, about 1.5 : 1: invisible. Nobody would have noticed that one until someone used the cover. The
test works the ratio out from theme.py, and doesn't only trust my list of pairs: it also reads the
pairs back from the real web page style sheet and from every colored piece of text in four built
decks. I planted 12 bad colors in a copy of the project and all 12 were caught. Status never relies
on color alone either: every cell also says "Tripped" or "Passed".
*Point to:* `theme.contrast_ratio`, `theme.TEXT_PAIRS`; `tests/test_contrast.py::test_every_colored_text_on_every_slide_passes_aa`.

**Q34p. How do you know a chart won't be unreadable next quarter?** (new, Task 19)
Because a test draws it at data it hasn't seen. The charts now share one set of rules: round whole-$K
ticks, zero always shown, every quarter the same width, and the latest value written just right of the
last bar, where nothing else is. Then a test renders both charts for the three companies and two made-up
extremes, 16 quarters in the billions and 2 quarters of under $1K, measures every label in pixels and
fails if any two touch. The old charts failed it: "data missing" labels ran into each other and a "0"
landed on the quarter labels. The fixes measure too: crowded quarter labels are thinned, a gap label
that won't fit is turned on its side. For color blindness, a quarter where ARR shrank is amber and
striped, and amber is far lighter than navy, so it reads even in grayscale. I planted 14 chart bugs in a
copy of the project and all 14 were caught.
*Point to:* `charts.finish_layout`, `charts.money_axis`; `tests/test_chart_layout.py::test_no_text_overlaps_anything_on_either_chart`.

**Q34q. A director wants every number in the deck, not just the latest quarter. What did you do?** (new, Task 20)
An opt-in appendix: `--appendix` adds one slide after the four with all 20 metrics for all 8 quarters.
Off by default, because the four slides are the update and the table is reference. It keeps the deck's
rules: `format_value` writes every number, and the same fit check applies, never below 12 pt. It didn't
fit at first: 20 rows at slide 1's padding need 403 pt and the slide has 392, and "n/a (no prior
period)" is wider than a quarter column. I kept 12 pt and gave up words instead. The cells are tighter,
the long reason words become "n/a", "n/m" and "∞", and a key under the table explains each one on the
slide. "data missing" keeps its words, because that's the one to chase. Gray and red cells match the
metrics workbook exactly, and tripped cells are bold too, so color isn't the only signal. `check_deck.py`
checks every cell against Excel. I planted 8 appendix bugs in a copy of the project and all 8 were caught.
*Point to:* `build_deck.appendix_slide`, `build_deck.appendix_key`; `check_deck.check_appendix_slide`.

**Q35. What breaks at 275 companies?** (new)
Not the math: Python is instant. Five things would:
1. **Time.** About 70 s of API time per company, so 5 hours one at a time. `--workers` now runs
   companies in parallel (Q34d); how many the account's rate limit allows is still to be measured.
2. **Formats.** Three companies prove three kinds of mess. 275 real workbooks will have header
   spellings and layouts clean.py has never seen, and it's built to stop rather than guess. New
   header spellings now get a proposed column to confirm once per company (`mapping.py`), so that
   queue is a review, not a code change. New layouts (no "Quarter" column, quarters across instead
   of down) still stop, and someone has to own that queue.
3. **Review.** A person reads every deck, and 275 decks a quarter is the real bottleneck. The
   portfolio rollup (Q34f) already ranks companies by flags tripped; the review queue should follow
   it, and the healthy companies get a lighter read.
4. **Output folder.** Everything goes to one output/ folder, and today the check scripts overwrite
   the real decks and manifests there. At scale that needs an output folder per run.
5. **Cost** isn't the problem: about $25 a quarter, roughly double for any company that needs the retry.

*Point to:* README.md "Next steps" and "Cost"; `main.quarter_mismatch_warning`; POLISH_REPORT.md Task 1 Unresolved.
The follow-ups ("which of your own choices would hurt?", "how would you know it worked?") are Q69 to Q72.

**Q36. What would you build next?** (guide Q25, replaced: the slide-fit and trend checks it listed are now built)
In order:
1. **Measure `--workers` against the real API**: parallel runs are built (Q34d), but how many
   workers the rate limit allows, and the real time for 275, are guesses until one live run.
2. **An output folder per run**, so a test run can never overwrite a real deck or manifest.
3. **A SharePoint or Power Automate trigger**: a company drops its workbook in a folder, the run
   starts, and the deck lands next to it, so nobody has to run a command.
4. **Close the claims code can't check yet**: whether a passing flag is really good news, and
   whether two figures are really "over the same period". Until then a person reads every deck.

*Point to:* README.md "Next steps".

---

## Working method

**Q37. How did you work with Claude Code on this?** (new)
Like managing an analyst whose work I have to be able to defend. CLAUDE.md is the spec: the rules
(Python does every number), every metric definition, each company's story, and "stop, don't guess".
I work in small steps: plan first, build one step, stop so I can review it, and have each file
explained to me in plain English, because I need to explain every line myself. Tests come before
code where possible, and the answer keys are worked out by hand, not copied from the code. Anything
that breaks goes into LEARNINGS.md with why and the fix, so mistakes turn into rules. And I read the
real output myself: the direction-claim failures were found by reading the summaries against the
data, not by any test.
*Point to:* CLAUDE.md; LEARNINGS.md; the git log (one commit per working piece).

**Q38. What about running it unattended?** (new)
For larger pieces of work I write a queue of tasks and let Claude Code run them without me, overnight
or during the day. Each task has explicit rules: don't call the paid API, don't edit config.yaml,
tests before code, commit after each working piece and push, log anything that breaks. The check
scripts run after every task, so a task that breaks the pipeline stops the queue instead of piling
more work on top. Permission rules limit which commands it can run, so it can't, say, start a server
or run arbitrary shell. The key is the report: every task ends with what it built, **every decision I
didn't specify**, what failed and how it was fixed, and what's unresolved. That's what I review. One
example of why that works: in one task a deliberate break test deleted a real saved analysis file.
It restored it from the session transcript, marked the file as restored, and said so in the report,
so I knew exactly what had happened.
*Point to:* OVERNIGHT_REPORT.md, DAY_REPORT.md, POLISH_REPORT.md; LEARNINGS.md rows on refused commands and the deleted analysis file.

**Q39. How do you know Claude Code didn't just write tests that pass?**
Three habits. First, the expected values are worked out by hand from the definitions in CLAUDE.md,
never copied from the code. The one time that rule was broken, a wrong formula passed every check
(Q27). Second, a new test has to fail on the old code before it counts: twice a new test passed
against the old behavior, which meant it proved nothing, and it was rewritten. Third, the code is
broken on purpose to see whether the tests notice: 35 of 36 deliberate breaks were caught in one
round, 18 of 18 in another, and 22 of 22 for the column mapping.
*Point to:* LEARNINGS.md rows on the font-size test and the "old analysis" test; OVERNIGHT_REPORT.md.

---

## Deep dives on the final run

The Scale and risk answers (Q34b to Q35) are the first answer on each topic. A good interviewer
then pushes: "why not just...?", "what if...?", "what can't it do?". These are those second
questions. Each names the file to open if they want to see it.

### Column mapping

**Q40. The tool proposes "Plan Burn" at high confidence. Why make a person click Confirm at all?**
Because the cost of being wrong is lopsided. Confirming takes a few seconds, once per company, and
every later quarter then runs unattended. A wrong mapping would put a wrong number on a board slide
with nothing to show it: if "Plan Burn" landed in net burn instead of budgeted burn, burn vs budget,
burn multiple and runway would all be worked out from the plan, and all would look plausible. The
confidence isn't a probability either. It's my heuristic's points, capped at 99%
(`mapping.MAX_CONFIDENCE`) because a heuristic is never certain. So confidence decides how a
proposal is shown, never whether it's used, and a test pins that a 99% proposal still stops the run.
*Point to:* `clean.clean_workbook`; `tests/test_mapping.py::test_a_high_confidence_proposal_still_needs_confirming`.

**Q41. Walk me through how it scores a header.**
Four steps, all in `mapping.propose`. One, candidates: only the standard columns no known header
already covers, so a column is never proposed twice. Two, the name: drop filler like "Total" and
"$K", swap the usual finance synonyms ("Opening" means starting, "Plan" means budget, "GP" means gross
profit), and compare with each column's words and known spellings, allowing for typos: "Revenu"
scores 92%. Three, the values: a number in the budget-only row rules out all 13 actual columns, ARR
and cash must still roll forward with the column in place, and gross profit can't be above revenue.
Evidence that fits adds points; evidence that breaks takes them off. Four, the strongest pair is
assigned first, and below 40% there's no proposal: the person picks. The proof: 29 headers renamed
across the three companies, and every one was proposed as the column it came from.
*Point to:* `mapping.propose`, `mapping.budget_row_evidence`, `mapping.SYNONYMS`; `tests/test_mapping.py::test_every_renamed_header_is_proposed_as_the_column_it_came_from`.

**Q42. Why word lists and arithmetic? Why not just ask Claude what the column means?**
I would add Claude, as a proposer, never a decider. The heuristics have real limits. "Cash Burn" is
about as close to ending cash as to net burn by name, and it only lands right because another header
took ending cash first. A word not in my synonym list, like "Bookings" or "Logos Lost", gets nothing
useful. A model reads "burn" as a flow and knows the vocabulary. The design: send the unknown headers,
their sample values and the 16 column definitions, ask for JSON, then run the same value checks in
Python, so a proposal that breaks the budget-row rule or a roll-forward is marked down whatever the
model says. A person still confirms. It's the same pattern as the commentary: the model suggests,
Python checks what's checkable, a person owns the rest. I built the heuristics first because they're
free, give the same answer every time, and can be tested.
*Point to:* README.md "Known limitations"; FINAL_REPORT.md Task 5, "Where a model call would improve it".

**Q43. Once a mapping is saved, what stops it quietly going wrong later?**
Three things. It's saved only if the workbook actually reads with it, tried first in a temporary
folder. A known header always wins over a saved line, so a hand-edited mapping file can't redefine
"Revenue". And the mapping file is an input, like the workbook: its SHA-256 hash goes in the
manifest, so changing it marks the company out of date and sends an approved deck back to "not
reviewed". One weakness I'd admit: the file is named after the workbook, so if a company renames its
file, the mapping is lost. Nothing wrong gets used; the review simply asks again.
*Point to:* `mapping.mapping_sha256`, `provenance.approval_status`; `tests/test_mapping.py::test_a_known_header_wins_over_a_saved_mapping`.

### The eval set

**Q44. You have over 1,300 unit tests. Why an eval set as well?**
They answer different questions. A unit test asks "does this function do what I think?", one function
at a time. The eval asks "does a whole company come out right?": a messy workbook goes through the
real clean, metrics, flags and gaps, and the result is compared with an answer key typed by hand. It's
built around the cases that break things, not the ones that demo well: exactly at every threshold, a
blank quarter first, second to last and last, zero revenue, negative budgeted burn, two quarters of
history, and two workbooks that must stop. And it's a scorecard, not just pass or fail: it collects
every mismatch by name across all 12, so after a change I see how many cases broke, not only the
first. It runs inside pytest, so nobody can forget to run it.
*Point to:* `eval/run_eval.py`; `tests/test_eval.py::test_every_company_matches_its_answer_key`.

**Q45. How do you know the answer keys are right, and not copied from the code?**
Values, flag statuses and every "not meaningful" cell are typed by hand from the definitions in
CLAUDE.md. Where a rule is applied, like which quarters a blank one spreads to, the prediction uses
hand-typed lists in check_northwind.py, not the inputs table in metrics.py, so the check shares no
code with what it checks. I learned that the hard way: an early bug passed because the answer key had
been written with the same wrong formula as the code (Q27). Every answer key must also tie out (ARR
and cash roll forward), and the threshold company's formulas must give each config.yaml threshold
exactly. Then I attacked the eval itself: 24 bugs planted one at a time, and each failed on the
company built for it.
*Point to:* `eval/make_eval_data.py`; `tests/test_eval.py::test_every_answer_key_ties_out`, `tests/test_eval.py::test_the_threshold_company_sits_exactly_on_every_threshold`.

**Q46. Did the eval ever miss something?**
Yes, and that was the most useful thing it did. In the first round of planted bugs, revenue growth
YoY computed three quarters back instead of four passed 11 of 12 companies. Only the eight flag
metrics were being checked for values, and revenue growth isn't one of them. So the healthy company
now checks all 20 metrics by hand formula, a test pins it, and the same bug then failed 4 companies.
The lesson: an eval only covers what it compares, and you find out what that is by attacking it. What
it still doesn't cover: a partly blank quarter, a workbook with no budget row, unknown headers. Those
have unit tests instead, and each could become one more eval company.
*Point to:* `tests/test_eval.py::test_the_healthy_answer_key_covers_every_metric_not_just_the_flags`; FINAL_REPORT.md Task 6.

**Q47. The eval scores the numbers. How would you score the AI commentary?**
It's designed, not built. Replay saved analyses through the checks, so no API call. Each eval company
gets a saved analysis plus copies with a planted error: an invented number, a dropped minus sign,
"fell" on a series that rose, a tripped flag left out of the risks. The existing checks run offline
(the number, fit and direction checks, the deck's and the memo's), plus new rules from the answer
keys: every flag that trips is named among the risks, and no number is quoted for a metric that has
none. Each planted copy must fail its rule and the clean one must pass. What stays human is whether
the commentary is actually insightful: that's the blind 1 to 5 score. Recording the analyses needs
one live run, about a dollar for ten companies.
*Point to:* FINAL_REPORT.md Task 6, "How the same harness would score AI commentary offline"; README.md "Next steps".

### Batch resilience

**Q48. The laptop dies halfway through a 275-company batch. What state is output/ in?**
Consistent, and that's the point. Each company is built in its own private folder under
`output/.staging` and moved into output/ only when every step succeeded, with the manifest moved
last. So a company has either all new files or all of last quarter's, never a new deck beside an old
memo. If the power goes during the move itself, the manifest still describes the older files, so the
next run sees they don't match and rebuilds that company. The next batch also clears whatever a
killed batch left in staging. Then `--resume` skips every company whose outputs were built from
exactly today's workbook, config.yaml and mapping, checked by hash, with every file present. The
rerun only redoes what's missing, and those reuse the saved analysis when the numbers haven't
changed, so they cost nothing.
*Point to:* `resilience.make_stage`, `resilience.commit_stage`, `resilience.resume_problem`; `tests/test_batch.py::test_the_manifest_is_moved_into_the_output_folder_last`, `tests/test_batch.py::test_the_next_batch_clears_what_a_killed_batch_left_in_staging`.

**Q49. Why retry only rate limits, and not every API error?**
Because a retry only helps if the next try can go differently. A rate limit is the API saying "not
now": wait and it clears. So it waits what the API asks for, else 5, 10, 20, then 40 seconds, never
more than 60 at once, and gives up after four retries. Anything else, a bad request, a bad key, a
server error, would most likely fail the same way again, and every wasted wait adds up over 275
companies. So those go straight to the placeholder slide: the deck is still built from the computed
numbers, the batch says "OK (AI failed)", and the reason is saved in the analysis JSON. The retry
for a failed validation is a different thing: once, with the list of problems, because Claude can
fix a specific mistake it's told about.
*Point to:* `resilience.RateLimitRetry`, `resilience.rate_limit_wait`; `tests/test_batch.py::test_other_api_errors_are_not_retried`, `tests/test_batch.py::test_rate_limits_that_never_clear_give_the_placeholder_not_a_failed_company`.

**Q50. What does --timeout actually do? Can it stop a hung API call?**
No, and I'd say so. Python can't safely stop a running step from outside. What `--timeout` does is
stop waiting: after that many seconds the batch gives up on the company, records it as timed out, and
starts the next one. The company's own thread notices at its next step and stops, and because it was
working in its private folder, nothing it built reaches output/: last quarter's files stay. But a
call already in flight carries on until it returns, up to the SDK's own 10-minute limit, and what it
cost still counts toward `--max-cost`. One planted bug survives the tests here ("a company that
finishes after its timeout moves its files in"), because a second guard throws its folder away first.
Two guards for one case, so no test can show either one alone doing the work.
*Point to:* README.md "Known limitations"; `tests/test_batch.py::test_a_company_that_wakes_after_its_timeout_while_the_batch_runs_on_lands_nothing`.

**Q51. How many workers would you run?**
I don't know yet, and I'd say that rather than guess. Workers help because the time is almost all
waiting for Claude: about 70 seconds a company, so 5 hours one at a time for 275, and about 1.5 hours
with four if the account allows it. The limit is the account's rate limit, and `--workers` has never
run against the real API. The plan: start with 2 on a real quarter, read the rate-limit waits each
company records, and step up until the waits start. The code is ready for it: the spend meter is
shared safely between companies, each company prints its lines together, and charts are drawn without
pyplot so two workers can draw at once. A test proves three workers really run three companies at the
same time, with a fake client that counts calls in flight.
*Point to:* `resilience.SpendMeter`; `tests/test_batch.py::test_three_workers_run_three_companies_at_the_same_time`; README.md "Cost".

### Cost ceilings

**Q52. How does --max-cost work? Can the batch spend more than the ceiling?**
It can, by a bounded amount, and that's deliberate. Every AI call's cost is worked out from its tokens
at the published price and added to a spend meter. Before each company starts, the batch checks it: at
or past the ceiling, that company is marked STOPPED, with the spend and the ceiling in its manifest,
the summary and the batch manifest, and the exit code is 1 because not everything was built. It checks
before a company starts, not halfway through, because stopping a company mid-run would waste what it
already spent and leave nothing to show for it. So the overshoot is at most what the companies already
running spend: about 9 cents each, double with a retry. Failed attempts count, because Claude charged
for them. A company skipped by `--resume` is never stopped: it costs nothing.
*Point to:* `main.start_or_settle`, `main.stopped_result`; `tests/test_batch.py::test_a_spend_exactly_at_the_ceiling_stops_the_batch`, `tests/test_batch.py::test_failed_validation_counts_toward_the_spend`.

**Q53. Apart from the ceiling, what keeps the AI bill down?**
Four things, and none of them lowers quality. Reuse: if the facts Claude would see are exactly the same
as a saved analysis, the saved one is used for free, after passing the same checks again. "Exactly"
means the whole payload, not just the company and quarter, so an edited workbook never gets an old
analysis. `--resume` skips companies that are already up to date. Only one retry for a failed answer:
if the second fails too, that points to a prompt problem that more retries would just pay for. And
`--skip-ai` builds every deck with no API call at all. What I didn't do is switch to the cheaper
model: Haiku cost about a third as much per run but scored 2.0 against Sonnet's 4.0 blind, and a false
claim on a board slide costs far more than the few cents saved.
*Point to:* `main.reusable_analysis`; `tests/test_batch.py::test_a_resume_rebuild_reuses_a_saved_analysis_of_the_same_numbers`; README.md "Cost".

**Q54. Where do your cost figures come from? Are they estimates?**
They're measured. Every run's manifest records the model, the prompt version and the input and output
tokens, a retry included. Cost is tokens times the published price, $2 in and $10 out per million for
Sonnet, typed into `compare_models.PRICES` with the date it was checked. The batch adds the companies
up, prints "AI spend this run", and saves the total in `output/batch_manifest.json`; the summary CSV
has each company's cost. The README's $0.09 a company and $25 a quarter for 275 come from a real live
run of all three companies, and the table says which one. The caveat I'd give: the price is typed in
with a date, so if Anthropic changes it, someone has to update that line, and three companies is a
small sample for an average.
*Point to:* `main.ai_cost`, `compare_models.PRICES`, `compare_models.PRICES_AS_OF`; README.md "Cost".

### Golden files

**Q55. Doesn't a golden file just freeze whatever bugs were there when you approved it?**
Yes, which is why it isn't the only check. A golden answers "is this exactly what a person
approved?", not "is it right?". If the approved deck had a wrong number, the golden would protect the
wrong number. So the jobs are split: the check scripts and the eval prove the numbers against hand
formulas, and the goldens hold everything else still: wording, sizes, colors, positions, page breaks,
Excel formats. That's where they earn their place. Of 17 planted bugs that change no number, like the
two charts swapping sides or the Excel header row no longer frozen, the goldens caught 17 and the
number checks caught 2. And approving means something: I read each one before it became the
definition of right, and an intended change means reading the diff before committing it.
*Point to:* `golden.compare`; FINAL_REPORT.md Task 8 (the 17-bug table); `tests/test_golden.py::test_the_output_matches_its_approved_golden`.

**Q56. Why compare text dumps instead of the files themselves?**
A .pptx, .docx or .xlsx is a zip file with timestamps inside, so its bytes change on every save: a
byte comparison would always fail and never say why. So `golden.py` turns each output into plain text,
one fact per line: every word with its size, bold and color, where each box sits, table fills, the
memo's page breaks from the PDF, every Excel cell's value and format. A difference is then a readable
diff: a minus line for what was approved, a plus line for now. Two things change by themselves, the
date and the git commit in the footer, so both are fixed while building. The AI text comes from
committed copies of saved analyses, so there's no API call and the text is the same every time.
Charts are compared by name, place and size, not pixels, or a matplotlib update would change every
golden.
*Point to:* `golden.dump_deck`, `golden.RUN_DATE`; `tests/test_golden.py::test_building_twice_gives_the_same_text`.

**Q57. What don't the goldens cover?**
A few things, on purpose, each covered somewhere else. Only the version with AI text has a golden: the
placeholder slide, the `--draft` watermark and the "reviewed by" footer have unit tests instead. The
memo's "what changed" section isn't in them, because the goldens are built with no earlier run;
check_diff.py checks those words exactly. The rollup has no golden, because its review column depends
on what's in output/ at the time. Chart pixels aren't compared. And one risk to know about: a package
upgrade, the PDF library say, can change a golden with no code change. That's worth reading when it
happens, but it isn't always a bug in this project.
*Point to:* FINAL_REPORT.md Task 8, "Unresolved"; `check_diff.py`, `check_rollup.py`.

### The approval gate

**Q58. What stops someone approving a deck they never opened?**
Nothing technical, and I wouldn't claim otherwise. The approval isn't proof that someone read the
deck; it's accountability: a named person and a time on the record, in the manifest and in the footer
of every slide. What the code does make sure of is that the name goes against the right numbers.
`approve.py` needs a name, given or taken from git. It refuses if there's no run, and it refuses if
the workbook, config.yaml or the column mapping has changed since the deck was built, because the
reviewer would be signing off numbers they never saw. After approval, if any of the three changes,
the deck goes back to "not reviewed" by itself. And approve.py never builds a deck, so making a deck
and vouching for it stay two separate acts.
*Point to:* `approve.approve`, `approve.reviewer_name`; `tests/test_approve.py::test_a_workbook_changed_since_the_run_stops`, `tests/test_approve.py::test_a_reviewer_name_is_required`.

**Q59. Why does a changed threshold undo an approval? The numbers didn't change.**
But the deck did. A threshold decides which flags trip, and flags are most of slide 3 and the top of
the memo. If a partner moves the runway threshold from 12 to 18 months, a company that passed now
trips, and the deck the reviewer approved said something different. So the approval stores the hashes
of the workbook, config.yaml and the mapping at the moment of signing, and `approval_status` compares
them every time it's asked. Any difference sends the deck back to "not reviewed" with the reason in
words: "config.yaml has changed since it was approved". A stale approval is worse than none, because
it tells the board someone checked what nobody checked. It's the same idea as re-signing a
reconciliation when the source numbers move.
*Point to:* `provenance.approval_status`; `tests/test_provenance.py::test_changed_thresholds_send_the_deck_back_to_draft`.

**Q60. Does approving the deck approve the memo too?**
Only if the memo was there to read. The approval lists the documents it covers: the deck, and the memo
if that run built one. The memo's footer says "reviewed" only when the memo is on that list, so an
approval recorded before memos existed leaves its memo "not reviewed": nobody can have read a memo
that wasn't there. The web page's Approve button only appears for files built from today's workbook.
And the `--draft` watermark follows the same judge: once a deck is approved, a `--draft` rebuild no
longer stamps it.
*Point to:* `approve.reviewed_documents`, `memo.memo_approval`; `tests/test_approve.py::test_approving_a_run_that_built_a_memo_covers_the_memo_too`.

### The run diff

**Q61. Why "more than 5 points or 10%"? Aren't those numbers arbitrary?**
The sizes are judgments, and I'd say so. What isn't arbitrary is that there are two. Percentages move
by points, everything else by percent of its old value, because one number can't mean both: NRR going
from 102% to 96.9% is 5.1 points but only 5% of itself, and a runway can't move "5 points". "More
than" means a move of exactly the setting isn't listed, the same rule as "exactly at a threshold
passes". config.yaml can override both (`diff_min_points`, `diff_min_relative`); I didn't add those
keys myself, because a new key changes config.yaml's hash and would have sent every approved deck back
to "not reviewed". And a flag that flips is always listed, whatever the size: Northwind's NRR moved 4.9
points, so it isn't listed as a move, but it's there as a flip.
*Point to:* `diff_runs.DEFAULT_MIN_POINTS`, `diff_runs.move_settings`; `tests/test_diff_runs.py::test_a_percentage_moves_by_points_and_exactly_the_setting_is_not_more_than_it`.

**Q62. What can't the diff tell you?**
Three limits. It keeps one step of history: this run and the one it was compared with, so "since two
quarters ago" would need a history file per company. It compares metrics only for each run's latest
quarter (data gaps are compared in every quarter), so if a company restates last year's numbers, the
diff won't list it unless a gap opens or closes. And a move's size comes from the exact values, not
the rounded ones shown, so someone checking by hand can be 0.1 out: net new ARR vs budget "down 17.6
pts", where subtracting the numbers on the page gives 17.5. It's the true move, but I'd warn a reader.
*Point to:* `diff_runs.move_text`; FINAL_REPORT.md Task 10, "Decisions you didn't specify" and "Unresolved".

**Q63. How did you test it without waiting a quarter for real data?**
`check_diff.py` rebuilds each company's workbook as it stood a quarter ago, from the answer key: the
latest quarter left off, and that quarter's budget turned into the budget-only row in the company's
own wording. It runs that, then today's workbook, into the same temporary folder, and compares every
line of the memo's section, in Word and in the PDF, with lists I worked out by hand, the formula
beside each line: burn multiple 3650/2020 = 1.81x becomes 3900/1660 = 2.35x. It also checks that the
page and the command line say the same, and that approving and rebuilding still compares with Q1, not
with a run ten seconds old. 25 planted bugs, all caught after one fix.
*Point to:* `check_diff.py`; `tests/test_diff_runs.py::test_a_rebuild_with_the_same_results_keeps_the_comparison_it_already_had`.

### The rollup

**Q64. Why doesn't the rollup have an AI summary of the whole portfolio?**
Because a portfolio page is where unreviewed text would do the most damage. The company decks carry AI
text with a label and a review gate. The rollup is a partner's one view of every company, and nobody
checks a portfolio summary line by line against 275 workbooks. So it's computed metrics only, and the
footer says so: "computed metrics only, no AI text". Every number is worked out from the workbooks at
the moment it's built, never read from last run's files, so it can't be stale; the only thing it reads
from output/ is each company's review status. If I added AI to it later, it would follow the decks'
pattern: numbers checked against the data, labelled, and reviewed.
*Point to:* `rollup.NO_AI_TEXT`, `rollup.company_entry`; `check_rollup.py`.

**Q65. What does the rollup do with a broken workbook, or two companies tied?**
A broken workbook doesn't stop it: that company is listed last, unranked, "Workbook can't be read" in
every output, with clean.py's reason in the workbook's Note column, and the others still rank. A tie
on flags tripped goes to the company whose worst flag is higher in the fixed order, then by name, so
the order never depends on which file happened to be read first. And it's sized for a real portfolio:
7 companies per ranking slide, running on to as many slides as needed, and at most 3 names per status
on the slide ("and 7 more"; the workbook lists them all), all measured with a 40-character name, the
longest the web page accepts. check_rollup.py builds a 12-company portfolio to prove it fits.
*Point to:* `rollup.rank_key`, `rollup.ROWS_PER_SLIDE`; `tests/test_rollup.py::test_an_unreadable_workbook_goes_last_with_no_rank_and_the_others_still_rank`, `tests/test_rollup.py::test_a_tie_on_flags_tripped_puts_the_worse_worst_flag_first_then_the_name`.

### Exports

**Q66. A data engineer loads your CSV. How do they know which version of the numbers they have?**
The JSON carries the SHA-256 hashes of the workbook, config.yaml and the column mapping it was built
from, the same hashes the manifest records, so a tool can check whether an export matches the inputs
it has. There's no run time in any export file, on purpose: the same workbook and thresholds give
byte-for-byte the same files, so a tool can tell a real change from a re-run. The JSON also has a
format version, raised whenever a field is renamed or removed, so a loader breaks loudly instead of
reading the wrong field. And the CSV is long, one row per quarter and metric, so two companies' files
stack without reshaping.
*Point to:* `export.export_record`, `export.FORMAT_VERSION`; `tests/test_export.py::test_the_same_numbers_always_give_the_same_files`, `tests/test_export.py::test_the_json_names_its_sources_and_has_no_nan_or_infinity`.

**Q67. Why is there no AI text in the email?**
Because an email is the easiest place for unreviewed text to escape. On the deck the AI commentary is
labelled "review before use", and the footer says whether anyone did. Paste a paragraph into an email
and forward it, and that label is gone. So the email is slide 1's table in the status colors, the flag
count, runway at budget and the data gaps, all computed, with a line saying so. And honestly: it
follows every rule I know Outlook breaks (inline styles only, a font on every cell or it falls back to
Times New Roman, fills repeated as bgcolor, 640 pixels wide), and check_export.py tests each one, but I
haven't pasted it into real Outlook. That's one paste I'd do before telling anyone it works.
*Point to:* `export.email_html`, `check_export.outlook_problems`; `tests/test_export.py::test_the_email_says_what_it_is_and_that_it_has_no_ai_text`.

**Q68. Why are exports made on demand instead of by every batch run?**
Four more files per company on every run would touch `--resume`'s list of expected files, the
manifests and the goldens, for files most runs don't need. So it's `python export.py` with a workbook
or `--all`, or the company page's Export button, which builds from today's workbook when clicked and
never writes into output/. Because the JSON carries its input hashes, a tool can always check which
inputs a file reflects, whenever it was made. Wiring exports into main.py is a small change if a data
team wants them after every run: that's a choice about their workflow, not a technical limit.
*Point to:* `export.save_exports`; FINAL_REPORT.md Task 11, "Decisions you didn't specify".

### What breaks at 275 companies

**Q69. Which of your own design choices would hurt first at 275?**
Q35 covers time, formats, review and cost. These are the ones in my own code, and I'd rather name them
than be asked:
1. **The mapping file is named after the workbook.** A company that sends a file called "Acme Q3 KPIs"
   one quarter and "Acme Q4 KPIs" the next gets asked to confirm its columns every quarter. Fix: key
   the mapping by company, not by file name.
2. **The rollup ranks 7 companies per slide.** For 275 that's about 40 ranking slides, and a partner
   reads the first one. It would need a "worst 20" slide, with the full list in the workbook.
3. **The diff remembers one run back**, so trends across quarters need a history file per company.
4. **The different-quarters warning becomes noise.** At 275, some companies are always a quarter late;
   that should be a "late reporters" list, not one warning line.
5. **The rollup and exports aren't part of the batch**: two more commands after every run.
6. **The web page has only been seen with a handful of companies.** It has search, but a table of 275
   rows hasn't been tried.

*Point to:* FINAL_REPORT.md "Unresolved" in Tasks 5, 9 and 10; `rollup.ROWS_PER_SLIDE`; `main.quarter_mismatch_warning`.

**Q70. The 275-company batch ran overnight. How do you know the next morning that it went right?**
Four places, from quickest to most detailed. The exit code: 0 means every company was built, 1 means
something needs a look, and a scheduler reads just that. The summary table and `batch_summary.csv`:
one row per company with its outcome (built, "OK (AI failed)", failed, timed out, stopped at the cost
ceiling, or skipped as up to date), its flags, the AI cost and a note saying why. The batch manifest:
the options, the commit, the total spend and every outcome. And the run log: one line per step per
company, with seconds and the exact error, so "Fernhollow failed at the memo step after 3 seconds" is
one search. Then the rollup is the morning's review queue: worst company first. What I'd add at 275 is
a one-paragraph morning summary sent to the team, counts by outcome, built from the batch manifest.
*Point to:* `main.EXIT_CODES`, `main.save_batch_manifest`, `run_log.CompanyLog.step`; `rollup.py`.

**Q71. On day one, 20 of the 275 workbooks fail. What do you do?**
Triage by the message, because each FAILED line says what's wrong, where, and what to do next. They
fall into three kinds. Unknown headers: that's the mapping review, a few minutes once per company, and
next quarter runs unattended. A layout clean.py can't read, like quarters across the top instead of
down: that stops, and the choice is to ask the company to use the standard layout or to add a reader
for it, and any new reader gets an eval company first, so it's proven before it's trusted. A real data
problem, like "TBD" typed in revenue: that goes back to the company. What I would never do is loosen
clean.py so it "just reads" the file. Every loosening is a guess, and a guess can reach a board.
*Point to:* `main.describe_error`, `mapping.propose`, `eval/run_eval.py`; CLAUDE.md "Stop, don't guess".

**Q72. What scales worse, the cost or the review?**
The review, by a long way. The AI is about $25 a quarter for 275 companies. The review is 275 decks: if
a careful read took 10 minutes, that's about 45 hours a quarter, and that 10 minutes is my assumption,
not a measurement. Three things shrink it without removing the person. The checks remove whole classes
of error: no number has to be re-added by hand, because Python did the math and code checked every
number Claude quoted, so the reviewer reads for judgment, not arithmetic. The rollup puts the worst
companies first, and a company with no flags gets a lighter read. And scoring the commentary offline
(Q47) would catch more before a person sees it. What I wouldn't do is approve automatically: the footer
saying "reviewed by" has to mean a person did.
*Point to:* README.md "Cost"; `rollup.WORST_FIRST`; `provenance.approval_status`.

---

## When you can't recall a detail

You will blank on something. The project's own rule applies to you: **never invent a number**. A
confident wrong answer does more damage than "I'd have to check", the same way one wrong number in a
deck undermines the rest.

- **Say what you do know, then where the rest lives.** "I don't remember the exact figure. It's in
  the Cost table in the README, and it came from the second live run." That shows you know the system,
  even without the number.
- **Give the shape, not a made-up precision.** "Roughly double, about 9 cents a company instead of
  5" is honest. "$0.0911" said with confidence when you aren't sure is not.
- **Explain how it's worked out instead.** If you can't recall runway, say the formula: cash divided
  by a month's burn. The method is what they're testing.
- **Be clear about who did what.** Claude Code wrote much of the code; you set the rules, the
  definitions, the decisions and the review. "I specified that, and I reviewed the test, but I'd
  have to look at how the function does it line by line" is a strong answer, not a weak one.
- **Correct yourself out loud.** If you realize mid-answer you got something wrong, say so straight
  away. That's the same habit as logging every break in LEARNINGS.md.
- **Offer to show it.** "I can open the manifest and show you" beats guessing.

**Know which kind of detail you've lost.** There are three, and each has its own honest answer:

- **A number** (a wait time, a count, a default). Give the shape and the file: "The waits double from
  a few seconds and never go past a minute. The exact values are constants at the top of
  resilience.py." Never round a guess into a precise-sounding figure.
- **A name** (a function, a test, a setting). Describe what it does and where it lives: "There's one
  function that decides whether an approval still holds. It's in provenance.py, and it compares three
  hashes." The name is the least important part.
- **A mechanism** (how exactly something works). Go back to the principle it follows, then say what
  you'd check: "It follows the stop-don't-guess rule, so I'm confident it stops rather than reads it.
  How it detects that case I'd want to look at in clean.py before I told you."

**Four situations the deep dives make likely:**

- **"Did you build that?" when you only designed it.** Say so plainly. Scoring the AI commentary
  offline (Q47), a model call for column mapping (Q42), an output folder per run and a SharePoint
  trigger are designs or next steps, not code. "That's designed, not built; the design is in the
  final report" is a strong answer. Claiming it's built, and then being asked to show it, is not.
- **"How many planted bugs?"** Every task in the final run planted bugs in a throwaway copy to prove
  its tests notice. You won't remember every count. Remember the pattern and the exceptions: nearly
  all were caught; each task's section in FINAL_REPORT.md has the table; and the survivors had a
  reason, like the timeout (Q50), where two guards protect one case so no test can show either one
  alone doing the work.
- **A decision you didn't make yourself.** Many small choices were made by Claude Code and flagged in
  each report's "Decisions you didn't specify". Own the review, not the authorship: "That was flagged
  for me, I checked it, and I kept it because..." If you can't remember why, say you'd reread the
  reasoning, which is written next to it.
- **A question about scale you haven't measured.** Workers against the real rate limit, review time
  per deck, the page with 275 rows: none is measured. Say "unmeasured", give the plan to measure it,
  and give any number as an assumption ("if a read takes 10 minutes..."), never as a finding.

**Useful sentences, in your own words:**
- "I don't want to give you a wrong number. It's in [file], and I can show you."
- "What I do know is the rule it follows: [rule]. The detail of how is in [file]."
- "I'd have to check that. It's exactly the kind of thing I'd look up rather than guess in a board pack."
- "Let me correct what I said a moment ago: [correction]."

**Numbers worth knowing cold** (from Northwind, Q2 2026, and the live runs):

| What | Number | Where it's from |
|---|---|---|
| Flags tripped | Northwind 6 of 9, Alderpeak 0, Fernhollow 7 (+1 cannot evaluate) | `check_companies.py` |
| Northwind NRR (annualized) | 97.1%, down from 108.0% | STUDY_GUIDE.md exercise 1 |
| Northwind runway | 11.0 months (13.0 if burn returns to plan) | STUDY_GUIDE.md exercises 2 and 7 |
| Northwind burn vs budget | 20.0% over | STUDY_GUIDE.md exercise 3 |
| AI cost per company | about $0.09 ($0.05 before the direction rules) | README.md "Cost" |
| 275 companies per quarter | about $25 ($12.81 before) | README.md "Cost" |
| API time per company | about 70 s, so 5 hours for 275 one at a time | README.md "Cost" |
| Blind model scores | Sonnet 4.0, Haiku 2.0 (rule: Haiku needs 4.0) | README.md "Model comparison" |
| Unit tests | 1,384 (say "over 1,300") | `python -m pytest -q` |

**And for the deep dives** (roughly is fine; the file is where to check):

| What | Number | Where it's from |
|---|---|---|
| Renamed headers proposed right | 29 of 29, across the three companies | `tests/test_mapping.py` |
| Proposal floor, and the cap | none below 40%; never above 99%, and even 99% needs confirming | `mapping.py` constants |
| Eval set | 12 companies, 12 of 12 match; 24 of 24 planted bugs caught | `eval/run_eval.py`; FINAL_REPORT.md Task 6 |
| Rate-limit waits | 5, 10, 20, 40 s (or what the API asks), no single wait over 60 s, then give up | `resilience.py` constants |
| Cost ceiling overshoot | at most what the companies already running spend, about 9 cents each | `main.start_or_settle` |
| Goldens | 9 files (3 companies × deck, memo, metrics); caught 17 of 17 look-only bugs, number checks 2 | FINAL_REPORT.md Task 8 |
| Approval is voided by | a changed workbook, config.yaml or column mapping (three hashes) | `provenance.approval_status` |
| Diff defaults | more than 5 points for percentages, more than 10% for everything else | `diff_runs.py` constants |
| Northwind, Q1 to Q2 | 5 flags flipped, 8 metrics moved, 1 gap closed | `check_diff.py` |
| Rollup | 7 companies per ranking slide; Fernhollow first, worst flag runway 6.0 months | `rollup.py`, `check_rollup.py` |
| Exports | 160 metric values per company (20 × 8 quarters), 16 significant digits | `check_export.py` |
