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
7. [When you can't recall a detail](#when-you-cant-recall-a-detail)

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
3. Over 500 pytest unit tests check each function, including every edge case, with the hand math in
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
were value-checked, so a wrong revenue YoY slipped through. The healthy company now checks all 19.
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
152 metric values per company with the workbook, and checks the email against each of those rules.
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
| Unit tests | over 500 | `python -m pytest -q` |
