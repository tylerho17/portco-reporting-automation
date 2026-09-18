# Final report

One section per task in the final run. Each says what was built, the decisions you didn't specify,
what failed and how it was fixed, and anything unresolved.

---

## Task 1: the board memo (memo.py, check_memo.py, main.py)

### What I built

- **`memo.py`** builds a 1 to 2 page board memo for one company and saves it twice, beside the deck:
  `output/<company>_board_memo.docx` (Word) and `output/<company>_board_memo.pdf`. Top to bottom:
  title and quarter; the AI headline under "AI-drafted from computed metrics - review before use";
  the key metrics table (latest, prior, budget or threshold, status in the Excel red / green / gray);
  runway at next quarter's budgeted burn; "Flags: 6 of 9 flags tripped" with each tripped flag's value
  and threshold, the flags that can't be evaluated and the combo rule; data gaps; the AI's 3 questions
  for management. The footer on every page is the deck's footer: fictional-data note, file, date,
  commit, model, review status.
- **How it works:** the memo is built once as a list of "blocks" (title, heading, paragraph, bullet
  list, table) from the same text helpers the deck uses (`value_text`, `threshold_text`,
  `status_label`, `gaps_lines`), then written twice (`write_docx`, `write_pdf`). The Word file and
  the PDF can't say different things, and check_memo.py proves they don't.
- **No number typed by hand:** `memo.py` is added to the deck's existing test
  (`tests/test_build_deck.py::test_no_digit_in_any_text_written_in_the_code`), which reads the code
  and fails on any digit in a string.
- **AI commentary unavailable:** a missing, failed, stale, tampered or too-long analysis gives
  "AI commentary unavailable" in both AI places, and every computed number still appears (a test
  proves the computed blocks are identical with and without the AI text).
- **`check_memo.py`** builds all three memos and proves, from the saved files, that every number in
  the Word body and the PDF body is in that company's saved metrics workbook, as Excel displays it.
  It also checks the table row by row against the Metrics and Flags sheets, the flag count against
  each company's story, every data gap, the AI text against the JSON (no wins or risks), that the PDF
  holds every piece of the Word text, no em dash anywhere, the PDF is 1 or 2 pages, and the footer
  (worked out independently from git, the analysis and the manifest). In a temporary folder it proves
  that a tampered analysis gives "AI commentary unavailable" and that a number planted in a copy of a
  saved memo is caught.
- **`main.py`** builds the memo after the deck for every company (`memo_step`), prints a ✓ Memo line
  (and why the AI text isn't in it, if it isn't), and records a `memo` part in the manifest
  (its two files, and whether it carries the AI text).
- **Tests:** 35 in `tests/test_memo.py`, 5 new in `tests/test_main.py`, 2 new in
  `tests/test_approve.py`. 575 tests in all, every check script passes.
- **Docs:** README (outputs table, commands, the memo's extra rule, approval), STUDY_GUIDE (function
  tables for memo.py and check_memo.py, main.py and approve.py rows, test counts), LOOM_SCRIPT test
  count, 6 LEARNINGS rows.

**New packages (requirements.txt), and why:**
- `python-docx`: writes the Word file. Nothing already installed can.
- `reportlab`: draws the PDF directly. Converting the Word file to PDF would need Word or LibreOffice
  installed, and neither is here or on a fresh clone.
- `pypdf`: reads a PDF's text back, so tests and check_memo.py can prove what is in the PDF. It is not
  used to build anything.

### Decisions you didn't specify

1. **The memo refuses AI text the deck may accept.** Every number in the memo must be in the metrics
   workbook, but Claude may quote anything in its payload, which includes raw inputs (net burn and
   ending cash in $K) the workbook doesn't show. So `memo_analysis` runs every check the deck makes,
   then checks the headline and questions against the workbook's numbers. None of today's three saved
   analyses is affected; a check with Northwind's ending cash in a question proves the deck takes it
   and the memo doesn't. `main.py` prints the reason and the manifest records `"ai_text": false` for
   the memo. The Result column stays about the deck.
2. **Same analysis checks as the deck, including slide fit.** An analysis too long for slide 4 is
   refused by the memo too, even though a page has room. I chose "the deck and the memo show the same
   AI text or neither does" over squeezing more AI text onto paper.
3. **Headline and questions only.** You listed those two; the wins and risks aren't in the memo (the
   deck shows risks). check_memo.py proves no win or risk leaks in.
4. **"6 of 9 flags tripped" counts as in the workbook** because it is the Flags sheet's rows counted
   by status. No cell shows it. Without this, Alderpeak's AI headline ("All 9 flags passed") would be
   refused. Stated in check_memo.py's docstring.
5. **The combo rule uses the Flags sheet's words** ("NRR falls at least 1 pt ... over the last 3
   quarters"), not the deck's ("1.0 pts"), so its numbers are the workbook's own.
6. **Questions are bullets, not numbered.** "1.", "2.", "3." would be numbers in the memo that aren't
   in the workbook.
7. **Em dashes.** Two labels shared with the deck and Excel carry one (the "Cannot evaluate" status
   and the "None" data gaps line). The memo shows a colon ("Cannot evaluate: missing input"); the deck and Excel are unchanged (that's
   Task 4). memo.py writes the character as `chr(0x2014)`, so the file itself contains none.
8. **Key metrics table:** the deck's three context rows plus Net new ARR ($K), then all 9 flags; the
   combo row says "see Flags below". US Letter, 0.7 inch margins; Arial in Word; DejaVu Sans in the
   PDF because the PDF's built-in fonts have no "∞" (Fernhollow's burn multiple is "∞ (ARR shrank)").
9. **Page breaks:** each PDF section (a heading and what follows) stays on one page when it fits;
   Word headings are "keep with next". A long footer wraps in the PDF instead of running off the page.
10. **Approval covers named documents.** approve.py now records `"documents": ["deck", "memo"]`
    (the memo only if that run built one), and the memo's footer says "reviewed by" only if its
    approval lists "memo". Northwind was approved on 2026-09-17, before memos existed, so its deck
    says "reviewed by Tyler Ho" and its memo says "not reviewed" until someone approves again.
11. **No DRAFT watermark on the memo.** `--draft` stamps the deck only; the memo's footer states the
    review status on every page.
12. **A memo that fails to build fails the company**, the same as a deck (no special result text).
    Old memo files are deleted before each build, so a failure never leaves last run's memo looking
    current. `python memo.py` on its own updates an existing manifest's memo part, as build_deck.py
    does for the deck.

### What failed and how I fixed it (all logged in LEARNINGS.md)

1. **A correct AI question was refused.** Fernhollow's first memo said the AI text quoted "15", a
   number not in the workbook; it was "threshold 15.0%". memo.py listed thresholds as stored (0.15),
   not as Excel shows them. Fixed with the table's own threshold wording, plus a test that fails on
   the old code.
2. **check_memo.py's first run failed on "2026,".** check_deck.py's number pattern keeps a comma after
   digits; the memo's "Q2 2026, compared with" hit it. check_memo.py drops a comma no digit follows;
   check_deck.py is unchanged.
3. **A heading stranded at the foot of page 1** (Northwind's "Questions for management"). Fixed by
   keeping sections together. My first test for it passed on the old code too (52 filler lines pushed
   the heading over either way); 44 lines fail on the old code, checked both ways.
4. **The memo claimed a review nobody gave.** Found by reading Northwind's rebuilt PDF; check_memo.py
   had agreed because it worked the status out the deck's way. Fixed as decision 10, with the check
   now working out the memo's status on its own.
5. **Proof by breaking code.** A script (kept in `output/memo_mutations.py`, git-ignored) copies the
   project to a temporary folder, breaks memo.py one way, runs check_memo.py there, and repeats. The
   real memo.py is never touched. Final result: **9 of 9 breakages caught**, each by the check meant
   to catch it: prior column showing the latest quarter; a number typed into a label; em dashes left in
   everywhere; an em dash left in the flag list only; the workbook gate removed; the PDF dropping the
   bullet lists; the footer losing the review status; a deck-only approval counted for the memo; risks
   shown instead of questions. On the first run the em-dash breakage failed at a different check (the
   data gaps wording), so I added the narrower one that only the em-dash check can see.
6. Some shell commands were refused (a loop, `sips`, a script in /tmp). I ran commands one at a time,
   read the PDFs with the file viewer, and kept the breakage script inside the project's `output/`.

### Unresolved

- **CLAUDE.md's Architecture list doesn't name memo.py or check_memo.py.** The rules say not to edit
  CLAUDE.md unless the task says so, and this one doesn't. Suggested lines: "memo.py: the board memo,
  output/<company>_board_memo.docx and .pdf, same numbers and analysis as the deck; AI text only if
  every number in it is in the metrics workbook" and "check_memo.py: every number in each memo is in
  that company's metrics workbook".
- **The Word file's layout has never been seen.** Nothing here renders .docx. Its text, table cells,
  colors and footer are proven by reading the file back, and the PDF (same blocks, same page size and
  margins) is 1 or 2 pages, but Word's own page count and look are unverified. Open one in Word.
- **No fit guarantee like the deck's.** The PDF is 2 pages for Northwind and Fernhollow with room to
  spare, and check_memo.py fails if any memo reaches 3, but memo.py itself doesn't shrink text. A
  company with many data gaps could reach 3 pages. (Task 21's `--length` may be the place for this.)
- **memo.py's list of workbook numbers is rebuilt from the data, not read from the saved xlsx.** It
  uses the same helpers as excel_output.py, and check_memo.py reads the real file, so a drift would be
  caught by the check, not by memo.py.
- **The flags appear twice** (the table's Status column and the Flags list). Deliberate (the list
  gives each tripped flag's value against its threshold in one line), but worth a reviewer's opinion.
- **Northwind's memo needs a fresh approval** to say "reviewed by" (`python approve.py northwind`,
  then `python memo.py data/northwind.xlsx`). I didn't record one: an approval is a person's act.
- **check_main.py still rewrites output/'s manifests as an "AI skipped" run** (existing behaviour). I
  rebuilt the decks and memos from the saved analyses afterwards, so output/ shows the AI text; the
  manifests' `ai` part says "skipped" until the next real run.
- **The web page (app.py) doesn't offer the memo yet.** Task 2 asks for a "Download memo" button.
