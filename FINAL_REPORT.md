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

---

## Task 2: the web page's home page (app.py, portfolio.py, main.py)

### What I built

- **`app.py` is now two pages.**
  - **Portfolio** (the first page): one row per company in `data/` with its name, latest quarter,
    flags tripped ("6 of 9 flags tripped"), data gaps ("19 metrics/flags (blank: Q1 2025)"), last run
    and deck status, and per row **Generate**, **Download deck**, **Download memo** and **Download
    Excel**. A search box narrows the table; a name with no workbook shows 'No KPI workbook found for
    "Bluefin". Upload one under "Add a company".' **Generate all** runs every company under a progress
    bar and reports each one. **Add a company** takes an .xlsx upload and a company name.
  - **Company** (click the company's name on its row): the flags table with value, threshold and
    status (the reason is in the status: "Cannot evaluate: missing input"), runway at next quarter's
    budget as context, the data gaps, the metrics table in the Excel red / green / gray, both charts
    (the deck's own `charts.py` functions), the AI commentary when present (headline, risks and
    questions, as on slide 4), and the buttons Generate, Download deck, Download memo (PDF and Word),
    Download Excel, and **Approve** with a reviewer-name box.
  - The AI box states the typical cost: "Ask Claude for AI commentary when no saved analysis matches
    (typically about $0.09 and 70 seconds per company)".
- **`portfolio.py`** (new) holds everything behind the buttons, with no Streamlit in it, so it is all
  tested without a browser. It reuses `main.run_company` (Generate: Excel, AI, deck, memo, manifest,
  exactly as `python main.py`), `main.write_summary_csv` (Generate all writes batch_summary.csv),
  `approve.approve` (Approve), `provenance.approval_status` / `deck_status` (the status column),
  `build_deck` text helpers, and `memo.no_em_dash`.
- **`main.py`:** `run_company(..., reuse_saved=True)` uses a saved analysis made from exactly these
  numbers before skipping or calling Claude, and records it in the manifest as "reused a saved analysis
  of exactly these numbers (no API call)" with no tokens and no cost. `reusable_analysis` moved here
  from app.py. The command line never passes `reuse_saved`, so `main.py` behaves as before.
- **Never a traceback, never an invented number:** a workbook clean.py can't read shows clean.py's
  own message (under its row, on its page, from Generate, and from Add a company) and "-" in place of
  every number; a non-Excel file gets the existing plain message; a bug gets "Something unexpected went
  wrong ..." with its traceback in the Terminal window only. Every loop over companies catches per
  company, so one failure never stops the others.
- **No API call anywhere in this task.** Tests use fake clients plus a guard that fails any test that
  creates a real Anthropic client, and every page test uses temporary data/ and output/ folders.
- **Tests:** `tests/test_portfolio.py` (32, new), `tests/test_app.py` rewritten for the two pages (20,
  including Streamlit AppTest runs that click Generate, Generate all, a company name and Approve),
  4 new in `tests/test_main.py`. 609 tests in all.
- **Docs:** README (how the page works, the file list, screenshot list), STUDY_GUIDE (app.py section
  rewritten, new portfolio.py section, main.py rows, test tables, counts), LOOM_SCRIPT and
  INTERVIEW_PREP web page lines, 5 LEARNINGS rows.

### Decisions you didn't specify

1. **Generate writes to `output/`.** The old page built in a temporary folder and never touched
   `output/`. A portfolio whose last run and deck status come from manifests, with an Approve button,
   has to work on the same files as `main.py` and `approve.py`, so it does. The protection moved
   elsewhere (decision 3).
2. **A saved analysis is reused even with the AI box unticked.** It costs nothing and passes every
   check against today's numbers, and without this, Generate would have rebuilt the three decks
   without their validated AI text. Unticked never calls the API; ticked calls Claude only when no
   saved analysis matches. The box was renamed to say exactly that.
3. **Out-of-date files are never offered.** If the workbook or config.yaml has changed since the last
   run (hashes in the manifest), the status reads "Out of date: the workbook has changed since the last
   run. Generate again." and the download buttons and Approve are greyed out. A file with no manifest
   is treated as not generated.
4. **Flags, gaps and latest quarter are worked out from the workbook when the page draws**, not read
   from the last run's summary, so they are never stale. Last run and deck status come from the
   manifest, as you specified.
5. **"Row click" is the company's name as a button.** Streamlit's tables can't hold buttons, so each
   row is a line of columns; clicking the name opens the company page, and "Back to portfolio"
   returns.
6. **Download memo gives the PDF** on the portfolio row (fixed layout, and Word's layout is still
   unseen, Task 1); the company page offers both the PDF and the Word file.
7. **Approve needs a typed name** and files built from today's workbook. approve.py falls back to
   git's user name; a web page shouldn't, because whoever is at the browser must say who they are.
   Approving builds nothing (as approve.py): the message says to click Generate so the footers say
   reviewed.
8. **Add a company:** the name is typed (suggested from the file name), must start with a letter and
   use only letters, digits, spaces and hyphens (at most 40), and becomes `data/<name in lower
   case>.xlsx` ("Blue River" → `data/blue river.xlsx`, shown as "Blue River"). The upload is checked
   in a temporary folder first and saved only if clean.py can read it. An existing company is replaced
   only if "Replace its workbook" is ticked; its old files then show as out of date. Adding doesn't
   generate: the row appears with "Not generated yet".
9. **AI commentary on the company page** shows only when a saved analysis was made from exactly
   today's numbers and still passes, and shows what slide 4 shows (headline, risks, questions; no
   wins). Em dashes in any text, including Claude's, are shown as colons, and `$` is kept as a dollar
   sign (markdown would read two of them as a formula).
10. **The deck status shows the manifest's own words** ("DRAFT - NOT REVIEWED", "approved by Tyler Ho
    on 2026-09-17T22:33:47"), so the page, the manifest and approve.py read the same.

### What failed and how I fixed it (all logged in LEARNINGS.md)

1. **Regenerating would have thrown away Claude's text** (decision 2): main.py had no reuse path.
   Fixed with `reuse_saved` and tests.
2. **Breaking code on purpose showed two weak tests.** A script (`output/task2_mutations.py`,
   git-ignored) copies the project to a temporary folder, breaks one thing there, and runs the tests
   for that file. The first run missed 2 of the first 2 I checked closely: `download()` ignoring "out
   of date" (the test only asked before any file existed) and removing the page's own "files must be
   current" check before Approve (approve.py refuses too, with a message that also says "changed").
   Both tests were strengthened. Final result: **18 of 18 breakages caught**, each by the test
   meant to catch it.
3. **Small slips:** the AppTest element for a chart is "image", not "imgs"; a README line with the
   literal "DRAFT" status tripped the watermark doc test; two test files got a literal em dash (now
   `chr(0x2014)`); the STUDY_GUIDE doc test failed until `reusable_analysis` was listed under main.py.
4. **Refused commands:** starting the Streamlit server, `sed -i`, `ps`, a long heredoc and `grep`
   chains. I rendered the real page (real data/ and output/, no clicks) with AppTest instead: all
   three companies listed, Fernhollow's page with its AI commentary, no error and no em dash.

### Unresolved

- **CLAUDE.md's Architecture line for app.py is now wrong** ("drag in an xlsx ... builds in a temp
  folder, never touches output/"), and portfolio.py isn't listed. I didn't edit CLAUDE.md (the task
  didn't say to). Suggested lines: "app.py: Streamlit web page, two pages. Portfolio: every company in
  data/ (latest quarter, flags, gaps, last run, deck status from manifests), per-row Generate and
  downloads, search, Generate all with progress, Add a company. Company: flags, gaps, colored metrics,
  charts, AI commentary, Generate, downloads, Approve. Writes to output/ like main.py; never offers
  out-of-date files" and "portfolio.py: the work behind app.py's buttons (no Streamlit): rows, search,
  generate, add a company, approve".
- **I haven't seen the page in a browser.** Starting the server was refused. AppTest proves every
  element is drawn with the right text and state, but not the layout: ten columns per row may wrap
  the button labels on a narrow window. Worth a look, and the README screenshots are still to take.
- **Speed at scale:** the table re-reads every workbook each time the page redraws (a click or a
  letter in the search box). Fine for 3 companies (well under a second); for 275 it would want caching.
- **One user at a time.** Two people clicking Generate on the same company at once would write the
  same files. It's a local tool, so I didn't add locking.
- **The live output/ manifests** were not changed by this task (no page button was clicked on the
  real folders), so Task 1's note about check_main.py's "skipped" AI records still stands.

---

## Task 3: styling (theme.py, the template, deck, charts, memo and web page)

### What I built

- **`theme.py`** (new) holds the only palette: navy 0B2545, navy dark 08192F, slate 334155, mid gray
  64748B, line E2E8F0, surface F8FAFC, white; red C0392B on FDE8E6, green 1E8449 on EAF6EF, gray fill
  EDF0F3. Also the font (Arial, then Helvetica, then DejaVu Sans), the sizes (title 28, section 20,
  body 15, caption 13, table 14, floor 12), the brand name, and the web page's style sheet
  (`streamlit_css`) and Streamlit settings (`streamlit_theme`). No other code file types a color;
  a test scans every one.
- **Excel keeps its fills:** `excel_output.STATUS_COLORS` is now `theme.EXCEL_STATUS_COLORS`, with
  the same values as before (FFC7CE / C6EFCE / D9D9D9).
- **`templates/base.pptx` regenerated** with `python make_template.py`: navy 28 pt titles, slate
  15 pt body (13 pt one level down), navy top bar, footer rule in the line color, surface in the
  theme, a 28 pt white cover title. A test fails if the committed file isn't rebuilt.
- **Deck:** title 28, headings and the AI headline 20, lists and AI text 15, the AI-drafted line and
  the placeholder note 13, table 14, footer 12. Navy table header, white and surface stripes, status
  cells in the palette's red, green and gray. The text still shrinks to fit and stops at 12 pt.
- **Charts:** Arial (else Helvetica, else DejaVu Sans) at 13 pt, navy bars and line, slate titles and
  labels, mid gray axes and "data missing".
- **Memo:** the palette, Arial in the Word file and embedded in the PDF (else DejaVu Sans), status
  cells in the palette's fills. Still "Example Capital"-neutral, like the deck.
- **Web page:** theme.py's style sheet on both pages; one column about 1100 px on the surface color;
  each section in a white card with a 1 px border; the flags and metrics tables as HTML with a navy
  header, white text and 40 px rows; the portfolio table with a navy header row and 40 px rows. One
  primary (navy) button per page (Generate all; Generate on a company's page); every other button
  white with navy text and a 1 px navy border; Approve is secondary; no button is red or green.
  `.streamlit/config.toml` gets a `[theme]` section (navy accents, Arial, 15 px, 6 px buttons).
- **Tests:** `tests/test_theme.py` (22, new), plus 6 in test_make_template.py, 2 in test_charts.py,
  2 in test_build_deck.py, 2 in test_memo.py, 5 in test_app.py. **648 tests pass.**
- **Proof the checks work:** `output/task3_mutations.py` broke the code 26 ways in temporary copies
  (a typed color, red primary buttons, a lost border, Excel's fill replaced, deck sizes and colors
  reverted, the chart font dropped, the old template put back, DejaVu in the PDF, no style sheet,
  two primary buttons, unescaped table text, config.toml drifting, a wider page, shorter rows, an 11 pt
  floor, another brand's name, Helvetica dropped from the order). **26 of 26 caught**, each by the
  test meant to catch it.
- **Rebuilt from saved analyses, no API call:** check_excel_output.py, check_deck.py, check_memo.py
  and check_main.py all pass; then the three decks and memos were rebuilt from their saved analyses,
  and all three still carry Claude's text (the bigger sizes still fit slide 4).
- **Docs:** README (the look, theme.py, design decision 14, screenshot list), STUDY_GUIDE (new
  theme.py section and table, every changed row, test table, counts), LOOM_SCRIPT count, 6
  LEARNINGS rows.

### Decisions you didn't specify

1. **Where the new red and green go.** The Excel workbook keeps Excel's fills, as you said. The deck,
   the memo and the web page take the palette's red, green and gray. So the web page's colors are no
   longer the workbook's exactly (they were before); they match the deck and memo instead.
2. **The memo keeps printed-page sizes** (title 18, body 9.5, table 8.5, footer 7). Your 28 / 20 / 15
   scale is a screen and slide scale; on US Letter it would turn a 1 to 2 page memo into about four.
   Colors and Arial do apply.
3. **Which slide text gets which size:** the AI headline and every heading are "section" (20); lists
   and the AI risks and questions are "body" (15); the AI-drafted line and the placeholder's note are
   "caption" (13); the footer line and the brand name stay at the 12 pt floor, because the footer
   already needs every point of width for the file name, commit, model and review status. Chart text
   is all caption size (13). The DRAFT watermark (--draft only) stays 48 pt.
4. **The template's cover title** went from 40 to 28 pt (title size). No deck uses the cover.
5. **The PDF font:** Helvetica on a Mac is a `.ttc` file the PDF library can't embed safely, so the
   memo's PDF goes Arial, then DejaVu Sans. Charts and the web page use the full order.
6. **config.toml repeats five colors and the font.** Streamlit sets its own accents (checkbox tick,
   progress bar, focus ring, links) only from that file; left alone they are Streamlit red. A test
   fails if the file and theme.py ever differ, so theme.py is still the one source.
7. **The check scripts now take their expected colors from theme.py** instead of typing them, and
   `test_theme.py` types every value by hand, so an independent check still exists (LEARNINGS).
8. **HTML tables on the web page.** Streamlit's own table can't make header text white, so a navy
   header would have been dark on navy. The HTML escapes every cell's text.
9. **A "Download" button per row** opens Download deck / memo / Excel. Four buttons plus six text
   columns don't fit 1100 px with 10 by 18 px padding. The company page still shows all four
   download buttons side by side.
10. **Buttons:** hover on a secondary button is the surface color with navy dark text; a greyed-out
    button of either kind is surface with mid gray text and a line-color border. A company's name in
    the table is a navy text link (Streamlit's "tertiary" button).
11. **Cards** are Streamlit containers with a `card-...` key; the style sheet finds them by that key.
    Messages (success green, error red) keep Streamlit's own colors: they are status, not buttons.

### What failed and how I fixed it (all logged in LEARNINGS.md)

1. **Helvetica fallback for the PDF:** the first test expected Helvetica when Arial is missing and got
   DejaVu Sans (`.ttc` file). Split into `first_installed_font` (charts, page) and `font_file` (PDF).
2. **The PDF-font test found a Helvetica** that reportlab names in an empty block on every page. The
   test now reads only fonts that draw text.
3. **"Only theme.py types a color" vs the check scripts' hand-typed colors.** Resolved as in
   decision 7.
4. **Layout:** white header text impossible in Streamlit's table (decision 8); ten columns too wide
   (decision 9).
5. **Small slips:** a STUDY_GUIDE row still named `charts.hex_color` (the doc test caught it); a CSS
   test helper that matched several rules; a size test looking for slide 4's AI-drafted line on a
   deck with no AI text; `cells[-1]` on a python-pptx row.
6. **Refused commands:** heredocs with `{"` or too long, `for` loops, `;` chains, `streamlit config
   show`, Quick Look to picture the deck. Worked around with the Edit tool, scripts in `output/`,
   Python, and single commands.

### Unresolved

- **Nobody has looked at the restyled deck or web page as a picture.** The chart PNGs look right
  (Arial, navy, slate). The deck passes every fit and size check, and AppTest draws both pages with no
  error, one primary button and no em dash, but AppTest can't show layout: whether the cards, the
  navy header row, the 40 px rows and the "Download" popover look right needs a browser. The style
  sheet targets Streamlit's own markers (`stBaseButton-primary`, `stPopoverButton`, `st-key-...`),
  which a future Streamlit version could rename. Worth ten minutes with `run_app.command` and
  PowerPoint before recording the Loom.
- **CLAUDE.md doesn't list theme.py** (I didn't edit it; the task didn't say to). Suggested line:
  "theme.py: the one palette, font (Arial, else Helvetica, else DejaVu Sans) and type sizes; the
  template, deck, charts, memo and web page import it; no other file types a color". Its app.py line
  is also still Task 2's (noted there).
- **The deck's slide titles still use an em dash** (between "key metrics" and "Q2 2026"), as before this task,
  and so do the "Cannot evaluate" status and the "None" data gaps label shared with Excel. I didn't
  change existing deck wording in a styling task; the memo and web page already show colons.
- **text_fit.py still measures with DejaVu Sans**, which is wider than Arial, so text shrinks a
  little earlier than it must. Safe, and unchanged.
- **The manifests' AI records** say "skipped" after check_main.py (as in Task 1): the decks and memos
  were rebuilt from the saved analyses afterwards, but the manifests' `ai` part stays "skipped" until
  the next real run.

---

## Task 4: no em dashes

### What I built

- **Three tests in `tests/test_docs.py`** (written first; all three failed before the fix):
  - `test_no_markdown_file_has_an_em_dash`: every `.md` file in the project and its subfolders
    (not `.venv` or `output/`). A failure lists the file and the start of each line.
  - `test_no_string_in_the_project_code_has_an_em_dash`: every string in every `.py` file in the
    project folder, read with Python's `ast` so f-string pieces and docstrings count and comments
    don't. That covers the files you named (app.py, build_deck.py, memo.py, main.py, analyze.py,
    metrics.py) and every other file whose words reach a person (excel_output.py, clean.py's stop
    messages, the check scripts).
  - `test_the_prompt_the_labels_and_the_flag_statuses_have_no_em_dash`: what the code builds when it
    runs: `analyze.SYSTEM_PROMPT`, every `METRIC_LABELS` and `INPUT_LABELS` value, and the
    "cannot evaluate" status for each of the three reasons. This one catches an em dash put together
    with `chr()`, which the string test can't see.
- **The fixes (19 in code, about 55 in docs):**
  - Flag status: "cannot evaluate: missing input" (metrics.py), and "Cannot evaluate: missing input"
    in Excel and on the deck. Data gaps: "None: every metric and flag has the data it needs".
  - Slide titles: "Northwind: key metrics, Q2 2026 vs Q1 2026", "ARR and cash, Q3 2024 to Q2 2026",
    "Risks and flags, Q2 2026", "AI commentary, Q2 2026".
  - Docs: colons for "label: explanation", full stops for two clauses, commas elsewhere. Every
    `.md` file, including the historical reports.
- **CLAUDE.md Rules** gets: "No em dashes in any .md file or user-facing text: use a comma, colon or
  full stop", naming the test that enforces it.
- **Proof the tests work:** `output/task4_mutations.py` made 14 changes in temporary copies: an em
  dash in README, in CLAUDE.md, in a new `.md` file in `data/`, a slide title, the empty table cell,
  the web page's legend, a memo label, main.py's reuse note, the system prompt (typed, and added with
  `chr()`), a metric label (with `chr()`), the flag status and Excel's None line. **13 of 13 caught**,
  and the control (an em dash in a comment) passes, as it should.
- **Rebuilt, no API call:** all 6 check scripts pass, and the three decks and memos were rebuilt from
  their saved analyses. A scan of the rebuilt decks finds no em dash and AI text on all three.
  **651 tests pass.**

### Decisions you didn't specify

1. **Wider than your list.** You named six files; the test covers every `.py` file in the project
   folder. The Excel workbook's "Cannot evaluate" and "None" lines had em dashes too and reach the
   same reader, and a narrower test would have left them. tests/ is left out, because
   tests/test_memo.py has to hold an em dash to look for one.
2. **Docstrings count, comments don't.** A docstring that quotes an output ("'cannot evaluate:
   <reason>'") should match the output, so it's checked. Nothing a user sees comes from a comment.
3. **Colon for statuses, comma for slide titles.** "Cannot evaluate: missing input" is the wording
   the memo and web page already showed. Slide 1's title already has a colon after the company name,
   so a second colon would read badly; the titles take a comma, all four the same way.
4. **The empty table cell is "-"** (non-flag rows' status, the combo row's values). No comma, colon
   or full stop fits an empty cell. "-" is what the memo already shows there (`memo.NOT_A_FLAG`), so
   the deck and memo now match. I considered a blank cell, but a board reader could take a blank for
   missing data.
5. **Historical reports were changed too** (OVERNIGHT_REPORT, DAY_REPORT, POLISH_REPORT), because the
   rule is "any md file". Where they quote the old wording they now show the new punctuation, and
   POLISH_REPORT's note "The titles use em dashes" now says Task 4 replaced them.
6. **`memo.no_em_dash` stays.** The project's own labels no longer need it, but the web page runs
   Claude's headline, risks and questions and other libraries' error messages through it, and those
   can still contain an em dash. Its docstring says so now.
7. **`check_memo.py` got stricter:** it compared the memo's status with the Excel status after
   swapping the em dash for a colon. Now both use the same words, so it compares them exactly.
8. **config.yaml keeps one em dash** in a comment (`rule_of_40_min`). You said not to edit it, and a
   comment isn't shown to anyone but a reader of the file.

### What failed and how I fixed it (all logged in LEARNINGS.md)

1. **My own docstring broke the rule.** The first docstring for `memo.no_em_dash` showed an em dash
   as its example. Caught on reread; it now shows the result instead.
2. **4 tests pinned the old wording** (test_analyze.py, test_excel_output.py, test_build_deck.py).
   Updated to the new words; test_build_deck.py now imports `NOT_APPLICABLE` rather than typing it.
3. **Fernhollow's saved analysis is no longer reused by the web page** (see Unresolved).
4. **Small slips:** one breakage's anchor text appeared 3 times in main.py (fixed to a unique line);
   `sed`, `awk` and chained shell commands were refused, so edits ran as small scripts in `output/`
   (`task4_code_edit.py`, `task4_tests_edit.py`, `task4_docs_edit.py`, `task4_checks.py`).

### Unresolved

- **The web page won't reuse Fernhollow's saved AI commentary.** Claude was sent "cannot evaluate"
  + em dash + "missing input" for Rule of 40; today's facts say "cannot evaluate: missing input", so
  `main.reusable_analysis` (rightly) sees different facts. Northwind and Alderpeak still match (no
  "cannot evaluate" flag in their latest quarter). The CLI decks and memos are unaffected: they show
  Fernhollow's AI text. Two ways out: one paid Fernhollow run (about $0.09), or editing that one
  string in `output/fernhollow_analysis.json`. I didn't do the second: that file is the record of
  what Claude was sent, and rewriting it would make the record untrue. Your call.
- **Claude's own text can contain an em dash.** None of the three saved analyses does, and the web
  page converts any it finds, but the deck and memo show Claude's words as written. A line in the
  system prompt ("no em dashes") would need a new prompt version (v5) and a paid run to test, so I
  left the prompt alone.
- **Task 3's Unresolved note about em dashes in the slide titles and the shared "Cannot evaluate"
  and "None" labels is now done.**


---

## Task 5: column mapping (mapping.py, clean.py, the web page's Review mapping step)

### What I built

- **`mapping.py` (new).** For every header that is neither a standard column nor in
  `clean.HEADER_ALIASES`, it proposes the standard column the header most likely means, with a
  confidence (shown as "85% (high)", "64% (medium)", "41% (low)"), a reason in plain words, the
  first 4 values under the header as written, and the list of columns it could mean. Heuristics only:
  1. **Candidates:** only the columns no known header already has.
  2. **Name:** the header's words, with filler dropped ("Total", "$K") and the usual FP&A synonyms
     swapped ("Opening" to starting, "Plan" to budget, "GP" to gross profit, "FTEs" to headcount),
     against each column's own words and its HEADER_ALIASES spellings. The score is the better of word
     overlap and letter-by-letter similarity (Python's `difflib`), so the typo "Revenu" scores 92%.
  3. **Values:** a value in the budget-only row rules out all 13 actual columns (clean.py would stop
     on them); ARR and cash must roll forward with the column in place; gross profit can't be above
     revenue; the only column left for the only unknown header gets points. A check that fits adds
     points, one that breaks takes them off.
  4. **Assignment:** strongest (header, column) pair first, no column proposed twice; below 40%, no
     proposal and the person chooses.
- **Nothing is guessed silently.** `clean.clean_workbook` now gathers every unknown header and stops
  with `UnconfirmedMappingError`: one line per header naming it, its cell, the proposal, the
  confidence and the reason, then how to confirm. This happens however high the confidence is: a 99%
  proposal stops too, and a test pins that.
- **Confirmed mappings are saved to `mappings/<company>.yaml`** (the header as written, the column,
  the company and the time, and a comment saying what the file is). `clean_workbook` reads the
  company's file on every run, so main.py, the deck, the memo, the Excel workbook and the web page
  all use a confirmed mapping with no change of their own, and next quarter's workbook with the same
  headers runs unattended.
- **Two places to confirm:**
  - the web page's **Review mapping** step, on a company's page and in Add a company: each pair
    shows the header and cell, the proposal and its confidence, the sample values, **Change** (a
    list of the columns it could mean, the proposal preselected) and **Confirm**, with the reason
    underneath. Save mapping (or Add company) stays greyed out until every pair is confirmed, and
    changing a column clears that pair's tick;
  - `python mapping.py data/acme.xlsx` lists the proposals; `--confirm` asks about each one (Enter
    accepts, or type another column; a typo is asked again; `q` stops and saves nothing).
- **A mapping is saved only if the workbook reads with it** (tried first on a temporary mappings
  folder). An upload is added to `data/` together with its mapping, or neither is saved.
- **The mapping is an input like the workbook.** The manifest records the mapping file and its
  SHA-256 hash; a changed mapping makes the portfolio row "Out of date: the column mapping has
  changed", takes the old files off the download buttons, stops `approve.py`, and sends an approved
  deck back to "not reviewed" (`provenance.approval_status` takes a fourth hash).
- **Tests (77 new, 728 in all, written before the code):** `tests/test_mapping.py` (55) uses copies
  of the three companies with their headers renamed to words clean.py doesn't know: all 16 of
  Northwind's ("Opening ARR", "Upsell ARR", "Downgrades", "GP", "Cash Burn", "Plan Burn" ...), 6 of
  Alderpeak's ("BoP ARR", "Churn", "ARR Target" ...) and 7 of Fernhollow's ("Expansion", "Gross
  Margin $", "Burn", "Pipe ($K)" ...). It proves:
  - **proposal:** every one of the 29 renamed headers is proposed as the column it came from;
  - **required confirmation:** each copy stops naming every header and its proposal; nothing is
    saved on the way; a partial confirmation still stops;
  - **identical metrics:** after confirming, each copy's cleaned numbers, budget row, metrics and
    flags equal the original's exactly; next quarter's workbook runs unasked; Change saves the
    person's column, not the proposal.
  Plus 12 in test_portfolio.py, 5 in test_app.py (AppTest clicks Confirm, Change and Save mapping),
  3 in test_provenance.py and 2 in test_approve.py.
- **Proof the tests work:** `output/task5_mutations.py` made 22 breakages in temporary copies of the
  project, and every one was caught (**22 of 22**; the control, a comment edit, passes). They cover
  clean.py using the proposals silently, or only the high-confidence ones; the stop message without
  the proposal; a saved line overriding a known header; a second confirmation forgetting the first;
  saved headers not normalized; a bad column accepted; the budget row not ruling anything out; one
  column proposed twice; no synonyms; a broken roll-forward counted as support; filler kept; the
  approval, approve.py, the manifest and the portfolio row each ignoring the mapping; a mapping saved
  before it's tried; an unchosen header let through; the page showing the command-line message;
  Change not clearing Confirm; Save mapping enabled early; the review step not drawn.
- **Rebuilt, no API call:** all 6 check scripts pass, and the three decks and memos were rebuilt
  from their saved analyses (slide 4 carries AI text on all three).
- **Docs:** README (commands, the page, data flow, design decision 15, a known limitation, next
  steps), STUDY_GUIDE (a mapping.py section with every function, the new clean.py, portfolio.py,
  app.py and provenance.py rows, test counts, `mappings/`), INTERVIEW_PREP (new Q34b, Q35 and Q39
  updated), LOOM_SCRIPT (test count). CLAUDE.md is not edited (see Unresolved).

### Where a model call would improve it

The heuristics are word lists and arithmetic. A Claude call would help in four places, and each
would still end in a person's confirmation:
1. **A name that fits two columns.** "Cash Burn" is 59% like net burn and 50% like ending cash by
   name; it is proposed correctly only because "Closing Cash Balance" took ending cash first
   (64%, medium). A model reads "burn" as a flow, not a balance.
2. **Words not in the synonym list.** "Bookings", "Logos Lost", "Net Retention $", "Opex - S&M",
   another language: each needs a line added to `SYNONYMS` today. A model knows the vocabulary
   without the list.
3. **Meaning in the values, not the name.** A model could look at the sample values and say "these
   are percentages, so 'Gross Margin' here is the ratio, not gross profit dollars", which the
   gross-profit-below-revenue check only half covers.
4. **The reason text.** A model can explain a proposal the way an analyst would ("the cash balance
   falls by exactly this amount every quarter"), where today's reasons list the checks that passed.

The design for it: send the unknown headers, their sample values and the 16 column definitions from
CLAUDE.md; ask for JSON (header, column, reason); then run the same value checks in Python, so a
proposal that breaks the budget-row rule or a roll-forward is marked down whatever the model says.
Python keeps the final say on anything checkable, and a person on everything else, as with the
commentary.

### Decisions you didn't specify

1. **The hook is inside `clean.clean_workbook`**, so every file that calls it (main.py, the deck,
   the memo, the Excel workbook, the web page, the check scripts) picks up confirmed mappings with no
   change. `clean.py` imports `mapping.py` inside that one function, not at the top,
   because `mapping.py` imports `clean.py` (the comment says so).
2. **Confidence informs, it never decides.** There is no level at which a proposal is used without a
   person. The levels (high from 80%, medium from 55%, low below) and the 40% floor for making a
   proposal at all are my choices; the scores are not probabilities.
3. **Only columns still without a header are candidates, and a column the budget row rules out is
   not offered under Change.** Mapping onto a column another header already has would stop anyway
   ("both mean ..."), and so would an actual column with a value in the budget-only row.
4. **All unknown headers stop at once.** clean.py used to stop at the first one. With a proposal for
   each, one review covers them all.
5. **A known header always wins over a saved line.** A hand-edited `mappings/<company>.yaml` can't
   redefine "Revenue".
6. **The file is keyed by the workbook's name** (`data/acme.xlsx` gives `mappings/acme.yaml`), holds
   headers as written, and matches them as clean.py matches HEADER_ALIASES (case, spaces and symbols
   ignored). Later confirmations are added to earlier ones. `mappings/` is not git-ignored: a
   confirmed mapping is a decision like config.yaml, worth keeping in history. It's empty today.
7. **No "not an input, ignore it" choice.** An extra column still has to be deleted from the
   workbook, as before: an ignored column would be data dropped without a trace in the numbers.
8. **The mapping's hash voids an approval**, as the workbook's and config.yaml's do. An approval
   from before this task has no mapping hash and still holds while the company has no mapping file,
   so Northwind's existing approval is unaffected.
9. **Confirm is a checkbox and Change a dropdown**, with one Save mapping button (secondary: the page
   keeps its one navy button). The checkbox's key includes the chosen column, which is what makes a
   change clear the tick.
10. **The page shows standard names** ("starting_arr"), the names in CLAUDE.md's input list and in
    clean.py's messages. Display names would mean adding 14 input labels to metrics.py's one label
    set for this screen alone; I left that for you to decide.
11. **Three different rename sets**, so the proposals can't pass by fitting one list, and Northwind
    renamed completely (every value check that needs a known neighbour is then unavailable, the
    hardest case).
12. **`--confirm` is all or nothing**: stopping halfway saves nothing, so a file never holds half a
    review.

### What failed and how I fixed it (all logged in LEARNINGS.md)

1. **NaN in a flag made identical flags compare unequal** (my test): compared as tables instead.
2. **"headcount headcount" in a reason**: an alias plus its synonym gave the same word twice; found
   by reading the printed proposals, fixed in `header_words`.
3. **Change listed columns clean.py would stop on**: removed, two tests updated, one added.
4. **5 existing tests failed** on the new stop: a pinned hint, and two fixtures whose "ARR" header
   now gets a proposal instead of the missing-columns stop they were written for. Updated.
5. **The web page would have shown a temporary file's path** in the command-line hint: the page has
   its own wording now.
6. **Two value checks aren't needed on these three companies.** Breaking the roll-forward and the
   filler list failed only their unit tests: every renamed header's name already decides its column
   once the others are taken. The checks are proved to work, not proved to matter here.
7. **Refused commands and one slip of mine** (a `git add` pathspec error skipped a commit; an edit
   script stopped halfway): scripts in `output/`, finished and checked before committing.

### Unresolved

- **CLAUDE.md doesn't list `mapping.py` or `mappings/`.** You said not to edit it. Suggested line
  for Architecture: "mapping.py: a header clean.py doesn't know gets a proposed column (name +
  values) with a confidence and reason; nothing runs until a person confirms it (web page Review
  mapping, or `python mapping.py <workbook> --confirm`); saved to mappings/<company>.yaml".
- **Not looked at in a browser.** The Review mapping step is tested with AppTest (it draws, Change
  and Confirm work, Save stays greyed out), but no one has seen it. Four columns per pair in a
  1100 px column should fit; worth a look.
- **A renamed "Quarter" header still stops without a proposal.** clean.py finds the KPI tab by that
  header, so "Period" means "no tab has a 'Quarter' header". Other layouts (quarters across the top)
  are out of scope too.
- **Renaming a company's workbook loses its mapping** (the file follows the workbook's name). The
  review then simply asks again; nothing wrong is used.
- **The synonym list is hand-written** from ordinary FP&A vocabulary. It will need lines for real
  companies' spellings until a model call replaces it (above).

---

## Task 6: the evaluation set (eval/make_eval_data.py, eval/run_eval.py)

### What I built

- **`eval/make_eval_data.py`** writes 12 fictional companies to `eval/data/`, each a messy workbook
  (the demo companies' header spellings, title rows, a Notes tab in front, money as "$1.2M", "850K",
  "5,090", and now negative and zero text: "$-0.2M", "-100K", "$0M", "0"), each with an answer key
  typed by hand:

  | Company | Case | What the answer key says |
  |---|---|---|
  | Larkspur | healthy | 0 of 9 trip, in every quarter; all 19 metrics checked by hand formula |
  | Quillmoor | distressed | 9 of 9 trip: burn multiple 28x (finite, unlike Fernhollow's ∞), and NRR falling while pipeline rises, so the combo trips too |
  | Tidewell | exactly at every threshold | NRR 100%, GRR 85%, burn multiple 2.0x, 15% over budget, runway 12.0 mo, CAC 24.0 mo, net new ARR -20% vs budget, Rule of 40 at 40: all 8 pass. NRR goes 102%, 101%, 100%, a drop of exactly 1 point each quarter, so the combo trips |
  | Brackenfield | blank quarter first (Q3 2024) | its own YoY is missing input, not no prior period; Q4 2024's QoQ and Q3 2025's YoY are gaps; the latest quarter is untouched |
  | Copperlane | blank second to last (Q1 2026) | latest QoQ metrics and net new ARR vs budget: missing; that flag and the combo: cannot evaluate, missing input |
  | Duskhaven | blank last (Q2 2026) | all 9 flags cannot evaluate: missing input; runway at next quarter's budget has no number either |
  | Emberfall | zero revenue (a year before launch) | 0 / 0 margins, NRR on no starting ARR, ARR vs a budget of 0, and growth from a zero base: 49 values "not meaningful", every one listed by quarter; burn multiple ∞ while burning with no ARR; Rule of 40 cannot evaluate: not meaningful |
  | Glenmarsh | negative budget | budgeted burn of 0, -100, -200 and a budget that plans ARR to shrink: both vs-budget flags cannot evaluate: not meaningful; next quarter's budgeted burn of -150 gives runway ∞ |
  | Hollowmere | NRR and pipeline both falling | the combo passes; NRR and Rule of 40 trip |
  | Ivywick | short history (2 quarters) | Rule of 40 and the combo: cannot evaluate, no prior period; **no data gaps** |
  | Kestrelwood | Q4 2025's row pasted twice | clean.py stops: "row 8: quarter 'Q4 2025' appears twice (also in row 7)" |
  | Lanternreach | Q1 2026's row missing | clean.py stops: "row 10: After Q4 2025 expected Q1 2026, found Q2 2026" |

- **`eval/run_eval.py`** runs each company through clean, metrics, flags and gaps and prints a
  scorecard: one row per company (Clean, Metrics, Flags, Gaps: ok, stopped, or MISMATCH (n)), then
  every mismatch by name ("Tidewell, Flags: Rule of 40: expected pass, got trip"), then "12 of 12
  companies match their answer keys". It exits 1 on any mismatch. Four checks:
  - **Clean:** every value back exactly, the blank quarter all empty, the budget-only row; for the two
    stop companies, the stop and its words (sheet, row and quarter).
  - **Metrics:** the latest quarter's values vs the hand formulas, runway at next quarter's budget,
    and **the reason for every missing value in every quarter** (19 metrics by 8 quarters each).
  - **Flags:** every status with its reason; Larkspur trips nothing in any quarter.
  - **Gaps:** `data_gaps()` equals the prediction, no more and no fewer.
- **It passes:** 12 of 12. It's in the checks: `tests/test_eval.py` runs it inside
  `python -m pytest`, README's "Prove it works" block and the study guide list it (a doc test fails
  if they stop), and the no-em-dash test now covers `eval/`.
- **Tests (36 new, 764 in all, written before the code):** 35 in `tests/test_eval.py`. They check
  that the set is what was asked for, that every answer key ties out, that Tidewell's formulas give
  each config.yaml threshold exactly, and that the saved workbooks are what `eval/make_eval_data.py` writes.
  Then 12 of 12 match. **Eleven tests break a copy of an answer key and require the
  scorecard to name it** (a wrong flag, reason, value, cleaned number, runway, gap, an unlisted
  not-meaningful cell, a trip in a never-trips company, a workbook that should stop but reads, a stop
  with the wrong words, a missing workbook). Plus 1 in `tests/test_docs.py`.
- **Proof the eval catches what it claims:** `output/task6_mutations.py` planted 24 bugs, one at a
  time, in temporary copies of the project, and ran `eval/run_eval.py` in each. **24 of 24** failed,
  each on the company built for that case; the control (a comment edit) passed 12 of 12. The bugs:
  exactly-at trips; no float rounding (Rule of 40 0.39999999999999997 < 0.40); a 1-point drop not
  counting; the combo ignoring pipeline; no prior period winning over a blank; a blank's look-back
  ignored; the combo's short-history check removed; no prior period counted as a gap; a flag losing
  its reason; flags on the wrong quarter; growth from zero shown as ∞; zero revenue giving an FCF
  margin of 0; burn vs a negative budget computed; net new ARR vs a shrinking plan computed; negative
  budgeted burn giving negative runway; YoY looking 3 back; ARR vs last quarter's budget; GRR not
  annualized; a repeated row, a missing row, "$-0.2M" unreadable, a blank read as 0, "$0M" read as
  blank.
- **Docs:** README (the command, design decision 16, next steps), STUDY_GUIDE (an eval section with
  every function, the test table, counts), INTERVIEW_PREP (Q34c), LOOM_SCRIPT (count, a proof row).

### How the same harness would score AI commentary offline

The eval scores the numbers. The commentary could be scored the same way, with no API call, by
replaying saved analyses (fixtures) through the checks. Nothing below is built.

1. **Fixtures.** One file per eval company in `eval/fixtures/<company>_analysis.json`, in the shape
   `analyze.save_analysis` already writes: the payload Claude saw, its answer, the model and
   `PROMPT_VERSION`. They come from one recorded live run (about $0.09 a company, so about $0.91 for
   the 10 that read; your call, since this task allowed no API calls) or are written by hand. Each
   also gets **planted-error copies**, the same idea as the 24 bugs: an invented number, a dropped
   minus sign, "fell" on a series that rose, "persistent" on one that zigzags, a tripped flag left out
   of the risks, and a number quoted for a metric that has none (Duskhaven's Rule of 40).
2. **Stale check first.** Rebuild the payload from the workbook with `analyze.build_payload` and
   compare it with the fixture's, as `main.reusable_analysis` does. Different numbers mean the fixture
   is stale and gets reported as stale. The same goes for a fixture recorded under an older
   `PROMPT_VERSION`: "re-record", never a pass.
3. **The checks that already exist, offline:** `analyze.validate_summary` (schema, every number in
   the payload with its sign, slide fit, direction words), `build_deck.load_analysis` (what the deck
   accepts) and the memo's stricter rule (`memo.unlisted_numbers`: only numbers the metrics workbook
   shows).
4. **New checks built from the answer keys that already exist:**
   - every flag the key expects to trip is named among the risks, by its label, since flag names are
     the metric labels;
   - no number is quoted for a metric whose latest value has none (`predicted_reason` already knows
     which, and why);
   - a company with predicted gaps says "data missing" somewhere, and quotes no value for a blank quarter;
   - Tidewell's "exactly at" values aren't called "below" or "above" their threshold;
   - no flag the key expects to pass is presented as a risk that tripped.
5. **An AI column on the same scorecard**, with mismatches by name ("Quillmoor, AI: CAC payback
   tripped but is not among the risks"). Each planted-error copy must fail the rule it was built
   for, and the clean fixture must pass: a mutation test for the text checks.
6. **The retry path with a fake client.** `tests/test_main.py`'s `FakeClient` pattern, fed fixture
   pairs: a bad answer then a good one must give one retry and the good text on the deck. Bad then
   bad must give the placeholder and "OK (AI failed)".
7. **What stays human:** whether the commentary is insightful. The 1 to 5 quality score stays a
   blind human score, as in `compare_models.py`, stored beside each fixture. Code can prove the
   commentary is grounded and complete, not that it is good.

### Decisions you didn't specify

1. **The twelve.** Your list has 11 items if "a blank quarter in each position" is one and
   "duplicate or missing rows" is one. I read "each position" as the positions that behave
   differently against the latest quarter: first (its own YoY), second to last (the latest QoQ and
   the combo window), last (every flag). Northwind (position 3) and Fernhollow (position 4) already
   cover the middle. I split "duplicate or missing" into one company each, because they stop in
   different places in clean.py. That makes 12.
2. **The blank-quarter and stop companies use Larkspur's numbers**, so the only thing that differs
   is the one each tests, and the 19 hand formulas are shared.
3. **The answer key is typed, the rules are applied.** Values, flag statuses and not-meaningful
   cells are typed by hand. "Missing input", "no prior period" and the gaps are predicted from the
   blank quarter's position, using the QoQ and YoY lists in check_northwind.py (typed by hand), not
   metrics.py's `METRIC_INPUTS`, so the prediction doesn't share code with what it checks.
4. **The not-meaningful list is complete, not sampled.** Any value with no number and no listed
   reason is a mismatch, so a new n/m anywhere shows up.
5. **Only Larkspur's numbers are checked for all 19 metrics** (Larkspur and the three blank-quarter
   companies built on it); the other six that read check the 8 flag metrics. After the first
   mutation round showed the gap (below), I covered every formula once rather than typing 19
   formulas for every company.
6. **No "one step past each threshold" company.** Quillmoor trips all nine, `test_metrics.py` covers
   real misses, and a thirteenth company wasn't asked for.
7. **The workbooks are committed** in `eval/data/`, like `data/`, so the eval needs no generating
   step. A test regenerates them into a temporary folder and compares every cell, so they can't go
   stale quietly.
8. **The scorecard collects every mismatch** instead of stopping at the first like the check
   scripts, because a scorecard is for seeing how many cases a change broke.
9. **`eval/` is a folder, not a package.** pytest.ini adds it to the import path; the two scripts
   add the project folder to theirs, so `python eval/run_eval.py` works from the project folder.
10. **The eval companies are not in `data/`,** so `main.py --all` and the web page don't list them.
11. **The mess reuses the demo companies' header spellings**, so no eval workbook stops for a column
    mapping (test_mapping.py covers that).

### What failed and how I fixed it (all logged in LEARNINGS.md)

1. **The eval missed a planted bug.** Revenue YoY computed 3 quarters back passed 11 of 12, because
   only the flag metrics were value-checked. Larkspur now checks all 19 (a test pins it). The bug
   then failed 4 companies, and a new ARR-vs-budget bug failed 10.
2. **One of my planted bugs changed nothing** (a "no prior period first" edit placed too late in
   the loop to matter). I replaced it with the real bug, which Brackenfield caught. A surviving bug
   isn't a hole until it's shown to change a result.
3. **Designing exact numbers took several passes:** the healthy Rule of 40 at 39.3%, the threshold
   company's ARR chain off, the distressed burn multiple infinite, an accidental 2.0x elsewhere. A
   scratch script in `output/` tried the designs; the answer key was then typed by hand.
4. **Refused commands,** including a refused edit that left the design file unchanged, so the next
   run printed old numbers, and a long heredoc for this report. The doc test caught two file names
   written without "eval/".

### Unresolved

- **CLAUDE.md doesn't list `eval/`.** You said not to edit it. Suggested Architecture line:
  "eval/make_eval_data.py, eval/run_eval.py: 12 edge-case companies (thresholds, blank positions,
  zero revenue, negative budget, short history, rows that must stop) with hand-typed answer keys;
  run_eval.py scores clean, metrics, flags and gaps and names every mismatch (run by pytest)".
- **Not covered by the eval** (covered by unit tests instead): a partly blank quarter, a workbook
  with no budget-only row, a company ending on a different quarter, unknown headers, two KPI tabs,
  footnote rows. Each would be one more company if you want them in the scorecard.
- **Two guards the eval can't tell apart.** A budgeted burn of exactly 0 is caught twice: by the
  "budget 0 or less" rule and by the "no infinity except the edge cases" rule. Removing either one
  leaves that case's output unchanged, so no eval or user can see which one did the work. I worked this out from the code and didn't
  plant it.
- **The AI scoring is a design only** (above). Recording real fixtures needs one API run, which this
  task didn't allow.
- **Not rerun:** `check_deck.py`, `check_main.py`, `check_memo.py`, `check_excel_output.py`. No code
  they use changed (only `eval/`, tests, `pytest.ini` and docs), and `check_main.py` would put the AI
  placeholder on the real decks. `check_northwind.py`, `check_companies.py`, the eval and all 764
  tests pass.

---

## Task 7: batch resilience (main.py, resilience.py)

This section was written at the start of Task 8. The Task 7 session ended after its code commit
(078d9a7) with its last tests and docs uncommitted and this section unwritten. Task 8 checked that
work, reran the proofs, closed one hole they found, and committed it.

### What I built

- **`main.py` options:** `--resume` (skip a company whose outputs are up to date), `--max-cost`
  (start no more companies once the AI spend reaches the ceiling), `--timeout` (give up on a
  company after that many seconds), `--workers` (companies side by side, default 1).
- **`resilience.py`:** rate-limit retries (`RateLimitRetry`: what the API's retry-after header asks,
  else 5, 10, 20, 40 s, never more than 60 s at once, then give up), the `--resume` check
  (`resume_problem`: workbook, config.yaml and mapping hashes, every file present, the same AI and
  `--draft` choice, no approval since), the private folder per company and the move into `output/`
  (`commit_stage`, manifest last), and the batch's cost meter.
- **All or nothing per company:** each company is built in `output/.staging/<company>_xxxx` and moved
  into `output/` only on success. A failure or timeout leaves last quarter's files as they were.
- **Every skip, timeout, stop and failure is recorded** in the company's manifest (`batch_events`,
  the last 100 kept), the summary table (a Notes column; the CSV also gains the AI cost) and
  `output/batch_manifest.json`.
- **Charts drawn without pyplot** (`Figure()` directly), so workers can draw at the same time.
- **Tests: 51 new, 815 in all,** in `tests/test_batch.py`, with fake clients (scripted, hanging,
  counting calls in flight) and a guard that fails any test creating a real client.
- **Proof:** `output/task7_mutations.py` planted 35 bugs one at a time in temporary copies. Rerun in
  full in Task 8: **34 of 35 caught** after one fix (below), control passes. All 7 check scripts and
  the eval pass in a temporary copy (`output/task7_run_checks.py`).

### Decisions you didn't specify

1. **Only rate limits are retried.** Any other API error would fail the same way again, so that
   company's deck gets the placeholder at once, like any AI failure.
2. **`--max-cost` is checked before each company starts**, so the batch can pass the ceiling by what
   the companies already running spend, never more. A spend exactly at the ceiling stops it.
3. **Failed attempts count toward the spend**: Claude charged for them.
4. **An up-to-date company is skipped, not stopped,** even past the ceiling: it costs nothing.
5. **A `--resume` rebuild reuses the saved analysis** when the numbers didn't change (free).
6. **A timeout or a stop makes the exit code 1**, like a failure: the batch didn't produce everything.
7. **A failed company keeps its old analysis** (it used to be deleted). The old rule protected
   against an old analysis beside a new deck, which the private folder now rules out.
8. **The manifest is moved last**, so a run killed during the move leaves a manifest that describes
   the older files, and `--resume` rebuilds that company.
9. **Default 1 worker**, so a plain run behaves as before.

### What failed and how I fixed it (all logged in LEARNINGS.md)

1. **Four existing tests broke** on the new CSV columns, the manifest's `draft` and the kept
   analysis; each was updated and the reason written in the test.
2. **A test wrote `output/batch_manifest.json` into the real output folder**; tests that call
   `main()` now replace that writer too.
3. **A fake answer with 0 output tokens had no cost**, because 0 tokens means "not counted". The fake
   now has 1.
4. **A planted bug was caught for the wrong reason** (a bare `AnthropicError` has no response); the
   test now raises a real HTTP 500 error as well.
5. **Found in Task 8's rerun:** "a timed-out company reports no error" survived: the exit-code test
   used a hand-typed result, so nothing checked the real one. The timeout test now checks the error
   text; the bug is caught.

### Unresolved

- **One planted bug survives, and can't be caught from outside:** "a company that finishes after its
  timeout moves its files into output/". The company's own thread already throws its private folder
  away when it was given up on, before the batch sees it, so the batch's second check has nothing
  left to move. Two guards for one case, like Task 6's budget guards.
- **`--workers` hasn't been run against the real API,** so how many workers a rate limit allows is
  unmeasured. A timeout stops waiting, not the work (README limitations).
