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
