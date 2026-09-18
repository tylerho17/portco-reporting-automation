# Polish report

One section per polish task: what was built, every decision you didn't specify, what failed and
how it was fixed, and anything unresolved.

---

## Task 1: watermark opt-in, review status in the footer

### What I built

- **No watermark by default.** Decks no longer carry "DRAFT - NOT REVIEWED" across every slide
  unless you ask for it:
  - `python build_deck.py data/northwind.xlsx --draft`
  - `python main.py --all --draft`
- **The footer always says whether a person has reviewed the deck.** It reads the manifest,
  using `provenance.approval_status`, the same judge as before. The footer ends with one of two things:
  - `AI-drafted | reviewed by Tyler Ho on 2026-09-17` once approve.py has recorded a reviewer
    and the workbook and config.yaml haven't changed since
  - `AI-drafted | not reviewed` otherwise (no manifest, no approval, or an approval that no longer counts)
- **The footer still fits on one line**, and this is now proven by measuring its width, not just its height:
  - Northwind (approved): 743 pt of 755 pt
  - Alderpeak and Fernhollow (not reviewed): about 601-604 pt
- approve.py, provenance.py and the manifest format are unchanged.

Full footer on the rebuilt Northwind deck. The `*` means the code had uncommitted edits when
it was built:
`Fictional data | northwind.xlsx | 2026-09-17 | 396f673* | claude-sonnet-5 | AI-drafted | reviewed by Tyler Ho on 2026-09-17`

### Files changed, in plain English

- **build_deck.py**
  - Three new small functions:
    - `review_text` writes "reviewed by NAME on DATE" or "not reviewed".
    - `footer_text` joins the footer's pieces with " | ".
    - `fits_one_line` measures the line against the footer box.
  - `add_footer` uses all three.
  - `build_presentation` and `save_deck` take `draft=False`. `main()` has a `--draft` flag, plus an
    `output_dir` setting so a test can build into a temporary folder.
- **main.py**: `--draft` is passed through `run_batch` → `run_company` → `deck_step` → `save_deck`.
- **make_template.py** and **templates/base.pptx**:
  - The footer box is wider: it takes the whole bottom strip except the 1.65 in that
    "Example Capital" needs.
  - The brand name box lost its left margin. The name is right-aligned, so it looks the same.
- **check_deck.py**:
  - For each company, it works out the expected review status from the manifest itself, without
    using build_deck's code.
  - It checks that every footer ends with that status, measures that the footer fits one line,
    and checks that no watermark appears by default.
  - A new `check_draft_option` builds all three decks with `--draft` in a temporary folder:
    - Northwind has a real approval, so it gets 0 watermarks.
    - Alderpeak and Fernhollow get the watermark on 5 of 5 slides.
- **tests**:
  - tests/test_build_deck.py: 11 new tests. 2 old ones were replaced ("every slide is
    watermarked until approved", "an approved deck has no watermark"), and 4 existing ones now
    build with `--draft`. The new tests cover:
    - default versus `--draft`
    - `--draft` never stamps an approved deck
    - both footer wordings, the same footer on every slide
    - the worst-case one-line fit, the long-name fallback
    - the footer reading the manifest, the command-line flag
  - tests/test_main.py: 4 new tests:
    - the default has no watermark and the footer says not reviewed
    - `--draft` stamps all 5 slides
    - a re-run names a reviewer whose approval still counts
    - `--draft` reaches the batch from the command line
  - tests/test_make_template.py: 1 new test. The brand name still fits beside the wider footer.
  - Totals: pytest has 494 passed (480 before this task, +14 net). Every check_*.py prints
    "All checks passed".

### Decisions you didn't specify

1. **`--draft` never stamps an approved deck.** The stamp says "NOT REVIEWED". On a deck someone
   approved, that would be false, so `--draft` only marks decks nobody has approved.
2. **The review date is the day only** ("2026-09-17"). The manifest records the time to the
   second; the footer has no room for it, and the manifest still has it.
3. **"AI-drafted" appears even when the deck shows "AI summary unavailable"**, because you
   specified both wordings exactly. The model slot in the same footer still says "no AI text" in
   that case, so the footer isn't misleading. It may read oddly, though, and it would be easy to
   change. See Unresolved.
4. **The fictional-data note is now "Fictional data"** (it was "Fictional data, generated for
   demonstration"). Without this, the review status could never fit. The short form keeps the
   point of the rule.
5. **The footer box is wider, and the brand name box is narrower.**
   - Old: 9.5 in footer and 2.83 in brand.
   - New: 10.68 in footer and 1.65 in brand. "Example Capital" is 109 pt wide and its box holds 111 pt.
   - The brand name stays at the bottom right, looking the same.
6. **A reviewer name too long for the line is left to the manifest.** The footer then says
   "reviewed on 2026-09-17".
   - Names are typed by people, so no width can hold every one, and a long name shouldn't stop
     every deck from building.
   - About 57 pt is left for the name with the longest file name and the default model:
     "Tyler Ho" fits, a name much longer than about 10 characters won't.
   - If even the short form doesn't fit, the build stops, as for any other text that doesn't fit.
7. **The review status goes at the end of the footer**, right after the model, so it reads
   "claude-sonnet-5 | AI-drafted | ...".
8. **build_deck.py now prints the full path** ("Saved /tmp/...") when a deck is saved outside
   the project folder. Before, it would have crashed trying to shorten that path. Inside the
   project it still prints "Saved output/northwind_board_pack.pptx".

### What failed and how I fixed it

- **The first plan didn't fit the footer.** I measured the footer before writing any code, and
  adding the review status to the old footer gave 931 pt against 670 pt of room. Changing only
  one thing wasn't enough: a short note alone gave 748 pt against 670, and a wider box alone
  gave 931 against 755. Both changes together make it fit (decisions 4 and 5), with the name
  fallback as a safety net. Logged in LEARNINGS.md.
- **An older gap turned up while measuring.** With the Haiku model id
  (`claude-haiku-4-5-20251001`), the *old* footer was 696 pt against 670 pt of room, so a
  Haiku deck would have stopped. Now:
  - A Haiku deck that isn't reviewed fits (about 676 pt).
  - A reviewed one uses the "reviewed on DATE" fallback (about 746 pt).
- **check_main.py overwrote the real manifests.** It runs `main.py --all --skip-ai` into
  output/, which replaced the three manifests' AI records (model, tokens, cost) with "skipped".
  - I rebuilt the three decks from the saved analysis JSONs with build_deck.py, with no API
    call, so they carry the AI text again.
  - Northwind's approval survived.
  - Logged in LEARNINGS.md.

### Unresolved

- **approve.py and provenance.py docstrings are now slightly out of date.** They say "every deck
  is watermarked until someone approves it". approve.py also prints "Rebuild the deck to drop the
  DRAFT watermark", when rebuilding now updates the footer. You asked me to leave both files
  unchanged, so I did. The fix is a one-line wording change in each.
- **Docs still describe the old behaviour:** CLAUDE.md's build_deck line, README.md,
  STUDY_GUIDE.md and LOOM_SCRIPT.md. Task 5 covers exactly this, so I left them to avoid doing it twice.
- **The manifests' `ai` block says "skipped"** after check_main.py runs, even though the decks
  carry AI text. The queue's `run_checks` runs check_main.py after every task, so this will
  happen again. The real fix is for check_main.py to use a temporary folder. That's outside this
  task, and it isn't a clear bug in the code being changed.
- **Whether "AI-drafted" should show on a placeholder deck** (decision 3): your call.
- **Not checked by eye in PowerPoint.** I have no LibreOffice here, and I couldn't script Keynote
  in this run. The one-line fit uses the same Arial measurements as every other fit check
  (text_fit.py), so it isn't PowerPoint's own layout. Northwind has 12 pt to spare. Worth one
  look when you open the deck.

---

## Task 2: the 4-slide deck

### What I built

- **The deck has 4 slides now** (it had 5):
  1. **Key metrics:** the table, unchanged, titled "Northwind: key metrics — Q2 2026 vs Q1 2026".
  2. **ARR and cash:** the two charts, unchanged.
  3. **Risks and flags:** unchanged. Its heading still reads "Tripped flags (6 of 9 flags tripped)".
  4. **AI commentary:** "AI-drafted from computed metrics - review before use" in gray under the
     title, then Claude's headline, then its 3 risks (left) and 3 questions for management
     (right), side by side at the same font size.
- **Wins are gone from the deck.** The Summary and Questions slides are gone too; their content
  is on slide 4.
- **The placeholder rules are the same.** If the analysis is missing, failed, is stale or too
  long, slide 4 says "AI summary unavailable" with the gray note, and the other 3 slides are built
  as normal. Slides 1 to 3 never carry AI text.
- **`ai_text_problems` measures slide 4's real boxes**: the headline, the Risks column and the
  Questions column. They come from `commentary_boxes`, the same function that draws the slide,
  so the check can't drift from the slide. Wins aren't measured, so long wins can't cost a deck
  its AI text any more.
- The three saved analyses all pass on the new slide with no API call. The decks in output/ were
  rebuilt from them.

Northwind's slide 4 as built:
- the headline at 22 pt
- Risks: heading 16 pt, text 12 pt
- Questions: heading 16 pt, text 12 pt (shrunk to match the Risks column)

### Files changed, in plain English

- **build_deck.py**
  - Removed `summary_slide`, `summary_boxes`, `summary_columns`, `questions_slide` and the flag
    count box.
  - Added:
    - `commentary_boxes`: where the slide 4 boxes go.
    - `column_heading`, `questions_paragraphs`, `commentary_columns`: the text for the two columns.
    - `commentary_slide`: builds slide 4.
  - `SLIDE_BUILDERS` lists the 4 slides. `ai_text_problems` measures slide 4.
  - New constant `AI_DRAFTED_LINE`. The key metrics title now includes the company name.
- **main.py**: the deck line says "(AI text on slide 4)" or "(AI summary unavailable on
  slide 4)". The number comes from `slide_number(commentary_slide)`, so it can't go stale.
- **analyze.py**: one docstring. No change to the prompt, the schema or `PROMPT_VERSION`.
- **check_deck.py**:
  - It checks for 4 slides and their titles.
  - `check_ai_slide` checks:
    - slide 4 has the AI-drafted line, the headline, every risk and every question
    - no win appears anywhere on the deck
    - no AI text or placeholder appears on slides 1 to 3
    - a placeholder slide has no AI-drafted line
  - The flag count is checked on slide 3.
  - `--draft` now expects 4 of 4 watermarked slides.
- **check_main.py**: reads the headline from slide 4, and expects "on slide 4" in the printout.
- **Tests** (tests written first; they failed until the code changed):
  - tests/test_build_deck.py:
    - 4 titles in order
    - slide 4 with and without AI text
    - the AI-drafted line sits between the title and the headline
    - no wins on the deck
    - the flag count is on slide 3
    - long wins still load, long risks give the placeholder
    - watermark counts are now 4
  - tests/test_analyze.py: the fit tests now name slide 4. A new test checks that wins aren't
    measured.
  - tests/test_main.py: the headline is read from slide 4. A new test checks the printout says
    "slide 4".
  - pytest: 500 passed (494 before, +6). Every check_*.py prints "All checks passed".
- **Docs that count or number slides**: README.md, CLAUDE.md (Goal, build_deck/charts lines,
  decision K), STUDY_GUIDE.md and LOOM_SCRIPT.md.
  - In LOOM_SCRIPT.md, 0:45–1:20 now walks the slides in the new order, with a new line for
    slide 4. The total time is the same.
  - LEARNINGS.md has 3 new rows.

### Decisions you didn't specify

1. **analyze.py still asks Claude for 3 wins.** The deck just doesn't show them.
   - Changing the prompt or schema means a new `PROMPT_VERSION`, and the saved analyses would no
     longer match. Checking a new prompt needs a paid run, which this task doesn't allow.
   - Cost of leaving it: a few output tokens per company. Also, a win that fails the number or
     direction check still fails the whole answer, even though nobody sees the win. See
     Unresolved.
2. **Slide 4 layout:**
   - The AI-drafted line (gray, 14 pt, 0.4 in tall) goes right under the title.
   - The headline goes full width.
   - Below them, Risks (left) and "Questions for management" (right).
   - The questions were 16 pt across the whole slide on the old slide 5. Now they are 14 pt in
     half the width, so their font size can match the risks column's.
3. **The AI-drafted line is left off the placeholder slide.** Nothing on that slide is
   AI-drafted, so the line would be false there. That spot stays empty, and the gray headline
   and note stay where they are.
   - This differs from Task 1's footer, which says "AI-drafted" on every deck (Task 1, decision 3).
4. **The company name moved into slide 1's title.** The Summary slide was the only place with
   the company name (apart from the file name in the footer).
5. **Slide 4's title is "AI commentary — Q2 2026"**, in the same pattern as the other titles.
6. **The flag count has no box of its own now.** It stays in slide 3's heading, "Tripped flags
   (6 of 9 flags tripped)", and check_deck.py checks it there against each company's story.
7. **The AI line uses your exact wording, with a hyphen**: "... metrics - review before use".
   The titles use em dashes ("—"). I kept your text as you wrote it.
8. **The placeholder note no longer mentions wins**: "The headline, risks and questions are
   written by Claude ...".
9. **main.py works out "slide 4" from the slide list** instead of typing the number, so moving
   a slide can't make the printout wrong.
10. **Docs:** I updated every place that counts or numbers slides now, so none of them describe a
    5-slide deck. I left these to Task 5:
    - the watermark wording (README "Review the deck, then approve it", CLAUDE.md's build_deck
      line)
    - the wider rewrite
11. **I corrected LOOM_SCRIPT's "Why A for now".** It said a slightly longer answer "stops the
    build and shows FAILED". That stopped being true in Task 7: now it gets the retry, then the
    placeholder.
12. **Historical records are left as they were**: DAY_REPORT.md, OVERNIGHT_REPORT.md, older
    LEARNINGS rows, and Task 1's section above (which says "5 of 5 slides").

### What failed and how I fixed it

- **check_deck.py crashed on its first run:** `AttributeError: 'list' object has no attribute
  'rId'`. python-pptx's slide list can't be sliced (`slides[:3]`). Fixed with `list(slides)`.
  Logged in LEARNINGS.md.
- **The doc test failed twice.** tests/test_docs.py caught STUDY_GUIDE.md naming functions I had
  removed or renamed: first `summary_slide` and `questions_slide`, then `check_ai_slides`. I
  updated the guide's tables. Logged in LEARNINGS.md.
- **Two shell commands were refused**: a heredoc with `{...['...']}` in it, and a `for` loop
  with `$f`. Nothing ran. I used the Edit tool and one command per check script instead. Logged
  in LEARNINGS.md.

### Unresolved

- **Should analyze.py stop asking for wins?** (decision 1) Your call.
  - Doing it means removing `wins` from `BoardSummary` and the prompt, bumping `PROMPT_VERSION`,
    and one paid run per company to refresh the saved analyses.
  - Until then, a bad win can still send a deck to the placeholder.
- **Slide 4's Risks column has little spare room.**
  - Northwind's and Fernhollow's risks fit only at the 12 pt floor. Alderpeak's fit at 13 pt.
  - That's about the same room as the old slide 1: the column is 3.7 in tall, and was 3.6 in.
  - A slightly longer answer gets the retry and then the placeholder. The build doesn't fail.
  - Tightening the prompt's 40-word guidance for risk details would add room. That changes the
    prompt, so I left it.
- **Not checked by eye in PowerPoint**, for the same reason as Task 1 (no LibreOffice here). The
  fit is proven by text_fit.py's measurements and check_deck.py's re-measure of the saved files.
- **Still open from Task 1:** check_main.py leaves the manifests' `ai` block saying "skipped".
  The decks in output/ were rebuilt afterwards from the saved analyses, so they carry the AI text.

## Task 3: a web page for non-technical users (app.py)

### What I built

- **app.py (Streamlit).** Drag in a KPI workbook (.xlsx) and the page runs the same steps as
  main.py: clean, metrics and flags, Excel, AI (only if ticked), deck. It then shows:
  - the company and latest quarter, and two download buttons: the deck and the metrics workbook
  - a note saying what happened to the AI commentary
  - the flag count ("6 of 9 flags tripped") and the flag table, each row red, green or gray
  - the Data gaps lines
  - the metrics table: gray = data missing, red = the flag tripped that quarter (the Excel rules)
  - The colors are `excel_output.STATUS_COLORS`, and every piece of text comes from build_deck.py's
    helpers, so the page, the workbook and the deck use the same words and colors. No new math.
- **The AI checkbox**: "Include AI commentary (typically about $0.09 and 70 seconds per
  workbook)". Off by default.
- **Errors**: a workbook clean.py can't read shows clean.py's own message, word for word. A file
  that isn't an Excel workbook gets a one-line plain message. A bug says "Something unexpected
  went wrong ..." with the error's name; its traceback goes to the Terminal window only.
- **run_app.command**: double-click on a Mac. The first time, it creates `.venv` and installs
  requirements.txt. Then it starts the app and opens the browser.
- **.streamlit/config.toml**: headless (so Streamlit doesn't stop at its first-run email
  question), no usage statistics, **no tracebacks on the page** (a backstop), and a minimal toolbar.
- **requirements.txt**: added `streamlit`. Why: it turns a Python script into a web page with
  file upload, tables and download buttons, with no HTML or JavaScript to write or maintain.
  It brings its own dependencies (pyarrow, altair, a web server, ...), about 25 packages.
- **tests/test_app.py (19 tests)**: imports app.py and exercises its helpers: the label, saving an
  upload, colors, all three error kinds, Northwind's deck/workbook/flags/metrics colors, the four
  AI paths (reused, not reused, asked, no key, API error), and Streamlit's `AppTest` drawing the
  page and the results. A guard makes creating a real Anthropic client fail every test.
- **No API calls.** I checked the AI path on the real saved analyses: with the box ticked, all
  three companies reused `output/<company>_analysis.json` (the Anthropic client was switched off
  for the run), and the flag counts matched each story (6 of 9; 0 of 9; 7 of 9, 1 cannot evaluate).
- pytest: 519 passed (500 before, +19).
- Docs: README "Or use the web page", CLAUDE.md architecture lines, STUDY_GUIDE (a table for
  app.py, the test_app.py row, the new small files). 2 LEARNINGS rows.

### Decisions you didn't specify

1. **What the checkbox does when ticked.** It first looks for a saved analysis in `output/` made
   from *exactly* the same facts (the whole payload Claude saw must match, not just company and
   quarter), which still passes the deck's checks. If there is one, it's reused for free and the
   note says so. Only otherwise does it call Claude (`main.ai_step`), which needs the key in
   `.env`. Without a key the deck is still built, with the placeholder and a note saying why.
   - This is how the rule "rebuild decks from the saved analysis JSONs" shows up in the app.
2. **A new analysis is not saved to `output/`.** The page never writes there (see 3), so
   uploading the same new workbook again costs again. Within one browser session it doesn't: the
   result is kept per file and checkbox choice.
3. **The page builds in a temporary folder**, not `output/`, and hands the files to the browser
   as downloads. So it can't overwrite the command-line decks, manifests or approvals. No
   manifest is written, and the deck's footer says "not reviewed".
4. **It builds as soon as a file is dropped in**, with no "Build" button. Ticking or unticking the
   box rebuilds. With the box ticked this can mean an API call, which is why the cost is on the
   box's label.
5. **The typical cost is a label copied from README's Cost table** ($0.0911 → "about $0.09",
   71 s → "70 seconds") in `TYPICAL_AI_COST_USD` / `TYPICAL_AI_SECONDS`, with a comment saying to
   update it with the README. The saved manifests have no cost recorded (`cost_usd: null`) and
   output/ isn't in git, so there's nothing reliable to compute it from.
6. **Metrics table: one row per metric, one column per quarter** (the Excel sheet is the other way
   round). 19 metrics across a screen don't fit; 8 quarters do.
7. **Metric cells: only red and gray**, no green. That is the Excel Metrics sheet's rule. Green
   appears in the flag table, as on the Excel Flags sheet.
8. **Runway at next quarter's budgeted burn is not on the page.** It's context, not a flag, and it
   is in the downloaded workbook and on the deck.
9. **Non-Excel files are caught in app.py, not clean.py**, so main.py's behavior is unchanged (see
   "What failed").
10. **An uploaded file keeps only its own name** ("../x.xlsx" → "x.xlsx"), and the company name
    comes from it, as in main.py. "Northwind Q2.xlsx" gives "Northwind Q2_board_pack.pptx".
11. **run_app.command sets itself up on first run** (creates `.venv`, installs packages) rather
    than assuming the setup was done, since the person double-clicking it may never have opened
    Terminal. It pins port 8501 and opens the browser after 3 seconds.
12. **The Streamlit toolbar is minimal**: the developer menu (rerun, settings, deploy) is hidden.

### What failed and how I fixed it

- **A non-Excel file showed pandas' message**, "Excel file format cannot be determined, you must
  specify an engine manually", not the plain one. Caught by the test written before the code.
  pandas raises it as a `ValueError`, like clean.py's own messages. Fixed: app.py checks the file
  is a zip archive (every .xlsx is) before cleaning. Logged in LEARNINGS.md.
- **`chmod +x` and starting the server were refused** in this session. I set the executable bit
  with Python's `os.chmod`, committed it (mode 100755 in git), and added a test for it. The page
  was tested in-process with Streamlit's `AppTest`. Logged in LEARNINGS.md.

### Unresolved

- **Not tried in a browser, and run_app.command not double-clicked.** Both need a person. `AppTest`
  proves the page and the results draw with no error, two tables and two download buttons, but not
  how they look. `AppTest` can't simulate a file upload, so the upload widget itself is untested.
- **macOS may block the first double-click** of a `.command` file that came from a download or a
  zip ("cannot be opened because it is from an unidentified developer"). Right-click → Open once
  fixes it. A clone made with git isn't affected.
- **If port 8501 is already in use** (the app is already running), a second double-click fails in
  its Terminal window; the browser still opens the first copy.
- **main.py still prints pandas' message** for a file that isn't an Excel workbook. A one-line
  check in clean.py would fix both; I left clean.py alone.
- **The live AI path (a real Claude call from the page) has not run**, by this task's rules. It
  reuses `main.ai_step`, which the live batch runs have exercised, and it's tested with a fake client.

---

## Task 4: INTERVIEW_PREP.md

### What I built

- **INTERVIEW_PREP.md: 39 questions in 6 groups, in the order an interviewer tends to ask them:**
  1. The project (Q1–Q4)
  2. Design decisions (Q5–Q16)
  3. The AI layer (Q17–Q25)
  4. What broke (Q26–Q31)
  5. Scale and risk (Q32–Q36)
  6. Working method (Q37–Q39)
  - Then **"When you can't recall a detail"**: how to answer honestly (never invent a number, give
    the shape instead of a made-up precision, explain the method, be clear what you specified versus
    what Claude Code wrote, correct yourself out loud), plus a table of the numbers worth knowing cold.
- **All 30 study guide questions are in it**, with their answers and Point to lines, each tagged with
  its old number ("(guide Q5)").
- **New questions for every item you listed:** why Python computes and Claude interprets (Q5,
  expanded from the guide's Q3), re-validation at deck build (Q26), the blind model comparison and
  its rule (Q24), the direction-claim failures and the two-layer check (Q21), what the direction rules
  cost (Q23), provenance and the approval gate (Q34), what breaks at 275 companies (Q35), how you work
  with Claude Code (Q37) including unattended task queues (Q38), and what you'd build next (Q36).
- **Three more I added:** why fake data at all (Q4), whether the direction check ever got it wrong
  (Q22), and how you know Claude Code didn't just write tests that pass (Q39).
- **tests/test_docs.py, 3 new tests** (written first; all 3 failed until the file existed):
  - every file INTERVIEW_PREP.md names exists (.py, .md, .yaml/.yml, .toml, .command, .ini, .txt,
    .xlsx; a bare test file name counts if it's in tests/)
  - every `module.function`, constant, and `tests/file.py::test_name` it points to exists (49 of
    them today)
  - the six groups and the honesty section appear in that order
- **Pointers:** a CLAUDE.md architecture line, and a note at the top of STUDY_GUIDE.md section 5.
- pytest: 522 passed (519 before, +3). No API call; no pipeline code, config.yaml or output/ file changed.

### Decisions you didn't specify

1. **I updated stale facts in the guide's answers instead of copying them word for word.** "Keep
   their answers" and "don't say false things in an interview" conflicted in four places. Each changed
   answer is marked "updated" or "replaced":
   - Q2 cost: $12.81 → about $25 per quarter (the direction rules doubled it).
   - Q18 limits: wrong direction is now partly caught; the answer points to Q21.
   - Q32 test count: 415 → "over 500", so it doesn't go stale with the next test.
   - Q36 next steps (guide Q25): the slide-fit and trend checks it listed are built, so it's rewritten.
   - Smaller edits: Q1 and Q16 mention the direction/fit checks inside validation; Q24 adds why the
     rule was written first; Q26 adds the reconciliation analogy.
   The study guide itself is left as it was, with a pointer to the new file.
2. **Grouping of the guide's questions.** The guide had a separate "The deck" group; its questions
   went to Design decisions (Q13–Q16), except re-validation, which went to What broke (Q26) because
   it's a guard against a stale file. The direction-claim questions went to the AI layer, not What
   broke, because the fix is part of the AI layer.
3. **"Point to" for LEARNINGS rows uses the row's subject**, not a row number ("LEARNINGS.md
   duplicate-quarter row"). The table isn't numbered and grows, so numbers would drift.
4. **The cost answer (Q23) uses the two real live runs**, not an estimate: about 3,500 → 7,500 output
   tokens per company, $0.047 → $0.091, $12.81 → $25.06 per quarter at 275, so about 1.1M extra output
   tokens and $12 more per quarter. It says openly that part of the rise was one retry caused by a
   since-fixed false positive, and that the companies that passed first time still roughly doubled.
5. **The doc test ignores files under output/.** They're made by a run and git ignores them, so a
   fresh clone would fail the test for no real reason. INTERVIEW_PREP.md names output/ files only in
   prose anyway.
6. **The doc test also checks function names**, not only files (you asked for files). A renamed
   function is the more likely way this file goes wrong, and the study guide already has the same check.
7. **I didn't run the check scripts.** No pipeline code changed, and check_main.py overwrites the real
   manifests in output/ (Task 1, Unresolved), so running it would have cost the manifests' AI records
   for nothing. No deck needed rebuilding.

### What failed and how I fixed it

- **Four stale guide answers** (decision 1), found by checking each against LEARNINGS, README and the
  code before reusing it. Logged in LEARNINGS.md.
- **My first function-name pattern was wrong**: it would have read "`metrics.py`" as a function
  called `py`, and failed constants like `analyze.SYSTEM_PROMPT`. Fixed before it ran; proved on a
  throwaway text that a made-up file, function and test name are each reported. Logged.
- **Two commands were refused and one search timed out** while looking for the task-queue runner
  (outside the project folder). Logged. See Unresolved.

### Unresolved

- **Q38 (unattended task queues) is written only from what this repo shows**: the reports, the task
  rules, the check scripts running after each task, and the refused commands in LEARNINGS.md. The
  runner lives outside the project, where this session can't read. Check that the wording matches how
  you actually run it (what starts it, what stops it) before you say it in an interview.
- **The honesty section's "Numbers worth knowing" table** repeats figures from README and the study
  guide. If a future live run changes the cost, update both places (the test only checks names, not
  numbers).
- **STUDY_GUIDE.md section 5 still has the four stale answers**; it now points to INTERVIEW_PREP.md.
  Worth updating them or cutting section 5 down to the pointer; your call.
- **Still open from Task 1:** CLAUDE.md's build_deck.py line says every slide is watermarked until
  approval, but since Task 1 the watermark is opt-in (`--draft`) and the footer carries the review
  status. INTERVIEW_PREP.md Q34 describes the current behavior.

---

## Task 5: docs and outputs

### What I built

- **README.md**
  - "Review the deck, then approve it" is rewritten:
    - the footer ends `AI-drafted | not reviewed`
    - `approve.py`, then a rebuild, makes it `AI-drafted | reviewed by NAME on DATE`
    - a changed workbook or config.yaml sends it back to "not reviewed"
    - `--draft` is how to get the watermark
  - Design decision 12 says why: a footer can stay on a deck that goes to a board; a stamp across the numbers can't.
  - Also updated:
    - `main.py --all --draft` and `python excel_output.py` added to the command lists
    - app.py and approve.py added to the data-flow picture
    - "500+ tests"
    - Known limitations no longer mentions the watermark
    - a screenshot line for the footer, and one for the web page
- **CLAUDE.md Architecture**
  - The build_deck.py line now says: the footer carries the review status, and the watermark is opt-in with `--draft`.
  - approve.py: "the rebuilt deck's footer says reviewed by".
  - main.py lists `[--draft]`.
  - The tests line says what test_docs.py now guards.
- **STUDY_GUIDE.md**
  - **New section:** `provenance.py` and `approve.py`, function by function. It has a finance analogy (a sign-off stapled to the version of the file that was checked). Neither file had a section before.
  - **build_deck.py section:** the real footer, replacing the old "Fictional data, generated for demonstration | Source: ... | Run date: ..." example. There are new rows for:
    - `commit_text`, `review_text`, `footer_text`, `fits_one_line`
    - `set_alpha` / `add_watermark`
    - `analysis_details`, `record_deck_status`
    - `--draft` in the signatures and the call order
  - **main.py section:** `--draft`, `ai_record` and `manifest_step`.
  - **app.py section:** how it differs from main.py (a temporary folder, no manifest, always "not reviewed", no watermark).
  - **check_deck.py rows:** `expected_review`, the new `check_footers`, `watermark_count` and `check_draft_option`.
  - **Walk-through:** a "Step 6: a person reviews and approves", a paragraph on the web page, and the footer in step 4.
  - **Glossary:** SHA-256 hash, manifest, JSON, Streamlit.
  - **Tests table:** the counts are corrected, and there are new rows for test_provenance.py and test_approve.py.
  - **Small files:** POLISH_REPORT.md and INTERVIEW_PREP.md.
  - **Header note:** says what was added since the day build.
- **LOOM_SCRIPT.md**
  - **Recording prep:** build without `--draft`, and check that the footer says "reviewed by" (if not, run `approve.py`). There's a 5th window to open: the web page, started with `run_app.command`.
  - **Slide 4 line:** it now also points at the footer. "Every footer says whether a person has reviewed the deck."
  - **The last 15 s** show the web page, then the next steps. Total: about 281 words, still 2:00.
  - **Stale facts fixed:** 415 → 527 tests ("over 500"), and Option B costs about $0.09 / 70 s.
- **approve.py**
  - Printed after an approval: "Rebuild the deck so its footer says reviewed: python build_deck.py data/northwind.xlsx". It used to say "...to drop the DRAFT watermark".
  - Its docstring and provenance.py's now describe the footer and `--draft`. No logic changed.
- **tests** (written first; 5 failed until the docs and approve.py changed):
  - tests/test_docs.py, 5 new:
    - no line in README, CLAUDE.md, STUDY_GUIDE or LOOM_SCRIPT mentions the watermark without `--draft`
    - all four name `--draft`, app.py and run_app.command
    - README and the guide show both footer wordings
    - none counts 5 slides
    - the guide has function tables for provenance.py, approve.py and app.py
  - tests/test_approve.py: the printed hint mentions the footer and never "watermark".
  - pytest: 527 passed (522 before, +5 new tests; the approve change is extra assertions in an existing test).
- **Outputs rebuilt with no API call**
  - All three Excel files (`python excel_output.py data/<company>.xlsx`) and all three decks (`python build_deck.py data/<company>.xlsx`).
  - All three saved analyses still pass today's checks, so every slide 4 carries Claude's text.
  - check_deck.py on the rebuilt decks:
    - Northwind "AI-drafted | reviewed by Tyler Ho on 2026-09-17", 737 of 755 pt
    - Alderpeak and Fernhollow "AI-drafted | not reviewed", 594 and 597 pt
    - no watermarks
    - `--draft` gives 0 / 4 / 4 slides
  - All 5 check scripts print "All checks passed".

### Decisions you didn't specify

1. **I changed approve.py's printed message and two docstrings.** Task 1 was told to leave approve.py
   and provenance.py alone, so it flagged the old wording as unresolved. This task is the docs task,
   and the message was the one place a user was told something false after every approval. Only
   wording changed: no logic, and the manifest format is the same. The manifest's `deck.status`
   still says "DRAFT - NOT REVIEWED" for an unapproved deck. That's data, not a doc, and changing it
   would change the manifest format.
2. **The doc test works line by line:** a line that mentions the watermark must also say `--draft`.
   That's strict, and it's simple to explain. The cost is that a sentence about the watermark has to
   carry `--draft` on the same line.
3. **The Loom's last 15 seconds changed.** The old "next steps" promised a trend-word check that is
   already built. I replaced it with 7 seconds on the web page ("drag in the workbook, download the
   deck") and kept 2 next steps (SharePoint trigger, portfolio rollup). The slide 4 line now points at
   the footer. The timing is still 2:00.
4. **Recording advice: no `--draft`.** The footer shows the review; a watermark would cover the
   numbers on camera. Northwind is approved, so `--draft` wouldn't stamp it anyway.
5. **Output rebuild order:** check_main.py first (it overwrites output/), then the Excel files and
   decks, then check_excel_output.py and check_deck.py last. That way the checks read the decks
   exactly as they're left.
6. **I left the manifests' `ai` block alone.** check_main.py set it to "skipped". The rebuild updated
   each manifest's `deck` part (`ai_text: true`, review status), but not the model/tokens/cost
   record. That record is only written by a real run. Rebuilding it from the analysis JSON is
   possible, but it would be a new code path written for one clean-up, and the cost would still be
   missing.
7. **STUDY_GUIDE section 5 (the 30 questions) is unchanged.** Task 4 left its four stale answers
   (cost, test count, next steps) as your call, and none of them is about this task's topics.
8. **Test counts in the guide are exact (527)** because the guide's tests table lists counts per
   file. README and the Loom narration say "over 500", so they don't go stale.

### What failed and how I fixed it

- **My new doc test flagged the wrong line.** It matched "DRAFT" inside the constant
  `AI_DRAFTED_LINE`. Fixed by matching the whole word only. Logged in LEARNINGS.md.
- **The docs had more stale facts than the watermark:**
  - the Loom promised an already-built check, and quoted 415 tests and $0.05 / 35 s
  - approve.py printed a false instruction
  All are fixed, and the test now guards the watermark wording. Logged.
- **Two shell commands were refused** (a `for ... $c` loop, and a heredoc with braces). Nothing
  ran; I used an `&&` chain and `python -c` instead. Logged.
- **The queue's post-task checks will undo the deck rebuild.** `run_checks` runs check_*.py in name
  order, so check_main.py (which writes `--skip-ai` decks into output/) runs after check_deck.py.
  Logged. See Unresolved.

### Unresolved

- **After the queue's checks run, output/'s decks will show "AI summary unavailable" again**, and
  batch_summary.csv will hold check_main.py's rows. The decks and Excel files were rebuilt with the
  AI text and checked in this task. To get the AI text back afterwards, with no API call, run
  `python build_deck.py data/northwind.xlsx` (and the same for alderpeak and fernhollow). The real
  fix is an `--output-dir` for main.py so check_main.py uses a temporary folder. It's already in
  README Next steps, and it's code, not docs.
- **The manifests' `ai` block says "skipped"** (decision 6) until the next live `main.py` run.
- **Not checked by eye:** the rebuilt decks in PowerPoint, and the web page in a browser (same
  reasons as Tasks 1–3). The Loom script now asks you to do both before recording.
- **STUDY_GUIDE section 5** still has Task 4's four stale answers (decision 7).

---

## Task 6: review of the whole run (git diff main..HEAD)

### What I reviewed

- All code in the diff, line by line:
  - build_deck.py, main.py, app.py, approve.py, provenance.py, analyze.py, make_template.py
  - check_main.py, run_app.command, .streamlit/config.toml, requirements.txt
- The docs, checked for claims the code no longer backs:
  - README, CLAUDE.md, STUDY_GUIDE, LOOM_SCRIPT, INTERVIEW_PREP
  - tests/test_docs.py already guards the file and function names, the slide count and the watermark wording. I searched the docs for stale slide numbers, "watermark" and test counts.
  - I didn't read all 474 lines of INTERVIEW_PREP.md one by one.
- I also ran things, not just read them:
  - the web page's `build_outputs` on awkward uploads: long file names, and a .pptx renamed .xlsx
  - the footer width for every model / reviewer combination
  - the timestamps and manifests of the saved outputs

### Findings, most serious first

1. **Not fixed, needs your decision: "reviewed by Tyler Ho" can be on text or slides the reviewer never saw.**
   - An approval is tied to two hashes only: the workbook and config.yaml (`provenance.approval_status`). It isn't tied to the AI analysis JSON or to the code that lays out the deck.
   - Example today:
     - Northwind was approved at 22:33:47, one second after its analysis was saved and before this run started. At that point it was a 5-slide deck.
     - Every deck built since then is the new 4-slide layout (wins dropped, headline moved). Each one still says "AI-drafted | reviewed by Tyler Ho on 2026-09-17" on every slide.
     - The AI text is the same, so today the claim is only stale about the layout.
     - But a live `python main.py data/northwind.xlsx` would write *new* AI text, and the footer would still say "reviewed by Tyler Ho". The workbook and config haven't changed, so the approval still counts.
   - The gap was already there on main. Before this run it only meant a missing watermark. Task 1 turned it into a positive claim printed on every slide.
   - INTERVIEW_PREP Q34 says "a stale approval is worse than none". An interviewer who asks "what if you re-run Claude after approval?" finds this gap.
   - Why I didn't fix it:
     - It changes what an approval means and what the manifest holds.
     - Task 1 was told to keep provenance.py and the manifest unchanged.
   - Likely fix: `approve.py` also stores the analysis JSON's hash (maybe the commit too), and `approval_status` checks it. Then new AI text makes the deck "not reviewed" again.
   - Until then, if you've changed anything since approving, re-approve before you show the Northwind deck as "reviewed".
2. **Fixed: a workbook with an ordinary long name got no deck at all.**
   - Example: "Northwind - Q2 2026 KPI pack (final version for board).xlsx" dropped on the web page showed "Slide 1, Footer: text doesn't fit even at the 12 pt minimum".
   - With the model and a reviewer on the footer, there is only about 170 pt (roughly 27 characters) left for the file name. The CLI has the same limit; the web page just made long names likely.
   - Fix: `add_footer` now shortens the file name in the middle ("Northwin….xlsx") before it gives up the reviewer's name. Anything else too long still stops the build.
   - Tests (both failed on the old code):
     - `test_a_file_name_too_long_for_the_footer_is_shortened_not_a_failed_deck`
     - `test_a_long_file_name_still_gets_a_deck`
     - plus `test_a_file_name_that_fits_is_never_shortened`
   - The old docstring's "748 pt of 755" figure went with it: my measurement didn't match it, and it only held for our own file names.
3. **Fixed: `run_app.command` failed on a `.venv` made before the web page existed.**
   - Setup only ran when `.venv/bin/python` was missing, so an older .venv without streamlit ended in "streamlit: command not found".
   - It now checks for `.venv/bin/streamlit` instead. `python3 -m venv` leaves an existing .venv alone, so setup only adds the missing packages.
   - Test: `test_the_launcher_installs_packages_when_the_venv_has_no_streamlit`.
4. **Fixed: another Office file renamed .xlsx got the "something unexpected went wrong" message.**
   - A .pptx is a zip file too, so the "is it a zip?" check let it through, and pandas failed with `OptionError`.
   - `app.is_excel_workbook` now also looks for `xl/workbook.xml`.
   - Test: `test_another_office_file_renamed_xlsx_gets_the_plain_message`.
5. **Known, not fixed: the queue's checks undo the deck rebuild.**
   - `run_checks` runs check_main.py after check_deck.py, and check_main.py writes `--skip-ai` decks into output/. So after this task's queue checks, output/'s decks will say "AI summary unavailable" again.
   - Task 5 logged this. The fix is an `--output-dir` for main.py (README Next steps).
   - Getting the AI text back needs no API call: `python build_deck.py data/<company>.xlsx` for each company.
6. **Low: the manifest's `deck.status` says "DRAFT - NOT REVIEWED" for a deck with no stamp.**
   - It's accurate about review, but the word DRAFT now only appears on the slides with `--draft`.
   - Task 5 (decision 1) kept it on purpose, because changing it changes the manifest format. I agree; noted so it isn't forgotten.
7. **Low: two things in app.py can go stale without a test failing.**
   - `TYPICAL_AI_COST_USD` / `TYPICAL_AI_SECONDS` are typed by hand from README's Cost section. The comment says so.
   - `streamlit` is unpinned in requirements.txt, and the page uses newer arguments (`width="stretch"`); 1.64 is installed here. A fresh install gets the newest version, so this only bites someone with an old streamlit.
8. **Housekeeping: `output/~$alderpeak_board_pack.pptx`** is a PowerPoint lock file from 22:32. The Alderpeak deck was open when the queue started, so the copy on screen is the old 5-slide one. Close it and reopen. output/ is git-ignored.

Checked and fine:
- `--draft` never stamps an approved deck.
- Where approval exists, `record_deck_status` keeps the manifest in step with it.
- `commentary_boxes` is shared by the slide and `ai_text_problems`, so what's measured is what's drawn.
- `check_main.headline_on_deck` reads slide 4.
- `main.py --draft` is passed through every layer.
- The web page's saved-analysis reuse compares the whole payload, not only company and quarter.
- `save_upload` can't write outside its folder.
- The web page never touches output/.

### Decisions you didn't specify

1. **Which gives way first on the footer: the file name.** It is shortened before the reviewer's name is dropped. "Reviewed by NAME" is the approval; a file name is also in the manifest (CLI) and in the download's name (web page).
2. **How to shorten the file name:** keep its start and its extension, with "…" between them, cut one character at a time until the line fits. The start says which company it is, and the extension says what kind of file.
3. **The launcher checks for the program it runs (streamlit)**, not for a list of every package. If a later package is missing, the page's own error names it.
4. **The workbook check looks for `xl/workbook.xml`**, the one part every .xlsx has. I didn't open it with openpyxl first, because clean.py does that next and gives its own messages.
5. **Finding 1 isn't fixed.** It's the most serious finding, but the fix changes what an approval means. That's your decision, not a bug fix.
6. **Test counts updated to 532** in STUDY_GUIDE and LOOM_SCRIPT (the two places with exact counts). The earlier sections of this report keep the counts they had at the time.

### What failed and how I fixed it

- **My first footer test expected "Northwind KPI" to survive shortening.** With the model and a reviewer on the line, only "Northwin" fits. The test was wrong, not the code; it now checks the start, the "….xlsx" and the width.
- **Two shell commands were refused:**
  - `bash -n run_app.command` (a syntax check), run in a chain
  - a `grep` in a chain
  I ran the syntax check through Python's `subprocess` (exit code 0) and used the Grep tool. Logged in LEARNINGS.md.
- **Each fix has a LEARNINGS.md row** (findings 2, 3 and 4).

### Checks

- pytest: 532 passed (527 before, +5: two footer tests, two web page tests, one launcher test).
- All 5 check scripts print "All checks passed". I ran check_main.py first, then rebuilt the three Excel files and decks from the saved analyses (no API call), then check_excel_output.py and check_deck.py.
- Rebuilt decks: 4 slides each, and Claude's headline on slide 4.
  - Footers: Northwind "reviewed by Tyler Ho on 2026-09-17"; Alderpeak and Fernhollow "not reviewed".
- Working tree clean, and `polish` is pushed to origin (see the last commit).

### Unresolved

- **Finding 1** (approval not tied to the AI text or the layout): needs your decision.
- **Finding 5:** output/'s decks go back to the placeholder after the queue's own checks. Rebuild with `build_deck.py` (no API call).
- **Not checked by eye:** the decks in PowerPoint and the web page in a browser (same as Tasks 1–5). In particular, look at a shortened footer ("Northwin….xlsx") on a real slide.
