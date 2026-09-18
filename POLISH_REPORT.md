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
