# Demo: a 5 minute walkthrough

A script for showing the Board Pack Generator to someone who doesn't write code: an investor, an
operating partner, a hiring manager. Every click is on the web page; nothing is typed into a
terminal once the demo starts. The words under **Say** are a guide, not lines to learn: say them
your own way.

All three companies are made up. Never show the tool with real company data.

## Before the demo (10 minutes ahead, not timed)

1. **Reset the output folder.** In Terminal, in the project folder:

   ```
   source .venv/bin/activate
   python demo_reset.py
   ```

   It takes about 10 seconds and must end with **Ready for the demo.** It deletes every file the
   tool built during practice (decks, memos, approvals, exports, run logs), keeps the saved AI commentary, and
   builds everything again from the three workbooks, with no call to Claude and no cost. It also builds
   last quarter's numbers first, so the page has "what changed since last quarter" to show.

   One line it prints is expected: `! Fernhollow: no saved analysis ... "AI summary unavailable"`.
   Fernhollow's saved commentary was written from older wording, so the tool won't reuse it (see
   FINAL_REPORT.md, Task 4). The demo shows AI commentary on Northwind only.

   If it ends **Not ready for the demo**, fix each line marked ✗ (each one says how) and run it again.

2. **Check the commit.** If `git status` shows uncommitted changes, every slide's footer shows the
   commit with a `*` after it ("built from code that wasn't committed"). Commit first, then run
   `python demo_reset.py` again, so the footer is clean.

3. **Open the web page.** Double-click `run_app.command` in Finder. A Terminal window opens (leave it
   open: closing it stops the page) and the browser opens the page at http://localhost:8501 after a
   few seconds.

4. **Set the screen.** Browser full screen, zoom 100%, PowerPoint (or Keynote) open in the background,
   notifications off. Leave the **Ask Claude for AI commentary** box unticked for the whole demo: ticked,
   Generate could call Claude and spend money; unticked, the saved commentary is reused for free.

## 1. The problem (0:00 to 0:45)

**Show:** the portfolio page, not scrolled.

**Say:** "Every quarter, each portfolio company sends a KPI spreadsheet, and every one is laid out
differently: different column names, numbers typed as text like '$1.2M', a quarter left blank. Someone
on the deal team spends hours turning each one into a board update. This tool does it in seconds: it
cleans the spreadsheet, computes the metrics an investor watches, flags what breaks our thresholds, and
builds the deck. Python does every calculation. Claude only writes commentary on numbers Python has
already checked."

## 2. The portfolio at a glance (0:45 to 1:30)

**Show:** the table of three companies. Point at the **Flags tripped** column, row by row.

**Say:** "Three made-up software companies. Alderpeak is healthy: 0 of 9 flags tripped. Northwind is
growing fast but has 6 of 9 flags tripped. Fernhollow is in trouble: 7 of 9 flags tripped, 1 cannot
evaluate, because a quarter of its data is missing and the tool won't guess a number that could
reach a board."

Point at the **Data gaps** column: "It says exactly which quarter is blank, instead of filling it in."

Point at **Deck status**, which reads **DRAFT - NOT REVIEWED**: "Nothing leaves here as final until a
person signs it off. We'll do that in a minute."

## 3. One company: Northwind (1:30 to 2:45)

**Click:** click **Northwind** (the company name in the first column).

**Show and say, top to bottom (scroll slowly):**

- **The flags table**, "6 of 9 flags tripped": "Red is a breach, green is fine. Net revenue retention
  is 97.1%, below 100%: existing customers are shrinking. Burn is 20% over budget. Runway is 11.0
  months, under our 12 month line. Each row shows the threshold, so you can argue with it."
- **What changed since the last run:** "Five flags flipped since last quarter's run: four were green,
  and one couldn't be checked for lack of data. It tells you what changed, so you don't reread the
  whole pack: runway went from 15.0 to 11.0 months."
- **Data gaps**, in gray: "Q1 2025 is blank in their spreadsheet. Every metric that needs it says
  'data missing' rather than a made-up number."
- **Metrics** and **ARR and cash**: "Every quarter, every metric, and the two charts that go on the
  deck. The blank quarter is a visible gap in the chart."
- **AI commentary:** "This is Claude's part: a headline, a short diagnosis, three risks and
  eight to ten questions for management, grouped by theme. It's written only from the computed numbers, and every number in it is checked against
  them before it's allowed on a slide. It's labelled 'review before use' because it's a draft for a
  person to check."

## 4. The deck (2:45 to 3:45)

**Click:** scroll back to the top and click **Download deck**. Open the downloaded file
(`northwind_board_pack.pptx`) in PowerPoint.

**Show:** the four slides, about 10 seconds each.

1. Key metrics, this quarter against last: "The table a board member reads first. Red is a breach."
2. ARR and cash charts: "Growth, and how long the cash lasts."
3. Risks and flags, with the data gaps listed: "Every tripped flag with its threshold, and what's missing."
4. AI commentary: "Claude's headline, risks and questions, marked AI-drafted."

Point at the footer: "Every slide says where it came from: the spreadsheet, the date, the version of
the code, the AI model, and 'AI-drafted | not reviewed'. Anyone holding this deck can see nobody has
signed it off yet."

## 5. Sign-off (3:45 to 4:30)

**Click:** back in the browser, scroll to **Approve** at the bottom of Northwind's page. Type your
name in **Reviewer name**, then click **Approve**.

**Say:** "Approving records who signed off, against exactly this spreadsheet and these thresholds. If
anyone changes the numbers afterwards, it goes back to 'not reviewed' by itself."

**Click:** click **Generate** at the top of the page (the navy button). The status line under the
title now reads "Deck status: approved by" your name.

**Say:** "Approving doesn't build anything; Generate rebuilds the deck, and now the footer says who
reviewed it and when." (Optional, if time allows: click **Download deck** again and show the footer
now reads "AI-drafted | reviewed by" your name.)

## 6. Close (4:30 to 5:00)

**Click:** click **Back to portfolio**.

**Say:** "So: a messy spreadsheet in, a checked board deck out, with the reasoning visible and a person
signing off. The same page does the whole portfolio at once with Generate all, and a new company is
added by dropping its spreadsheet under Add a company. What would you want it to show you?"

Stop there and take questions.

## If something goes wrong

| What you see | What to do |
|---|---|
| The browser shows "This site can't be reached" | Wait five seconds and reload: the page takes a moment to start. |
| Download deck is greyed out | The files are out of date. Click **Generate** on that company (box unticked), then download. |
| Northwind shows "No AI commentary" | Its saved commentary is missing. Say "the AI part is optional; the numbers never depend on it", carry on, and run `python demo_reset.py` after the demo to see what it says. |
| A red error box | Read it out: every error says what's wrong in plain words. Then click **Back to portfolio**. |
| You approved the wrong company, or clicked something twice | Carry on. `python demo_reset.py` afterwards puts everything back. |

## Questions they may ask

- **"Could the AI make up a number?"** It can write one, but a check compares every number in its text
  with the computed metrics, and it's retried once, then left off the deck if it still fails. The
  deck's numbers never come from the AI.
- **"What does it cost?"** About 9 cents and 70 seconds per company for the AI commentary; everything
  else is free and takes seconds.
- **"What if the spreadsheet is laid out differently?"** It learns the new column names once: a person
  confirms what each unknown header means, and next quarter's spreadsheet runs without asking.
- **"Is the data real?"** No. All three companies and every number are made up.
- **"How would I know if last night's run went wrong?"** Scroll to **Recent runs** at the bottom of the
  portfolio page. Every run is listed, from the page or a scheduled command-line run; open one to see
  each company's steps, how long each took, and the exact error of anything that failed, in red. The
  Generate click from this demo is the one at the top.

## After the demo

Run `python demo_reset.py` again. It clears your approval and anything else the demo changed, so the
next demo starts from the same place.
