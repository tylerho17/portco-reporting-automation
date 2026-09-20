# Target question set: Northwind

Company profile: NRR fell from 108% to 97% over three quarters, pipeline rising while bookings are
flat, burn 20% over budget, runway 11 months. Northwind's own figures are NRR 97.1% (annualized) and
runway 11.0 mo, Q2 2026.
Source: docs/RESEARCH_BOARD_PACKS.md, section 6, "Ten questions for a company at 97% NRR and
11 months runway". Themes, order and provenance markers are section 6's own.

This is the gold standard the analyze.py question prompt is tuned against. Nothing in it is
generated. Read it with these limits in mind:

- **Only the sourced parts are sourced.** A `verbatim` question is quoted from the source named
  under it. A `constructed` question is section 6's own construction on a sourced finding, quotable
  to no one. Where section 6 says no sourced question exists (burn budget variance, a changed
  metric definition), that is a gap in the public record, not an oversight here.
- **The target does not pass its own checks at 100%.** Section 6's generation rule says every
  question names a metric and its value, but questions 3, 7, 8 and 9 quote no figure (8 names
  pipeline and no number), so it is 6 of 10. And five of the ten ask for a comparison, a
  sensitivity or a consistency check rather than a cut of the data (2, 6, 7, 9 and 10), so it is 5
  of 10 on decomposition. The scorer prints these two rates beside every set's, so a prompt is read
  against the standard and not against a perfect score.
- **One marker per question.** Section 6 marks two questions with two provenances. Question 3 keeps
  the table's own `[V]`. Question 10 is a constructed clause plus a verbatim clause with no lead
  named, so it takes `constructed`, the weaker one. The source line under each says which part is
  which.
- **Left out on purpose:** section 6 also lists two reserve questions ("Is there liquidity behind
  the compliance?", "Does the sponsor have our back?") and one generic follow-up ("Are these
  challenges specific to one region or product line, or are they systemic?"). They are for a
  covenant or ownership context Northwind does not have, so the scored set stays at ten.
- **Two edits to the wording, both forced.** Section 6's em dashes are a comma, a colon or a full
  stop here (the project has none in a .md file). And question 3's scoping clause, which section 6
  writes as a trailing note ("applied to the accounts comprising the eleven-point decline"), is
  written as a sentence ("Apply this to the accounts comprising the eleven-point decline"), so it
  reads as part of the question. No other word was changed.

The markers are section 6's own: `verbatim` is its **[V]**, verbatim from the cited source;
`verbatim with values` is its **[V+]**, a verbatim question or instruction with the company's values
inserted; `constructed` is its **[C]**, constructed by that report on a sourced finding, quotable
to no one.

## Theme: Retention decomposition

1. [verbatim with values] NRR fell from 108% to 97% over three quarters. Rebuild the monthly waterfall for the last eight quarters so the components reconcile to ending ARR, then tell us which month and which cohort the eleven points actually live in.
   Source: the instruction is verbatim, the values and the closing clause are built on "isolate the month, isolate the cohort" (ORM, https://orm-tech.com/blog/why-net-revenue-retention-dropped).

2. [constructed] What has gross retention done over the same three quarters in which net fell to 97%? If GRR held while NRR fell eleven points, we have a pricing and expansion problem; if GRR fell too, the retention engine is broken and the sales motion is paying for it.
   Source: constructed; the second clause paraphrases the verbatim escalation pattern "Falling gross with flat net is the pattern to escalate on" (ORM, https://orm-tech.com/blog/why-net-revenue-retention-dropped).

3. [verbatim] How do churn and downgrades distribute across your customer base? Are losses concentrated in specific segments or cohorts? Apply this to the accounts comprising the eleven-point decline.
   Source: the two questions are verbatim (G-Squared CFO, https://www.gsquaredcfo.com/blog/arr-quality); the scoping clause, the last sentence, is constructed.

## Theme: Burn variance and efficiency

4. [constructed] Burn is 20% over budget. Split the burn multiple: did net burn rise, did net new ARR fall, or both, and by how much each, in dollars, versus plan?
   Source: constructed; no verbatim board question on burn budget variance exists in public. Built on the Net Burn ÷ Net New ARR definition (The SaaS CFO, https://www.thesaascfo.com/how-to-calculate-the-burn-multiple/).

5. [verbatim with values] Is the 20% overrun a gross margin problem, a sales efficiency problem, a churn problem or a growth problem? Show us which of the four moved.
   Source: derived directly from "any serious problem will eventually impact the Burn Multiple" and its named four (David Sacks, https://sacks.substack.com/p/the-burn-multiple-51a7e43cb200).

## Theme: Liquidity and downside

6. [verbatim with values] At 11 months of runway, how much cushion remains after downside sensitivity, and what happens if EBITDA misses by a further 5–10%?
   Source: both clauses are verbatim, the runway figure is inserted (A Faster Exit, https://afasterexit.com/guides/covenant-headroom/cfo-protocol/).

7. [verbatim] What is the worst outcome here, and how likely is that outcome?
   Source: verbatim (CFO Secrets, https://www.cfosecrets.io/p/cfo-role-boardroom).

## Theme: Pipeline versus bookings

8. [verbatim with values] Pipeline is up while bookings are flat. What is conversion at each stage entered, and does the decline survive segmentation by segment, lead source, rep tenure and deal size?
   Source: both analysis instructions are verbatim, rendered as a question (ORM, https://orm-tech.com/blog/why-is-my-win-rate-dropping).

9. [verbatim] Are close dates realistic or repeatedly pushed? On the deals carrying this quarter's forecast, is there a scheduled next step with clear buyer action, and has the cost of inaction been quantified?
   Source: all three checks are verbatim (A Sales Growth Company, https://salesgrowth.com/sales-pipeline-review-playbook/).

## Theme: Definitions and assumptions

10. [constructed] Is the 97% computed on exactly the same basis as the 108%: same per-customer cap on GRR, same treatment of contraction versus churn, same monthly-versus-year-ago period basis? And what are we assuming in the re-forecast that, if wrong, would materially change this decision?
    Source: the first clause is constructed; no sourced question on definitional change exists, and it is built on the consistency norm (Mostly Metrics, https://www.mostlymetrics.com/p/your-complete-guide-to-board-meetings-d92) and the GRR cap definition (SaaS Capital, https://www.saas-capital.com/blog-posts/grrumbling-about-retention-metrics-or-pitfalls-in-measuring-saas-churn/). The second clause is verbatim (Aspirations Group, https://www.aspirations-group.com/post/the-board-question-that-separates-good-governance-from-theater).
