"""One valid v5 answer for the tests that need Claude to have said something sensible.

Not a test file: the tests import it (pytest puts tests/ on the path). It lives here rather than
being typed out in each test file, because the shape and the question rules are in one place in
analyze.py and a copy per file would drift from them.

Every number in it is a threshold from config.yaml, which every company's payload holds, so the
same answer passes analyze.validate_summary for Northwind, Alderpeak, Fernhollow and the made-up
Testco of tests/test_build_deck.py. The questions carry a metric, a value and a demand
(a decomposition, a reconciliation or a distance to breach), and the first retention question is
the gross versus net divergence test, so they pass whether or not NRR has moved.
"""

# 69 words, and no number in it, so only the length and fit checks can fail it.
PLAIN_DIAGNOSIS = ("Retention is the question this quarter. " + "The installed base shrank while new sales held "
                   "up, and spend stayed where the plan put it. ") * 3

QUESTIONS = [
    {"theme": "Retention decomposition",
     "question": "How much of NRR (annualized) below 100.0% sits in GRR (annualized) below its 85.0% threshold?"},
    {"theme": "Retention decomposition",
     "question": "Which accounts make up GRR (annualized) against its 85.0% threshold?"},
    {"theme": "Burn and budget variance",
     "question": "Which cost lines carry net burn vs budget against its 15.0% threshold?"},
    {"theme": "Burn and budget variance",
     "question": "How much of net new ARR vs budget sits above its -20.0% threshold?"},
    {"theme": "Liquidity and runway",
     "question": "How far is runway at current burn from its 12.0 mo threshold?"},
    {"theme": "Liquidity and runway",
     "question": "Which months of runway at current burn, against 12.0 mo, depend on collections?"},
    {"theme": "Pipeline and sales efficiency",
     "question": "Which segments carry CAC payback against its 24.0 mo threshold?"},
    {"theme": "Definitions and assumptions",
     "question": "Which definition gives the burn multiple against its 2.00x threshold?"},
]


def summary_dict(headline="Retention is the main question for the board.", questions=None):
    """A valid v5 answer as a dictionary: headline, diagnosis, 3 risks, 8 themed questions."""
    point = {"title": "Steady base", "detail": "Customers stayed."}
    return {"headline": headline, "diagnosis": PLAIN_DIAGNOSIS, "risks": [point] * 3,
            "questions": list(questions if questions is not None else QUESTIONS)}
