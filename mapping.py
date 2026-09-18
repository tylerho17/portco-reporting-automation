"""Column mapping: what to do with a header clean.py doesn't know.

A new company's workbook might say "Opening ARR" where clean.py expects starting_arr. For each header
that is neither a standard column nor in clean.HEADER_ALIASES, this file proposes the standard
column it most likely means, with a confidence and a reason in plain words. A person then confirms
or changes each one, and the answer is saved as mappings/<company>.yaml, so next quarter's workbook
(same headers, new numbers) runs with nobody asked.

Nothing is guessed silently. However confident a proposal is, clean.py stops until a person has
confirmed it (clean.UnconfirmedMappingError names the header and the proposal).

How a proposal is made (heuristics only, no API call):
1. Candidates: only the standard columns no known header has. If "Beginning ARR" is already there,
   "Opening ARR" can't be starting ARR too.
2. Name: the header's words, after dropping filler ("Total", "$K") and swapping the usual synonyms
   ("Opening" -> starting, "Plan" -> budget), against each column's words. The score is the best
   of word overlap and letter-by-letter similarity (so the typo "Revenu" still scores high).
3. Values: a value in the budget-only row rules out every actual column (clean.py would stop on it);
   ARR and cash must roll forward; gross profit can't be above revenue; and a column that is the
   only one still without a header gets points for that.
4. Each header gets its best candidate, strongest pair first, and no column is proposed twice.

Run:  python mapping.py data/acme.xlsx             show the proposals
      python mapping.py data/acme.xlsx --confirm   confirm or change each one, then save
"""

import argparse
import sys
from dataclasses import dataclass, field
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path

import yaml

from clean import (ACTUAL_COLUMNS, HEADER_ALIASES, LABEL_HEADER, STANDARD_COLUMNS, UnknownHeadersError, excel_column,
                   excel_row, find_kpi_sheet, is_blank, is_budget_label, name_headers, normalize_header, parse_number)
from provenance import file_sha256

PROJECT_DIR = Path(__file__).parent
MAPPINGS_DIR = PROJECT_DIR / "mappings"

SAMPLE_COUNT = 4          # values shown beside each proposal, as written in the workbook
MIN_CONFIDENCE = 0.4      # below this, no proposal: the person picks the column
HIGH, MEDIUM = 0.8, 0.55  # confidence levels as words: high, medium, low
MAX_CONFIDENCE = 0.99     # a heuristic is never certain
TOLERANCE = 0.5           # $K: how close a roll-forward must come (values are whole $K)

# Points the values add to (or take from) the name score.
ONLY_MISSING_POINTS = 0.3   # the only header and the only column left
ROLL_FORWARD_POINTS = 0.25  # ARR or cash rolls forward with this column in place
BELOW_REVENUE_POINTS = 0.1  # gross profit at or below revenue every quarter
BUDGET_ROW_POINTS = 0.05    # filled (or empty) in the budget-only row, as this kind of column is
BROKEN_POINTS = -0.3        # a roll-forward or gross-profit check that fails

# A few words in a row that mean one word, on the normalized header ("S&M" -> "s_m" -> "sm").
PHRASES = {"sales_and_marketing": "sm", "sales_marketing": "sm", "s_and_m": "sm", "s_m": "sm"}

# Words that say nothing about which column is meant ("Total Revenue ($K)" is revenue).
FILLER_WORDS = {"k", "m", "usd", "total", "of", "the", "and", "in", "qtr", "quarter", "quarterly",
                "balance", "amount", "sales", "000", "000s"}

# The usual other words for each part of a column name (FP&A vocabulary, lowercase).
SYNONYMS = {
    "opening": ["starting"], "beginning": ["starting"], "bop": ["starting"], "boq": ["starting"],
    "start": ["starting"],
    "closing": ["ending"], "end": ["ending"], "eop": ["ending"], "eoq": ["ending"],
    "upsell": ["expansion"], "upsells": ["expansion"],
    "downgrade": ["contraction"], "downgrades": ["contraction"], "downsell": ["contraction"],
    "downsells": ["contraction"],
    "churn": ["churned"], "lost": ["churned"], "cancelled": ["churned"], "cancellations": ["churned"],
    "rev": ["revenue"], "revenues": ["revenue"],
    "gp": ["gross", "profit"],
    "expense": ["spend"], "expenses": ["spend"], "spending": ["spend"], "cost": ["spend"], "costs": ["spend"],
    "logos": ["customers"], "logo": ["customers"], "clients": ["customers"], "client": ["customers"],
    "customer": ["customers"], "won": ["new"], "added": ["new"],
    "fte": ["headcount"], "ftes": ["headcount"], "employees": ["headcount"], "staff": ["headcount"],
    "pipe": ["pipeline"],
    "plan": ["budget"], "planned": ["budget"], "bud": ["budget"], "target": ["budget"], "budgeted": ["budget"],
}

# The two roll-forwards: ARR (starting + new + expansion - contraction - churn = next starting) and
# cash (last quarter's cash - this quarter's burn = this quarter's cash).
ARR_PARTS = ["starting_arr", "new_arr", "expansion_arr", "contraction_arr", "churned_arr"]
CASH_PARTS = ["ending_cash", "net_burn"]


@dataclass
class Proposal:
    """One unknown header, and what it most likely means."""
    header: str            # as written in the workbook, e.g. "Opening ARR"
    cell: str              # where it is, e.g. "B3"
    column: str | None     # the proposed standard column, or None (no proposal)
    confidence: float      # 0 to 0.99; 0 when there's no proposal
    reason: str            # why, in plain words
    samples: list          # the first few values under it, as written
    choices: list = field(default_factory=list)   # the columns it could mean, best first (for Change)


@dataclass
class ColumnValues:
    """One column's numbers: one per actual quarter (None when blank or unreadable), and the budget-only row."""
    actuals: list
    budget_filled: bool


# ---------------------------------------------------------------------------
# Names
# ---------------------------------------------------------------------------

def header_words(header):
    """'Total Revenue ($K)' -> ['revenue']; 'Opening ARR' -> ['starting', 'arr']; 'GP' -> ['gross', 'profit']."""
    text = f"_{normalize_header(header)}_"
    for phrase, word in PHRASES.items():
        text = text.replace(f"_{phrase}_", f"_{word}_")
    words = []
    for word in text.strip("_").split("_"):
        if word and word not in FILLER_WORDS:
            words += [w for w in SYNONYMS.get(word, [word]) if w not in words]   # "Headcount (FTE)" -> one word
    return words


def column_names():
    """{standard column: every known way of writing it, as words}: its own name and its HEADER_ALIASES."""
    names = {column: [header_words(column)] for column in STANDARD_COLUMNS}
    for alias, column in HEADER_ALIASES.items():
        names[column].append(header_words(alias))
    return names


def similarity(words, name):
    """How alike two word lists are, 0 to 1: the better of word overlap and letter-by-letter similarity."""
    if not words or not name:
        return 0.0
    overlap = 2 * len(set(words) & set(name)) / (len(set(words)) + len(set(name)))
    letters = SequenceMatcher(None, " ".join(words), " ".join(name)).ratio()
    return max(overlap, letters)


def name_score(words, column):
    """(score 0 to 1, the closest way of writing the column): how much the header reads like it."""
    return max((similarity(words, name), " ".join(name)) for name in column_names()[column])


# ---------------------------------------------------------------------------
# Values
# ---------------------------------------------------------------------------

def readable_number(value):
    """A cell as a number in $K, or None if it's blank or can't be read (the checks then skip it)."""
    try:
        number = parse_number(value)
    except ValueError:
        return None
    return None if number != number else number   # NaN (a blank) is the only value not equal to itself


def table_rows(sheet, header_row, label_position):
    """(actual row indexes, budget-only row index or None) under the header: rows with a label."""
    actual_rows, budget_row = [], None
    for row_index in range(header_row + 1, len(sheet)):
        label = sheet.iat[row_index, label_position]
        if is_blank(label):
            continue
        if is_budget_label(str(label)):
            budget_row = row_index
        else:
            actual_rows.append(row_index)
    return actual_rows, budget_row


def column_values(sheet, rows, position):
    """The ColumnValues of one column, given table_rows' (actual rows, budget row)."""
    actual_rows, budget_row = rows
    actuals = [readable_number(sheet.iat[row_index, position]) for row_index in actual_rows]
    budget_filled = budget_row is not None and not is_blank(sheet.iat[budget_row, position])
    return ColumnValues(actuals, budget_filled)


def sample_values(sheet, header_row, position):
    """The first SAMPLE_COUNT filled cells under a header, as written ('$1.2M' stays '$1.2M'; 850.0 shows as '850')."""
    samples = []
    for value in sheet.iloc[header_row + 1:, position]:
        if is_blank(value):
            continue
        if isinstance(value, float) and value.is_integer():
            value = int(value)
        samples.append(str(value).strip())
        if len(samples) == SAMPLE_COUNT:
            break
    return samples


def all_present(*values):
    """True if none of the values is None (a blank or unreadable cell)."""
    return all(value is not None for value in values)


def ties_out(results):
    """(points, word) from a list of True/False checks: none to judge -> 0; all true -> support; any false -> against."""
    if not results:
        return 0.0, None
    return (ROLL_FORWARD_POINTS, "fits") if all(results) else (BROKEN_POINTS, "breaks")


def arr_checks(parts):
    """True/False per pair of quarters: does starting ARR + the flows equal next quarter's starting ARR?"""
    start, new, expansion, contraction, churn = (parts[name] for name in ARR_PARTS)
    return [abs(start[i] + new[i] + expansion[i] - contraction[i] - churn[i] - start[i + 1]) <= TOLERANCE
            for i in range(len(start) - 1)
            if all_present(start[i], new[i], expansion[i], contraction[i], churn[i], start[i + 1])]


def cash_checks(parts):
    """True/False per pair of quarters: is last quarter's cash minus this quarter's burn this quarter's cash?"""
    cash, burn = parts["ending_cash"], parts["net_burn"]
    return [abs(cash[i - 1] - burn[i] - cash[i]) <= TOLERANCE
            for i in range(1, len(cash)) if all_present(cash[i - 1], burn[i], cash[i])]


def roll_forward_evidence(column, values, known):
    """(points, reason) from the ARR or cash roll-forward, with this column's values in place of `column`.

    Only when every other part of that roll-forward has a known header. known = {column: numbers}.
    """
    for parts, checks, what in ((ARR_PARTS, arr_checks, "ARR"), (CASH_PARTS, cash_checks, "Cash")):
        if column in parts and all(part in known for part in parts if part != column):
            points, word = ties_out(checks({**known, column: values.actuals}))
            if word == "fits":
                return points, f"{what} rolls forward exactly with it as {column}"
            if word == "breaks":
                return points, f"{what} doesn't roll forward with it as {column}"
    return 0.0, None


def gross_profit_evidence(column, values, known):
    """(points, reason): gross profit is at or below revenue in every quarter (either one may be the new header)."""
    pairs = {"gross_profit": "revenue", "revenue": "gross_profit"}
    if column not in pairs or pairs[column] not in known:
        return 0.0, None
    profit, revenue = ((values.actuals, known["revenue"]) if column == "gross_profit"
                       else (known["gross_profit"], values.actuals))
    results = [p <= r for p, r in zip(profit, revenue) if all_present(p, r)]
    if not results:
        return 0.0, None
    if all(results):
        return BELOW_REVENUE_POINTS, f"with it as {column}, gross profit is below revenue in every quarter"
    return BROKEN_POINTS, f"with it as {column}, gross profit would be above revenue in some quarter"


def budget_row_evidence(column, values):
    """(points or None, reason): None rules the column out.

    A value in the budget-only row means a budget column: clean.py stops if an actual column has one.
    An empty cell there leans towards an actual column but proves nothing (a budget can be blank).
    """
    if values.budget_filled:
        if column in ACTUAL_COLUMNS:
            return None, "it has a value in the budget-only row, which only budget columns may"
        return BUDGET_ROW_POINTS, "it has a value in the budget-only row, as budget columns do"
    return (BUDGET_ROW_POINTS, "it is empty in the budget-only row, as actual columns are") \
        if column in ACTUAL_COLUMNS else (0.0, None)


# ---------------------------------------------------------------------------
# Proposals
# ---------------------------------------------------------------------------

def score_candidate(header, column, values, known, only_one_left):
    """(confidence, reasons) for one header meaning one column, or (None, reason) if it can't."""
    points, why = budget_row_evidence(column, values)
    if points is None:
        return None, why
    score, closest = name_score(header_words(header), column)
    reasons = [f"name is {score:.0%} like '{closest}'"]
    for extra, reason in [(points, why), roll_forward_evidence(column, values, known),
                          gross_profit_evidence(column, values, known)]:
        score += extra
        if reason:
            reasons.append(reason)
    if only_one_left:
        score += ONLY_MISSING_POINTS
        reasons.append("it is the only input column still without a header")
    return round(min(max(score, 0.0), MAX_CONFIDENCE), 2), reasons


def best_first_assignment(scores):
    """{header position: column}: strongest (header, column) pair first; each header and column used once."""
    taken, chosen = set(), {}
    for (position, column), (confidence, _) in sorted(scores.items(), key=lambda item: -item[1][0]):
        if position not in chosen and column not in taken and confidence >= MIN_CONFIDENCE:
            chosen[position] = column
            taken.add(column)
    return chosen


def no_proposal_reason(open_columns):
    """Why a header has no proposal: no column left, or nothing close enough."""
    if not open_columns:
        return ("every input column already has a header, so this looks like an extra column: "
                "if it isn't one of the 16 inputs, delete it from the workbook")
    return f"no column scores {MIN_CONFIDENCE:.0%} or more: choose the column it means"


def propose(sheet, header_row, unknown, known_names):
    """A Proposal for each unknown header.

    unknown = the positions of the headers clean.py doesn't know; known_names = {position: name}
    for the ones it does (including the 'quarter' label column).
    """
    label_position = next(p for p, name in known_names.items() if name == LABEL_HEADER)
    rows = table_rows(sheet, header_row, label_position)
    known = {name: column_values(sheet, rows, p).actuals for p, name in known_names.items() if name != LABEL_HEADER}
    open_columns = [column for column in STANDARD_COLUMNS if column not in known]
    only_one_left = len(unknown) == 1 and len(open_columns) == 1

    headers = {position: str(sheet.iat[header_row, position]).strip() for position in unknown}
    values = {position: column_values(sheet, rows, position) for position in unknown}
    scores = {}
    for position in unknown:
        for column in open_columns:
            confidence, reasons = score_candidate(headers[position], column, values[position], known, only_one_left)
            if confidence is not None:   # a ruled-out column is never proposed or offered
                scores[(position, column)] = (confidence, reasons)
    chosen = best_first_assignment(scores)

    proposals = []
    for position in unknown:
        column = chosen.get(position)
        choices = sorted((c for c in open_columns if (position, c) in scores), key=lambda c: -scores[(position, c)][0])
        if column is None:
            confidence, reason = 0.0, no_proposal_reason(open_columns)
        else:
            confidence, reasons = scores[(position, column)]
            reason = "; ".join(reasons)
        proposals.append(Proposal(headers[position], f"{excel_column(position)}{excel_row(header_row)}", column,
                                  confidence, reason, sample_values(sheet, header_row, position), choices))
    return proposals


def review_workbook(workbook_path, mappings_dir=None):
    """The proposals for every header in the workbook that neither clean.py nor a saved mapping knows ([] if none)."""
    _, sheet, header_row = find_kpi_sheet(workbook_path)
    try:
        name_headers(sheet.iloc[header_row], header_row, confirmed_aliases(workbook_path, mappings_dir))
    except UnknownHeadersError as error:
        return propose(sheet, header_row, error.unknown, error.known)
    return []


# ---------------------------------------------------------------------------
# What the person sees
# ---------------------------------------------------------------------------

def confidence_text(confidence):
    """0.95 -> '95% (high)'; 0.6 -> '60% (medium)'; 0.41 -> '41% (low)'."""
    level = "high" if confidence >= HIGH else "medium" if confidence >= MEDIUM else "low"
    return f"{confidence:.0%} ({level})"


def proposal_line(proposal):
    """One header in the stop message: where it is, what it is, and the proposal (or why there is none)."""
    start = f"cell {proposal.cell} (header): Unknown column header {proposal.header!r}"
    if proposal.column is None:
        return f"{start}, no proposal: {proposal.reason}"
    return (f"{start}, proposed {proposal.column}, confidence {confidence_text(proposal.confidence)}: "
            f"{proposal.reason}")


def stop_message(proposals, workbook_path):
    """clean.py's stop: every unknown header and its proposal, then how to confirm."""
    lines = [proposal_line(proposal) for proposal in proposals]
    lines.append(f"Nothing is guessed: confirm or change each column with 'python mapping.py {shown(workbook_path)} "
                 f"--confirm' or the web page's Review mapping step (saved to {shown(mapping_path(workbook_path))} "
                 f"for next quarter). A column that isn't an input: delete it from the workbook.")
    return "\n".join(lines)


def shown(path):
    """A path as printed: relative to the project folder when it's inside it (mappings/...), else in full."""
    path = Path(path)
    return path.relative_to(PROJECT_DIR) if path.is_relative_to(PROJECT_DIR) else path


# ---------------------------------------------------------------------------
# The saved file: mappings/<company>.yaml
# ---------------------------------------------------------------------------

FILE_COMMENT = ("# Confirmed column mappings: workbook header -> standard input column (clean.py).\n"
                "# Written by mapping.py after a person confirmed each line. Headers match as clean.py\n"
                "# matches them (case, spaces and symbols ignored). Edit only to correct a mapping.\n")


def mapping_path(workbook_path, mappings_dir=None):
    """data/acme.xlsx -> mappings/acme.yaml (one file per company, named like its workbook)."""
    return Path(mappings_dir or MAPPINGS_DIR) / f"{Path(workbook_path).stem}.yaml"


def check_columns(columns, where, fix="choose one of them"):
    """Stop unless every mapping points at one of the 16 standard columns. `fix` says what to do instead."""
    for header, column in columns.items():
        if column not in STANDARD_COLUMNS:
            raise ValueError(f"{where}{header!r} -> {column!r} isn't one of the {len(STANDARD_COLUMNS)} input "
                             f"columns ({', '.join(STANDARD_COLUMNS)}): {fix}")


def confirm_again(workbook_path):
    """The command that rebuilds a company's mapping file, for the "what to do next" of a stop."""
    return f"python mapping.py {shown(workbook_path)} --confirm"


def read_mapping_file(path, workbook_path):
    """The mapping file as YAML. Broken YAML stops naming the file and line, never with a YAML library error."""
    try:
        return yaml.safe_load(path.read_text()) or {}
    except yaml.YAMLError as error:
        mark = getattr(error, "problem_mark", None)   # where the YAML library stopped reading, if it knows
        line = f", line {mark.line + 1}" if mark else ""
        raise ValueError(f"{shown(path)}{line}: this line isn't in 'workbook header: input column' form. Fix it, "
                         f"or delete the file and confirm the headers again with {confirm_again(workbook_path)}") \
            from None


def saved_columns(workbook_path, mappings_dir=None):
    """{header as written: column} from the company's mapping file, or {} if it has none. Stops on a bad file."""
    path = mapping_path(workbook_path, mappings_dir)
    if not path.exists():
        return {}
    saved = read_mapping_file(path, workbook_path)
    columns = saved.get("columns") if isinstance(saved, dict) else None
    if not isinstance(columns, dict):
        raise ValueError(f"{shown(path)} isn't a mapping file this tool can read: it needs a 'columns:' line, then "
                         f"one 'workbook header: input column' line per mapping. Fix it, or delete the file and "
                         f"confirm the headers again with {confirm_again(workbook_path)}")
    columns = {str(header): str(column) for header, column in columns.items()}
    check_columns(columns, f"{shown(path)}: ", f"fix that line in the file, or delete it and confirm the header "
                                               f"again with {confirm_again(workbook_path)}")
    return columns


def confirmed_aliases(workbook_path, mappings_dir=None):
    """{normalized header: column} from the saved mapping, ready for clean.py (like HEADER_ALIASES)."""
    return {normalize_header(header): column for header, column in saved_columns(workbook_path, mappings_dir).items()}


def save_mapping(workbook_path, choices, mappings_dir=None, now=None):
    """Save confirmed {header: column} choices to mappings/<company>.yaml, keeping earlier confirmations.

    Call this only with what a person confirmed. Returns the file's path.
    """
    check_columns(choices, "")
    path = mapping_path(workbook_path, mappings_dir)
    columns = {**saved_columns(workbook_path, mappings_dir), **choices}
    record = {"company": Path(workbook_path).stem,
              "confirmed_at": now or datetime.now().isoformat(timespec="seconds"),
              "columns": columns}
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(FILE_COMMENT + yaml.safe_dump(record, sort_keys=False, allow_unicode=True))
    return path


def mapping_sha256(workbook_path, mappings_dir=None):
    """The mapping file's SHA-256 hash (for the manifest), or None if the company has no mapping file."""
    path = mapping_path(workbook_path, mappings_dir)
    return file_sha256(path) if path.exists() else None


def mapping_record(workbook_path, mappings_dir=None):
    """What the manifest says about the mapping: its file and hash, or None if the company has none."""
    path = mapping_path(workbook_path, mappings_dir)
    return {"file": str(shown(path)), "sha256": file_sha256(path)} if path.exists() else None


# ---------------------------------------------------------------------------
# Command line
# ---------------------------------------------------------------------------

def print_proposals(proposals):
    """Each unknown header, its proposal, confidence, reason and sample values."""
    for proposal in proposals:
        proposed = proposal.column or "no proposal"
        confidence = confidence_text(proposal.confidence) if proposal.column else "-"
        print(f"  {proposal.cell} {proposal.header!r} -> {proposed}, confidence {confidence}")
        print(f"      why: {proposal.reason}")
        print(f"      values: {', '.join(proposal.samples) or 'none'}")


def ask_column(proposal, ask):
    """One header's confirmed column, asked until the answer is valid; None if the person quits ('q')."""
    default = f"Enter = {proposal.column}, " if proposal.column else ""
    while True:
        reply = ask(f"{proposal.header!r} means which column? ({default}or type one; q = stop, save nothing): ")
        reply = reply.strip()
        if reply.lower() == "q":
            return None
        if not reply and proposal.column:
            return proposal.column
        if reply in proposal.choices:
            return reply
        print(f"  {reply!r} isn't one of the columns still without a header: {', '.join(proposal.choices)}")


def confirm(workbook_path, proposals, ask):
    """Ask about every proposal, then save them all; nothing is saved if the person stops. Returns 0 or 1."""
    choices = {}
    for proposal in proposals:
        column = ask_column(proposal, ask)
        if column is None:
            print("Stopped: nothing saved.")
            return 1
        choices[proposal.header] = column
    path = save_mapping(workbook_path, choices)
    print(f"Saved {len(choices)} confirmed mapping(s) to {shown(path)}. Next quarter's workbook uses them unasked.")
    return 0


def main(argv=None, ask=input):
    parser = argparse.ArgumentParser(description="Propose, confirm and save mappings for unknown workbook headers.")
    parser.add_argument("workbook", help="the KPI workbook, e.g. data/acme.xlsx")
    parser.add_argument("--confirm", action="store_true", help="confirm or change each proposal, then save")
    args = parser.parse_args(argv)

    proposals = review_workbook(args.workbook)
    if not proposals:
        print("Nothing to confirm: every header is a known input column or already confirmed.")
        return 0
    print(f"{len(proposals)} header(s) in {args.workbook} need a confirmed mapping:")
    print_proposals(proposals)
    if args.confirm:
        return confirm(args.workbook, proposals, ask)
    print(f"Nothing saved. To confirm: python mapping.py {args.workbook} --confirm")
    return 1


if __name__ == "__main__":
    sys.exit(main())
