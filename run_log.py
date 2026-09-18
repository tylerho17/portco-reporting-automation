"""The run log (Task 15): output/logs/run_<timestamp>.jsonl, one line per step per company.

Writing (main.py and the web page's Generate buttons):
- A run is one `python main.py ...`, one Generate click or one Generate all click. Its file is named
  after the time its first line was written, e.g. output/logs/run_20260918-141503.jsonl.
- Each line is one JSON object: the run, where it came from (command line or web page), the
  company, the step, when it started, how many seconds it took, its result, and the error (or null).
- The steps, in order: clean, metrics and flags, what changed, metrics workbook, AI commentary,
  deck, memo, manifest. A step that raises is "failed" with the error, and no later step runs.
- The last line for each company is "whole company": what the batch made of it (built, failed,
  skipped, stopped, timed out), and how long it took from start to finish.
- A line is written the moment its step ends, so a run that crashes still leaves every line up to
  the crash. Companies built side by side (--workers) share one file; a lock keeps each line whole.
- A log that can't be written prints a warning once and never stops a company: the board pack
  matters more than its log.

Reading (the web page's Recent runs panel): recent_runs() summarises the newest logs. A line that
can't be read is counted and skipped, never shown as a traceback.

Every log is kept (they are small); the page shows the newest five.
"""

import json
import threading
import time
from collections import Counter
from contextlib import contextmanager
from datetime import datetime, timedelta
from pathlib import Path

from memo import no_em_dash
from resilience import Cancelled
from text_fit import TextDoesNotFitError

LOG_FOLDER = "logs"          # inside output/
FILE_PREFIX = "run_"
STAMP_FORMAT = "%Y%m%d-%H%M%S"

# Where a run came from.
COMMAND_LINE, WEB_PAGE = "command line", "web page"

# The steps, in the order run_company does them, and the batch's line for the whole company.
CLEAN, METRICS, CHANGES, EXCEL = "clean", "metrics and flags", "what changed", "metrics workbook"
AI, DECK, MEMO, MANIFEST = "AI commentary", "deck", "memo", "manifest"
WHOLE_COMPANY = "whole company"

# A step's result. The whole-company line uses main.py's outcome words instead (built, failed, ...).
OK, FAILED, SKIPPED, REUSED, STOPPED = "ok", "failed", "skipped", "reused", "stopped"
GIVEN_UP = "the batch gave up on this company (--timeout) before this step finished"
UNFINISHED = "unfinished"    # a company with no whole-company line: the run was killed (Ctrl+C, say)

RECENT_RUNS = 5              # how many runs the web page shows

# Errors caused by a bad input file (clean.py raises ValueError with a clear message; a missing
# or unreadable file raises OSError), or by text that can't fit a slide (text_fit.py names the slide
# and the box). Anything else is probably a bug. main.py and the web page use this list too.
INPUT_ERRORS = (ValueError, OSError, TextDoesNotFitError)


# ---------------------------------------------------------------------------
# How an error reads (Task 16): the log, main.py's Result column and its printout
# ---------------------------------------------------------------------------

def error_text(error):
    """An error in plain words: what is wrong, where, and what to do next. No Python error name.

    A bug keeps its Python name and message, so whoever fixes it knows what to look for.
    """
    if isinstance(error, PermissionError) and error.filename:   # e.g. the deck is open in PowerPoint
        return (f"can't open or save {error.filename}: if it's open in Excel, PowerPoint or Word, close it, "
                f"then run again")
    if isinstance(error, FileNotFoundError) and error.filename:
        return f"can't find {error.filename}: check the file name and folder, then run again"
    if isinstance(error, INPUT_ERRORS):   # clean.py's and text_fit.py's messages already say all three
        return str(error)
    return (f"unexpected problem, probably a bug in this tool rather than the workbook ({type(error).__name__}: "
            f"{error}): the Terminal window shows where it happened")


# ---------------------------------------------------------------------------
# Writing
# ---------------------------------------------------------------------------

def log_folder(output_dir):
    """output/logs, for any output folder (tests use a temporary one). None stays None (no_log)."""
    return None if output_dir is None else Path(output_dir) / LOG_FOLDER


def new_log_file(folder, started):
    """Create an empty log file named after `started`, and return its path.

    Two runs in the same second get run_<time>.jsonl and run_<time>_2.jsonl: the file is created
    with "x" (fails if it exists), so two runs can never pick the same name.
    """
    folder.mkdir(parents=True, exist_ok=True)
    stem, number = FILE_PREFIX + started.strftime(STAMP_FORMAT), 1
    while True:
        path = folder / (f"{stem}.jsonl" if number == 1 else f"{stem}_{number}.jsonl")
        try:
            path.open("x").close()
            return path
        except FileExistsError:
            number += 1


class RunLog:
    """One run's log file. The file is created with the first line, so a run that logs nothing leaves no file.

    clock and now are for tests: a stopwatch (seconds) and the time of day.
    """

    def __init__(self, output_dir, source, clock=time.monotonic, now=datetime.now):
        self.folder, self.source, self.clock, self.now = log_folder(output_dir), source, clock, now
        self.path = None
        self.lock = threading.Lock()   # companies side by side (--workers) write to the same file
        self.warned = False

    def write(self, company, step, seconds, result, error=None, started=None):
        """Add one line. Never raises: a problem is a warning.

        started = when the step began; by default `seconds` before now, since a line is written when
        its step ends (the whole-company line is written long after the company started).
        """
        if self.folder is None:   # no_log(): a run that is deliberately not logged
            return
        started = started or self.now() - timedelta(seconds=seconds)
        with self.lock:
            try:
                if self.path is None:
                    self.path = new_log_file(self.folder, started)
                line = {"run": self.path.stem, "source": self.source, "company": company, "step": step,
                        "started_at": started.isoformat(timespec="milliseconds"), "seconds": round(seconds, 3),
                        "result": result, "error": error}
                with open(self.path, "a") as file:   # opened for each line, so every line is on disk at once
                    file.write(json.dumps(line) + "\n")
            except OSError as problem:
                self.warn(problem)

    def warn(self, problem):
        """Say once that the log isn't being written; the run carries on."""
        if not self.warned:
            print(f"⚠ Run log not written ({problem}); the run carries on")
            self.warned = True

    def company(self, name):
        """The log for one company's steps."""
        return CompanyLog(self, name)


def no_log():
    """A run log that writes nothing: demo_reset.py's rebuild, so a demo starts with no runs listed."""
    return RunLog(None, "")


class StepOutcome:
    """What a step reports about itself, when "ok" isn't the whole story (the AI step: skipped, reused, failed)."""

    def __init__(self):
        self.result, self.error = OK, None


class CompanyLog:
    """One company's steps in a run. run_log=None logs nothing (a caller that doesn't want a log)."""

    def __init__(self, run_log, company):
        self.run_log, self.company = run_log, company
        self.clock = run_log.clock if run_log else time.monotonic

    @contextmanager
    def step(self, name):
        """Time the code inside `with`, then write its line: ok, the result it set, stopped, or failed.

        An error is written and then raised again, so the company still stops exactly as before.
        """
        started, start, outcome = self.now(), self.clock(), StepOutcome()
        try:
            yield outcome
        except Cancelled:
            self.write(name, self.clock() - start, STOPPED, GIVEN_UP, started)
            raise
        except Exception as error:
            self.write(name, self.clock() - start, FAILED, error_text(error), started)
            raise
        self.write(name, self.clock() - start, outcome.result, outcome.error, started)

    def now(self):
        """The time of day, from the run log's clock (tests fix it)."""
        return self.run_log.now() if self.run_log else datetime.now()

    def write(self, step, seconds, result, error=None, started=None):
        """One line for this company (nothing without a run log)."""
        if self.run_log is not None:
            self.run_log.write(self.company, step, seconds, result, error, started)


# ---------------------------------------------------------------------------
# Reading (the web page's Recent runs panel)
# ---------------------------------------------------------------------------

FIELDS = ("company", "step", "started_at", "seconds", "result", "error")


def read_lines(path):
    """(the readable lines as dicts, how many lines couldn't be read). Never raises on a bad line."""
    lines, unreadable = [], 0
    for text in Path(path).read_text(errors="replace").splitlines():
        try:
            line = json.loads(text)
            datetime.fromisoformat(line["started_at"])   # every field the panel uses must be there
            float(line["seconds"])
        except (ValueError, TypeError, KeyError, IndexError):
            unreadable += 1
            continue
        if isinstance(line, dict) and all(field in line for field in FIELDS):
            lines.append(line)
        else:
            unreadable += 1
    return lines, unreadable


def outcomes_text(lines):
    """'2 built, 1 failed': each company's whole-company result, in the order first seen ('unfinished' if none)."""
    companies = dict.fromkeys(line["company"] for line in lines)
    for line in lines:
        if line["step"] == WHOLE_COMPANY:
            companies[line["company"]] = line["result"]
    counts = Counter(result or UNFINISHED for result in companies.values())
    return ", ".join(f"{count} {result}" for result, count in counts.items())


def problem_lines(lines):
    """'Company, step: error' for every step that went wrong (the whole-company line repeats it, so it's left out)."""
    return [plain(f"{line['company']}, {line['step']}: {line['error']}") for line in lines
            if line["error"] and line["step"] != WHOLE_COMPANY]


def wall_seconds(lines):
    """From the first step's start to the last step's end: the run's length, even with companies side by side."""
    starts = [datetime.fromisoformat(line["started_at"]).timestamp() for line in lines]
    ends = [start + float(line["seconds"]) for start, line in zip(starts, lines)]
    return round(max(ends) - min(starts), 1)


def run_summary(path):
    """One run as the panel shows it, or None if no line of it can be read."""
    lines, unreadable = read_lines(path)
    if not lines:
        return None
    return {"run": Path(path).stem, "started": lines[0]["started_at"].replace("T", " ")[:16],
            "source": lines[0].get("source", ""), "companies": len({line["company"] for line in lines}),
            "outcomes": outcomes_text(lines), "seconds": wall_seconds(lines), "problems": problem_lines(lines),
            "lines": lines, "unreadable": unreadable}


def file_order(path):
    """Sort key for log files: the time in the name, then the _2, _3 of a second run in the same second."""
    stamp, _, number = path.stem[len(FILE_PREFIX):].partition("_")
    return stamp, int(number) if number.isdigit() else 1


def recent_runs(output_dir, limit=RECENT_RUNS):
    """The newest `limit` runs with at least one readable line, newest first ([] when there are none)."""
    files = sorted(log_folder(output_dir).glob(f"{FILE_PREFIX}*.jsonl"), key=file_order, reverse=True)
    runs = []
    for path in files:
        summary = run_summary(path)
        if summary:
            runs.append(summary)
        if len(runs) == limit:
            break
    return runs


def plain(text):
    """Text as the page shows it: no em dash (memo.no_em_dash), since errors can come from other libraries."""
    return no_em_dash(str(text))


def run_label(run):
    """'2026-09-18 14:15 · command line · 3 companies: 2 built, 1 failed · 6.0 s' (and any unreadable lines)."""
    companies = "1 company" if run["companies"] == 1 else f"{run['companies']} companies"
    label = f"{run['started']} · {run['source']} · {companies}: {run['outcomes']} · {run['seconds']:.1f} s"
    if run["unreadable"]:
        lines = "1 line" if run["unreadable"] == 1 else f"{run['unreadable']} lines"
        label += f" · {lines} of this log couldn't be read"
    return plain(label)


def step_rows(run):
    """The run's lines as table rows: company, step, seconds (one decimal), result, error (blank if none).

    Grouped by company, in the order each company first appears: companies built side by side
    (--workers) write their lines interleaved, which is hard to read. Each company's steps keep their order.
    """
    companies = dict.fromkeys(line["company"] for line in run["lines"])   # each once, in the order first seen
    order = {company: position for position, company in enumerate(companies)}
    lines = sorted(run["lines"], key=lambda line: order[line["company"]])   # sorted() keeps equal keys in order
    return [[line["company"], line["step"], f"{float(line['seconds']):.1f}", line["result"],
             plain(line["error"]) if line["error"] else ""] for line in lines]
