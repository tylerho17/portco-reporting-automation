"""The parts that keep a long batch run going: rate-limit retries, the spend ceiling, timeouts, workers, resume.

main.py's run_batch uses these. Nothing here calls the API itself or builds an output.

- RateLimitRetry wraps the Claude client. A rate-limited call (HTTP 429) waits and tries again: the
  wait the API asks for (its retry-after header), else 5, 10, 20, 40 s; never more than 60 s, and
  at most 4 retries per company. Only rate limits are retried here: any other API error still ends
  the AI step ("OK (AI failed)"), as before. (The SDK also makes its own quick retries first.)
- SpendMeter adds up what the AI calls cost, across companies running side by side (--max-cost).
- CompanyControl is how the batch tells a running company to stop (--timeout): the company checks
  it between steps and during a rate-limit wait, and stops at the next one.
- Every company is built in its own private folder (output/.staging/<company>_xxxx) and its files
  are moved into output/ only when it succeeds. So a company that fails or times out leaves its
  earlier outputs exactly as they were: never half old and half new.
- ThreadOutput keeps each company's printed lines together when several run at once (--workers).
- resume_problem says whether a company's saved outputs are still up to date (--resume).
"""

import contextlib
import io
import os
import shutil
import sys
import tempfile
import threading
from pathlib import Path

import anthropic

from build_deck import analysis_path, deck_path
from excel_output import output_path as metrics_path
from mapping import mapping_sha256
from provenance import file_sha256, manifest_path, read_manifest, save_manifest, timestamp

FIRST_WAIT_SECONDS = 5         # the first rate-limit wait when the API doesn't say how long; doubles each time
MAX_WAIT_SECONDS = 60          # no single wait is longer than this, whatever the API asks
MAX_RATE_LIMIT_RETRIES = 4     # per company: 5 + 10 + 20 + 40 s is over a minute of waiting, then give up
STAGING_FOLDER = ".staging"    # inside output/: each running company's private folder
MAX_BATCH_EVENTS = 100         # a manifest keeps the last 100 skips, timeouts, stops and failures
UP_TO_DATE = "outputs match the workbook and config.yaml"


# ---------------------------------------------------------------------------
# Rate limits
# ---------------------------------------------------------------------------

def rate_limit_wait(error, retry_number):
    """Seconds to wait before retry 1, 2, 3...: the API's retry-after if it sent one, else 5, 10, 20, 40."""
    try:
        seconds = float(error.response.headers.get("retry-after"))
    except (TypeError, ValueError):   # no header, or not a number of seconds
        seconds = FIRST_WAIT_SECONDS * 2 ** (retry_number - 1)
    return min(seconds, MAX_WAIT_SECONDS)


class RateLimitRetry:
    """Stands in for the Claude client: the same client.messages.parse(...), but a rate limit waits and tries again.

    `wait(seconds)` does the waiting (a batch passes CompanyControl.wait, so a timeout ends the wait).
    `waits` lists every wait, for the manifest and the summary table.
    """

    def __init__(self, client, wait):
        self.client, self.wait, self.waits = client, wait, []
        self.messages = self   # analyze.py calls client.messages.parse(...)

    def parse(self, **kwargs):
        """One API call, repeated after a wait each time it's rate limited, up to MAX_RATE_LIMIT_RETRIES times."""
        while True:
            try:
                return self.client.messages.parse(**kwargs)
            except anthropic.RateLimitError as error:
                if len(self.waits) >= MAX_RATE_LIMIT_RETRIES:
                    raise   # still rate limited: the AI step fails as any API error does
                seconds = rate_limit_wait(error, len(self.waits) + 1)
                self.waits.append(seconds)
                self.wait(seconds)


def rate_limit_note(waits):
    """'2 rate-limit retries (waited 15 s)' for the summary table, or None if there were none."""
    if not waits:
        return None
    retries = "retry" if len(waits) == 1 else "retries"
    return f"{len(waits)} rate-limit {retries} (waited {round(sum(waits), 2):g} s)"


# ---------------------------------------------------------------------------
# Spend, and stopping a company
# ---------------------------------------------------------------------------

class SpendMeter:
    """The AI spend so far in this batch, in dollars. Safe to add to from companies running side by side."""

    def __init__(self):
        self.lock, self.total = threading.Lock(), 0.0

    def add(self, usd):
        with self.lock:   # two companies finishing at once must not lose one of the additions
            self.total = round(self.total + usd, 6)


class Cancelled(Exception):
    """The batch gave up on this company (--timeout): it stops at its next step."""


class CompanyControl:
    """What a running company checks with the batch: has it been given up on, and where its spend goes."""

    def __init__(self, spend=None):
        self.cancel = threading.Event()   # set by the batch when the company runs out of time
        self.spend = spend

    def checkpoint(self):
        """Called between steps: stop here if the batch has given up on this company."""
        if self.cancel.is_set():
            raise Cancelled()

    def wait(self, seconds):
        """Wait (for a rate limit), but stop at once if the company is given up on meanwhile."""
        if self.cancel.wait(seconds):
            raise Cancelled()

    def add_cost(self, usd):
        """Count what an AI call cost toward the batch's --max-cost ceiling."""
        if self.spend is not None and usd:
            self.spend.add(usd)


# ---------------------------------------------------------------------------
# Printing: each company's lines together
# ---------------------------------------------------------------------------

class ThreadOutput(io.TextIOBase):
    """Stands in for sys.stdout: a company's print() goes to its own buffer, anything else straight through."""

    def __init__(self, real):
        self.real, self.local = real, threading.local()   # threading.local: each thread sees its own buffer

    def capture(self):
        """From now on, this thread's print() lines go to a buffer (returned) instead of the screen."""
        self.local.buffer = io.StringIO()
        return self.local.buffer

    def write(self, text):
        return (getattr(self.local, "buffer", None) or self.real).write(text)

    def flush(self):
        self.real.flush()


@contextlib.contextmanager
def routed_stdout():
    """Put a ThreadOutput in front of sys.stdout for the batch, and take it away afterwards."""
    real = sys.stdout
    sys.stdout = ThreadOutput(real)
    try:
        yield sys.stdout
    finally:
        sys.stdout = real


# ---------------------------------------------------------------------------
# Private folders: a company's files reach output/ only when it succeeds
# ---------------------------------------------------------------------------

def clear_old_stages(output_dir):
    """Delete output/.staging: anything there is left from a batch that was killed while running."""
    shutil.rmtree(Path(output_dir) / STAGING_FOLDER, ignore_errors=True)


def make_stage(workbook_path, output_dir):
    """A new private folder for one company's run, with its last manifest and analysis copied in.

    The manifest carries any approval and the batch events over; the analysis can be reused (--resume).
    """
    staging = Path(output_dir) / STAGING_FOLDER
    staging.mkdir(parents=True, exist_ok=True)
    stage = Path(tempfile.mkdtemp(prefix=f"{Path(workbook_path).stem}_", dir=staging))
    for path in (manifest_path(workbook_path, output_dir), analysis_path(workbook_path, output_dir)):
        if path.exists():
            shutil.copy2(path, stage / path.name)
    return stage


def commit_stage(stage, output_dir):
    """Move every file the company built into output/, the manifest last, then delete the private folder.

    Manifest last: if the run is killed halfway through the move, the old manifest still describes
    the old files, so --resume can't mistake a half-moved company for an up-to-date one.
    """
    files = sorted((path for path in Path(stage).rglob("*") if path.is_file()),
                   key=lambda path: path.name.endswith("_manifest.json"))   # False sorts first
    for path in files:
        target = Path(output_dir) / path.relative_to(stage)
        target.parent.mkdir(parents=True, exist_ok=True)
        os.replace(path, target)   # a move within one disk: the file is either old or new, never half written
    discard_stage(stage)


def discard_stage(stage):
    """Delete a private folder, and output/.staging too once nothing else is in it."""
    shutil.rmtree(stage, ignore_errors=True)
    with contextlib.suppress(OSError):   # another company is still using .staging
        Path(stage).parent.rmdir()


# ---------------------------------------------------------------------------
# Recording skips, timeouts, stops and failures in the company's manifest
# ---------------------------------------------------------------------------

def record_batch_event(workbook_path, output_dir, event, detail):
    """Add {when, event, detail} to the company's manifest. False if there is no manifest to add it to.

    A company that has never been built has no manifest; its outcome is still in the summary table
    and output/batch_manifest.json.
    """
    path = manifest_path(workbook_path, output_dir)
    manifest = read_manifest(path)
    if manifest is None:
        return False
    history = manifest.get("batch_events", []) + [{"at": timestamp(), "event": event, "detail": detail}]
    manifest["batch_events"] = history[-MAX_BATCH_EVENTS:]
    save_manifest(path, manifest)
    return True


# ---------------------------------------------------------------------------
# --resume: are this company's saved outputs still up to date?
# ---------------------------------------------------------------------------

def expected_files(workbook_path, output_dir, manifest):
    """Every output the last run wrote: the metrics workbook, the deck, and the memo if it built one."""
    output_dir = Path(output_dir)
    files = [metrics_path(workbook_path, output_dir), deck_path(workbook_path, output_dir)]
    return files + [output_dir / name for name in (manifest.get("memo") or {}).get("files", [])]


def hash_problem(workbook_path, config_path, manifest):
    """Why the saved outputs were built from other inputs (workbook, thresholds, column mapping), or None."""
    if manifest["input"]["sha256"] != file_sha256(workbook_path):
        return "the workbook has changed"
    if manifest["config"]["sha256"] != file_sha256(config_path):
        return "config.yaml has changed"
    if (manifest.get("mapping") or {}).get("sha256") != mapping_sha256(workbook_path):
        return "the column mapping has changed"
    return None


def deck_problem(manifest, want_ai, draft):
    """Why the saved deck isn't the one this run would build from the same inputs, or None."""
    deck, approval = manifest.get("deck"), manifest.get("approval")
    if not deck:
        return "the manifest has no deck"
    if want_ai and not deck.get("ai_text"):
        return "the deck has no AI text"
    if bool(deck.get("draft")) != draft:
        return "the deck was built with a different --draft setting"
    if approval and approval.get("approved_at", "") > manifest.get("run_at", ""):   # ISO times sort as text
        return "approved after the deck was built, so its footer still says not reviewed"
    return None


def resume_problem(workbook_path, output_dir, config_path, want_ai, draft):
    """None if the saved outputs are what this run would build, else why the company has to be rebuilt.

    Up to date means: built from today's workbook, config.yaml and column mapping (by hash), every
    output file still there, AI text on the deck if this run asks for it, the same --draft setting,
    and no approval recorded since the deck was built.
    """
    manifest = read_manifest(manifest_path(workbook_path, output_dir))
    if manifest is None:
        return "no earlier run"
    problem = hash_problem(workbook_path, config_path, manifest)
    missing = [path.name for path in expected_files(workbook_path, output_dir, manifest) if not path.exists()]
    if problem is None and missing:
        problem = f"{missing[0]} is missing"
    return problem or deck_problem(manifest, want_ai, draft)
