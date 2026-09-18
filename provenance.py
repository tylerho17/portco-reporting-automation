"""Where a deck came from, and whether a human has approved it.

Every run writes output/<company>_manifest.json: which workbook (by SHA-256 hash), which thresholds
(by hash), which code (git commit), which model and prompt wording, when it ran, what it cost, and
whether the deck carries Claude's text or the "AI summary unavailable" placeholder. Someone holding
a printed slide can trace it back to the exact inputs that produced it.

The manifest also holds the approval: a deck is watermarked "DRAFT - NOT REVIEWED" until a person
records their name with approve.py. `approval_status` is the only place that decides whether an
approval still counts, so main.py, build_deck.py and approve.py can never disagree about it. An
approval is void as soon as the workbook or config.yaml changes, because the reviewer approved what
the deck said, and either change can change that.

Nothing here calls an API or needs one.
"""

import hashlib
import json
import subprocess
from datetime import datetime
from pathlib import Path

PROJECT_DIR = Path(__file__).parent
READ_CHUNK = 65536          # hash big files a piece at a time instead of loading them whole
UNKNOWN_COMMIT = "unknown"  # a copy of the project without git still runs; it just can't say which commit
NOT_REVIEWED = "DRAFT - NOT REVIEWED"   # the watermark, and the deck status when nobody has approved


# ---------------------------------------------------------------------------
# Hashing the inputs, and naming the code
# ---------------------------------------------------------------------------

def file_sha256(path):
    """The SHA-256 hash of a file: the same bytes always give the same hash, one byte different never does."""
    digest = hashlib.sha256()
    with open(path, "rb") as file:
        while chunk := file.read(READ_CHUNK):
            digest.update(chunk)
    return digest.hexdigest()


def git_output(arguments, folder):
    """Run a read-only git command in `folder` and return its output, or None if git can't answer."""
    try:
        finished = subprocess.run(["git", *arguments], cwd=folder, capture_output=True, text=True, timeout=10)
    except (OSError, subprocess.SubprocessError):  # no git installed, or it didn't finish
        return None
    return finished.stdout.strip() if finished.returncode == 0 else None


def git_commit(folder=PROJECT_DIR):
    """Which commit built this deck, and whether the folder had uncommitted changes at the time.

    "Uncommitted changes" matters for an audit trail: the commit alone wouldn't describe the code
    that actually ran.
    """
    commit = git_output(["rev-parse", "--short", "HEAD"], folder)
    if commit is None:
        return {"commit": UNKNOWN_COMMIT, "uncommitted_changes": False}
    return {"commit": commit, "uncommitted_changes": bool(git_output(["status", "--porcelain"], folder))}


def timestamp():
    """Now, to the second, as text ("2026-09-17T14:03:11")."""
    return datetime.now().isoformat(timespec="seconds")


# ---------------------------------------------------------------------------
# Reading and writing the manifest
# ---------------------------------------------------------------------------

def manifest_path(workbook_path, output_dir):
    """data/northwind.xlsx -> output/northwind_manifest.json (beside the deck and the metrics workbook)."""
    return Path(output_dir) / f"{Path(workbook_path).stem}_manifest.json"


def read_manifest(path):
    """The saved manifest, or None if there isn't one or it can't be read.

    A missing manifest is normal (the first run), and an unreadable one must not stop a run: it only
    means this deck has nothing proving it was approved, so it stays a draft.
    """
    try:
        saved = json.loads(Path(path).read_text())
    except (OSError, json.JSONDecodeError):
        return None
    return saved if isinstance(saved, dict) else None


def approval_status(manifest, input_hash, config_hash):
    """(the approval, None) if this deck is approved, else (None, why it isn't).

    The reviewer approved a deck built from one workbook and one set of thresholds. If either has
    changed since, the deck in front of them isn't the deck they signed off, so it goes back to draft.
    """
    approval = (manifest or {}).get("approval")
    if not approval:
        return None, NOT_REVIEWED
    if approval.get("input_sha256") != input_hash:
        return None, f"{NOT_REVIEWED}: the workbook has changed since it was approved"
    if approval.get("config_sha256") != config_hash:
        return None, f"{NOT_REVIEWED}: config.yaml has changed since it was approved"
    return approval, None


def deck_status(approval):
    """What the deck says about review: who approved it and when, or that nobody has."""
    if not approval:
        return NOT_REVIEWED
    return f"approved by {approval['reviewer']} on {approval['approved_at']}"


def build_manifest(company, workbook_path, config_path, deck_file, ai, ai_text, approval=None, run_at=None):
    """Everything needed to trace one deck back to its inputs.

    `ai` holds the model, prompt version, attempts, tokens, seconds, cost and validation result
    (empty when there was no AI step). `approval` is carried over from the last manifest; whether it
    still counts is decided by approval_status, never assumed.
    """
    input_hash, config_hash = file_sha256(workbook_path), file_sha256(config_path)
    still_valid, _ = approval_status({"approval": approval}, input_hash, config_hash)
    return {
        "company": company,
        "run_at": run_at or timestamp(),
        "input": {"file": Path(workbook_path).name, "sha256": input_hash},
        "config": {"file": Path(config_path).name, "sha256": config_hash},
        "code": git_commit(),
        "ai": ai,
        "deck": {"file": deck_file, "ai_text": ai_text, "status": deck_status(still_valid)},
        "approval": approval,
    }


def save_manifest(path, manifest):
    """Write the manifest as readable JSON (it's meant to be opened and checked by a person)."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(manifest, indent=2, ensure_ascii=False))
    return path
