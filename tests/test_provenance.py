"""Unit tests for provenance.py: file hashes, the git commit, and whether a deck is approved.

Everything here works on files in a temporary folder, so the real output/ folder is untouched.
Run from the project folder:  pytest
"""

import json

import pytest

from provenance import (NOT_REVIEWED, UNKNOWN_COMMIT, approval_status, build_manifest, file_sha256, git_commit,
                        manifest_path, read_manifest, save_manifest)

# The SHA-256 of an empty file, from `shasum -a 256 /dev/null`. Written out by hand so the test
# proves the hash is the standard one, not just "whatever the code produced".
EMPTY_FILE_SHA256 = "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"


# ---------------------------------------------------------------------------
# Hashing the inputs
# ---------------------------------------------------------------------------

def test_the_hash_is_the_standard_sha256(tmp_path):
    empty = tmp_path / "empty.xlsx"
    empty.write_bytes(b"")
    assert file_sha256(empty) == EMPTY_FILE_SHA256


def test_the_hash_changes_when_the_bytes_change_and_not_otherwise(tmp_path):
    first, second = tmp_path / "a.xlsx", tmp_path / "b.xlsx"
    first.write_bytes(b"Q2 2026,27470")
    second.write_bytes(b"Q2 2026,27470")
    assert file_sha256(first) == file_sha256(second)  # same bytes, same hash, whatever the name

    second.write_bytes(b"Q2 2026,27471")  # one digit different
    assert file_sha256(first) != file_sha256(second)


def test_a_missing_file_says_so(tmp_path):
    with pytest.raises(OSError):
        file_sha256(tmp_path / "nothing.xlsx")


# ---------------------------------------------------------------------------
# The code the deck was built from
# ---------------------------------------------------------------------------

def test_git_commit_of_this_project():
    commit = git_commit()
    assert set(commit) <= {"commit", "uncommitted_changes"}
    assert isinstance(commit["uncommitted_changes"], bool)
    assert commit["commit"] == UNKNOWN_COMMIT or commit["commit"].isalnum()


def test_a_folder_that_is_not_a_repository_says_unknown(tmp_path):
    # A copy of the project without git still has to run; it just can't say which commit it is.
    assert git_commit(tmp_path) == {"commit": UNKNOWN_COMMIT, "uncommitted_changes": False}


# ---------------------------------------------------------------------------
# Reading and writing the manifest
# ---------------------------------------------------------------------------

def test_manifest_path_sits_next_to_the_other_outputs(tmp_path):
    assert manifest_path("data/northwind.xlsx", tmp_path) == tmp_path / "northwind_manifest.json"


def test_reading_a_manifest_that_is_missing_or_broken(tmp_path):
    assert read_manifest(tmp_path / "nothing.json") is None
    broken = tmp_path / "broken_manifest.json"
    broken.write_text("{not json")
    assert read_manifest(broken) is None


def manifest_for(tmp_path, approval=None, run_at="2026-09-17T14:03:11"):
    """A manifest for a made-up company, written to tmp_path. Returns (path, manifest)."""
    workbook, config = tmp_path / "testco.xlsx", tmp_path / "config.yaml"
    workbook.write_bytes(b"quarter,arr")
    config.write_bytes(b"nrr_min: 1.00")
    ai = {"model": "claude-sonnet-5", "prompt_version": "v4", "attempts": 1, "input_tokens": 100,
          "output_tokens": 50, "seconds": 1.0, "cost_usd": 0.0007, "validation": "passed"}
    manifest = build_manifest("Testco", workbook, config, "testco_board_pack.pptx", ai,
                              ai_text=True, approval=approval, run_at=run_at)
    return save_manifest(manifest_path(workbook, tmp_path), manifest), manifest


def test_a_manifest_records_every_input_and_output(tmp_path):
    path, manifest = manifest_for(tmp_path)
    saved = json.loads(path.read_text())
    assert saved == manifest
    assert saved["company"] == "Testco" and saved["run_at"] == "2026-09-17T14:03:11"
    assert saved["input"] == {"file": "testco.xlsx", "sha256": file_sha256(tmp_path / "testco.xlsx")}
    assert saved["config"] == {"file": "config.yaml", "sha256": file_sha256(tmp_path / "config.yaml")}
    assert set(saved["code"]) == {"commit", "uncommitted_changes"}
    assert saved["ai"]["model"] == "claude-sonnet-5" and saved["ai"]["prompt_version"] == "v4"
    assert saved["deck"] == {"file": "testco_board_pack.pptx", "ai_text": True, "status": NOT_REVIEWED}
    assert saved["approval"] is None


# ---------------------------------------------------------------------------
# Is this deck approved?
# ---------------------------------------------------------------------------

APPROVAL = {"reviewer": "Tyler Ho", "approved_at": "2026-09-17T15:00:00",
            "input_sha256": "input-hash", "config_sha256": "config-hash"}


def test_no_manifest_and_no_approval_mean_not_reviewed():
    assert approval_status(None, "input-hash", "config-hash") == (None, NOT_REVIEWED)
    assert approval_status({"approval": None}, "input-hash", "config-hash") == (None, NOT_REVIEWED)


def test_an_approval_for_todays_files_holds():
    approval, why = approval_status({"approval": APPROVAL}, "input-hash", "config-hash")
    assert approval == APPROVAL and why is None


def test_a_changed_workbook_sends_the_deck_back_to_draft():
    approval, why = approval_status({"approval": APPROVAL}, "a-new-input-hash", "config-hash")
    assert approval is None and "workbook has changed" in why


def test_changed_thresholds_send_the_deck_back_to_draft():
    # A threshold edit changes which flags trip, so the reviewer approved something else.
    approval, why = approval_status({"approval": APPROVAL}, "input-hash", "a-new-config-hash")
    assert approval is None and "config.yaml has changed" in why


def test_an_approved_manifest_says_who_approved_it(tmp_path):
    workbook, config = tmp_path / "testco.xlsx", tmp_path / "config.yaml"
    workbook.write_bytes(b"quarter,arr")
    config.write_bytes(b"nrr_min: 1.00")
    approval = {"reviewer": "Tyler Ho", "approved_at": "2026-09-17T15:00:00",
                "input_sha256": file_sha256(workbook), "config_sha256": file_sha256(config)}
    manifest = build_manifest("Testco", workbook, config, "testco_board_pack.pptx", {}, ai_text=True,
                              approval=approval)
    assert manifest["deck"]["status"] == "approved by Tyler Ho on 2026-09-17T15:00:00"
