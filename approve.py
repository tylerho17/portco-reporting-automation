"""Record that a person has reviewed a company's board deck (the human approval gate).

Every deck's footer says "AI-drafted | not reviewed" until someone approves it here (and a deck built
with --draft is also stamped "DRAFT - NOT REVIEWED"). Approval is a name and a time written into
output/<company>_manifest.json, next to the hashes of the workbook and the thresholds the deck was
built from. Rebuilding the deck then makes the footer say "reviewed by NAME on DATE"; if the workbook,
config.yaml or the company's confirmed column mapping changes afterwards, it goes back to "not
reviewed" on its own, because what the reviewer read is no longer what the deck says
(provenance.approval_status).

This never builds or edits a deck. Approving and producing are two separate acts: the last word on
whether numbers reach a board belongs to a person, and this file is only the record of it.

Run: python approve.py northwind
     python approve.py northwind --reviewer "Tyler Ho"
"""

import argparse
import sys
from pathlib import Path

from mapping import mapping_sha256
from metrics import CONFIG_PATH
from provenance import (approval_status, deck_status, file_sha256, git_output, manifest_path, read_manifest,
                        save_manifest, timestamp)

PROJECT_DIR = Path(__file__).parent
DATA_DIR = PROJECT_DIR / "data"
OUTPUT_DIR = PROJECT_DIR / "output"

NO_REVIEWER = ("no reviewer name: pass --reviewer \"Your Name\", or set one for git with "
               "git config user.name \"Your Name\"")


def reviewer_name(given, folder=PROJECT_DIR):
    """The reviewer's name: the one given, else git's user.name. An approval needs a person's name on it."""
    name = (given or git_output(["config", "user.name"], folder) or "").strip()
    if not name:
        raise ValueError(NO_REVIEWER)
    return name


def reviewed_documents(manifest):
    """What this approval covers: the deck, and the memo if this run built one.

    An approval recorded before memos existed has no list, and memo.py then treats the memo as not
    reviewed: nobody can have read a memo that wasn't there.
    """
    return ["deck", "memo"] if manifest.get("memo") else ["deck"]


def approve(company, reviewer=None, data_dir=DATA_DIR, output_dir=OUTPUT_DIR, config_path=CONFIG_PATH, now=None):
    """Write the approval into the company's manifest and return the saved manifest.

    Stops (ValueError) if there is no manifest to approve, or if the workbook, thresholds or column
    mapping have changed since that run: the deck on disk was built from the old ones, so approving it would put
    a reviewer's name against numbers they never saw.
    """
    name = reviewer_name(reviewer)
    workbook = Path(data_dir) / f"{company}.xlsx"
    path = manifest_path(workbook, output_dir)
    manifest = read_manifest(path)
    if manifest is None:
        raise ValueError(f"no run to approve for {company!r}: {path.name} isn't there - run main.py first")

    input_hash, config_hash = file_sha256(workbook), file_sha256(config_path)
    if manifest["input"]["sha256"] != input_hash:
        raise ValueError(f"the workbook has changed since this deck was built - run main.py {workbook} again "
                         f"before approving")
    if manifest["config"]["sha256"] != config_hash:
        raise ValueError("config.yaml has changed since this deck was built - run main.py again before approving")
    mapping_hash = mapping_sha256(workbook)
    if (manifest.get("mapping") or {}).get("sha256") != mapping_hash:
        raise ValueError("the column mapping has changed since this deck was built - run main.py again before "
                         "approving")

    manifest["approval"] = {"reviewer": name, "approved_at": now or timestamp(),
                            "input_sha256": input_hash, "config_sha256": config_hash, "mapping_sha256": mapping_hash,
                            "documents": reviewed_documents(manifest)}
    manifest["deck"]["status"] = deck_status(approval_status(manifest, input_hash, config_hash, mapping_hash)[0])
    save_manifest(path, manifest)
    return manifest


def main(argv=None, data_dir=DATA_DIR, output_dir=OUTPUT_DIR, config_path=CONFIG_PATH):
    parser = argparse.ArgumentParser(description="Record that a person has reviewed a company's board deck.")
    parser.add_argument("company", help="company name as in data/, e.g. northwind")
    parser.add_argument("--reviewer", help="who reviewed it (default: git config user.name)")
    args = parser.parse_args(argv)

    try:
        manifest = approve(args.company, args.reviewer, data_dir, output_dir, config_path)
    except (ValueError, OSError) as error:
        print(f"Not approved: {error}")
        return 1

    approval = manifest["approval"]
    print(f"Approved {manifest['company']} by {approval['reviewer']} on {approval['approved_at']}")
    print(f"Rebuild the deck so its footer says reviewed: python build_deck.py data/{args.company}.xlsx")
    if "memo" in approval["documents"]:
        print(f"and the memo: python memo.py data/{args.company}.xlsx")
    return 0


if __name__ == "__main__":
    sys.exit(main())
