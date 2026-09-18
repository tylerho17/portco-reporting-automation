"""Put output/ back to a known good state before a demo (Task 12), keeping every saved AI analysis.

A practice run leaves things a demo shouldn't start with: an approval ("reviewed by ..."), a deck
built with --skip-ai or --draft, exports, a rollup, a half-finished batch. This script:
1. Takes a copy of every saved analysis (output/<company>_analysis.json). They cost money to make
   and are never deleted; if anything changed one, the copy is put back and the demo isn't ready.
2. Deletes every file the tool builds (decks, memos, metrics workbooks, manifests, exports, charts,
   the batch summary, the rollup, a killed batch's private folders). Any other file in output/ is
   left alone and listed.
3. Builds each demo company's workbook as it stood a quarter ago (check_diff.py's
   last_quarter_workbook), so the page's "What changed since the last run" has a run to compare with.
4. Builds today's files for every company in data/ exactly as the page's Generate all does with the AI
   box unticked: a saved analysis of exactly today's numbers goes on slide 4, otherwise "AI summary
   unavailable". The API is never called: the client passed in refuses any use.
5. Checks the result: every company's files match today's workbook, nobody has approved them, the
   demo company (Northwind) has AI text and something to compare with. Prints "Ready for the demo." or
   what to fix, and exits 1 if not ready. Notes (not problems): another company without AI text,
   a workbook added in practice, code with uncommitted changes (every footer's commit then has a "*").

Deleting the manifests clears approvals and the batch history on purpose: the demo starts with
"not reviewed" so it can show approving. Everything deleted is rebuilt from data/ and the saved analyses.
Run: python demo_reset.py
"""

import contextlib
import hashlib
import io
import shutil
import sys
import tempfile
from pathlib import Path

import make_data
import make_data_alderpeak
import make_data_fernhollow
from build_deck import CHART_FOLDER, PLACEHOLDER_TEXT, deck_path
from check_diff import last_quarter_workbook
from excel_output import output_path as excel_path
from export import export_paths
from main import BATCH_MANIFEST_PATH, DATA_DIR, OUTPUT_DIR, SUMMARY_CSV_PATH, company_name, find_workbooks, \
    run_company, shown_path
from memo import memo_paths
from metrics import load_config
from portfolio import generate_all, run_state
from provenance import NOT_REVIEWED, git_commit, manifest_path, read_manifest
from resilience import STAGING_FOLDER
from rollup import rollup_paths

DEMO_COMPANY = "northwind"   # the company DEMO.md walks through: it must have AI text
ANSWER_KEYS = {"northwind": make_data, "alderpeak": make_data_alderpeak, "fernhollow": make_data_fernhollow}

READY = "Ready for the demo."
NOT_READY = "Not ready for the demo: fix the lines marked ✗, then run python demo_reset.py again."
SHOWN_LEFT_ALONE = 5   # how many untouched files to name before "and N more"


class NoApiClient:
    """Stands in for the Anthropic client: any use fails, so a reset can never spend money."""

    def __getattr__(self, name):
        raise RuntimeError("demo_reset.py never calls the API")


# ---------------------------------------------------------------------------
# 1 and 2: the saved analyses, and the files the tool builds
# ---------------------------------------------------------------------------

def built_files(stem, output_dir):
    """Every file the tool builds for one company in output_dir. Never its saved analysis."""
    workbook = Path(f"{stem}.xlsx")   # the path helpers only use the name
    return [deck_path(workbook, output_dir), *memo_paths(workbook, output_dir), excel_path(workbook, output_dir),
            *export_paths(workbook, output_dir).values(), manifest_path(workbook, output_dir)]


def shared_built(output_dir):
    """The files and folders built for the whole portfolio: batch summary, rollup, charts, private folders."""
    output_dir = Path(output_dir)
    return [output_dir / SUMMARY_CSV_PATH.name, output_dir / BATCH_MANIFEST_PATH.name,
            *rollup_paths(output_dir).values(), output_dir / CHART_FOLDER, output_dir / STAGING_FOLDER]


def company_stems(data_dir, output_dir):
    """Every company with a workbook in data/ or a manifest in output/ (one added in practice and since deleted)."""
    from_data = {path.stem for path in find_workbooks(data_dir)}
    from_output = {path.name.removesuffix("_manifest.json") for path in Path(output_dir).glob("*_manifest.json")
                   if path.name != BATCH_MANIFEST_PATH.name}
    return sorted(from_data | from_output)


def analysis_hashes(output_dir):
    """{file name: SHA-256} of every saved analysis in output_dir."""
    return {path.name: hashlib.sha256(path.read_bytes()).hexdigest()
            for path in sorted(Path(output_dir).glob("*_analysis.json"))}


def recorded_approvals(stems, output_dir):
    """'Northwind: approved by Tyler Ho' for every manifest holding an approval (the reset clears them)."""
    approvals = []
    for stem in stems:
        approval = (read_manifest(manifest_path(Path(f"{stem}.xlsx"), output_dir)) or {}).get("approval")
        if approval:
            approvals.append(f"{company_name(Path(stem))}: approved by {approval['reviewer']}")
    return approvals


def remove_path(path):
    """Delete a file or a whole folder; True if there was one."""
    if path.is_dir():
        shutil.rmtree(path)
    elif path.exists():
        path.unlink()
    else:
        return False
    return True


def remove_built(stems, output_dir):
    """Delete every built file. Returns (how many were deleted, the files left alone, relative to output_dir)."""
    targets = [path for stem in stems for path in built_files(stem, output_dir)] + shared_built(output_dir)
    removed = sum(remove_path(path) for path in targets)
    kept = Path(output_dir).glob("*_analysis.json")
    left = set(Path(output_dir).rglob("*")) - set(kept)
    left_alone = sorted(str(path.relative_to(output_dir)) for path in left if path.is_file())
    return removed, left_alone


def changed_analyses(before, after):
    """The saved analyses (file names) that are missing or different after the reset."""
    return [name for name in sorted(before) if after.get(name) != before[name]]


def copy_files(source, target, names):
    """Copy the named files from one folder to another (the analyses' copy, and putting it back)."""
    for name in names:
        shutil.copy2(Path(source) / name, Path(target) / name)


# ---------------------------------------------------------------------------
# 3 and 4: last quarter's run, then today's
# ---------------------------------------------------------------------------

def quietly(function, *args, **kwargs):
    """Call a function with its printout hidden (each build prints a line per step)."""
    with contextlib.redirect_stdout(io.StringIO()):
        return function(*args, **kwargs)


def last_quarter_run(stem, config, output_dir):
    """Build the company's workbook as it stood a quarter ago, so today's run has an earlier run to compare with.

    No AI step and no saved analysis read: that quarter's numbers have none.
    """
    with tempfile.TemporaryDirectory() as folder:
        workbook = last_quarter_workbook(ANSWER_KEYS[stem], folder)
        quietly(run_company, workbook, config, True, client=NoApiClient(), output_dir=output_dir)


def rebuild(config, data_dir, output_dir):
    """Today's files for every company, as the page's Generate all with the AI box unticked. Returns its outcomes."""
    return quietly(generate_all, config, False, data_dir, output_dir, client=NoApiClient())


# ---------------------------------------------------------------------------
# 5: is it ready?
# ---------------------------------------------------------------------------

def ai_line(name, stem):
    """Why a company's deck has no AI text, and how to get it back."""
    return (f"{name}: no saved analysis of exactly today's numbers passes the deck's checks "
            f"(output/{stem}_analysis.json), so its deck and memo say \"{PLACEHOLDER_TEXT}\". One paid run "
            f"(about $0.09) makes a new one: python main.py data/{stem}.xlsx, then python demo_reset.py again.")


def company_readiness(workbook, output_dir):
    """(problems, notes) for one company whose files were built."""
    name, stem = company_name(workbook), workbook.stem
    state = run_state(workbook, output_dir)
    manifest = read_manifest(manifest_path(workbook, output_dir)) or {}
    problems, notes = [], []
    if not state["current"] or state["status"] != NOT_REVIEWED:
        problems.append(f"{name}: its files aren't a fresh, unreviewed build of today's workbook ({state['status']})")
    if not manifest.get("deck", {}).get("ai_text"):
        (problems if stem == DEMO_COMPANY else notes).append(ai_line(name, stem))
    if stem not in ANSWER_KEYS:
        notes.append(f"{name}: data/{stem}.xlsx isn't a demo company (added in a practice run?). It's on the "
                     f"page with nothing to compare with; delete the file to take it off.")
    elif manifest.get("previous_run") is None:
        problems.append(f"{name}: no earlier run to compare with, so What changed will be empty")
    return problems, notes


def code_note(commit):
    """A note when the code had uncommitted changes (provenance.git_commit's answer), else None.

    Every slide's footer then shows the commit with a "*", which a viewer may ask about.
    """
    if not commit["uncommitted_changes"]:
        return None
    return (f"The code has uncommitted changes, so every footer shows {commit['commit']}*: commit, then run "
            f"python demo_reset.py again for a clean footer.")


def readiness(outcomes, data_dir, output_dir):
    """(problems, notes): every company that failed to build, then each built company's own check."""
    problems = [outcome["message"] for outcome in outcomes if not outcome["ok"]]
    notes = []
    built = {outcome["company"] for outcome in outcomes if outcome["ok"]}
    for workbook in find_workbooks(data_dir):
        if company_name(workbook) in built:
            more_problems, more_notes = company_readiness(workbook, output_dir)
            problems, notes = problems + more_problems, notes + more_notes
    if DEMO_COMPANY not in {workbook.stem for workbook in find_workbooks(data_dir)}:
        problems.append(f"data/{DEMO_COMPANY}.xlsx is missing: DEMO.md walks through it (git checkout data/)")
    return problems, notes


# ---------------------------------------------------------------------------
# The whole reset, and the command line
# ---------------------------------------------------------------------------

def reset(data_dir=DATA_DIR, output_dir=OUTPUT_DIR):
    """Steps 1 to 5. Returns what happened: removed, left_alone, cleared_approvals, rebuilt, problems, notes."""
    data_dir, output_dir = Path(data_dir), Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    config = load_config()
    stems = company_stems(data_dir, output_dir)
    before = analysis_hashes(output_dir)
    with tempfile.TemporaryDirectory() as copies:
        copy_files(output_dir, copies, before)
        cleared = recorded_approvals(stems, output_dir)
        removed, left_alone = remove_built(stems, output_dir)
        for stem in ANSWER_KEYS:
            if (data_dir / f"{stem}.xlsx").exists():
                last_quarter_run(stem, config, output_dir)
        outcomes = rebuild(config, data_dir, output_dir)
        changed = changed_analyses(before, analysis_hashes(output_dir))
        copy_files(copies, output_dir, changed)   # put back any the rebuild touched
    problems, notes = readiness(outcomes, data_dir, output_dir)
    notes += [note for note in [code_note(git_commit())] if note]
    problems += [f"{name} changed during the reset: the copy taken before it was put back" for name in changed]
    return {"removed": removed, "left_alone": left_alone, "cleared_approvals": cleared, "kept": sorted(before),
            "rebuilt": [outcome["company"] for outcome in outcomes if outcome["ok"]],
            "problems": problems, "notes": notes}


def some_names(names):
    """'a, b, c' or 'a, b, c, d, e and 3 more'."""
    shown = ", ".join(names[:SHOWN_LEFT_ALONE])
    return shown + (f" and {len(names) - SHOWN_LEFT_ALONE} more" if len(names) > SHOWN_LEFT_ALONE else "")


def print_report(answer, output_dir):
    """What the reset did, each problem (✗) and note (!), then Ready or Not ready."""
    print(f"Demo reset: {shown_path(output_dir)}/")
    print(f"  ✓ Deleted {answer['removed']} built files and folders; kept {len(answer['kept'])} saved AI analyses")
    for line in answer["cleared_approvals"]:
        print(f"  ✓ Cleared an approval: {line}")
    print(f"  ✓ Built last quarter's run for each demo company, then today's files: {', '.join(answer['rebuilt'])}")
    if answer["left_alone"]:
        print(f"  - Left alone {len(answer['left_alone'])} files the tool didn't build: "
              f"{some_names(answer['left_alone'])}")
    for line in answer["notes"]:
        print(f"  ! {line}")
    for line in answer["problems"]:
        print(f"  ✗ {line}")
    print(NOT_READY if answer["problems"] else READY)


def main(data_dir=DATA_DIR, output_dir=OUTPUT_DIR):
    """Reset, print the report, and return the exit code: 0 ready, 1 not ready."""
    answer = reset(data_dir, output_dir)
    print_report(answer, output_dir)
    return 1 if answer["problems"] else 0


if __name__ == "__main__":
    sys.exit(main())
