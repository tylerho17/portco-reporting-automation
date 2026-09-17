"""Tests for the docs (no API calls): README.md, CLAUDE.md, STUDY_GUIDE.md and LOOM_SCRIPT.md stay in step with the code.

- compare_models.py rewrites the README's model comparison between two marker comments,
  so the markers must stay, and a rewrite must not touch the rest of the README.
- Every .py file the README, CLAUDE.md, STUDY_GUIDE.md or LOOM_SCRIPT.md names must exist
  (a renamed file breaks the docs quietly).
- Every function or class in STUDY_GUIDE.md's function tables must exist in the file its
  heading names (a removed function would otherwise stay in the guide as if it were real).

Run from the project folder:  pytest
"""

import re
from pathlib import Path

import compare_models

PROJECT_DIR = Path(__file__).parent.parent
README = PROJECT_DIR / "README.md"
CLAUDE_MD = PROJECT_DIR / "CLAUDE.md"
STUDY_GUIDE = PROJECT_DIR / "STUDY_GUIDE.md"
LOOM_SCRIPT = PROJECT_DIR / "LOOM_SCRIPT.md"


def python_files_named(text):
    """Every 'something.py' mentioned in a piece of text, e.g. 'main.py' or 'tests/test_docs.py'."""
    return set(re.findall(r"[\w/]+\.py\b", text))


def study_guide_tables():
    """[(files in the heading, names in the first column of its table)] for section 4 of STUDY_GUIDE.md.

    A heading like "### `main.py`: the batch runner" names the file; a row like
    "| `run_company(workbook_path, ...)` | ..." names a function in it. Headings without a
    .py file (config.yaml, tests/, small files) are skipped.
    """
    section = STUDY_GUIDE.read_text().split("## 4.", 1)[1].split("\n## 5.", 1)[0]
    tables = []
    for chunk in re.split(r"\n(?=#{3,4} )", section):
        heading, _, body = chunk.partition("\n")
        files = python_files_named(heading)
        rows = [line[1:].split("|")[0] for line in body.splitlines() if line.startswith("| `")]
        names = {name for cell in rows for name in re.findall(r"`(\w+)[(`]", cell)}
        if files and names:
            tables.append((files, names))
    return tables


def defined_in(files, name):
    """True if `def name` or `class name` appears in any of the files."""
    pattern = re.compile(rf"^\s*(def|class) {name}\b", re.MULTILINE)
    return any(pattern.search((PROJECT_DIR / file).read_text()) for file in files)


def test_readme_keeps_one_model_comparison_block():
    text = README.read_text()
    assert text.count(compare_models.README_START) == 1
    assert text.count(compare_models.README_END) == 1
    assert text.index(compare_models.README_START) < text.index(compare_models.README_END)


def test_rewriting_the_comparison_keeps_the_rest_of_the_readme(tmp_path, monkeypatch):
    # A copy in a temporary folder, so the real README is never touched.
    copy = tmp_path / "README.md"
    copy.write_text(README.read_text())
    monkeypatch.setattr(compare_models, "README_PATH", copy)

    compare_models.write_readme_section("NEW SECTION")

    before, after = README.read_text(), copy.read_text()
    start, end = compare_models.README_START, compare_models.README_END
    assert "NEW SECTION" in after
    assert after.split(start)[0] == before.split(start)[0]          # everything above the block
    assert after.split(end, 1)[1] == before.split(end, 1)[1]        # everything below the block


def test_readme_names_only_files_that_exist():
    missing = [name for name in python_files_named(README.read_text()) if not (PROJECT_DIR / name).exists()]
    assert missing == []


def test_claude_md_architecture_names_only_files_that_exist():
    text = CLAUDE_MD.read_text()
    architecture = text.split("## Architecture", 1)[1].split("\n## ", 1)[0]
    missing = [name for name in python_files_named(architecture) if not (PROJECT_DIR / name).exists()]
    assert missing == []


def test_study_guide_and_loom_script_name_only_files_that_exist():
    # The guide's tests table names test files without their folder ("test_metrics.py"), so tests/ counts too.
    for doc in (STUDY_GUIDE, LOOM_SCRIPT):
        missing = [name for name in python_files_named(doc.read_text())
                   if not (PROJECT_DIR / name).exists() and not (PROJECT_DIR / "tests" / name).exists()]
        assert missing == [], f"{doc.name} names files that don't exist"


def test_study_guide_tables_cover_the_deck_files():
    # The deck files must have their own function tables, not just a mention.
    covered = set().union(*(files for files, _ in study_guide_tables()))
    assert {"build_deck.py", "make_template.py", "charts.py", "text_fit.py", "check_deck.py"} <= covered


def test_study_guide_names_only_functions_that_exist():
    missing = [f"{name} ({', '.join(sorted(files))})" for files, names in study_guide_tables()
               for name in sorted(names) if not defined_in(files, name)]
    assert missing == []
