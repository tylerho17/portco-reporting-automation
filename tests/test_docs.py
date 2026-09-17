"""Tests for the docs (no API calls): README.md and CLAUDE.md stay in step with the code.

- compare_models.py rewrites the README's model comparison between two marker comments,
  so the markers must stay, and a rewrite must not touch the rest of the README.
- Every .py file the README or CLAUDE.md names must exist (a renamed file breaks the docs quietly).

Run from the project folder:  pytest
"""

import re
from pathlib import Path

import compare_models

PROJECT_DIR = Path(__file__).parent.parent
README = PROJECT_DIR / "README.md"
CLAUDE_MD = PROJECT_DIR / "CLAUDE.md"


def python_files_named(text):
    """Every 'something.py' mentioned in a piece of text, e.g. 'main.py' or 'tests/test_docs.py'."""
    return set(re.findall(r"[\w/]+\.py\b", text))


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
