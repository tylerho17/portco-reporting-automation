"""Tests for the docs (no API calls): README.md, CLAUDE.md, STUDY_GUIDE.md and LOOM_SCRIPT.md stay in step with the code.

- compare_models.py rewrites the README's model comparison between two marker comments,
  so the markers must stay, and a rewrite must not touch the rest of the README.
- Every .py file the README, CLAUDE.md, STUDY_GUIDE.md or LOOM_SCRIPT.md names must exist
  (a renamed file breaks the docs quietly).
- Every function or class in STUDY_GUIDE.md's function tables must exist in the file its
  heading names (a removed function would otherwise stay in the guide as if it were real).
- INTERVIEW_PREP.md: every file it names exists, every `module.function` and
  `tests/file.py::test_name` it points to exists, and its sections come in interview order.

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
INTERVIEW_PREP = PROJECT_DIR / "INTERVIEW_PREP.md"


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


def files_named(text):
    """Every project file a piece of text names, e.g. 'metrics.py', 'config.yaml', 'data/northwind.xlsx'.

    Files under output/ are left out: they are made by a run and git ignores them, so a fresh
    clone doesn't have them.
    """
    pattern = r"[\w./-]+\.(?:py|md|ya?ml|toml|command|ini|txt|xlsx)\b"
    return {name for name in re.findall(pattern, text) if not name.startswith("output/")}


def functions_named(text):
    """Every `module.function` and `tests/file.py::test_name` a piece of text names, as (file, name)."""
    pairs = {(f"{module}.py", name) for module, name in re.findall(r"`(\w+)\.(\w+)", text)
             if name != "py"}  # `metrics.py` is a file name, not a function called "py"
    pairs |= set(re.findall(r"([\w/]+\.py)::(\w+)", text))
    return {(file, name) for file, name in pairs if (PROJECT_DIR / file).exists()}


def defined_or_assigned_in(file, name):
    """True if the file defines `name` as a function or class, or sets it as a constant (NAME = ...)."""
    assigned = re.compile(rf"^{name}\s*=", re.MULTILINE)
    return defined_in([file], name) or bool(assigned.search((PROJECT_DIR / file).read_text()))


def test_interview_prep_names_only_files_that_exist():
    # Like the study guide, a bare test file name ("test_metrics.py") counts if it is in tests/.
    missing = [name for name in files_named(INTERVIEW_PREP.read_text())
               if not (PROJECT_DIR / name).exists() and not (PROJECT_DIR / "tests" / name).exists()]
    assert missing == []


def test_interview_prep_names_only_functions_that_exist():
    missing = [f"{file}: {name}" for file, name in sorted(functions_named(INTERVIEW_PREP.read_text()))
               if not defined_or_assigned_in(file, name)]
    assert missing == []


def test_interview_prep_groups_the_questions_in_interview_order():
    headings = re.findall(r"^## (.+)$", INTERVIEW_PREP.read_text(), re.MULTILINE)
    expected = ["The project", "Design decisions", "The AI layer", "What broke",
                "Scale and risk", "Working method", "When you can't recall a detail"]
    assert [h for h in headings if h in expected] == expected


def test_study_guide_tables_cover_the_deck_files():
    # The deck files must have their own function tables, not just a mention.
    covered = set().union(*(files for files, _ in study_guide_tables()))
    assert {"build_deck.py", "make_template.py", "charts.py", "text_fit.py", "check_deck.py"} <= covered


def test_study_guide_tables_cover_approval_and_the_web_page():
    # The approval gate and the web page are explained function by function, like the rest.
    covered = set().union(*(files for files, _ in study_guide_tables()))
    assert {"provenance.py", "approve.py", "app.py"} <= covered


# The four docs a reader follows: the watermark is opt-in, the deck has 4 slides, and the web page exists.
USER_DOCS = (README, CLAUDE_MD, STUDY_GUIDE, LOOM_SCRIPT)


def watermark_lines(text):
    """Every line that talks about the watermark ("watermark", or the capitalised word "DRAFT" of its wording).

    Whole word only, so the constant AI_DRAFTED_LINE (slide 4's "AI-drafted ..." line) doesn't count.
    """
    return [line for line in text.splitlines() if re.search(r"[Ww]atermark|\bDRAFT\b", line)]


def test_docs_say_the_watermark_needs_draft():
    # Since polish Task 1 a deck has no watermark unless it is built with --draft; the footer says
    # "not reviewed" instead. A line about the watermark that doesn't say --draft reads as the old default.
    for doc in USER_DOCS:
        stale = [line.strip()[:80] for line in watermark_lines(doc.read_text()) if "--draft" not in line]
        assert stale == [], f"{doc.name} describes the watermark without --draft"


def test_docs_name_the_draft_option_the_web_page_and_its_launcher():
    for doc in USER_DOCS:
        missing = [name for name in ("--draft", "app.py", "run_app.command") if name not in doc.read_text()]
        assert missing == [], f"{doc.name} doesn't mention {missing}"


def test_readme_and_study_guide_show_both_footer_review_wordings():
    for doc in (README, STUDY_GUIDE):
        text = doc.read_text()
        assert "AI-drafted | not reviewed" in text and "AI-drafted | reviewed by" in text, doc.name


def test_docs_count_four_slides():
    for doc in USER_DOCS:
        assert not re.search(r"(?i)\b(5|five)[- ]slide|\bslide 5\b", doc.read_text()), doc.name


def test_study_guide_names_only_functions_that_exist():
    missing = [f"{name} ({', '.join(sorted(files))})" for files, names in study_guide_tables()
               for name in sorted(names) if not defined_in(files, name)]
    assert missing == []
