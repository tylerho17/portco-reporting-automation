"""Tests for main.py's command line (Task 14): --version, --list-companies, --help, and every exit code.

Written before the code. No test calls the API or builds a deck: --version and --list-companies
never do, and the batch is replaced by a stand-in wherever a test only needs main()'s exit code.
--list-companies runs on the real data/ and output/ once, and checks it wrote nothing.
Run from the project folder:  pytest
"""

import re
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

import build_deck
import main
from analyze import DEFAULT_MODEL, PROMPT_VERSION
from config_schema import ConfigError
from metrics import load_config

PROJECT_DIR = Path(__file__).parent.parent
README = PROJECT_DIR / "README.md"
COMPANIES = ["Alderpeak", "Fernhollow", "Northwind"]


def run_cli(*arguments):
    """Run `python main.py ...` as a user would: (exit code, what it printed)."""
    finished = subprocess.run([sys.executable, "main.py", *arguments], cwd=PROJECT_DIR, capture_output=True,
                              text=True, timeout=120)
    return finished.returncode, finished.stdout + finished.stderr


def no_batch(monkeypatch):
    """main() may not build anything: the batch, the summary and the API key check fail the test if reached."""
    monkeypatch.setattr(main, "run_batch", lambda *args, **kwargs: pytest.fail("the batch ran"))
    monkeypatch.setattr(main, "api_key_problem", lambda: pytest.fail("the API key was looked for"))


def files_in(folder):
    """Every file under folder with its size and modified time: to prove a command wrote nothing."""
    return {path: (path.stat().st_size, path.stat().st_mtime_ns) for path in Path(folder).rglob("*") if path.is_file()}


# ---------------------------------------------------------------------------
# --version
# ---------------------------------------------------------------------------

def test_version_names_the_code_the_model_and_the_prompt(monkeypatch, capsys):
    no_batch(monkeypatch)
    # A made-up commit, so the test can't pass on "unknown" in a folder with no git (found by a planted bug).
    monkeypatch.setattr(build_deck, "git_commit", lambda: {"commit": "abc1234", "uncommitted_changes": True})
    assert main.main(["--version"]) == main.EXIT_OK
    printed = capsys.readouterr().out
    assert "code abc1234*" in printed       # the same words every deck footer shows (build_deck.commit_text)
    assert DEFAULT_MODEL in printed
    assert f"prompt {PROMPT_VERSION}" in printed
    assert f"Python {sys.version_info.major}.{sys.version_info.minor}" in printed


def test_version_works_with_a_broken_config(monkeypatch, capsys):
    # Asking which version this is must never depend on the thresholds being right.
    no_batch(monkeypatch)
    monkeypatch.setattr(main, "load_config", lambda: pytest.fail("--version read config.yaml"))
    assert main.main(["--version"]) == main.EXIT_OK


def test_version_from_the_shell():
    code, printed = run_cli("--version")
    assert code == 0
    assert printed.startswith(main.PRODUCT_NAME)


# ---------------------------------------------------------------------------
# --list-companies
# ---------------------------------------------------------------------------

def test_list_companies_shows_every_company_in_data(monkeypatch, capsys):
    no_batch(monkeypatch)
    before = files_in(PROJECT_DIR / "output")
    assert main.main(["--list-companies"]) == main.EXIT_OK
    printed = capsys.readouterr().out
    for name in COMPANIES:
        assert name in printed
        assert f"data/{name.lower()}.xlsx" in printed
    assert "Q2 2026" in printed
    assert "7 of 9 flags tripped, 1 cannot evaluate" in printed   # Fernhollow, as the page says it
    assert files_in(PROJECT_DIR / "output") == before              # a listing writes nothing


@pytest.fixture
def small_portfolio(tmp_path):
    """(data folder, output folder): Northwind, a file that isn't a workbook, and an Excel lock file."""
    data_dir, output_dir = tmp_path / "data", tmp_path / "output"
    data_dir.mkdir()
    output_dir.mkdir()
    shutil.copy(PROJECT_DIR / "data" / "northwind.xlsx", data_dir)
    (data_dir / "notes.xlsx").write_text("not a workbook")
    (data_dir / "~$northwind.xlsx").write_text("Excel's lock file")
    return data_dir, output_dir


def test_list_companies_uses_the_pages_words(small_portfolio, capsys):
    import portfolio
    data_dir, output_dir = small_portfolio
    assert main.list_companies(load_config(), data_dir, output_dir) == main.EXIT_OK
    printed = capsys.readouterr().out
    northwind = next(row for row in portfolio.portfolio_rows(load_config(), data_dir, output_dir)
                     if row["stem"] == "northwind")
    line = next(line for line in printed.splitlines() if line.startswith("Northwind"))
    for words in (northwind["latest"], northwind["flags"], northwind["last_run"], northwind["status"]):
        assert words in line
    assert portfolio.NOT_GENERATED in line


def test_a_workbook_that_cannot_be_read_is_listed_with_why(small_portfolio, capsys):
    import portfolio
    data_dir, output_dir = small_portfolio
    assert main.list_companies(load_config(), data_dir, output_dir) == main.EXIT_OK   # the listing still worked
    printed = capsys.readouterr().out
    assert "Notes" in printed
    assert portfolio.NOT_A_WORKBOOK in printed
    assert "~$" not in printed    # Excel's lock file is not a company


def test_list_companies_says_how_many_and_how_to_build_them(small_portfolio, capsys):
    data_dir, output_dir = small_portfolio
    main.list_companies(load_config(), data_dir, output_dir)
    printed = capsys.readouterr().out
    assert "2 companies" in printed
    assert "python main.py --all" in printed


def test_one_company_is_1_company(tmp_path, capsys):
    # Found by a planted bug: "1 companies" passed every other test.
    shutil.copy(PROJECT_DIR / "data" / "northwind.xlsx", tmp_path)
    main.list_companies(load_config(), tmp_path, tmp_path)
    assert "1 company in" in capsys.readouterr().out


def test_list_companies_writes_nothing(small_portfolio):
    data_dir, output_dir = small_portfolio
    before = files_in(data_dir.parent)
    main.list_companies(load_config(), data_dir, output_dir)
    assert files_in(data_dir.parent) == before


def test_list_companies_with_no_workbooks_is_exit_1(tmp_path, capsys):
    assert main.list_companies(load_config(), tmp_path, tmp_path) == main.EXIT_PROBLEM
    assert "No .xlsx workbooks found" in capsys.readouterr().out


def test_list_companies_stops_on_a_broken_config(monkeypatch, capsys):
    # The flag counts need the thresholds: a broken config.yaml says why, as a batch would.
    no_batch(monkeypatch)
    monkeypatch.setattr(main, "load_config", lambda: (_ for _ in ()).throw(ConfigError("config.yaml: nrr_min is missing")))
    assert main.main(["--list-companies"]) == main.EXIT_PROBLEM
    assert "nrr_min is missing" in capsys.readouterr().out


def test_list_companies_reads_the_data_folder_at_run_time(monkeypatch, small_portfolio, capsys):
    data_dir, output_dir = small_portfolio
    monkeypatch.setattr(main, "DATA_DIR", data_dir)
    monkeypatch.setattr(main, "OUTPUT_DIR", output_dir)
    assert main.main(["--list-companies"]) == main.EXIT_OK
    assert "Alderpeak" not in capsys.readouterr().out


# ---------------------------------------------------------------------------
# --help
# ---------------------------------------------------------------------------

def help_text(capsys):
    with pytest.raises(SystemExit) as stop:
        main.parse_args(["--help"])
    assert stop.value.code == main.EXIT_OK
    return capsys.readouterr().out


def test_help_has_every_example_and_what_it_does(capsys):
    text = help_text(capsys)
    assert "Examples:" in text
    for command, meaning in main.EXAMPLES:
        assert command in text
        assert meaning in " ".join(text.split())


@pytest.mark.parametrize("command", [command for command, _ in main.EXAMPLES])
def test_every_help_example_is_a_valid_command(command):
    main.parse_args(command.split()[2:])   # a usage error would raise SystemExit


def test_help_lists_every_exit_code(capsys):
    text = " ".join(help_text(capsys).split())
    assert "Exit codes:" in text
    for code, meaning in main.EXIT_CODES.items():
        assert f"{code} {meaning}" in text


def test_help_explains_every_option_on_its_own_line(capsys):
    # On its own indented line in the option groups, not just named somewhere: the usage line and the
    # examples name --version too, so a hidden help line would otherwise pass (found by a planted bug).
    options_part = help_text(capsys).split("Examples:")[0]
    for option in ("workbook", "--all", "--skip-ai", "--draft", "--resume", "--max-cost", "--timeout", "--workers",
                   "--list-companies", "--version", "-h, --help"):
        assert re.search(rf"^  {re.escape(option)}\b", options_part, flags=re.MULTILINE), option
    assert "python main.py --list-companies | --version | --help" in options_part   # the usage line's second way


def test_help_from_the_shell():
    code, printed = run_cli("--help")
    assert code == 0
    assert "Examples:" in printed and "Exit codes:" in printed


# ---------------------------------------------------------------------------
# Usage errors: exit 2, and nothing runs
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("arguments", [
    [],                                               # neither a workbook nor --all
    ["data/northwind.xlsx", "--all"],                 # both
    ["--version", "--all"],
    ["--list-companies", "data/northwind.xlsx"],
    ["--list-companies", "--version"],
    ["--list-companies", "--skip-ai"],                # an option it would silently ignore
    ["--all", "--no-such-option"],
    ["--all", "--workers", "0"],
])
def test_a_wrong_command_is_exit_2_and_runs_nothing(monkeypatch, arguments, capsys):
    no_batch(monkeypatch)
    with pytest.raises(SystemExit) as stop:
        main.main(arguments)
    assert stop.value.code == main.EXIT_USAGE
    assert "usage:" in capsys.readouterr().err


def test_version_or_list_with_something_else_says_they_run_on_their_own(capsys):
    with pytest.raises(SystemExit):
        main.parse_args(["--version", "--draft"])
    assert "--version runs on its own" in capsys.readouterr().err


def test_a_wrong_command_from_the_shell_is_exit_2():
    code, printed = run_cli("--all", "data/northwind.xlsx")
    assert code == 2
    assert "give one workbook path, or --all" in printed


# ---------------------------------------------------------------------------
# Every exit code main() returns is a documented one
# ---------------------------------------------------------------------------

def test_the_exit_codes_are_0_1_2_and_130():
    assert (main.EXIT_OK, main.EXIT_PROBLEM, main.EXIT_USAGE, main.EXIT_STOPPED) == (0, 1, 2, 130)
    assert set(main.EXIT_CODES) == {0, 1, 2, 130}


def test_ctrl_c_is_exit_130_in_plain_words(monkeypatch, capsys):
    def interrupted(*args, **kwargs):
        raise KeyboardInterrupt
    monkeypatch.setattr(main, "run_batch", interrupted)
    monkeypatch.setattr(main, "write_summary_csv", lambda *args: pytest.fail("a summary was written after Ctrl+C"))
    assert main.main(["--all", "--skip-ai"]) == main.EXIT_STOPPED
    printed = capsys.readouterr().out
    assert printed.strip() == main.STOPPED_MESSAGE
    assert "Traceback" not in printed


def test_ctrl_c_while_listing_is_exit_130(monkeypatch):
    def interrupted(*args, **kwargs):
        raise KeyboardInterrupt
    monkeypatch.setattr(main, "list_companies", interrupted)
    assert main.main(["--list-companies"]) == main.EXIT_STOPPED


# ---------------------------------------------------------------------------
# README documents every exit code, and its main.py commands are real
# ---------------------------------------------------------------------------

def readme_exit_codes():
    """{code: meaning} from README's exit-code table (rows like '| `0` | meaning |')."""
    rows = re.findall(r"^\| `(\d+)` \| (.+?) \|$", README.read_text(), flags=re.MULTILINE)
    return {int(code): meaning for code, meaning in rows}


def test_readme_documents_every_exit_code_in_mains_words():
    assert readme_exit_codes() == main.EXIT_CODES


def readme_main_commands():
    """Every `python main.py ...` in README: in a code block (up to its # comment) or between backticks."""
    text = README.read_text()
    in_blocks = re.findall(r"^python main\.py([^#\n]*)", text, flags=re.MULTILINE)
    in_backticks = re.findall(r"`python main\.py([^`]*)`", text)
    # A bare `python main.py` is prose naming the tool ("exactly as `python main.py` does"), not a command.
    return sorted({arguments.strip() for arguments in in_blocks + in_backticks} - {""})


def test_readme_has_main_py_commands_to_check():
    assert "--list-companies" in readme_main_commands()
    assert "--version" in readme_main_commands()


@pytest.mark.parametrize("arguments", readme_main_commands())
def test_every_main_py_command_in_readme_is_valid(arguments):
    try:
        main.parse_args(arguments.split())
    except SystemExit as stop:   # --help prints and exits 0; a usage error exits 2
        assert stop.code == main.EXIT_OK, f"python main.py {arguments}"
