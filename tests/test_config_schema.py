"""Tests for config_schema.py (Task 13): config.yaml is checked against a schema when it is loaded.

Every problem message must name the key, say what is wrong, and show an example line to copy.
The tests write their own config files in a temporary folder; the project's config.yaml is only
read, never edited. Run from the project folder:  pytest
"""

import math

import pytest
import yaml
from streamlit.testing.v1 import AppTest

import main
import metrics
from config_schema import SETTINGS, ConfigError, config_problems, is_number, read_config
from metrics import FLAG_THRESHOLDS, evaluate_flags, load_config, validate_config

# A config that passes, written out here so tuning config.yaml never breaks these tests.
GOOD = {
    "nrr_min": 1.00, "grr_min": 0.85, "burn_multiple_max": 2.0, "burn_over_budget_max": 0.15,
    "runway_min_months": 12, "cac_payback_max_months": 24, "net_new_arr_vs_budget_min": -0.20,
    "rule_of_40_min": 0.40, "nrr_falling_pipeline_rising_flag": True, "combo_lookback_quarters": 3,
    "combo_min_nrr_drop": 0.01,
}
REQUIRED = [key for key, setting in SETTINGS.items() if setting.required]
OPTIONAL = [key for key, setting in SETTINGS.items() if not setting.required]


def without(key):
    """GOOD with one key taken out."""
    return {name: value for name, value in GOOD.items() if name != key}


def one_problem(config):
    """The only problem config_problems finds (fails the test if there are none, or several)."""
    problems = config_problems(config)
    assert len(problems) == 1, problems
    return problems[0]


def write_config(tmp_path, text):
    """A config.yaml in the test's own folder."""
    path = tmp_path / "config.yaml"
    path.write_text(text)
    return path


def as_yaml(config):
    """{'nrr_min': 1.0} -> 'nrr_min: 1.0' lines, true/false the YAML way."""
    return "".join(f"{key}: {str(value).lower() if isinstance(value, bool) else value}\n"
                   for key, value in config.items())


# ---------------------------------------------------------------------------
# The schema itself
# ---------------------------------------------------------------------------

def test_the_real_config_passes_and_has_every_required_key():
    # Reads config.yaml (never edits it).
    assert config_problems(load_config()) == []
    assert set(REQUIRED) <= set(load_config())


def test_good_passes_and_names_exactly_the_required_keys():
    assert config_problems(GOOD) == []
    assert set(GOOD) == set(REQUIRED)


def test_every_flag_threshold_is_in_the_schema():
    # A new flag rule whose threshold the schema doesn't know would pass any value, or be "unknown".
    assert {key for _, key, _ in FLAG_THRESHOLDS} <= set(REQUIRED)


def test_the_diff_settings_are_optional():
    assert set(OPTIONAL) == {"diff_min_points", "diff_min_relative"}
    assert config_problems({**GOOD, "diff_min_points": 0.05, "diff_min_relative": 0.10}) == []


@pytest.mark.parametrize("key", list(SETTINGS))
def test_every_example_passes_its_own_check(key):
    # An example that fails its own rule would teach the wrong fix. Examples are the text a person types.
    assert config_problems({**GOOD, key: yaml.safe_load(SETTINGS[key].example)}) == []


# ---------------------------------------------------------------------------
# A missing key
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("key", REQUIRED)
def test_a_missing_key_is_named_with_an_example(key):
    problem = one_problem(without(key))
    assert problem.startswith(f"config.yaml: {key} is missing")
    assert f"Example: {key}: " in problem


def test_a_missing_optional_key_is_fine():
    assert config_problems(GOOD) == []


# ---------------------------------------------------------------------------
# A wrong type
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("key", [key for key in SETTINGS if SETTINGS[key].kind != "true/false"])
@pytest.mark.parametrize("bad", ["15%", "twelve", None, True, [1, 2]])
def test_a_number_setting_given_something_else_is_named(key, bad):
    problem = one_problem({**GOOD, key: bad})
    assert problem.startswith(f"config.yaml: {key} must be")
    assert f"(got {bad!r})" in problem or "(got nothing)" in problem
    assert f"Example: {key}: {SETTINGS[key].example}" in problem


@pytest.mark.parametrize("bad", [math.nan, math.inf, -math.inf])
def test_nan_and_infinity_are_not_numbers_here(bad):
    # YAML reads .nan and .inf as floats. Today's ranges reject them too; is_number is the check that
    # still would for a future setting with no upper limit (found by a planted bug).
    problem = one_problem({**GOOD, "nrr_min": bad})
    assert problem.startswith("config.yaml: nrr_min must be")
    assert not is_number(bad)


def test_is_number_takes_only_real_finite_numbers():
    assert is_number(3) and is_number(-0.2) and is_number(0)
    assert not any(is_number(value) for value in (True, False, None, "1", math.nan, math.inf))


@pytest.mark.parametrize("bad", ["yes", "true", 1, 0, None])
def test_the_combo_switch_must_be_true_or_false(bad):
    problem = one_problem({**GOOD, "nrr_falling_pipeline_rising_flag": bad})
    assert problem.startswith("config.yaml: nrr_falling_pipeline_rising_flag must be true or false")
    assert "Example: nrr_falling_pipeline_rising_flag: true" in problem


def test_a_whole_number_setting_rejects_a_fraction():
    problem = one_problem({**GOOD, "combo_lookback_quarters": 2.5})
    assert "combo_lookback_quarters must be a whole number of at least 2 (got 2.5)" in problem


def test_a_whole_number_written_with_a_decimal_point_passes():
    # YAML reads "3.0" as a float; it is still 3 quarters.
    assert config_problems({**GOOD, "combo_lookback_quarters": 3.0}) == []


# ---------------------------------------------------------------------------
# An out of range value
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("key, bad", [
    ("nrr_min", 2.5), ("nrr_min", -0.1), ("grr_min", 1.2), ("burn_multiple_max", -1),
    ("burn_multiple_max", 11), ("burn_over_budget_max", 1.5), ("burn_over_budget_max", -0.1),
    ("runway_min_months", 0.5), ("runway_min_months", 61), ("cac_payback_max_months", 0.5),
    ("cac_payback_max_months", 121), ("net_new_arr_vs_budget_min", -1.5), ("net_new_arr_vs_budget_min", 1.5),
    ("rule_of_40_min", 1.5), ("rule_of_40_min", -1.5), ("combo_min_nrr_drop", -0.01),
    ("combo_min_nrr_drop", 1.5), ("diff_min_points", -0.05), ("diff_min_relative", 2),
])
def test_an_out_of_range_value_is_named_with_its_range_and_an_example(key, bad):
    problem = one_problem({**GOOD, key: bad})
    assert problem.startswith(f"config.yaml: {key} must be")
    assert f"(got {bad!r})" in problem
    assert f"Example: {key}: {SETTINGS[key].example}" in problem


@pytest.mark.parametrize("lookback", [1, 0, -3])
def test_the_combo_lookback_must_be_at_least_2(lookback):
    # With 1 quarter there are no steps to compare, and "every step fell" would be true of nothing.
    problem = one_problem({**GOOD, "combo_lookback_quarters": lookback})
    assert f"combo_lookback_quarters must be a whole number of at least 2 (got {lookback})" in problem
    assert "quarter-to-quarter step" in problem


@pytest.mark.parametrize("key, bad, meant", [
    ("nrr_min", 100, "1"), ("rule_of_40_min", 40, "0.4"), ("burn_over_budget_max", 15, "0.15"),
    ("net_new_arr_vs_budget_min", -20, "-0.2"), ("combo_min_nrr_drop", 5, "0.05"),
])
def test_a_percent_typed_as_a_whole_number_gets_the_decimal_it_meant(key, bad, meant):
    # The likeliest slip: "40" for a 40% Rule of 40, when ratios are decimals.
    problem = one_problem({**GOOD, key: bad})
    assert f"{bad}% is written {meant}" in problem


def test_the_range_ends_themselves_pass():
    for key, setting in SETTINGS.items():
        for end in (setting.low, setting.high):
            if end is not None:
                assert config_problems({**GOOD, key: end}) == [], (key, end)


# ---------------------------------------------------------------------------
# An unknown key
# ---------------------------------------------------------------------------

def test_an_unknown_key_is_named_with_the_closest_real_one():
    problem = one_problem({**GOOD, "nrr_minimum": 1.0})
    assert problem.startswith("config.yaml: nrr_minimum is not a setting this tool reads")
    assert "Did you mean nrr_min?" in problem
    assert "Example: nrr_min: 1.00" in problem


def test_a_misspelt_key_reports_both_the_unknown_and_the_missing_one():
    config = {**without("runway_min_months"), "runway_min_month": 12}
    problems = config_problems(config)
    assert len(problems) == 2
    assert any("runway_min_month is not a setting" in problem and "runway_min_months" in problem
               for problem in problems)
    assert any(problem.startswith("config.yaml: runway_min_months is missing") for problem in problems)


def test_an_unknown_key_like_nothing_lists_the_settings():
    problem = one_problem({**GOOD, "colour_scheme": "navy"})
    assert "colour_scheme is not a setting this tool reads" in problem
    assert "Did you mean" not in problem
    assert "nrr_min" in problem   # the settings it does read


# ---------------------------------------------------------------------------
# Several problems at once; checks from code
# ---------------------------------------------------------------------------

def test_every_problem_is_listed_at_once_in_the_schema_order():
    config = {**without("grr_min"), "nrr_min": "100%", "combo_lookback_quarters": 1, "extra": 1}
    problems = config_problems(config)
    assert [problem.split()[1] for problem in problems] == [
        "nrr_min", "grr_min", "combo_lookback_quarters", "extra"]


def test_validate_config_raises_one_error_with_every_problem_one_per_line():
    config = {**GOOD, "nrr_min": 100, "grr_min": "85%"}
    with pytest.raises(ConfigError) as caught:
        validate_config(config)
    assert str(caught.value).splitlines() == config_problems(config)   # run together, they'd be unreadable
    assert len(config_problems(config)) == 2
    assert isinstance(caught.value, ValueError)   # every caller that catches ValueError still does


def test_code_built_configs_are_checked_without_the_file_only_rules():
    # evaluate_flags checks what the flags use, but a config built in code may carry other keys.
    from test_metrics import full_actuals
    actuals = full_actuals()
    table = metrics.compute_metrics(actuals)
    reasons = metrics.metric_reasons(actuals, table)
    evaluate_flags(table, reasons, {**GOOD, "diff_min_points": "5%", "note": "x"})   # no error
    with pytest.raises(ConfigError, match="grr_min is missing"):
        evaluate_flags(table, reasons, without("grr_min"))


# ---------------------------------------------------------------------------
# Reading the file
# ---------------------------------------------------------------------------

def test_load_config_reads_a_good_file(tmp_path):
    assert load_config(write_config(tmp_path, as_yaml(GOOD))) == GOOD


def test_load_config_stops_on_a_bad_file_naming_the_key(tmp_path):
    path = write_config(tmp_path, as_yaml({**GOOD, "rule_of_40_min": 40}))
    with pytest.raises(ConfigError, match="rule_of_40_min must be"):
        load_config(path)


def test_an_empty_file_says_so(tmp_path):
    with pytest.raises(ConfigError, match="config.yaml is empty") as caught:
        read_config(write_config(tmp_path, "# only a comment\n"))
    assert "Example: nrr_min: 1.00" in str(caught.value)


def test_a_file_that_is_not_key_value_lines_says_so(tmp_path):
    with pytest.raises(ConfigError, match="key: value"):
        read_config(write_config(tmp_path, "- nrr_min\n- grr_min\n"))


def test_a_yaml_syntax_error_names_its_line(tmp_path):
    with pytest.raises(ConfigError, match="line 2") as caught:
        read_config(write_config(tmp_path, "nrr_min: 1.00\ngrr_min: [0.85\n"))
    assert "Example: nrr_min: 1.00" in str(caught.value)


def test_a_key_written_twice_is_named(tmp_path):
    # PyYAML alone keeps the last one silently, so the threshold a reader sees first isn't the one used.
    text = as_yaml(GOOD) + "nrr_min: 0.90\n"
    with pytest.raises(ConfigError, match="nrr_min is written twice") as caught:
        read_config(write_config(tmp_path, text))
    assert "only one" in str(caught.value)


def test_the_file_named_is_the_file_read(tmp_path):
    # A message about a copy elsewhere must not send the reader to the project's config.yaml.
    path = tmp_path / "settings.yaml"
    path.write_text(as_yaml(without("nrr_min")))
    with pytest.raises(ConfigError, match="settings.yaml: nrr_min is missing"):
        load_config(path)


# ---------------------------------------------------------------------------
# Where a person sees the message
# ---------------------------------------------------------------------------

def test_main_prints_the_problem_and_exits_1_without_a_traceback(tmp_path, monkeypatch, capsys):
    monkeypatch.setattr(metrics, "CONFIG_PATH", write_config(tmp_path, as_yaml(without("runway_min_months"))))
    assert main.main(["--all", "--skip-ai"]) == 1
    printed = capsys.readouterr().out
    assert "config.yaml: runway_min_months is missing" in printed
    assert "Traceback" not in printed and "Summary saved" not in printed   # nothing was built


def render():
    """The portfolio page (AppTest runs this function as a Streamlit script)."""
    import tempfile
    from pathlib import Path

    import app
    app.main(Path(tempfile.mkdtemp()), Path(tempfile.mkdtemp()))


def test_the_web_page_shows_the_problem_in_plain_words(tmp_path, monkeypatch):
    monkeypatch.setattr(metrics, "CONFIG_PATH", write_config(tmp_path, as_yaml({**GOOD, "nrr_min": "100%"})))
    test = AppTest.from_function(render, default_timeout=60).run()
    errors = [error.value for error in test.error]
    assert any("config.yaml: nrr_min must be" in error and "Example: nrr_min: 1.00" in error for error in errors)
    assert not test.exception
