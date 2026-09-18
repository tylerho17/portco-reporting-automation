"""Check config.yaml against a schema when it is loaded, with messages a person can act on.

Every setting the tool reads is in SETTINGS: what kind of value it takes, the range that makes
sense, whether it must be there, and an example line to copy. A problem message always says three
things: the key, what is wrong with it, and an example, e.g.

    config.yaml: rule_of_40_min must be a decimal from -1 to 1 (got 40). 40% is written 0.4.
    Example: rule_of_40_min: 0.40

Every problem in the file is listed at once, so fixing it takes one pass, not one run per mistake.
No new package: the checks are plain Python, and "Did you mean" uses difflib from the standard library.

Used by metrics.load_config (the whole file) and metrics.evaluate_flags (a config built in code).
"""

import difflib
import math
from collections import namedtuple
from pathlib import Path

import yaml

# kind: "decimal" (a ratio, 0.15 = 15%), "multiple" (x), "months", "whole" (a count), "true/false".
# low / high: the allowed range, ends included (None: no limit that side).
# required: False for a setting that has a default elsewhere (diff_runs.py's).
Setting = namedtuple("Setting", "kind low high required example meaning")

SETTINGS = {
    "nrr_min": Setting("decimal", 0, 2, True, "1.00", "the lowest NRR that passes (1.00 = 100%)"),
    "grr_min": Setting("decimal", 0, 1, True, "0.85", "the lowest GRR that passes; GRR can't pass 100%"),
    "burn_multiple_max": Setting("multiple", 0, 10, True, "2.0", "the highest burn multiple that passes"),
    "burn_over_budget_max": Setting("decimal", 0, 1, True, "0.15", "how far over the burn budget passes"),
    "runway_min_months": Setting("months", 1, 60, True, "12", "the shortest runway that passes"),
    "cac_payback_max_months": Setting("months", 1, 120, True, "24", "the longest CAC payback that passes"),
    "net_new_arr_vs_budget_min": Setting("decimal", -1, 1, True, "-0.20",
                                         "how far under the net new ARR plan passes"),
    "rule_of_40_min": Setting("decimal", -1, 1, True, "0.40", "the lowest Rule of 40 that passes (0.40 = 40%)"),
    "nrr_falling_pipeline_rising_flag": Setting("true/false", None, None, True, "true",
                                                "whether the combo rule is checked"),
    "combo_lookback_quarters": Setting("whole", 2, None, True, "3", "how many quarters the combo rule looks at"),
    "combo_min_nrr_drop": Setting("decimal", 0, 1, True, "0.01", "how far NRR must fall at each step (0.01 = 1 point)"),
    "diff_min_points": Setting("decimal", 0, 1, False, "0.05", "how far a percentage must move to be listed"),
    "diff_min_relative": Setting("decimal", 0, 1, False, "0.10", "how far anything else must move, as a share"),
}

# Why a setting has the limit it has, where the bare range doesn't say it.
WHY = {
    "combo_lookback_quarters": "the combo rule needs at least one quarter-to-quarter step",
}

FILE_NAME = "config.yaml"
KEY_VALUE_EXAMPLE = f"Example: nrr_min: {SETTINGS['nrr_min'].example}"


class ConfigError(ValueError):
    """config.yaml can't be used. A ValueError, so every caller that already catches one still does."""


# ---------------------------------------------------------------------------
# One setting
# ---------------------------------------------------------------------------

def example_text(key):
    """'Example: nrr_min: 1.00', the line to copy."""
    return f"Example: {key}: {SETTINGS[key].example}"


def got_text(value):
    """'(got 40)', '(got '15%')', or '(got nothing)' for a key with no value after the colon."""
    return "(got nothing)" if value is None else f"(got {value!r})"


def is_number(value):
    """True for an int or float that is a real, finite number (not true/false, NaN or infinity)."""
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


def in_range(value, setting):
    """True if value is inside the setting's range, ends included."""
    return ((setting.low is None or value >= setting.low)
            and (setting.high is None or value <= setting.high))


def rule_text(setting):
    """What the setting must be: 'a decimal from 0 to 2', 'a whole number of at least 2', 'true or false'."""
    if setting.kind == "true/false":
        return "true or false"
    if setting.kind == "whole":
        return f"a whole number of at least {setting.low}"
    noun = {"decimal": "a decimal", "multiple": "a number", "months": "a number of months"}[setting.kind]
    return f"{noun} from {setting.low:g} to {setting.high:g}"


def percent_hint(value, setting):
    """' 40% is written 0.4.' when a decimal setting got a percent typed as a whole number, else ''."""
    if setting.kind != "decimal" or not is_number(value) or not in_range(value / 100, setting):
        return ""
    return f" {value:g}% is written {value / 100:g}."


def value_problem(key, value):
    """What is wrong with one setting's value, or None if it's fine."""
    setting = SETTINGS[key]
    if setting.kind == "true/false":
        fine = isinstance(value, bool)
    elif setting.kind == "whole":   # 3.0 is still 3 quarters
        fine = is_number(value) and float(value).is_integer() and in_range(value, setting)
    else:
        fine = is_number(value) and in_range(value, setting)
    if fine:
        return None
    why = f": {WHY[key]}" if key in WHY else ""
    return (f"{FILE_NAME}: {key} must be {rule_text(setting)} {got_text(value)}{why}."
            f"{percent_hint(value, setting)} {example_text(key)}")


# ---------------------------------------------------------------------------
# The whole config
# ---------------------------------------------------------------------------

def missing_problem(key):
    """A required setting that isn't in the file."""
    return f"{FILE_NAME}: {key} is missing ({SETTINGS[key].meaning}). Add it as its own line. {example_text(key)}"


def unknown_problem(key):
    """A key the tool doesn't read: most likely a typo of one it does."""
    close = difflib.get_close_matches(str(key), SETTINGS, n=1, cutoff=0.6)
    if close:
        return (f"{FILE_NAME}: {key} is not a setting this tool reads. Did you mean {close[0]}? "
                f"{example_text(close[0])}")
    return (f"{FILE_NAME}: {key} is not a setting this tool reads. Remove it, or use one of: "
            f"{', '.join(SETTINGS)}. {KEY_VALUE_EXAMPLE}")


def config_problems(config, whole_file=True):
    """Every problem with a config, in SETTINGS order, then unknown keys. [] means it's fine.

    whole_file=False checks only the settings the flags need (a config built in code may carry other
    keys, e.g. a test's): no unknown keys, no optional ones.
    """
    problems = []
    for key, setting in SETTINGS.items():
        if key not in config:
            if setting.required:
                problems.append(missing_problem(key))
        elif setting.required or whole_file:
            problems.append(value_problem(key, config[key]))
    if whole_file:
        problems += [unknown_problem(key) for key in config if key not in SETTINGS]
    return [problem for problem in problems if problem]


def check_config(config, whole_file=True, file_name=FILE_NAME):
    """Stop with every problem, one per line, naming the file they are in."""
    problems = config_problems(config, whole_file)
    if problems:
        raise ConfigError("\n".join(problem.replace(FILE_NAME, file_name, 1) for problem in problems))


# ---------------------------------------------------------------------------
# Reading the file
# ---------------------------------------------------------------------------

class KeyCountingLoader(yaml.SafeLoader):
    """PyYAML's safe loader, noting any key written twice (on its own it silently keeps the last)."""

    def construct_mapping(self, node, deep=False):
        keys = [self.construct_object(key_node, deep=deep) for key_node, _ in node.value]
        self.repeated = getattr(self, "repeated", []) + sorted({str(key) for key in keys if keys.count(key) > 1})
        return super().construct_mapping(node, deep)


def parse_yaml(text, file_name):
    """(the file's contents, keys written twice). A YAML syntax error stops naming its line."""
    loader = KeyCountingLoader(text)
    try:
        return loader.get_single_data(), getattr(loader, "repeated", [])
    except yaml.YAMLError as error:
        mark = getattr(error, "context_mark", None) or getattr(error, "problem_mark", None)
        where = f" near line {mark.line + 1}" if mark else ""
        raise ConfigError(f"{file_name} can't be read{where}: {getattr(error, 'problem', error)}. "
                          f"Each line should be a key, a colon and a value. {KEY_VALUE_EXAMPLE}") from None
    finally:
        loader.dispose()


def read_config(path):
    """config.yaml as a dictionary, after checking it against SETTINGS. Stops with ConfigError."""
    file_name = Path(path).name
    config, repeated = parse_yaml(Path(path).read_text(), file_name)
    if config is None:
        raise ConfigError(f"{file_name} is empty: it needs the flag thresholds. {KEY_VALUE_EXAMPLE}")
    if not isinstance(config, dict):
        raise ConfigError(f"{file_name} must be key: value lines, one setting per line. {KEY_VALUE_EXAMPLE}")
    if repeated:
        raise ConfigError("\n".join(f"{file_name}: {key} is written twice; keep only one, since only the "
                                    f"last would be used. {example_text(key) if key in SETTINGS else ''}".rstrip()
                                    for key in repeated))
    check_config(config, file_name=file_name)
    return config
