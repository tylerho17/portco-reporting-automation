"""Unit tests for diff_runs.py (Task 10): what changed since the last run.

The results of a run are small dicts written out here by hand, so every expected flip, move and
data gap can be checked by eye. A few tests build results from a small made-up table (the one
tests/test_memo.py uses) to prove what a real run records. No API calls.
Run from the project folder:  pytest
"""

import json

import pytest

import diff_runs
from build_deck import collect_deck_data
from diff_runs import (HEADING, NO_EARLIER_RUN, NOT_CHECKED, baseline, change_sections, changes_since_last_run,
                       compare, compared_with_text, move_settings, run_results, summary_text)
from metrics import METRIC_LABELS
from provenance import save_manifest
from test_memo import TEST_CONFIG, three_quarters

SETTINGS = {"min_points": 0.05, "min_relative": 0.10}


def metric(value, shown):
    return {"value": value, "shown": shown}


def results(quarter="Q2 2026", metrics=None, flags=None, gaps=None):
    """A run's results, as the manifest stores them. Defaults: NRR 100.0%, ending ARR 1,000, runway passed."""
    return {
        "quarter": quarter,
        "metrics": metrics if metrics is not None else {"nrr": metric(1.0, "100.0%"),
                                                        "ending_arr": metric(1000.0, "1,000")},
        "flags": flags if flags is not None else {"Runway at current burn": "Passed"},
        "gaps": gaps if gaps is not None else {},
    }


# ---------------------------------------------------------------------------
# How far a metric must move: defaults, config.yaml, the command line
# ---------------------------------------------------------------------------

def test_the_default_move_is_five_points_or_ten_percent():
    assert move_settings(TEST_CONFIG) == {"min_points": 0.05, "min_relative": 0.10}


def test_config_yaml_may_set_the_move_and_an_argument_beats_it():
    config = {**TEST_CONFIG, "diff_min_points": 0.02, "diff_min_relative": 0.25}
    assert move_settings(config) == {"min_points": 0.02, "min_relative": 0.25}
    assert move_settings(config, min_points=0.1) == {"min_points": 0.1, "min_relative": 0.25}


@pytest.mark.parametrize("bad", [-0.01, "5%", True, None])
def test_a_move_setting_that_is_not_a_number_of_zero_or_more_stops_with_its_name(bad):
    with pytest.raises(ValueError, match="diff_min_points"):
        move_settings({**TEST_CONFIG, "diff_min_points": bad})


# ---------------------------------------------------------------------------
# What a run records
# ---------------------------------------------------------------------------

def test_a_run_records_its_latest_quarter_every_metric_every_flag_and_every_gap():
    data = collect_deck_data("Testco", "testco.xlsx", three_quarters(blank="Q1 2026"), None, TEST_CONFIG)
    recorded = run_results(data)
    assert recorded["quarter"] == "Q2 2026"
    assert list(recorded["metrics"]) == list(METRIC_LABELS)
    assert list(recorded["flags"]) == [flag["flag"] for flag in data["flags"]]
    assert recorded["gaps"] == data["gaps"]
    # Every input 100: net new ARR = 100 + 100 - 100 - 100 = 0, burning 100 -> burn multiple is ∞ (ARR shrank).
    assert recorded["metrics"]["net_new_arr"] == {"value": 0.0, "shown": "0"}
    assert recorded["metrics"]["burn_multiple"] == {"value": None, "shown": "∞ (ARR shrank)"}
    # QoQ growth needs Q1 2026, which is blank: no number, and the words say why.
    assert recorded["metrics"]["arr_qoq"] == {"value": None, "shown": "data missing"}
    assert recorded["flags"]["Burn multiple"] == "Tripped"


def test_a_run_s_results_survive_a_trip_through_json_unchanged():
    # The manifest is JSON: results that changed on the way through would look like a change every run.
    recorded = run_results(collect_deck_data("Testco", "testco.xlsx", three_quarters(), None, TEST_CONFIG))
    assert json.loads(json.dumps(recorded)) == recorded


def test_values_are_rounded_so_float_noise_is_never_a_change():
    data = collect_deck_data("Testco", "testco.xlsx", three_quarters(), None, TEST_CONFIG)
    data["metrics"].loc["Q2 2026", "gross_margin"] = 0.1 + 0.2        # 0.30000000000000004
    assert run_results(data)["metrics"]["gross_margin"]["value"] == 0.3


# ---------------------------------------------------------------------------
# Which earlier run to compare with
# ---------------------------------------------------------------------------

def test_no_manifest_or_one_from_before_task_10_has_nothing_to_compare_with():
    assert baseline(None, results()) is None
    assert baseline({"run_at": "2026-06-18T09:00:00"}, results()) is None


def test_different_results_are_compared_with_the_manifest_s_own_run():
    manifest = {"run_at": "2026-06-18T09:00:00", "results": results(quarter="Q1 2026")}
    assert baseline(manifest, results()) == {"run_at": "2026-06-18T09:00:00", "results": results(quarter="Q1 2026")}


def test_a_rebuild_with_the_same_results_keeps_the_comparison_it_already_had():
    # After approve.py the deck is rebuilt from the same numbers: the memo must still say what
    # changed since last quarter, not "nothing changed since ten minutes ago".
    earlier = {"run_at": "2026-06-18T09:00:00", "results": results(quarter="Q1 2026")}
    manifest = {"run_at": "2026-09-18T10:00:00", "results": results(), "previous_run": earlier}
    assert baseline(manifest, results()) == earlier


def test_a_first_run_rebuilt_still_has_nothing_to_compare_with():
    assert baseline({"run_at": "2026-09-18T10:00:00", "results": results()}, results()) is None


def test_results_that_are_not_a_run_s_results_are_ignored():
    # A hand-edited manifest must not stop a run: it only has nothing to compare with.
    assert baseline({"run_at": "2026-06-18T09:00:00", "results": {"quarter": "Q1 2026"}}, results()) is None
    assert baseline({"run_at": "2026-06-18T09:00:00", "results": "garbage"}, results()) is None


# ---------------------------------------------------------------------------
# Flags that flipped
# ---------------------------------------------------------------------------

def test_a_flag_whose_status_changed_is_listed_with_both_statuses():
    before = results(flags={"Runway at current burn": "Passed", "NRR (annualized)": "Tripped"})
    after = results(flags={"Runway at current burn": "Tripped", "NRR (annualized)": "Tripped"})
    assert compare(before, after, SETTINGS)["flags"] == [("Runway at current burn", "Passed", "Tripped")]


def test_a_change_of_reason_counts_as_a_flip():
    before = results(flags={"Rule of 40": "Cannot evaluate: no prior period"})
    after = results(flags={"Rule of 40": "Cannot evaluate: missing input"})
    assert compare(before, after, SETTINGS)["flags"] == [
        ("Rule of 40", "Cannot evaluate: no prior period", "Cannot evaluate: missing input")]


def test_a_flag_only_one_run_checked_says_not_checked_for_the_other():
    # The combo rule can be switched off in config.yaml.
    before = results(flags={"Runway at current burn": "Passed", "NRR falling while pipeline rising": "Tripped"})
    after = results(flags={"Runway at current burn": "Passed"})
    assert compare(before, after, SETTINGS)["flags"] == [
        ("NRR falling while pipeline rising", "Tripped", NOT_CHECKED)]


# ---------------------------------------------------------------------------
# Metrics that moved more than the setting
# ---------------------------------------------------------------------------

def moves(before_metrics, after_metrics, settings=SETTINGS):
    return compare(results(metrics=before_metrics), results(metrics=after_metrics), settings)["moved"]


def test_a_percentage_moves_by_points_and_exactly_the_setting_is_not_more_than_it():
    before = {"nrr": metric(1.02, "102.0%")}
    assert moves(before, {"nrr": metric(0.969, "96.9%")}) == [("NRR (annualized)", "102.0%", "96.9%", "down 5.1 pts")]
    assert moves(before, {"nrr": metric(0.97, "97.0%")}) == []      # 5.0 points: not more than 5
    assert moves(before, {"nrr": metric(1.069, "106.9%")}) == []    # 4.9 points up


def test_other_metrics_move_by_percent_of_their_old_value():
    before = {"ending_arr": metric(1000.0, "1,000"), "runway_months": metric(15.0, "15.0 mo")}
    after = {"ending_arr": metric(1101.0, "1,101"), "runway_months": metric(13.5, "13.5 mo")}   # +10.1%, -10.0%
    assert moves(before, after) == [("Ending ARR ($K)", "1,000", "1,101", "up 10.1%")]


def test_a_move_away_from_zero_always_counts():
    # No percent of zero exists; any move from zero is more than any setting.
    assert moves({"net_new_arr": metric(0.0, "0")}, {"net_new_arr": metric(10.0, "10")}) == [
        ("Net new ARR ($K)", "0", "10", "up from zero")]


def test_a_value_that_became_or_stopped_being_a_number_is_listed_without_a_size():
    before = {"burn_multiple": metric(1.8, "1.80x"), "arr_qoq": metric(None, "data missing")}
    after = {"burn_multiple": metric(None, "∞ (ARR shrank)"), "arr_qoq": metric(0.05, "5.0%")}
    assert moves(before, after) == [("Burn multiple", "1.80x", "∞ (ARR shrank)", None),
                                    ("ARR growth QoQ", "data missing", "5.0%", None)]


def test_the_same_words_both_runs_are_not_a_move():
    same = {"burn_multiple": metric(None, "∞ (ARR shrank)"), "arr_yoy": metric(None, "data missing")}
    assert moves(same, same) == []


def test_a_bigger_setting_hides_smaller_moves():
    before, after = {"nrr": metric(1.02, "102.0%")}, {"nrr": metric(0.969, "96.9%")}
    assert moves(before, after, {"min_points": 0.10, "min_relative": 0.10}) == []


# ---------------------------------------------------------------------------
# New and resolved data gaps
# ---------------------------------------------------------------------------

def test_gaps_are_compared_quarter_by_quarter():
    before = results(gaps={"arr_qoq": ["Q1 2025", "Q2 2025"], "flag: Rule of 40": ["Q1 2026"]})
    after = results(gaps={"arr_qoq": ["Q1 2025", "Q2 2025"], "revenue_yoy": ["Q2 2026"]})
    changes = compare(before, after, SETTINGS)
    assert changes["new_gaps"] == {"revenue_yoy": ["Q2 2026"]}
    assert changes["resolved_gaps"] == {"flag: Rule of 40": ["Q1 2026"]}


def test_a_gap_that_spread_to_another_quarter_is_new_only_in_that_quarter():
    before = results(gaps={"arr_qoq": ["Q1 2025"]})
    after = results(gaps={"arr_qoq": ["Q1 2025", "Q2 2025"]})
    assert compare(before, after, SETTINGS)["new_gaps"] == {"arr_qoq": ["Q2 2025"]}


# ---------------------------------------------------------------------------
# The words the page, the memo and the command line show
# ---------------------------------------------------------------------------

def report(before, after, run_at="2026-06-18T09:05:41"):
    manifest = {"run_at": run_at, "results": before}
    return changes_since_last_run_from(manifest, after)


def changes_since_last_run_from(manifest, after):
    """changes_since_last_run with the results already worked out (the data step is tested above)."""
    found = baseline(manifest, after)
    return None if found is None else diff_runs.report_for(found, after, SETTINGS)


def test_the_compared_with_line_names_the_run_and_both_quarters():
    shown = compared_with_text(report(results(quarter="Q1 2026"), results()))
    assert shown == "Compared with the run of 2026-06-18 09:05, whose latest quarter was Q1 2026 (now Q2 2026)."


def test_the_compared_with_line_for_the_same_quarter_says_so():
    shown = compared_with_text(report(results(metrics={}), results()))
    assert shown == "Compared with the run of 2026-06-18 09:05, on the same latest quarter (Q2 2026)."


def test_sections_list_flips_moves_and_gaps_in_that_order_leaving_out_empty_ones():
    before = results(quarter="Q1 2026", metrics={"nrr": metric(1.02, "102.0%")},
                     flags={"Runway at current burn": "Passed"}, gaps={"flag: Rule of 40": ["Q1 2026"]})
    after = results(metrics={"nrr": metric(0.9, "90.0%")}, flags={"Runway at current burn": "Tripped"})
    assert change_sections(report(before, after), SETTINGS) == [
        ("Flags that flipped", ["Runway at current burn: Tripped (was Passed)"]),
        ("Metrics that moved more than 5.0 pts (percentages) or 10.0% (other metrics)",
         ["NRR (annualized): 102.0% to 90.0% (down 12.0 pts)"]),
        ("Resolved data gaps", ["Q1 2026: Flag: Rule of 40"]),
    ]


def test_a_metric_at_zero_both_runs_has_not_moved():
    # Found writing these tests: zero to zero was listed as "down from zero".
    assert moves({"net_new_arr": metric(0.0, "0")}, {"net_new_arr": metric(0.0, "0")}) == []


def test_a_flip_reads_today_s_status_first_so_a_reason_s_colon_stays_clear():
    assert diff_runs.flip_line("Rule of 40", "Cannot evaluate: missing input", "Tripped") == (
        "Rule of 40: Tripped (was Cannot evaluate: missing input)")


def test_nothing_changed_says_what_was_checked():
    # Different results (NRR moved 1 point), but nothing the comparison lists.
    before = results(metrics={"nrr": metric(1.0, "100.0%")})
    after = results(metrics={"nrr": metric(1.01, "101.0%")})
    assert change_sections(report(before, after), SETTINGS) == [
        ("Nothing changed", ["No flag flipped, no metric moved more than 5.0 pts (percentages) or 10.0% "
                             "(other metrics), and no data gap opened or closed."])]


def test_the_one_line_summary_counts_each_kind():
    before = results(flags={"Runway at current burn": "Passed"}, gaps={"arr_qoq": ["Q1 2025"]})
    after = results(flags={"Runway at current burn": "Tripped"})
    assert summary_text(report(before, after)) == (
        "1 flag flipped, 0 metrics moved, 0 new data gaps, 1 resolved (since the run of 2026-06-18 09:05)")
    assert summary_text(None) == NO_EARLIER_RUN


# ---------------------------------------------------------------------------
# From a company's data and manifest, and the command line
# ---------------------------------------------------------------------------

def test_changes_since_last_run_compares_today_s_data_with_the_manifest():
    data = collect_deck_data("Testco", "testco.xlsx", three_quarters(), None, TEST_CONFIG)
    earlier = run_results(data)
    earlier["flags"]["Runway at current burn"] = "Tripped"
    manifest = {"run_at": "2026-06-18T09:05:41", "results": earlier}
    found = changes_since_last_run(data, manifest, SETTINGS)
    assert found["flags"] == [("Runway at current burn", "Tripped", "Passed")]
    assert found["moved"] == [] and found["new_gaps"] == {} and found["resolved_gaps"] == {}
    assert changes_since_last_run(data, None, SETTINGS) is None


def test_the_command_line_prints_the_changes(tmp_path, capsys, monkeypatch):
    workbook = tmp_path / "northwind.xlsx"
    workbook.write_bytes((diff_runs.PROJECT_DIR / "data" / "northwind.xlsx").read_bytes())
    output_dir = tmp_path / "output"
    assert diff_runs.main([str(workbook), "--output-dir", str(output_dir)]) == 0
    assert capsys.readouterr().out.splitlines() == ["Northwind: " + HEADING, NO_EARLIER_RUN]

    today = changes_input(workbook)
    earlier = json.loads(json.dumps(today))
    earlier["flags"]["Runway at current burn"] = "Passed"
    save_manifest(output_dir / "northwind_manifest.json", {"run_at": "2026-06-18T09:05:41", "results": earlier})
    assert diff_runs.main([str(workbook), "--output-dir", str(output_dir), "--min-points", "0.5"]) == 0
    printed = capsys.readouterr().out.splitlines()
    assert printed[:2] == ["Northwind: " + HEADING, compared_with_text(
        {"since": "2026-06-18T09:05:41", "quarter_before": "Q2 2026", "quarter_now": "Q2 2026"})]
    assert "Flags that flipped:" in printed and "  - Runway at current burn: Tripped (was Passed)" in printed


def changes_input(workbook):
    """Today's results for a workbook, as the manifest would store them."""
    from clean import clean_workbook
    from metrics import load_config
    actuals, next_budget = clean_workbook(workbook)
    return run_results(collect_deck_data("Northwind", workbook.name, actuals, next_budget, load_config()))


def test_no_text_has_an_infinite_or_missing_number_written_as_nan():
    data = collect_deck_data("Testco", "testco.xlsx", three_quarters(blank="Q1 2026"), None, TEST_CONFIG)
    text = json.dumps(run_results(data))
    assert "NaN" not in text and "Infinity" not in text   # not JSON: other programs can't read either
