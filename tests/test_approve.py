"""Unit tests for approve.py: recording a human's approval, and refusing when it wouldn't mean anything.

Everything runs in a temporary folder with made-up files, so no real deck or manifest is touched.
Run from the project folder:  pytest
"""

import pytest

import approve
from provenance import build_manifest, file_sha256, manifest_path, read_manifest, save_manifest


@pytest.fixture
def company(tmp_path):
    """A made-up company with a workbook, a config and a manifest from a finished run."""
    data_dir, output_dir = tmp_path / "data", tmp_path / "output"
    data_dir.mkdir()
    output_dir.mkdir()
    workbook, config = data_dir / "testco.xlsx", tmp_path / "config.yaml"
    workbook.write_bytes(b"quarter,arr")
    config.write_bytes(b"nrr_min: 1.00")
    manifest = build_manifest("Testco", workbook, config, "testco_board_pack.pptx", {}, ai_text=True,
                              approval=None)
    save_manifest(manifest_path(workbook, output_dir), manifest)
    return {"workbook": workbook, "config": config, "data_dir": data_dir, "output_dir": output_dir,
            "manifest": manifest_path(workbook, output_dir)}


def approve_testco(company, reviewer="Tyler Ho", now="2026-09-17T15:00:00"):
    return approve.approve("testco", reviewer, data_dir=company["data_dir"], output_dir=company["output_dir"],
                           config_path=company["config"], now=now)


def test_approving_records_the_reviewer_and_when(company):
    approve_testco(company)
    saved = read_manifest(company["manifest"])
    assert saved["approval"] == {
        "reviewer": "Tyler Ho", "approved_at": "2026-09-17T15:00:00",
        "input_sha256": file_sha256(company["workbook"]), "config_sha256": file_sha256(company["config"])}
    assert saved["deck"]["status"] == "approved by Tyler Ho on 2026-09-17T15:00:00"


def test_approving_leaves_the_rest_of_the_manifest_alone(company):
    before = read_manifest(company["manifest"])
    after = approve_testco(company)
    assert {key: value for key, value in after.items() if key not in ("approval", "deck")} == \
           {key: value for key, value in before.items() if key not in ("approval", "deck")}


def test_a_company_with_no_manifest_stops(company):
    with pytest.raises(ValueError, match="no run to approve"):
        approve.approve("othercorp", "Tyler Ho", data_dir=company["data_dir"],
                        output_dir=company["output_dir"], config_path=company["config"])


def test_a_workbook_changed_since_the_run_stops(company):
    # The deck on disk was built from the old numbers, so approving it would approve the wrong thing.
    company["workbook"].write_bytes(b"quarter,arr,more")
    with pytest.raises(ValueError, match="workbook has changed"):
        approve_testco(company)


def test_thresholds_changed_since_the_run_stop(company):
    company["config"].write_bytes(b"nrr_min: 0.90")
    with pytest.raises(ValueError, match="config.yaml has changed"):
        approve_testco(company)


def test_a_reviewer_name_is_required(company):
    with pytest.raises(ValueError, match="reviewer"):
        approve.approve("testco", "  ", data_dir=company["data_dir"], output_dir=company["output_dir"],
                        config_path=company["config"])


def test_the_command_line_prints_what_to_do_next(company, capsys):
    exit_code = approve.main(["testco", "--reviewer", "Tyler Ho"], data_dir=company["data_dir"],
                             output_dir=company["output_dir"], config_path=company["config"])
    printed = capsys.readouterr().out
    assert exit_code == 0
    assert "Tyler Ho" in printed and "build_deck.py" in printed  # how to get "reviewed by" on the deck
    # The watermark is opt-in (--draft), so the hint is about the footer, not about removing a watermark.
    assert "footer" in printed and "watermark" not in printed.lower()
    assert read_manifest(company["manifest"])["approval"]["reviewer"] == "Tyler Ho"


def test_the_command_line_prints_one_line_when_it_refuses(company, capsys):
    company["workbook"].write_bytes(b"quarter,arr,more")
    exit_code = approve.main(["testco"], data_dir=company["data_dir"], output_dir=company["output_dir"],
                             config_path=company["config"])
    printed = capsys.readouterr().out
    assert exit_code == 1 and "workbook has changed" in printed
    assert read_manifest(company["manifest"])["approval"] is None
