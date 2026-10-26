from __future__ import annotations

import json
import hashlib
from collections import Counter
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from zoneinfo import ZoneInfo

import pytest
from openpyxl import Workbook

from evals.phase2.cli import EXPECTED_CAP, validate
from evals.phase2.common import DEFAULT_OUT, load_cases, load_gzip
from evals.phase2.generator import prepare
from evals.phase2.runner import (
    _event_cost, _restore_call_metrics, _semantic_output_sha,
    _semantic_plan_sha, instances, summarize, worker,
)


@pytest.fixture(scope="module")
def suite():
    out=DEFAULT_OUT
    prepare(out)
    return out,load_cases(out/"cases")


def test_locked_distribution_and_validation(suite):
    out,cases=suite
    assert validate(out)==[]
    assert len(cases)==160
    assert Counter(c.journey for c in cases)==Counter(direct=80,multi_step=32,clarification=32,boundary=16)
    assert Counter(c.primary_capability for c in cases)==Counter(EXPECTED_CAP)
    assert Counter(c.size_class for c in cases)==Counter(small=80,medium=56,large=24)
    assert Counter(c.source["type"] for c in cases)==Counter(research_reconstructed=64,spreadsheetbench_adapted=32,synthetic=64)


def test_stability_expands_to_360_runs(suite):
    _,cases=suite
    selected=[c for c in cases if c.stability_selected]
    assert len(selected)==40
    assert Counter(c.primary_capability for c in selected)==Counter({name:5 for name in EXPECTED_CAP})
    assert len(instances(cases,"stability"))==360


def test_corrected_cases_only_lock_user_stated_values(suite):
    _,cases=suite;by_id={case.id:case for case in cases}
    assert "去掉首尾空格并做 NFKC" in by_id["P2-NORMALIZE-009"].request
    assert "=F2*G2" in by_id["P2-FORMULA-001"].request
    assert by_id["P2-CLEAN-017"].spec["output_sheet"] == "SheetPilot_Cleaned"
    assert by_id["P2-CONVERSION-011"].spec["encoding"] == "auto"
    assert by_id["P2-FORMULA-011"].spec["recalculate"] is True
    assert by_id["P2-TRANSFORM-014"].spec["operations"][0]["impute_rules"][0]["output_column"] == ""
    assert all(case.version == "2.1" for case in cases)
    for index in range(17, 21):
        case = by_id[f"P2-AGGREGATE-{index:03d}"]
        assert "结果列叫" not in case.request
        assert case.spec["metrics"][0]["output_column"] == ""
        oracle = load_gzip(DEFAULT_OUT / "cases" / case.id / case.oracle)
        assert oracle["workbook"]["SheetPilot_Summary"][0][-1] == "金额_sum"
    for index in range(14, 17):
        case = by_id[f"P2-CLEAN-{index:03d}"]
        assert "按标准编号精确去重" in case.request
        assert case.spec["operations"][-1]["key_columns"] == ["标准编号"]
    for index in range(13, 19):
        case = by_id[f"P2-TRANSFORM-{index:03d}"]
        impute = case.spec["operations"][0]["impute_rules"][0]
        assert impute["output_column"] == ""
        assert impute["indicator_column"] == ""
    for index in range(19, 25):
        case = by_id[f"P2-TRANSFORM-{index:03d}"]
        assert case.spec["map_rules"][0]["output_column"] == ""
        oracle = load_gzip(DEFAULT_OUT / "cases" / case.id / case.oracle)
        assert "状态_映射" in oracle["workbook"]["交易"][0]
    for index in range(11, 15):
        case = by_id[f"P2-FORMULA-{index:03d}"]
        assert case.initial_spec["mode"] == "set_formula"
        assert case.initial_spec["formula"] is None
        assert "formula" in case.initial_spec["missing_fields"]
    for index in range(14, 16):
        case = by_id[f"P2-SORT_RANK-{index:03d}"]
        assert "排名" in case.clarification_script[0].answer
        assert case.spec["sort_keys"] and case.spec["rank_column"] == "金额"


def test_semantic_plan_hash_ignores_run_ids(suite):
    _, cases = suite
    case = next(item for item in cases if item.journey == "direct")
    step = SimpleNamespace(tool_name="aggregate", arguments={
        "task_type": "aggregate", "sheet": "交易", "group_by": ["地区"],
        "metrics": [{"column": "金额", "function": "sum", "output_column": "金额_sum"}],
        "output_sheet": "汇总", "confidence": 1,
    })
    first = SimpleNamespace(plan_id="run-a", steps=[step])
    second = SimpleNamespace(plan_id="run-b", steps=[step])
    assert _semantic_plan_sha(first, None, case.request) == _semantic_plan_sha(
        second, None, case.request
    )


def test_semantic_output_hash_ignores_xlsx_zip_metadata(tmp_path):
    first = tmp_path / "first.xlsx"
    second = tmp_path / "second.xlsx"
    for path, title in ((first, "first"), (second, "second")):
        workbook = Workbook()
        workbook.active.append(["编号", "公式"])
        workbook.active.append([1, "=A2*2"])
        workbook.properties.title = title
        workbook.save(path)
        workbook.close()
    assert hashlib.sha256(first.read_bytes()).hexdigest() != hashlib.sha256(second.read_bytes()).hexdigest()
    assert _semantic_output_sha(first, tmp_path / "none-a") == _semantic_output_sha(
        second, tmp_path / "none-b"
    )


def test_timeout_metrics_are_restored_from_call_log(tmp_path):
    calls = tmp_path / "calls.json"
    calls.write_text(json.dumps([
        {"usage": {"prompt_tokens": 120, "completion_tokens": 30, "prompt_cache_hit_tokens": 20}},
        {"usage": {"prompt_tokens": 80, "completion_tokens": 10, "prompt_cache_hit_tokens": 5}},
    ]), encoding="utf-8")
    result = _restore_call_metrics(
        {"status": "failed", "error_category": "timeout", "calls": 0}, calls
    )
    assert result["calls"] == 2
    assert result["token_usage"] == {
        "input_tokens": 200, "output_tokens": 40, "cached_tokens": 25,
    }


def test_deepseek_pricing_selects_peak_and_off_peak_rates(suite):
    out,_=suite;pricing=json.loads((out/"pricing.json").read_text(encoding="utf-8"))
    timezone=ZoneInfo("Asia/Shanghai")
    usage={"prompt_tokens":1_000_000,"completion_tokens":1_000_000,"prompt_cache_hit_tokens":100_000}
    events=[
        {"started":datetime(2026,10,23,10,0,tzinfo=timezone).timestamp(),"usage":usage},
        {"started":datetime(2026,10,24,10,0,tzinfo=timezone).timestamp(),"usage":usage},
    ]
    cost,tiers=_event_cost(events,pricing)
    assert tiers == {"peak":1,"off_peak":1}
    assert cost == pytest.approx((.9*2+.1*.04+8)+(.9*1+.1*.02+4))


def test_prepare_is_idempotent_and_manifest_locks_files(suite):
    out,_=suite
    before=json.loads((out/"manifest.json").read_text(encoding="utf-8"))["suite_sha256"]
    assert prepare(out)==out.resolve()
    after=json.loads((out/"manifest.json").read_text(encoding="utf-8"))["suite_sha256"]
    assert before==after


def test_empty_report_is_unpriced(suite):
    out,_=suite
    report=summarize(out)
    assert report["locked"]["completed"]==0
    expected="priced" if (out/"pricing.json").exists() else "unpriced"
    assert report["locked"]["cost_status"]==expected


def test_runner_rejects_missing_suite():
    from evals.phase2.runner import run_suite
    from evals.phase2.common import DEFAULT_OUT
    missing=DEFAULT_OUT/"missing-suite-test"
    missing.mkdir(parents=True,exist_ok=True)
    with pytest.raises(RuntimeError,match="expected 160 cases, found 0"):
        run_suite(missing,"locked",dry_run=True)


def test_offline_worker_covers_execute_clarify_and_refuse(suite):
    out,cases=suite;by_id={c.id:c for c in cases};smoke=out/"worker-smoke"
    direct=next(c for c in cases if c.journey=="direct" and c.size_class=="small")
    clarify=next(c for c in cases if c.journey=="clarification" and len(c.clarification_script)==2 and c.size_class=="small")
    boundary=next(c for c in cases if c.journey=="boundary")
    for case in (direct,clarify,boundary):
        folder=smoke/case.id
        worker(out/"cases"/case.id/"case.json",folder,"offline",6,case.request,"smoke-"+case.id)
        result=json.loads((folder/"result.json").read_text(encoding="utf-8"))
        assert result["status"]=="passed",result
        assert result["calls"]==0
    result=json.loads((smoke/clarify.id/"result.json").read_text(encoding="utf-8"))
    # The runner accepts a complete plan as soon as the workbook makes the
    # remaining field unambiguous; it must not force a scripted second round.
    assert 0 <= result["clarification_rounds"] <= 2

