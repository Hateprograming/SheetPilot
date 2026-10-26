from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import load_workbook

from sheetpilot.config import Settings
from sheetpilot.llm import FakeLLMClient
from sheetpilot.planner import validate_plan_structure
from sheetpilot.schemas import (
    CleanOperation,
    CleanTaskSpec,
    CompositeTaskSpec,
    ConversionDirection,
    ConversionTaskSpec,
    LookupTaskSpec,
    TaskEnvelope,
)
from sheetpilot.service import SheetPilotService
from sheetpilot.skills import SkillRegistry, capability_refusal


def make_service(tmp_path: Path, envelope: TaskEnvelope | None = None) -> SheetPilotService:
    settings = Settings(
        project_root=tmp_path,
        data_dir=tmp_path / "data",
        database_path=tmp_path / "data" / "sheetpilot.sqlite3",
        llm_provider="fake",
        api_key="",
        base_url="",
        model="fake",
        max_file_bytes=20 * 1024 * 1024,
        max_sheet_rows=100_000,
    )
    return SheetPilotService(settings, FakeLLMClient(envelope) if envelope else None)


def test_skill_registry_preserves_v1_contracts_and_adds_transform():
    registry = SkillRegistry()
    manifest = registry.manifest()
    assert len(manifest) == 8
    assert {item["id"] for item in manifest} == {
        "lookup_skill", "normalization_skill", "aggregation_skill",
        "sort_rank_skill", "cleaning_skill", "formula_skill", "conversion_skill",
        "transform_skill",
    }
    versions = {item["id"]: item["version"] for item in manifest}
    assert all(
        versions[item] == "1.0"
        for item in versions
        if item not in {"normalization_skill", "transform_skill", "formula_skill", "conversion_skill"}
    )
    assert versions["normalization_skill"] == "1.3"
    assert versions["formula_skill"] == "1.1"
    assert versions["conversion_skill"] == "1.1"
    assert versions["transform_skill"] == "2.5"
    assert all(item["allowed_tools"] for item in manifest)


def test_conversion_cannot_be_composed_with_other_skills():
    envelope = TaskEnvelope(
        goal="先转换再清洗",
        spec=CompositeTaskSpec(
            operations=[
                ConversionTaskSpec(direction=ConversionDirection.CSV_TO_XLSX, confidence=1),
                CleanTaskSpec(
                    sheet="CSV_Data", operation=CleanOperation.REMOVE_DUPLICATES,
                    key_columns=["编号"], confidence=1,
                ),
            ],
            confidence=1,
        ),
    )
    assert "不能放入复合任务" in "；".join(SkillRegistry().validate_composition(envelope))


def test_skill_contract_rejects_tool_substitution(workbook_factory, tmp_path):
    service = make_service(tmp_path)
    profile = service.create_task("input.xlsx", workbook_factory().read_bytes())
    task_id = service.task_id_from_profile(profile)
    envelope = TaskEnvelope(
        goal="补充客户等级",
        spec=LookupTaskSpec(
            main_sheet="订单", source_sheet="客户",
            main_key_column="客户编号", source_key_column="客户编号",
            return_columns=["客户等级"], output_column_names=["客户等级"], confidence=1,
        ),
    )
    plan, validation = service.make_task_plan(task_id, envelope)
    assert validation.valid
    plan.steps[0].tool_name = "clean"
    errors = validate_plan_structure(plan, service.tools, service.skills)
    assert any("不允许工具" in item for item in errors)


def test_unsupported_request_returns_structured_refusal(workbook_factory, tmp_path):
    envelope = TaskEnvelope(
        goal="占位",
        spec=CleanTaskSpec(sheet="订单", confidence=1),
    )
    service = make_service(tmp_path, envelope)
    profile = service.create_task("input.xlsx", workbook_factory().read_bytes())
    task_id = service.task_id_from_profile(profile)
    turn = service.submit_message(task_id, "帮我写一个 VBA 宏删除重复数据")
    assert turn.refusal is not None
    assert turn.plan is None
    assert "VBA" in turn.refusal.reason
    assert capability_refusal("执行任意代码") is not None


@pytest.mark.parametrize(
    ("encoding", "delimiter"),
    [("utf-8", ","), ("utf-8-sig", "\t"), ("gb18030", ";")],
)
def test_csv_to_xlsx_preserves_text_and_can_continue_with_cleaning(
    tmp_path, encoding, delimiter
):
    csv_bytes = (f"编号{delimiter}名称\n001{delimiter}张三\n001{delimiter}李四\n").encode(encoding)
    conversion = TaskEnvelope(
        goal="转换为 XLSX",
        spec=ConversionTaskSpec(
            direction=ConversionDirection.CSV_TO_XLSX,
            encoding=encoding,
            delimiter=delimiter,
            preserve_text=True,
            confidence=1,
        ),
    )
    service = make_service(tmp_path, conversion)
    profile = service.create_task("customers.csv", csv_bytes)
    task_id = service.task_id_from_profile(profile)
    turn = service.submit_message(task_id, "转换为 XLSX")
    assert turn.plan and turn.plan.skill_ids == ["conversion_skill"]
    executed = service.approve_current_plan(task_id)
    workbook = load_workbook(executed.output_path, data_only=True)
    assert workbook["CSV_Data"]["A2"].value == "001"
    workbook.close()

    clean = TaskEnvelope(
        goal="标记重复编号",
        spec=CleanTaskSpec(
            sheet="CSV_Data", operation=CleanOperation.MARK_DUPLICATES,
            key_columns=["编号"], output_sheet="重复检查", confidence=1,
        ),
    )
    plan, validation = service.make_task_plan(task_id, clean)
    assert validation.valid, validation.errors
    service.approve(task_id, "approve")
    _, report, result = service.execute_task_plan(task_id, clean, plan)
    assert report.status == "passed"
    assert "重复检查" in result.created_sheets


def test_xlsx_to_csv_requires_sheet_and_creates_downloadable_artifact(
    workbook_factory, tmp_path
):
    service = make_service(tmp_path)
    profile = service.create_task("input.xlsx", workbook_factory().read_bytes())
    task_id = service.task_id_from_profile(profile)
    missing = TaskEnvelope(
        goal="导出 CSV",
        spec=ConversionTaskSpec(
            direction=ConversionDirection.XLSX_TO_CSV,
            confidence=1,
        ),
    )
    _, validation = service.make_task_plan(task_id, missing)
    assert not validation.valid
    assert any("必须指定" in item for item in validation.errors)

    export = TaskEnvelope(
        goal="导出订单表",
        spec=ConversionTaskSpec(
            direction=ConversionDirection.XLSX_TO_CSV,
            source_sheet="订单",
            confidence=1,
        ),
    )
    plan, validation = service.make_task_plan(task_id, export)
    assert validation.valid
    service.approve(task_id, "approve")
    _, report, result = service.execute_task_plan(task_id, export, plan)
    assert report.status == "passed"
    step = result.details["steps"][0]
    name, content = service.get_artifact_download(task_id, step["details"]["artifact_path"])
    assert name.endswith(".csv")
    assert content.startswith(b"\xef\xbb\xbf")
