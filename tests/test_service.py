from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook

from sheetpilot.config import Settings
from sheetpilot.llm import FakeLLMClient
from sheetpilot.schemas import LookupTaskSpec
from sheetpilot.service import SheetPilotService


def make_settings(tmp_path: Path) -> Settings:
    return Settings(
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


def test_end_to_end_service(workbook_factory, tmp_path):
    input_path = workbook_factory(main_rows=[("001", 10), ("999", 20)])
    task_spec = LookupTaskSpec(
        main_sheet="订单",
        source_sheet="客户",
        main_key_column="客户编号",
        source_key_column="客户编号",
        return_columns=["客户等级"],
        output_column_names=["客户等级"],
        confidence=1.0,
    )
    service = SheetPilotService(make_settings(tmp_path), FakeLLMClient(task_spec))
    profile = service.create_task("input.xlsx", input_path.read_bytes())
    task_id = service.task_id_from_profile(profile)
    parsed = service.parse_request(task_id, "补充客户等级")
    plan, validation = service.make_plan(task_id, parsed)
    assert validation.valid
    assert plan.estimated_rows == 2
    service.approve(task_id, "approve")
    output, report = service.execute(task_id, parsed)
    assert report.status == "passed_with_warnings"
    assert output.exists()
    wb = load_workbook(output, data_only=True)
    assert wb["订单"]["C2"].value == "A"
    assert wb["订单"]["C3"].value is None
    wb.close()
    stored = service.result_payload(task_id)
    assert stored["task"]["status"] == "completed"
    assert stored["changeset"]["unmatched_count"] == 1

