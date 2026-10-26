from __future__ import annotations

from sheetpilot.profiler import profile_workbook


def test_profiles_sheets_headers_and_preview(workbook_factory):
    path = workbook_factory()
    profile = profile_workbook(path, "客户订单.xlsx", 100_000)
    assert profile.file_name == "客户订单.xlsx"
    assert [sheet.name for sheet in profile.sheets] == ["订单", "客户"]
    assert profile.sheets[0].header_row == 1
    assert [column.name for column in profile.sheets[0].columns] == ["客户编号", "订单额"]
    assert len(profile.sheets[0].preview_rows) == 3


def test_marks_duplicate_headers_as_risk(workbook_factory):
    path = workbook_factory(main_headers=("编号", "编号"))
    profile = profile_workbook(path, "dup.xlsx", 100_000)
    assert any("重复列名" in risk for risk in profile.risks)

