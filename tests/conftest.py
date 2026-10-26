from __future__ import annotations

import sys
from pathlib import Path

import pytest
from openpyxl import Workbook


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC = PROJECT_ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))


@pytest.fixture
def workbook_factory(tmp_path):
    def create(
        *,
        main_rows=None,
        source_rows=None,
        main_headers=("客户编号", "订单额"),
        source_headers=("客户编号", "客户等级", "地区"),
    ) -> Path:
        path = tmp_path / "input.xlsx"
        wb = Workbook()
        main = wb.active
        main.title = "订单"
        main.append(list(main_headers))
        for row in main_rows or [("001", 10), ("002", 20), ("003", 30)]:
            main.append(list(row))
        source = wb.create_sheet("客户")
        source.append(list(source_headers))
        for row in source_rows or [("001", "A", "华东"), ("002", "B", "华北"), ("003", "C", "华南")]:
            source.append(list(row))
        wb.save(path)
        wb.close()
        return path
    return create
