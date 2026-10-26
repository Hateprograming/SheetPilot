from __future__ import annotations

import pyarrow as pa

from sheetpilot.ui_utils import display_dataframe


def test_mixed_numeric_and_text_column_is_arrow_safe():
    frame = display_dataframe([
        {"年龄": 60, "BMI": 22.5},
        {"年龄": "61岁", "BMI": "未测量"},
        {"年龄": None, "BMI": None},
    ])

    table = pa.Table.from_pandas(frame, preserve_index=False)

    assert table.column("年龄").to_pylist() == ["60", "61岁", None]
    assert table.column("BMI").to_pylist() == ["22.5", "未测量", None]
