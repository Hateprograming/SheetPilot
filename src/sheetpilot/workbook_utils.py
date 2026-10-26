from __future__ import annotations

from typing import Any


def worksheet_dimensions(
    sheet: Any,
    *,
    fallback_max_row: int = 1,
    fallback_max_column: int = 1,
) -> tuple[int, int]:
    """Return usable dimensions for normal and read-only openpyxl worksheets.

    Some valid XLSX producers omit the worksheet ``dimension`` element. In
    openpyxl read-only mode that leaves ``max_row`` and ``max_column`` as None
    until the sheet dimensions are explicitly calculated.
    """

    if sheet.max_row is None or sheet.max_column is None:
        try:
            sheet.calculate_dimension(force=True)
        except (AttributeError, TypeError, ValueError):
            pass

    def positive_int(value: Any, fallback: int) -> int:
        try:
            resolved = int(value)
        except (TypeError, ValueError):
            resolved = fallback
        return max(1, resolved)

    return (
        positive_int(sheet.max_row, fallback_max_row),
        positive_int(sheet.max_column, fallback_max_column),
    )
