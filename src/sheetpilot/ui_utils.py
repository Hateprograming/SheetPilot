from __future__ import annotations

import json
from datetime import date, datetime
from typing import Any

import pandas as pd


def display_dataframe(rows: Any) -> pd.DataFrame:
    """Return an Arrow-safe display copy without changing workbook values."""
    frame = rows.copy() if isinstance(rows, pd.DataFrame) else pd.DataFrame(rows)

    def missing(value: Any) -> bool:
        if value is None:
            return True
        if isinstance(value, (list, dict, tuple, set)):
            return False
        try:
            return bool(pd.isna(value))
        except (TypeError, ValueError):
            return False

    for column in frame.columns:
        values = [value for value in frame[column].tolist() if not missing(value)]
        if not values:
            continue
        categories = set()
        for value in values:
            if isinstance(value, str):
                categories.add("text")
            elif isinstance(value, (bool, int, float)):
                categories.add("number")
            elif isinstance(value, (date, datetime)):
                categories.add("temporal")
            else:
                categories.add("other")
        if len(categories) > 1 or "other" in categories:
            frame[column] = frame[column].map(
                lambda value: None
                if missing(value)
                else json.dumps(value, ensure_ascii=False, default=str)
                if isinstance(value, (list, dict))
                else str(value)
            ).astype("string")
    return frame
