from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


def _load_dotenv(path: Path) -> None:
    if not path.exists():
        return
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


@dataclass(frozen=True)
class Settings:
    project_root: Path
    data_dir: Path
    database_path: Path
    llm_provider: str
    api_key: str
    base_url: str
    model: str
    max_file_bytes: int
    max_sheet_rows: int
    session_secret: str = "sheetpilot-development-secret-change-me"
    retention_days: int = 7
    execution_budget_seconds: float = 600
    planning_budget_seconds: float = 180

    @classmethod
    def load(cls, project_root: Path | None = None) -> "Settings":
        root = (project_root or Path.cwd()).resolve()
        _load_dotenv(root / ".env")
        data_raw = Path(os.getenv("SHEETPILOT_DATA_DIR", "./data"))
        data_dir = data_raw if data_raw.is_absolute() else root / data_raw
        max_mb = int(os.getenv("SHEETPILOT_MAX_FILE_MB", "20"))
        return cls(
            project_root=root,
            data_dir=data_dir.resolve(),
            database_path=(data_dir / "sheetpilot.sqlite3").resolve(),
            llm_provider=os.getenv("SHEETPILOT_LLM_PROVIDER", "deepseek"),
            api_key=os.getenv("SHEETPILOT_API_KEY", ""),
            base_url=os.getenv("SHEETPILOT_BASE_URL", "").rstrip("/"),
            model=os.getenv("SHEETPILOT_MODEL", "deepseek-flash"),
            max_file_bytes=max_mb * 1024 * 1024,
            max_sheet_rows=int(os.getenv("SHEETPILOT_MAX_SHEET_ROWS", "100000")),
            session_secret=os.getenv(
                "SHEETPILOT_SESSION_SECRET", "sheetpilot-development-secret-change-me"
            ),
            retention_days=max(1, int(os.getenv("SHEETPILOT_RETENTION_DAYS", "7"))),
            execution_budget_seconds=max(1, float(os.getenv("SHEETPILOT_EXECUTION_BUDGET_SECONDS", "600"))),
            planning_budget_seconds=max(1, float(os.getenv("SHEETPILOT_PLANNING_BUDGET_SECONDS", "180"))),
        )

    def ensure_dirs(self) -> None:
        self.data_dir.mkdir(parents=True, exist_ok=True)
        (self.data_dir / "tasks").mkdir(parents=True, exist_ok=True)

    @property
    def llm_configured(self) -> bool:
        return bool(self.api_key and self.base_url and self.model)

