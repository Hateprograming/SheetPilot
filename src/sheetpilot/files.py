from __future__ import annotations

import hashlib
import shutil
import uuid
import zipfile
from pathlib import Path

from .errors import FileRejectedError


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


class TaskFileManager:
    def __init__(self, data_dir: Path, max_file_bytes: int):
        self.tasks_dir = data_dir / "tasks"
        self.tasks_dir.mkdir(parents=True, exist_ok=True)
        self.max_file_bytes = max_file_bytes

    def ingest(self, original_name: str, content: bytes) -> tuple[str, Path]:
        suffix = Path(original_name).suffix.casefold()
        if suffix not in {".xlsx", ".csv"}:
            raise FileRejectedError("v1.0 仅支持 .xlsx 和 .csv 文件。")
        if not content:
            raise FileRejectedError("上传文件为空。")
        if len(content) > self.max_file_bytes:
            raise FileRejectedError("文件超过 v1.0 的大小限制。")
        task_id = uuid.uuid4().hex
        task_dir = self.tasks_dir / task_id
        original_dir = task_dir / "original"
        (task_dir / "versions").mkdir(parents=True)
        (task_dir / "artifacts").mkdir(parents=True)
        original_dir.mkdir(parents=True)
        original_path = original_dir / f"input{suffix}"
        original_path.write_bytes(content)
        if suffix == ".xlsx" and not zipfile.is_zipfile(original_path):
            shutil.rmtree(task_dir, ignore_errors=True)
            raise FileRejectedError("文件不是有效的 XLSX 压缩包，可能已损坏或被加密。")
        return task_id, original_path

    def staging_path(self, task_id: str) -> Path:
        return self.task_dir(task_id) / "versions" / "v0.xlsx"

    def task_dir(self, task_id: str) -> Path:
        path = (self.tasks_dir / task_id).resolve()
        if path.parent != self.tasks_dir.resolve() or not path.exists():
            raise FileRejectedError("任务目录不存在。")
        return path

    def version_path(self, task_id: str, version: int) -> Path:
        return self.task_dir(task_id) / "versions" / f"v{version}.xlsx"

    def version_candidate_path(self, task_id: str, version: int, commit_id: str) -> Path:
        safe_commit_id = Path(commit_id).name
        return self.task_dir(task_id) / "versions" / f".v{version}.{safe_commit_id}.candidate.xlsx"

    def quarantine_path(self, task_id: str, name: str) -> Path:
        directory = self.task_dir(task_id) / "quarantine"
        directory.mkdir(parents=True, exist_ok=True)
        return directory / Path(name).name

    def create_version_copy(self, task_id: str, source: Path, version: int) -> Path:
        target = self.version_path(task_id, version)
        shutil.copy2(source, target)
        return target

    def artifact_path(self, task_id: str, name: str) -> Path:
        safe_name = Path(name).name
        return self.task_dir(task_id) / "artifacts" / safe_name

    def run_dir(self, task_id: str, run_id: str) -> Path:
        safe_run_id = Path(run_id).name
        path = self.task_dir(task_id) / "runs" / safe_run_id
        (path / "checkpoints").mkdir(parents=True, exist_ok=True)
        (path / "reports").mkdir(parents=True, exist_ok=True)
        return path

    def checkpoint_path(self, task_id: str, run_id: str, step_order: int) -> Path:
        return self.run_dir(task_id, run_id) / "checkpoints" / f"step_{step_order:02d}.xlsx"

    def run_report_path(self, task_id: str, run_id: str, name: str) -> Path:
        return self.run_dir(task_id, run_id) / "reports" / Path(name).name

    def remove_run_files(self, task_id: str, run_id: str) -> None:
        shutil.rmtree(self.run_dir(task_id, run_id), ignore_errors=True)

    def remove_run_checkpoints(self, task_id: str, run_id: str) -> None:
        shutil.rmtree(self.run_dir(task_id, run_id) / "checkpoints", ignore_errors=True)
