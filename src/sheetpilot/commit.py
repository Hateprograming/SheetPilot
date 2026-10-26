from __future__ import annotations

import os
import json
import shutil
import time
from pathlib import Path
from uuid import uuid4

from .errors import IntentValidationError, VerificationError
from .files import TaskFileManager, sha256_file
from .integrity import FinalVerifier
from .schemas import TaskPlan, ToolResult, VerificationReport, WorkbookProfile
from .storage import Repository


class VersionCommitter:
    def __init__(
        self,
        repository: Repository,
        files: TaskFileManager,
        verifier: FinalVerifier,
    ):
        self.repository = repository
        self.files = files
        self.verifier = verifier

    @staticmethod
    def _durable_copy(source: Path, candidate: Path) -> None:
        if candidate.exists():
            candidate.unlink()
        shutil.copy2(source, candidate)
        with candidate.open("r+b") as handle:
            handle.flush()
            os.fsync(handle.fileno())

    @staticmethod
    def _replace_with_retry(candidate: Path, final_path: Path) -> None:
        last_error: OSError | None = None
        for delay in (0.0, 0.2, 0.5, 1.0):
            if delay:
                time.sleep(delay)
            try:
                os.replace(candidate, final_path)
                return
            except OSError as exc:
                last_error = exc
        assert last_error is not None
        raise last_error

    def _reverify_recovery_candidate(self, record: dict, path: Path) -> None:
        """Re-run final verification before completing an interrupted commit."""
        task = self.repository.get_task(record["task_id"])
        if not task or not task.get("plan_json") or not record.get("verification_json"):
            raise ValueError("恢复提交缺少任务计划或终检记录。")
        delivery = json.loads(record.get("changeset_json") or "{}").get("details", {}).get("delivery")
        if delivery:
            export = Path(delivery["path"]).resolve()
            expected = self.files.version_path(record["task_id"], int(record["target_version"])).with_suffix(".csv").resolve()
            if export != expected or not export.is_file() or sha256_file(export) != delivery["sha256"]:
                raise VerificationError("恢复时 CSV 交付文件完整性校验失败。")
        plan = TaskPlan.model_validate_json(task["plan_json"])
        prior_report = VerificationReport.model_validate_json(record["verification_json"])
        report, _, _ = self.verifier.verify(
            task_id=record["task_id"],
            original_path=Path(task["original_path"]),
            original_sha256=task.get("original_sha256") or "",
            base_path=Path(task["current_path"]),
            base_version=int(task["current_version"]),
            base_sha256=plan.base_sha256,
            candidate_path=path,
            original_name=task["original_name"],
            plan=plan,
            step_reports=[prior_report],
            artifact_dir=self.files.task_dir(record["task_id"]) / "artifacts",
        )
        if report.status == "failed_rolled_back":
            raise ValueError("恢复时独立终检未通过。")
        if report.final_sha256 != record["candidate_sha256"]:
            raise ValueError("恢复时终检哈希与 prepared 记录不一致。")

    def commit(
        self,
        *,
        task: dict,
        plan: TaskPlan,
        run_id: str,
        fencing_token: str,
        source_path: Path,
        step_reports: list[VerificationReport],
        combined_result: ToolResult,
    ) -> tuple[Path, VerificationReport, WorkbookProfile]:
        target_version = int(task["current_version"]) + 1
        commit_id = uuid4().hex
        candidate = self.files.version_candidate_path(task["task_id"], target_version, commit_id)
        final_path = self.files.version_path(task["task_id"], target_version)
        if final_path.exists():
            raise IntentValidationError("目标正式版本文件已经存在，已停止提交以避免覆盖。")
        self._durable_copy(source_path, candidate)
        report, profile, _ = self.verifier.verify(
            task_id=task["task_id"],
            original_path=Path(task["original_path"]),
            original_sha256=task.get("original_sha256") or "",
            base_path=Path(task["current_path"]),
            base_version=int(task["current_version"]),
            base_sha256=plan.base_sha256,
            candidate_path=candidate,
            original_name=task["original_name"],
            plan=plan,
            step_reports=step_reports,
            artifact_dir=self.files.task_dir(task["task_id"]) / "artifacts",
        )
        if report.status == "failed_rolled_back":
            raise VerificationError(
                "最终结果包含未批准的变更，正式版本未提交。",
                detail=report.model_dump(mode="json"),
            )
        # Copy the export into the immutable version directory before publication.
        exports = [item for item in combined_result.details.get("steps", [])
                   if item.get("details", {}).get("output_format") == "csv"]
        if exports:
            export = Path(exports[-1]["details"]["artifact_path"])
            if export.resolve().parent != (self.files.task_dir(task["task_id"]) / "artifacts").resolve():
                raise VerificationError("CSV 交付文件不在受管目录中。")
            delivery_path = final_path.with_suffix(".csv")
            if self.repository.get_version(task["task_id"], target_version):
                raise VerificationError("目标版本已经发布，不能覆盖其 CSV 文件。")
            delivery_candidate = candidate.with_suffix(".csv")
            self._durable_copy(export, delivery_candidate)
            export_hash = sha256_file(export)
            if export_hash != exports[-1]["details"].get("artifact_sha256") or sha256_file(delivery_candidate) != export_hash:
                raise VerificationError("CSV 交付复制核验失败。")
            self._replace_with_retry(delivery_candidate, delivery_path)
            combined_result.details["delivery"] = {
                "path": str(delivery_path), "name": export.name,
                "format": "csv", "sha256": export_hash,
            }
        report.commit_id = commit_id
        report.output_path = str(final_path)
        self.repository.create_version_commit(
            commit_id=commit_id,
            run_id=run_id,
            task_id=task["task_id"],
            target_version=target_version,
            candidate_path=candidate,
            final_path=final_path,
            candidate_sha256=report.final_sha256 or sha256_file(candidate),
            fencing_token=fencing_token,
            profile_json=profile.model_dump_json(),
            changeset_json=combined_result.model_dump_json(),
            verification_json=report.model_dump_json(),
        )
        if not self.repository.validate_execution_lease(task["task_id"], run_id, fencing_token):
            self.repository.update_version_commit(
                commit_id, "failed", {"message": "执行租约已经失效。"}
            )
            candidate.unlink(missing_ok=True)
            raise IntentValidationError("执行租约已经失效，正式版本未提交。")
        try:
            self._replace_with_retry(candidate, final_path)
        except OSError as exc:
            self.repository.update_version_commit(
                commit_id, "failed", {"type": type(exc).__name__, "message": str(exc)}
            )
            raise IntentValidationError(
                "正式版本文件暂时被占用，候选结果已保留但未提交。"
            ) from exc
        if sha256_file(final_path) != (report.final_sha256 or ""):
            self.repository.update_version_commit(
                commit_id, "failed", {"message": "正式路径哈希与候选哈希不一致。"}
            )
            raise VerificationError("正式版本复制后哈希不一致，已停止数据库提交。")
        self.repository.update_version_commit(commit_id, "file_committed")
        try:
            self.repository.complete_version_commit(commit_id=commit_id)
        except Exception:
            # The database transaction is the publication boundary. If a cancel
            # request wins after the candidate was copied, remove that unlisted
            # version so a later plan cannot mistake it for a formal release.
            task_after = self.repository.get_task(task["task_id"])
            if (
                final_path.is_file()
                and task_after
                and int(task_after["current_version"]) < target_version
            ):
                final_path.unlink(missing_ok=True)
                self.repository.update_version_commit(commit_id, "failed", {"message": "数据库未发布，允许从检查点重试提交。"})
            raise
        return final_path, report, profile

    def recover_incomplete_commits(self) -> None:
        for record in self.repository.list_incomplete_version_commits():
            task_id = record["task_id"]
            try:
                task_root = self.files.task_dir(task_id).resolve()
                candidate = Path(record["candidate_path"]).resolve()
                final_path = Path(record["final_path"]).resolve()
                expected_final = self.files.version_path(task_id, int(record["target_version"])).resolve()
                versions_dir = (task_root / "versions").resolve()
                if candidate.parent != versions_dir or final_path != expected_final:
                    raise ValueError("提交路径不在受管版本目录中。")
                expected_hash = record["candidate_sha256"]
                if final_path.is_file():
                    if sha256_file(final_path) != expected_hash:
                        quarantine = self.files.quarantine_path(
                            task_id, f"{record['commit_id']}-conflict.xlsx"
                        )
                        shutil.move(str(final_path), str(quarantine))
                        raise ValueError("正式路径存在哈希冲突，文件已隔离。")
                    if not self.repository.validate_execution_lease(
                        task_id, record["run_id"], record["fencing_token"]
                    ):
                        raise ValueError("执行租约已经失效。")
                    self._reverify_recovery_candidate(record, final_path)
                    self.repository.update_version_commit(record["commit_id"], "file_committed")
                    self.repository.complete_version_commit(commit_id=record["commit_id"])
                    continue
                if not candidate.is_file() or sha256_file(candidate) != expected_hash:
                    raise ValueError("候选文件不存在或哈希不一致。")
                if not self.repository.validate_execution_lease(
                    task_id, record["run_id"], record["fencing_token"]
                ):
                    raise ValueError("执行租约已经失效。")
                self._reverify_recovery_candidate(record, candidate)
                self._replace_with_retry(candidate, final_path)
                if sha256_file(final_path) != expected_hash:
                    raise ValueError("恢复提交后的正式文件哈希不一致。")
                self.repository.update_version_commit(record["commit_id"], "file_committed")
                self.repository.complete_version_commit(commit_id=record["commit_id"])
            except Exception as exc:
                self.repository.update_version_commit(
                    record["commit_id"],
                    "failed",
                    {"type": type(exc).__name__, "message": str(exc)},
                )
                self.repository.log_event(
                    task_id,
                    "version_commit_recovery_failed",
                    {"commit_id": record["commit_id"], "message": str(exc)},
                )
