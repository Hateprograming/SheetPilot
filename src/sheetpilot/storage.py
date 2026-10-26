from __future__ import annotations

import json
import sqlite3
from hashlib import sha256
from contextlib import contextmanager
from datetime import UTC, datetime
from pathlib import Path
from typing import Any, Iterator
from uuid import uuid4


def utc_now() -> str:
    return datetime.now(UTC).isoformat()


class Repository:
    def __init__(self, path: Path):
        self.path = path
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._backup_if_legacy()
        self._init_schema()

    def _backup_if_legacy(self) -> None:
        if not self.path.is_file() or self.path.stat().st_size == 0:
            return
        source = sqlite3.connect(self.path)
        try:
            has_migrations = source.execute(
                "SELECT 1 FROM sqlite_master WHERE type='table' AND name='schema_migrations'"
            ).fetchone()
            current = (
                source.execute("SELECT MAX(version) FROM schema_migrations").fetchone()[0]
                if has_migrations else 0
            ) or 0
            if int(current) >= 4:
                return
            backup_dir = self.path.parent / "backups"
            backup_dir.mkdir(parents=True, exist_ok=True)
            stamp = datetime.now(UTC).strftime("%Y%m%dT%H%M%SZ")
            label = "pre-v2" if int(current) < 2 else "pre-v3" if int(current) < 3 else "pre-v4"
            backup_path = backup_dir / (
                f"{self.path.stem}-{label}-{stamp}-{uuid4().hex[:8]}{self.path.suffix}"
            )
            target = sqlite3.connect(backup_path)
            try:
                source.backup(target)
            finally:
                target.close()
        finally:
            source.close()

    @contextmanager
    def connect(self) -> Iterator[sqlite3.Connection]:
        connection = sqlite3.connect(self.path)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA busy_timeout=5000")
        try:
            yield connection
            connection.commit()
        finally:
            connection.close()

    def _init_schema(self) -> None:
        with self.connect() as db:
            db.executescript(
                """
                PRAGMA journal_mode=WAL;
                BEGIN EXCLUSIVE;
                CREATE TABLE IF NOT EXISTS tasks (
                    task_id TEXT PRIMARY KEY,
                    owner_id TEXT,
                    status TEXT NOT NULL,
                    original_name TEXT NOT NULL,
                    original_path TEXT NOT NULL,
                    current_path TEXT NOT NULL,
                    current_version INTEGER NOT NULL DEFAULT 0,
                    profile_json TEXT,
                    spec_json TEXT,
                    plan_json TEXT,
                    changeset_json TEXT,
                    verification_json TEXT,
                    error_json TEXT,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS file_versions (
                    task_id TEXT NOT NULL,
                    version INTEGER NOT NULL,
                    path TEXT NOT NULL,
                    sha256 TEXT NOT NULL,
                    parent_version INTEGER,
                    profile_json TEXT,
                    created_at TEXT NOT NULL,
                    PRIMARY KEY (task_id, version)
                );
                CREATE TABLE IF NOT EXISTS events (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    task_id TEXT NOT NULL,
                    event_type TEXT NOT NULL,
                    payload_json TEXT NOT NULL,
                    created_at TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS approvals (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    task_id TEXT NOT NULL,
                    decision TEXT NOT NULL,
                    payload_json TEXT NOT NULL,
                    created_at TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS tool_runs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    task_id TEXT NOT NULL,
                    plan_step_id TEXT NOT NULL,
                    tool_name TEXT NOT NULL,
                    status TEXT NOT NULL,
                    arguments_json TEXT NOT NULL,
                    result_json TEXT,
                    error_json TEXT,
                    started_at TEXT NOT NULL,
                    finished_at TEXT
                );
                CREATE TABLE IF NOT EXISTS plan_runs (
                    run_id TEXT PRIMARY KEY,
                    task_id TEXT NOT NULL,
                    plan_id TEXT NOT NULL,
                    plan_revision INTEGER NOT NULL,
                    base_version INTEGER NOT NULL,
                    base_sha256 TEXT NOT NULL,
                    status TEXT NOT NULL,
                    current_step INTEGER NOT NULL DEFAULT 0,
                    pause_requested INTEGER NOT NULL DEFAULT 0,
                    cancel_requested INTEGER NOT NULL DEFAULT 0,
                    last_checkpoint_path TEXT,
                    error_json TEXT,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS conversation_messages (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    task_id TEXT NOT NULL,
                    role TEXT NOT NULL,
                    kind TEXT NOT NULL,
                    content TEXT NOT NULL,
                    metadata_json TEXT NOT NULL DEFAULT '{}',
                    created_at TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS schema_migrations (
                    version INTEGER PRIMARY KEY,
                    applied_at TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS version_commits (
                    commit_id TEXT PRIMARY KEY,
                    run_id TEXT NOT NULL,
                    task_id TEXT NOT NULL,
                    target_version INTEGER NOT NULL,
                    candidate_path TEXT NOT NULL,
                    final_path TEXT NOT NULL,
                    candidate_sha256 TEXT NOT NULL,
                    fencing_token TEXT NOT NULL,
                    status TEXT NOT NULL,
                    profile_json TEXT,
                    changeset_json TEXT,
                    verification_json TEXT,
                    error_json TEXT,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    UNIQUE(task_id, target_version)
                );
                CREATE TABLE IF NOT EXISTS derived_column_lineage (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    task_id TEXT NOT NULL,
                    version INTEGER NOT NULL,
                    sheet_name TEXT NOT NULL,
                    output_column TEXT NOT NULL,
                    output_column_index INTEGER,
                    source_columns_json TEXT NOT NULL,
                    operation TEXT NOT NULL,
                    parameters_json TEXT NOT NULL,
                    status TEXT NOT NULL DEFAULT 'active',
                    created_at TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS planning_runs (
                    job_id TEXT PRIMARY KEY,
                    task_id TEXT NOT NULL,
                    request_message_id INTEGER NOT NULL,
                    fencing_token TEXT NOT NULL,
                    status TEXT NOT NULL,
                    stage TEXT NOT NULL,
                    cancel_requested INTEGER NOT NULL DEFAULT 0,
                    result_message_id INTEGER,
                    error_json TEXT,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS activity_events (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    task_id TEXT NOT NULL,
                    run_kind TEXT NOT NULL,
                    run_id TEXT NOT NULL,
                    stage TEXT NOT NULL,
                    label TEXT NOT NULL,
                    status TEXT NOT NULL,
                    metadata_json TEXT NOT NULL DEFAULT '{}',
                    created_at TEXT NOT NULL
                );
                CREATE INDEX IF NOT EXISTS idx_conversation_task
                ON conversation_messages(task_id, id);
                CREATE INDEX IF NOT EXISTS idx_lineage_task_version
                ON derived_column_lineage(task_id, version, sheet_name);
                CREATE INDEX IF NOT EXISTS idx_planning_task
                ON planning_runs(task_id, created_at);
                CREATE INDEX IF NOT EXISTS idx_activity_run
                ON activity_events(run_kind, run_id, id);
                """
            )
            version_columns = {
                row["name"] for row in db.execute("PRAGMA table_info(file_versions)").fetchall()
            }
            if "profile_json" not in version_columns:
                db.execute("ALTER TABLE file_versions ADD COLUMN profile_json TEXT")
            if "delivery_json" not in version_columns:
                db.execute("ALTER TABLE file_versions ADD COLUMN delivery_json TEXT")
            tool_columns = {
                row["name"] for row in db.execute("PRAGMA table_info(tool_runs)").fetchall()
            }
            additions = {
                "run_id": "TEXT",
                "step_order": "INTEGER",
                "attempt": "INTEGER NOT NULL DEFAULT 1",
                "input_sha256": "TEXT",
                "output_sha256": "TEXT",
                "checkpoint_path": "TEXT",
                "failure_category": "TEXT",
            }
            for name, sql_type in additions.items():
                if name not in tool_columns:
                    db.execute(f"ALTER TABLE tool_runs ADD COLUMN {name} {sql_type}")

            task_columns = {
                row["name"] for row in db.execute("PRAGMA table_info(tasks)").fetchall()
            }
            task_additions = {
                "original_sha256": "TEXT",
                "integrity_level": "TEXT NOT NULL DEFAULT 'verified'",
                "owner_id": "TEXT",
            }
            for name, sql_type in task_additions.items():
                if name not in task_columns:
                    db.execute(f"ALTER TABLE tasks ADD COLUMN {name} {sql_type}")

            plan_columns = {
                row["name"] for row in db.execute("PRAGMA table_info(plan_runs)").fetchall()
            }
            if "fencing_token" not in plan_columns:
                db.execute("ALTER TABLE plan_runs ADD COLUMN fencing_token TEXT")
            duplicate_active = db.execute(
                """SELECT task_id FROM plan_runs
                WHERE status IN ('running','paused','failed')
                GROUP BY task_id HAVING COUNT(*) > 1"""
            ).fetchall()
            for duplicate in duplicate_active:
                active_rows = db.execute(
                    """SELECT run_id FROM plan_runs WHERE task_id=?
                    AND status IN ('running','paused','failed')
                    ORDER BY updated_at DESC, created_at DESC""",
                    (duplicate["task_id"],),
                ).fetchall()
                for stale in active_rows[1:]:
                    db.execute(
                        "UPDATE plan_runs SET status='superseded',updated_at=? WHERE run_id=?",
                        (utc_now(), stale["run_id"]),
                    )
            db.execute(
                """CREATE UNIQUE INDEX IF NOT EXISTS idx_one_active_plan_run
                ON plan_runs(task_id)
                WHERE status IN ('running','paused','failed')"""
            )
            db.execute(
                "CREATE INDEX IF NOT EXISTS idx_version_commits_status ON version_commits(status)"
            )
            commit_columns = {
                row["name"] for row in db.execute("PRAGMA table_info(version_commits)").fetchall()
            }
            for name in ("profile_json", "changeset_json", "verification_json"):
                if name not in commit_columns:
                    db.execute(f"ALTER TABLE version_commits ADD COLUMN {name} TEXT")
            db.execute(
                "INSERT OR IGNORE INTO schema_migrations(version,applied_at) VALUES (2,?)",
                (utc_now(),),
            )
            db.execute(
                "INSERT OR IGNORE INTO schema_migrations(version,applied_at) VALUES (3,?)",
                (utc_now(),),
            )
            db.execute(
                "INSERT OR IGNORE INTO schema_migrations(version,applied_at) VALUES (4,?)",
                (utc_now(),),
            )
        self._backfill_integrity_baselines()

    @staticmethod
    def _hash_path(path: Path) -> str:
        digest = sha256()
        with path.open("rb") as handle:
            for block in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()

    def _backfill_integrity_baselines(self) -> None:
        """Idempotently establish a baseline for databases created before schema v2."""
        with self.connect() as db:
            rows = db.execute(
                "SELECT task_id,original_path FROM tasks WHERE original_sha256 IS NULL"
            ).fetchall()
            for row in rows:
                original = Path(row["original_path"])
                if not original.is_file():
                    db.execute(
                        "UPDATE tasks SET integrity_level='missing',updated_at=? WHERE task_id=?",
                        (utc_now(), row["task_id"]),
                    )
                    continue
                version = db.execute(
                    "SELECT path,sha256 FROM file_versions WHERE task_id=? AND version=0",
                    (row["task_id"],),
                ).fetchone()
                same_file = False
                if version:
                    try:
                        same_file = Path(version["path"]).resolve() == original.resolve()
                    except OSError:
                        same_file = False
                baseline = version["sha256"] if version and same_file else self._hash_path(original)
                level = "verified" if version and same_file else "legacy_rebased"
                db.execute(
                    """UPDATE tasks SET original_sha256=?,integrity_level=?,updated_at=?
                    WHERE task_id=?""",
                    (baseline, level, utc_now(), row["task_id"]),
                )
                db.execute(
                    "INSERT INTO events(task_id,event_type,payload_json,created_at) VALUES (?,?,?,?)",
                    (
                        row["task_id"],
                        "integrity_baseline_migrated",
                        json.dumps({"integrity_level": level}, ensure_ascii=False),
                        utc_now(),
                    ),
                )

    def create_task(
        self, task_id: str, original_name: str, original_path: Path,
        original_sha256: str | None = None,
        owner_id: str | None = None,
    ) -> None:
        now = utc_now()
        with self.connect() as db:
            db.execute(
                """INSERT INTO tasks
                (task_id,owner_id,status,original_name,original_path,current_path,current_version,
                 original_sha256,integrity_level,created_at,updated_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                (
                    task_id, owner_id, "uploaded", original_name, str(original_path), str(original_path), 0,
                    original_sha256 or self._hash_path(original_path), "verified", now, now,
                ),
            )

    def update_task(self, task_id: str, **fields: Any) -> None:
        allowed = {
            "status", "current_path", "current_version", "profile_json", "spec_json",
            "plan_json", "changeset_json", "verification_json", "error_json",
            "original_sha256", "integrity_level", "original_name",
        }
        updates = {key: value for key, value in fields.items() if key in allowed}
        if not updates:
            return
        updates["updated_at"] = utc_now()
        sql = ", ".join(f"{name}=?" for name in updates)
        with self.connect() as db:
            db.execute(f"UPDATE tasks SET {sql} WHERE task_id=?", (*updates.values(), task_id))

    def get_task(self, task_id: str) -> dict[str, Any] | None:
        with self.connect() as db:
            row = db.execute("SELECT * FROM tasks WHERE task_id=?", (task_id,)).fetchone()
        return dict(row) if row else None

    def recent_tasks(
        self, limit: int | None = 20, *, owner_id: str | None = None
    ) -> list[dict[str, Any]]:
        with self.connect() as db:
            if owner_id is None:
                rows = db.execute(
                    "SELECT * FROM tasks ORDER BY updated_at DESC LIMIT ?", (limit if limit is not None else -1,)
                ).fetchall()
            else:
                rows = db.execute(
                    "SELECT * FROM tasks WHERE owner_id=? ORDER BY updated_at DESC LIMIT ?",
                    (owner_id, limit if limit is not None else -1),
                ).fetchall()
        return [dict(row) for row in rows]

    def delete_task(self, task_id: str) -> None:
        related_tables = (
            "p0_feedback", "p0_grants", "p0_measurements", "p0_links", "p0_operations",
            "activity_events", "planning_runs", "derived_column_lineage",
            "version_commits", "conversation_messages", "plan_runs", "tool_runs",
            "approvals", "events", "file_versions",
        )
        with self.connect() as db:
            tables = {row[0] for row in db.execute("SELECT name FROM sqlite_master WHERE type='table'")}
            for table in related_tables:
                if table in tables:
                    db.execute(f"DELETE FROM {table} WHERE task_id=?", (task_id,))
            db.execute("DELETE FROM tasks WHERE task_id=?", (task_id,))

    def add_version(
        self,
        task_id: str,
        version: int,
        path: Path,
        sha256: str,
        parent_version: int | None,
        profile_json: str | None = None,
    ) -> None:
        with self.connect() as db:
            db.execute(
                """INSERT INTO file_versions
                (task_id,version,path,sha256,parent_version,profile_json,created_at)
                VALUES (?,?,?,?,?,?,?)""",
                (task_id, version, str(path), sha256, parent_version, profile_json, utc_now()),
            )

    def get_version(self, task_id: str, version: int) -> dict[str, Any] | None:
        with self.connect() as db:
            row = db.execute(
                "SELECT * FROM file_versions WHERE task_id=? AND version=?",
                (task_id, version),
            ).fetchone()
        return dict(row) if row else None

    def list_versions(self, task_id: str) -> list[dict[str, Any]]:
        with self.connect() as db:
            rows = db.execute(
                "SELECT * FROM file_versions WHERE task_id=? ORDER BY version",
                (task_id,),
            ).fetchall()
        return [dict(row) for row in rows]

    def update_version_profile(self, task_id: str, version: int, profile_json: str) -> None:
        with self.connect() as db:
            db.execute(
                "UPDATE file_versions SET profile_json=? WHERE task_id=? AND version=?",
                (profile_json, task_id, version),
            )

    def copy_lineage(self, task_id: str, source_version: int, target_version: int) -> None:
        """Carry active dependencies forward to the newly committed workbook version."""
        with self.connect() as db:
            db.execute(
                """INSERT INTO derived_column_lineage
                (task_id,version,sheet_name,output_column,output_column_index,
                 source_columns_json,operation,parameters_json,status,created_at)
                SELECT task_id,?,sheet_name,output_column,output_column_index,
                       source_columns_json,operation,parameters_json,status,?
                FROM derived_column_lineage
                WHERE task_id=? AND version=? AND status='active'""",
                (target_version, utc_now(), task_id, source_version),
            )

    def add_lineage(
        self,
        task_id: str,
        version: int,
        *,
        sheet_name: str,
        output_column: str,
        output_column_index: int | None,
        source_columns: list[dict[str, Any]],
        operation: str,
        parameters: dict[str, Any],
        status: str = "active",
    ) -> None:
        with self.connect() as db:
            db.execute(
                """UPDATE derived_column_lineage SET status='superseded'
                WHERE task_id=? AND version=? AND sheet_name=? AND output_column=?
                AND status='active'""",
                (task_id, version, sheet_name, output_column),
            )
            db.execute(
                """INSERT INTO derived_column_lineage
                (task_id,version,sheet_name,output_column,output_column_index,
                 source_columns_json,operation,parameters_json,status,created_at)
                VALUES (?,?,?,?,?,?,?,?,?,?)""",
                (
                    task_id, version, sheet_name, output_column, output_column_index,
                    json.dumps(source_columns, ensure_ascii=False, default=str), operation,
                    json.dumps(parameters, ensure_ascii=False, default=str), status, utc_now(),
                ),
            )

    def list_lineage(
        self, task_id: str, version: int, *, sheet_name: str | None = None
    ) -> list[dict[str, Any]]:
        sql = (
            "SELECT * FROM derived_column_lineage "
            "WHERE task_id=? AND version=? AND status='active'"
        )
        params: list[Any] = [task_id, version]
        if sheet_name is not None:
            sql += " AND sheet_name=?"
            params.append(sheet_name)
        sql += " ORDER BY id"
        with self.connect() as db:
            rows = db.execute(sql, params).fetchall()
        result: list[dict[str, Any]] = []
        for raw in rows:
            item = dict(raw)
            item["source_columns"] = json.loads(item.pop("source_columns_json"))
            item["parameters"] = json.loads(item.pop("parameters_json"))
            result.append(item)
        return result

    def log_event(self, task_id: str, event_type: str, payload: dict[str, Any]) -> None:
        with self.connect() as db:
            db.execute(
                "INSERT INTO events(task_id,event_type,payload_json,created_at) VALUES (?,?,?,?)",
                (task_id, event_type, json.dumps(payload, ensure_ascii=False, default=str), utc_now()),
            )

    def add_approval(self, task_id: str, decision: str, payload: dict[str, Any]) -> None:
        with self.connect() as db:
            db.execute(
                "INSERT INTO approvals(task_id,decision,payload_json,created_at) VALUES (?,?,?,?)",
                (task_id, decision, json.dumps(payload, ensure_ascii=False), utc_now()),
            )

    def start_tool_run(
        self,
        task_id: str,
        plan_step_id: str,
        tool_name: str,
        arguments: dict[str, Any],
        *,
        run_id: str | None = None,
        step_order: int | None = None,
        attempt: int = 1,
        input_sha256: str | None = None,
    ) -> int:
        with self.connect() as db:
            cursor = db.execute(
                """INSERT INTO tool_runs
                (task_id,plan_step_id,tool_name,status,arguments_json,started_at,
                 run_id,step_order,attempt,input_sha256)
                VALUES (?,?,?,?,?,?,?,?,?,?)""",
                (
                    task_id,
                    plan_step_id,
                    tool_name,
                    "running",
                    json.dumps(arguments, ensure_ascii=False, default=str),
                    utc_now(),
                    run_id,
                    step_order,
                    attempt,
                    input_sha256,
                ),
            )
            return int(cursor.lastrowid)

    def finish_tool_run(
        self,
        run_id: int,
        status: str,
        *,
        result: dict[str, Any] | None = None,
        error: dict[str, Any] | None = None,
        output_sha256: str | None = None,
        checkpoint_path: str | None = None,
        failure_category: str | None = None,
    ) -> None:
        with self.connect() as db:
            db.execute(
                """UPDATE tool_runs
                SET status=?,result_json=?,error_json=?,finished_at=?,output_sha256=?,
                    checkpoint_path=?,failure_category=? WHERE id=?""",
                (
                    status,
                    json.dumps(result, ensure_ascii=False, default=str) if result else None,
                    json.dumps(error, ensure_ascii=False, default=str) if error else None,
                    utc_now(),
                    output_sha256,
                    checkpoint_path,
                    failure_category,
                    run_id,
                ),
            )

    def create_plan_run(
        self,
        run_id: str,
        task_id: str,
        plan_id: str,
        revision: int,
        base_version: int,
        base_sha256: str,
    ) -> str:
        now = utc_now()
        fencing_token = uuid4().hex
        with self.connect() as db:
            db.execute(
                """INSERT INTO plan_runs
                (run_id,task_id,plan_id,plan_revision,base_version,base_sha256,status,
                 current_step,fencing_token,created_at,updated_at)
                VALUES (?,?,?,?,?,?,?,0,?,?,?)""",
                (
                    run_id, task_id, plan_id, revision, base_version, base_sha256,
                    "running", fencing_token, now, now,
                ),
            )
        return fencing_token

    def acquire_plan_run(
        self,
        run_id: str,
        task_id: str,
        plan_id: str,
        revision: int,
        base_version: int,
        base_sha256: str,
        current_sha256: str,
    ) -> dict[str, Any]:
        """Atomically acquire the sole execution lease for a task."""
        now = utc_now()
        fencing_token = uuid4().hex
        with self.connect() as db:
            db.execute("BEGIN IMMEDIATE")
            task = db.execute("SELECT * FROM tasks WHERE task_id=?", (task_id,)).fetchone()
            if not task:
                raise ValueError("任务不存在。")
            if task["status"] != "approved":
                raise ValueError("任务尚未获得批准，或已由其他运行占用。")
            if int(task["current_version"]) != base_version or current_sha256 != base_sha256:
                raise ValueError("文件版本已经变化，请重新生成并批准执行计划。")
            if task["plan_json"]:
                stored = json.loads(task["plan_json"])
                if stored.get("plan_id") != plan_id or int(stored.get("revision", 0)) != revision:
                    raise ValueError("计划已经修订，旧审批不再有效。")
            active = db.execute(
                """SELECT run_id FROM plan_runs WHERE task_id=?
                AND status IN ('running','paused','failed') LIMIT 1""",
                (task_id,),
            ).fetchone()
            if active:
                raise ValueError("该任务已有运行正在处理。")
            db.execute(
                """INSERT INTO plan_runs
                (run_id,task_id,plan_id,plan_revision,base_version,base_sha256,status,
                 current_step,fencing_token,created_at,updated_at)
                VALUES (?,?,?,?,?,?, 'running',0,?,?,?)""",
                (
                    run_id, task_id, plan_id, revision, base_version, base_sha256,
                    fencing_token, now, now,
                ),
            )
            db.execute(
                "UPDATE tasks SET status='executing',error_json=NULL,updated_at=? WHERE task_id=?",
                (now, task_id),
            )
        return {
            "run_id": run_id,
            "task_id": task_id,
            "fencing_token": fencing_token,
            "plan_id": plan_id,
            "plan_revision": revision,
            "base_version": base_version,
            "base_sha256": base_sha256,
        }

    def validate_execution_lease(
        self, task_id: str, run_id: str, fencing_token: str
    ) -> bool:
        with self.connect() as db:
            row = db.execute(
                """SELECT 1 FROM plan_runs WHERE task_id=? AND run_id=? AND fencing_token=?
                AND cancel_requested=0 AND status IN ('running','paused','failed')""",
                (task_id, run_id, fencing_token),
            ).fetchone()
        return row is not None

    def update_plan_run(self, run_id: str, **fields: Any) -> None:
        allowed = {
            "status", "current_step", "pause_requested", "cancel_requested",
            "last_checkpoint_path", "error_json", "plan_revision", "fencing_token",
        }
        updates = {key: value for key, value in fields.items() if key in allowed}
        if not updates:
            return
        updates["updated_at"] = utc_now()
        sql = ", ".join(f"{name}=?" for name in updates)
        with self.connect() as db:
            db.execute(f"UPDATE plan_runs SET {sql} WHERE run_id=?", (*updates.values(), run_id))

    def get_plan_run(self, run_id: str) -> dict[str, Any] | None:
        with self.connect() as db:
            row = db.execute("SELECT * FROM plan_runs WHERE run_id=?", (run_id,)).fetchone()
        return dict(row) if row else None

    def latest_plan_run(self, task_id: str) -> dict[str, Any] | None:
        with self.connect() as db:
            row = db.execute(
                "SELECT * FROM plan_runs WHERE task_id=? ORDER BY created_at DESC LIMIT 1",
                (task_id,),
            ).fetchone()
        return dict(row) if row else None

    def request_run_control(self, run_id: str, action: str) -> None:
        if action not in {"pause", "cancel", "resume"}:
            raise ValueError("invalid run control action")
        values = {
            "pause": {"pause_requested": 1},
            "cancel": {"cancel_requested": 1},
            "resume": {"pause_requested": 0, "status": "running"},
        }[action]
        self.update_plan_run(run_id, **values)

    def list_tool_runs(self, task_id: str) -> list[dict[str, Any]]:
        with self.connect() as db:
            rows = db.execute(
                "SELECT * FROM tool_runs WHERE task_id=? ORDER BY id",
                (task_id,),
            ).fetchall()
        return [dict(row) for row in rows]

    def create_version_commit(
        self,
        *,
        commit_id: str,
        run_id: str,
        task_id: str,
        target_version: int,
        candidate_path: Path,
        final_path: Path,
        candidate_sha256: str,
        fencing_token: str,
        profile_json: str,
        changeset_json: str,
        verification_json: str,
    ) -> dict[str, Any]:
        now = utc_now()
        with self.connect() as db:
            db.execute("BEGIN IMMEDIATE")
            previous = db.execute(
                "SELECT * FROM version_commits WHERE task_id=? AND target_version=?",
                (task_id, target_version),
            ).fetchone()
            if previous:
                published = db.execute("SELECT 1 FROM file_versions WHERE task_id=? AND version=?", (task_id, target_version)).fetchone()
                lease = db.execute("SELECT 1 FROM plan_runs WHERE task_id=? AND run_id=? AND fencing_token=? AND status='running' AND cancel_requested=0", (task_id, run_id, fencing_token)).fetchone()
                if previous["status"] != "failed" or published or not lease:
                    raise ValueError("目标版本已有提交记录，不能覆盖已发布或待恢复的提交。")
                # Keep the failed attempt in the audit log before freeing the
                # unique target-version slot for a verified retry.
                db.execute("INSERT INTO events(task_id,event_type,payload_json,created_at) VALUES (?,?,?,?)", (task_id, "version_commit_retried", json.dumps(dict(previous), ensure_ascii=False), now))
                db.execute("DELETE FROM version_commits WHERE commit_id=?", (previous["commit_id"],))
            db.execute(
                """INSERT INTO version_commits
                (commit_id,run_id,task_id,target_version,candidate_path,final_path,
                 candidate_sha256,fencing_token,status,profile_json,changeset_json,
                 verification_json,created_at,updated_at)
                VALUES (?,?,?,?,?,?,?,?, 'prepared',?,?,?,?,?)""",
                (
                    commit_id, run_id, task_id, target_version, str(candidate_path),
                    str(final_path), candidate_sha256, fencing_token, profile_json,
                    changeset_json, verification_json, now, now,
                ),
            )
        return self.get_version_commit(commit_id) or {}

    def update_version_commit(
        self, commit_id: str, status: str, error: dict[str, Any] | None = None
    ) -> None:
        with self.connect() as db:
            db.execute(
                """UPDATE version_commits SET status=?,error_json=?,updated_at=?
                WHERE commit_id=?""",
                (
                    status,
                    json.dumps(error, ensure_ascii=False, default=str) if error else None,
                    utc_now(),
                    commit_id,
                ),
            )

    def get_version_commit(self, commit_id: str) -> dict[str, Any] | None:
        with self.connect() as db:
            row = db.execute(
                "SELECT * FROM version_commits WHERE commit_id=?", (commit_id,)
            ).fetchone()
        return dict(row) if row else None

    def list_incomplete_version_commits(self) -> list[dict[str, Any]]:
        with self.connect() as db:
            rows = db.execute(
                """SELECT * FROM version_commits
                WHERE status IN ('prepared','file_committed') ORDER BY created_at"""
            ).fetchall()
        return [dict(row) for row in rows]

    def complete_version_commit(
        self,
        *,
        commit_id: str,
        profile_json: str | None = None,
        changeset_json: str | None = None,
        verification_json: str | None = None,
    ) -> None:
        """Atomically publish DB metadata after the file has reached its final path."""
        with self.connect() as db:
            db.execute("BEGIN IMMEDIATE")
            record = db.execute(
                "SELECT * FROM version_commits WHERE commit_id=?", (commit_id,)
            ).fetchone()
            if not record:
                raise ValueError("提交记录不存在。")
            if record["status"] == "completed":
                return
            lease = db.execute(
                """SELECT * FROM plan_runs WHERE run_id=? AND task_id=?
                AND fencing_token=? AND cancel_requested=0
                AND status IN ('running','paused','failed')""",
                (record["run_id"], record["task_id"], record["fencing_token"]),
            ).fetchone()
            if not lease:
                raise ValueError("执行租约已经失效。")
            task = db.execute(
                "SELECT * FROM tasks WHERE task_id=?", (record["task_id"],)
            ).fetchone()
            expected_parent = int(record["target_version"]) - 1
            if not task or int(task["current_version"]) != expected_parent:
                raise ValueError("正式版本已经变化，无法提交候选文件。")
            existing = db.execute(
                "SELECT sha256 FROM file_versions WHERE task_id=? AND version=?",
                (record["task_id"], record["target_version"]),
            ).fetchone()
            if existing:
                if existing["sha256"] == record["candidate_sha256"]:
                    db.execute(
                        "UPDATE version_commits SET status='completed',updated_at=? WHERE commit_id=?",
                        (utc_now(), commit_id),
                    )
                    return
                raise ValueError("目标版本已经存在且哈希不一致。")
            delivery = json.loads(changeset_json or record["changeset_json"]).get("details", {}).get("delivery")
            db.execute(
                """INSERT INTO file_versions
                (task_id,version,path,sha256,parent_version,profile_json,delivery_json,created_at)
                VALUES (?,?,?,?,?,?,?,?)""",
                (
                    record["task_id"], record["target_version"], record["final_path"],
                    record["candidate_sha256"], expected_parent,
                    profile_json or record["profile_json"],
                    json.dumps(delivery, ensure_ascii=False) if delivery else None,
                    utc_now(),
                ),
            )
            db.execute(
                """UPDATE tasks SET status='completed',current_path=?,current_version=?,
                profile_json=?,changeset_json=?,verification_json=?,error_json=NULL,updated_at=?
                WHERE task_id=?""",
                (
                    record["final_path"], record["target_version"],
                    profile_json or record["profile_json"],
                    changeset_json or record["changeset_json"],
                    verification_json or record["verification_json"],
                    utc_now(), record["task_id"],
                ),
            )
            db.execute(
                """UPDATE plan_runs SET status='completed',
                current_step=COALESCE((SELECT MAX(step_order) FROM tool_runs WHERE run_id=?),current_step),
                updated_at=?
                WHERE run_id=? AND fencing_token=?""",
                (record["run_id"], utc_now(), record["run_id"], record["fencing_token"]),
            )
            db.execute(
                """UPDATE version_commits SET status='completed',updated_at=?
                WHERE commit_id=?""",
                (utc_now(), commit_id),
            )
            # Publish the operation's version in the same transaction as the
            # formal file version. A crash cannot leave a committed result
            # looking like an unrelated unfinished operation.
            if db.execute("SELECT 1 FROM sqlite_master WHERE name='p0_links'").fetchone():
                link = db.execute("SELECT operation_id FROM p0_links WHERE kind='execution' AND link_id=?", (record['run_id'],)).fetchone()
                if link:
                    from .p0 import P0Store
                    result = json.loads(changeset_json or record['changeset_json'])
                    verification = json.loads(verification_json or record['verification_json'])
                    summary = P0Store.summary(result, verification)
                    db.execute("UPDATE p0_operations SET status='completed',result_version=?,summary_json=?,updated_at=? WHERE operation_id=?",
                               (record['target_version'], json.dumps(summary,ensure_ascii=False),utc_now(),link['operation_id']))

    def add_message(
        self,
        task_id: str,
        role: str,
        kind: str,
        content: str,
        metadata: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        created_at = utc_now()
        metadata_json = json.dumps(metadata or {}, ensure_ascii=False, default=str)
        with self.connect() as db:
            cursor = db.execute(
                """INSERT INTO conversation_messages
                (task_id,role,kind,content,metadata_json,created_at)
                VALUES (?,?,?,?,?,?)""",
                (task_id, role, kind, content, metadata_json, created_at),
            )
            message_id = int(cursor.lastrowid)
        return {
            "id": message_id,
            "task_id": task_id,
            "role": role,
            "kind": kind,
            "content": content,
            "metadata": metadata or {},
            "created_at": created_at,
        }

    def create_planning_run(
        self,
        task_id: str,
        role: str,
        kind: str,
        content: str,
        *,
        waiting_for_execution: bool = False,
    ) -> dict[str, Any]:
        """Persist the user message and supersede any older planning request atomically."""
        now = utc_now()
        job_id = uuid4().hex
        fencing_token = uuid4().hex
        status = "waiting_for_execution" if waiting_for_execution else "queued"
        with self.connect() as db:
            db.execute("BEGIN IMMEDIATE")
            task_before = db.execute(
                "SELECT * FROM tasks WHERE task_id=?", (task_id,)
            ).fetchone()
            if not task_before or dict(task_before).get("cleanup_status") == "pending":
                raise ValueError("工作簿不存在或正在清理。")
            resuming_clarification = bool(
                task_before and task_before["status"] == "needs_clarification"
            )
            superseded = db.execute(
                """SELECT COUNT(*) FROM planning_runs WHERE task_id=?
                AND status IN ('queued','waiting_for_execution','running','waiting_for_user')""",
                (task_id,),
            ).fetchone()[0]
            db.execute(
                """UPDATE planning_runs
                SET status='superseded',cancel_requested=1,updated_at=?
                WHERE task_id=? AND status IN ('queued','waiting_for_execution','running','waiting_for_user')""",
                (now, task_id),
            )
            cursor = db.execute(
                """INSERT INTO conversation_messages
                (task_id,role,kind,content,metadata_json,created_at)
                VALUES (?,?,?,?,?,?)""",
                (task_id, role, kind, content, "{}", now),
            )
            message_id = int(cursor.lastrowid)
            db.execute(
                """INSERT INTO planning_runs
                (job_id,task_id,request_message_id,fencing_token,status,stage,
                 cancel_requested,created_at,updated_at)
                VALUES (?,?,?,?,?,?,0,?,?)""",
                (job_id, task_id, message_id, fencing_token, status, "received", now, now),
            )
            db.execute(
                "UPDATE tasks SET status='analyzing_request',error_json=NULL,updated_at=? WHERE task_id=?",
                (now, task_id),
            )
            db.execute(
                """INSERT INTO activity_events
                (task_id,run_kind,run_id,stage,label,status,metadata_json,created_at)
                VALUES (?,?,?,?,?,?,?,?)""",
                (
                    task_id, "planning", job_id, "received",
                    "已收到补充规则，继续完善当前计划"
                    if resuming_clarification else "已收到你的要求",
                    "completed", "{}", now,
                ),
            )
            if superseded:
                db.execute(
                    """INSERT INTO activity_events
                    (task_id,run_kind,run_id,stage,label,status,metadata_json,created_at)
                    VALUES (?,?,?,?,?,?,?,?)""",
                    (
                        task_id, "planning", job_id, "superseded_previous",
                        (
                            "已收到回答，继续完善上一份计划草案"
                            if resuming_clarification
                            else "已停止采用旧请求，正在按新要求规划"
                        ),
                        "completed",
                        json.dumps({
                            "superseded_runs": int(superseded),
                            "resumed_clarification": resuming_clarification,
                        }),
                        now,
                    ),
                )
        return {
            "job_id": job_id,
            "task_id": task_id,
            "request_message_id": message_id,
            "fencing_token": fencing_token,
            "status": status,
            "stage": "received",
            "created_at": now,
            "updated_at": now,
        }

    def get_message(self, message_id: int) -> dict[str, Any] | None:
        with self.connect() as db:
            row = db.execute(
                "SELECT * FROM conversation_messages WHERE id=?", (message_id,)
            ).fetchone()
        if not row:
            return None
        return {**dict(row), "metadata": json.loads(row["metadata_json"] or "{}")}

    def get_planning_run(self, job_id: str) -> dict[str, Any] | None:
        with self.connect() as db:
            row = db.execute("SELECT * FROM planning_runs WHERE job_id=?", (job_id,)).fetchone()
        return dict(row) if row else None

    def latest_planning_run(self, task_id: str) -> dict[str, Any] | None:
        with self.connect() as db:
            row = db.execute(
                "SELECT * FROM planning_runs WHERE task_id=? ORDER BY created_at DESC LIMIT 1",
                (task_id,),
            ).fetchone()
        return dict(row) if row else None

    def update_planning_run(self, job_id: str, **fields: Any) -> None:
        allowed = {"status", "stage", "cancel_requested", "result_message_id", "error_json"}
        updates = {key: value for key, value in fields.items() if key in allowed}
        if not updates:
            return
        updates["updated_at"] = utc_now()
        sql = ", ".join(f"{name}=?" for name in updates)
        with self.connect() as db:
            db.execute(f"UPDATE planning_runs SET {sql} WHERE job_id=?", (*updates.values(), job_id))

    def planning_run_is_current(self, job_id: str, fencing_token: str) -> bool:
        with self.connect() as db:
            row = db.execute(
                """SELECT 1 FROM planning_runs WHERE job_id=? AND fencing_token=?
                AND cancel_requested=0 AND status IN ('queued','waiting_for_execution','running')""",
                (job_id, fencing_token),
            ).fetchone()
        return row is not None

    def add_activity_event(
        self,
        task_id: str,
        run_kind: str,
        run_id: str,
        stage: str,
        label: str,
        status: str,
        metadata: dict[str, Any] | None = None,
    ) -> None:
        with self.connect() as db:
            db.execute(
                """INSERT INTO activity_events
                (task_id,run_kind,run_id,stage,label,status,metadata_json,created_at)
                VALUES (?,?,?,?,?,?,?,?)""",
                (
                    task_id, run_kind, run_id, stage, label, status,
                    json.dumps(metadata or {}, ensure_ascii=False, default=str), utc_now(),
                ),
            )

    def list_activity_events(self, task_id: str, limit: int = 200) -> list[dict[str, Any]]:
        with self.connect() as db:
            rows = db.execute(
                """SELECT * FROM (
                    SELECT * FROM activity_events WHERE task_id=? ORDER BY id DESC LIMIT ?
                ) ORDER BY id""",
                (task_id, limit),
            ).fetchall()
        return [
            {**dict(row), "metadata": json.loads(row["metadata_json"] or "{}")}
            for row in rows
        ]

    def publish_planning_result(
        self,
        *,
        job_id: str,
        fencing_token: str,
        task_status: str,
        spec_json: str | None,
        plan_json: str | None,
        message_role: str,
        message_kind: str,
        message_content: str,
        message_metadata: dict[str, Any] | None = None,
        error_json: str | None = None,
        run_status: str = "completed",
        run_stage: str = "completed",
    ) -> dict[str, Any] | None:
        """Publish only if this planning job still owns the task's planning lease."""
        now = utc_now()
        metadata_json = json.dumps(message_metadata or {}, ensure_ascii=False, default=str)
        with self.connect() as db:
            db.execute("BEGIN IMMEDIATE")
            run = db.execute(
                """SELECT * FROM planning_runs WHERE job_id=? AND fencing_token=?
                AND cancel_requested=0 AND status='running'""",
                (job_id, fencing_token),
            ).fetchone()
            if not run:
                return None
            newer = db.execute(
                """SELECT 1 FROM planning_runs WHERE task_id=? AND created_at>?
                AND status IN ('queued','waiting_for_execution','running','completed') LIMIT 1""",
                (run["task_id"], run["created_at"]),
            ).fetchone()
            if newer:
                db.execute(
                    "UPDATE planning_runs SET status='superseded',updated_at=? WHERE job_id=?",
                    (now, job_id),
                )
                return None
            cursor = db.execute(
                """INSERT INTO conversation_messages
                (task_id,role,kind,content,metadata_json,created_at)
                VALUES (?,?,?,?,?,?)""",
                (run["task_id"], message_role, message_kind, message_content, metadata_json, now),
            )
            message_id = int(cursor.lastrowid)
            db.execute(
                """UPDATE tasks SET status=?,spec_json=?,plan_json=?,error_json=?,updated_at=?
                WHERE task_id=?""",
                (task_status, spec_json, plan_json, error_json, now, run["task_id"]),
            )
            db.execute(
                """UPDATE planning_runs SET status=?,stage=?,
                result_message_id=?,updated_at=? WHERE job_id=?""",
                (run_status, run_stage, message_id, now, job_id),
            )
        return {
            "id": message_id,
            "task_id": run["task_id"],
            "role": message_role,
            "kind": message_kind,
            "content": message_content,
            "metadata": message_metadata or {},
            "created_at": now,
        }

    def list_messages(self, task_id: str, limit: int = 200) -> list[dict[str, Any]]:
        with self.connect() as db:
            rows = db.execute(
                """SELECT * FROM (
                    SELECT * FROM conversation_messages
                    WHERE task_id=? ORDER BY id DESC LIMIT ?
                ) ORDER BY id""",
                (task_id, limit),
            ).fetchall()
        return [
            {
                **dict(row),
                "metadata": json.loads(row["metadata_json"] or "{}"),
            }
            for row in rows
        ]
