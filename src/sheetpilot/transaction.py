from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

from .schemas import TaskEnvelope, TaskPlan, ToolResult, VerificationReport

if TYPE_CHECKING:
    from .service import SheetPilotService


class TransactionExecutor:
    """Stable execution boundary used by every current multi-skill workflow."""

    def __init__(self, service: "SheetPilotService"):
        self.service = service

    def execute(
        self,
        task_id: str,
        envelope: TaskEnvelope,
        plan: TaskPlan,
        *,
        run_id: str | None = None,
    ) -> tuple[Path, VerificationReport, ToolResult]:
        return self.service._execute_task_plan_impl(
            task_id, envelope, plan, run_id=run_id
        )
