from __future__ import annotations

from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterator, TypedDict

from .schemas import LookupTaskSpec, TaskEnvelope, TaskPlan, TaskType


class WorkflowState(TypedDict, total=False):
    task_id: str
    user_request: str
    spec: dict[str, Any]
    validation: dict[str, Any]
    plan: dict[str, Any]
    approval: dict[str, Any]
    status: str
    result: dict[str, Any]
    error: dict[str, Any]
    run_id: str
    skill_hint: str | None
    skill_context: dict[str, Any]
    previous_spec: dict[str, Any] | None
    validation_errors: list[str]
    conversation: list[dict[str, Any]]


def build_planning_workflow(service, checkpointer=None):
    """Skill selection → bounded context → plan → composition validation."""
    try:
        from langgraph.checkpoint.memory import MemorySaver
        from langgraph.graph import END, START, StateGraph
    except ImportError as exc:
        raise RuntimeError("Install project dependencies to use the LangGraph workflow.") from exc

    from .context import build_context, guess_task_type

    def select_skill(state: WorkflowState):
        hint = guess_task_type(state["user_request"])
        return {"skill_hint": hint.value if hint else None, "status": "skill_selected"}

    def build_skill_context(state: WorkflowState):
        profile = service.load_profile(state["task_id"])
        hint = state.get("skill_hint")
        context = build_context(profile, state["user_request"])
        if hint and hint != "composite":
            try:
                context = service.skills.get(TaskType(hint)).build_context(profile, state["user_request"])
            except (KeyError, ValueError):
                pass
        return {"skill_context": context, "status": "context_built"}

    def parse_with_contract(state: WorkflowState):
        profile = service.load_profile(state["task_id"])
        previous = TaskEnvelope.model_validate(state["previous_spec"]) if state.get("previous_spec") else None
        envelope = service.configured_llm().parse_task(
            state["user_request"],
            profile,
            previous_envelope=previous,
            validation_errors=state.get("validation_errors", []),
            conversation=state.get("conversation", []),
            capability_manifest=service.skills.manifest(),
            workbook_context=state.get("skill_context"),
        )
        source_text = "\n".join(
            item.get("content", "") for item in state.get("conversation", [])
            if item.get("role") == "user"
        )
        envelope = service.prepare_model_envelope(
            envelope, source_text or state["user_request"], profile
        )
        return {"spec": envelope.model_dump(mode="json"), "status": "intent_parsed"}

    def plan_with_skill(state: WorkflowState):
        envelope = TaskEnvelope.model_validate(state["spec"])
        row = service.repository.get_task(state["task_id"])
        if (
            row
            and Path(row["original_path"]).suffix.casefold() == ".csv"
            and int(row["current_version"]) == 0
            and envelope.spec.task_type != TaskType.CONVERSION
        ):
            raise ValueError("CSV 必须先通过 conversion_skill 转换为 XLSX。")
        service.save_task_envelope(state["task_id"], envelope)
        source_text = "\n".join(
            item.get("content", "") for item in state.get("conversation", [])
            if item.get("role") == "user"
        )
        plan, validation = service.make_task_plan(
            state["task_id"], envelope,
            policy_source_text=source_text or state.get("user_request"),
        )
        return {
            "plan": plan.model_dump(mode="json"),
            "validation": validation.model_dump(mode="json"),
            "status": "plan_validated" if validation.valid else "needs_clarification",
        }

    graph = StateGraph(WorkflowState)
    graph.add_node("select_skill", select_skill)
    graph.add_node("build_skill_context", build_skill_context)
    graph.add_node("plan_with_skill", plan_with_skill)
    graph.add_node("parse_with_contract", parse_with_contract)
    graph.add_edge(START, "select_skill")
    graph.add_edge("select_skill", "build_skill_context")
    graph.add_edge("build_skill_context", "parse_with_contract")
    graph.add_edge("parse_with_contract", "plan_with_skill")
    graph.add_edge("plan_with_skill", END)
    return graph.compile(checkpointer=checkpointer if checkpointer is not None else False)


def build_workflow(service, checkpointer=None):
    """Build the resumable v0.1 graph.

    The Streamlit UI uses the same service methods explicitly so it can render each
    review screen. This graph is the programmatic orchestration entrypoint and keeps
    the approval boundary as a LangGraph interrupt.
    """
    try:
        from langgraph.checkpoint.memory import MemorySaver
        from langgraph.graph import END, START, StateGraph
        from langgraph.types import interrupt
    except ImportError as exc:
        raise RuntimeError("Install project dependencies to use the LangGraph workflow.") from exc

    def parse_intent(state: WorkflowState):
        spec = service.parse_request(state["task_id"], state["user_request"])
        return {"spec": spec.model_dump(mode="json"), "status": "intent_parsed"}

    def validate_and_plan(state: WorkflowState):
        spec = LookupTaskSpec.model_validate(state["spec"])
        plan, validation = service.make_plan(state["task_id"], spec)
        if validation.missing_fields or validation.ambiguities or validation.output_conflicts:
            return {"validation": validation.model_dump(mode="json"), "status": "needs_clarification"}
        return {
            "validation": validation.model_dump(mode="json"),
            "plan": plan.model_dump(mode="json"),
            "status": "awaiting_approval",
        }

    def route_after_validation(state: WorkflowState):
        return "stop" if state["status"] == "needs_clarification" else "approval"

    def approval(state: WorkflowState):
        decision = interrupt({"kind": "approval", "plan": state["plan"], "validation": state["validation"]})
        service.approve(state["task_id"], decision.get("decision", "reject"), decision)
        return {"approval": decision, "status": decision.get("decision", "reject")}

    def route_after_approval(state: WorkflowState):
        return "execute" if state["status"] == "approve" else "stop"

    def execute(state: WorkflowState):
        spec = LookupTaskSpec.model_validate(state["spec"])
        _, report = service.execute(state["task_id"], spec)
        return {"result": report.model_dump(mode="json"), "status": "completed"}

    graph = StateGraph(WorkflowState)
    graph.add_node("parse_intent", parse_intent)
    graph.add_node("validate_and_plan", validate_and_plan)
    graph.add_node("approval", approval)
    graph.add_node("execute", execute)
    graph.add_edge(START, "parse_intent")
    graph.add_edge("parse_intent", "validate_and_plan")
    graph.add_conditional_edges(
        "validate_and_plan", route_after_validation, {"approval": "approval", "stop": END}
    )
    graph.add_conditional_edges("approval", route_after_approval, {"execute": "execute", "stop": END})
    graph.add_edge("execute", END)
    return graph.compile(checkpointer=checkpointer or MemorySaver())


def build_v2_workflow(service, checkpointer=None):
    """Build the v0.2 router → controlled tool workflow.

    Both clarification and approval use durable interrupts. An edited envelope is
    always routed back through planning and validation; execution cannot silently
    expand an approved plan.
    """
    try:
        from langgraph.checkpoint.memory import MemorySaver
        from langgraph.graph import END, START, StateGraph
        from langgraph.types import interrupt
    except ImportError as exc:
        raise RuntimeError("Install project dependencies to use the LangGraph workflow.") from exc

    def parse_task(state: WorkflowState):
        envelope = service.parse_task(state["task_id"], state["user_request"])
        return {"spec": envelope.model_dump(mode="json"), "status": "intent_parsed"}

    def plan_task(state: WorkflowState):
        envelope = TaskEnvelope.model_validate(state["spec"])
        plan, validation = service.make_task_plan(state["task_id"], envelope)
        return {
            "plan": plan.model_dump(mode="json"),
            "validation": validation.model_dump(mode="json"),
            "status": "awaiting_approval" if validation.valid else "needs_clarification",
        }

    def route_after_plan(state: WorkflowState):
        return "approval" if state["status"] == "awaiting_approval" else "clarification"

    def clarification(state: WorkflowState):
        response = interrupt(
            {
                "kind": "clarification",
                "validation": state["validation"],
                "envelope": state["spec"],
            }
        )
        envelope = TaskEnvelope.model_validate(response["envelope"])
        service.save_task_envelope(state["task_id"], envelope)
        return {"spec": envelope.model_dump(mode="json"), "status": "intent_parsed"}

    def approval_v2(state: WorkflowState):
        decision = interrupt(
            {"kind": "approval", "plan": state["plan"], "validation": state["validation"]}
        )
        action = decision.get("decision", "reject")
        service.approve(state["task_id"], action, decision)
        if action == "edit" and decision.get("envelope"):
            envelope = TaskEnvelope.model_validate(decision["envelope"])
            service.save_task_envelope(state["task_id"], envelope)
            return {"spec": envelope.model_dump(mode="json"), "approval": decision, "status": "editing"}
        return {"approval": decision, "status": action}

    def route_after_approval_v2(state: WorkflowState):
        if state["status"] == "approve":
            return "execute"
        if state["status"] == "editing":
            return "plan"
        return "stop"

    def execute_v2(state: WorkflowState):
        envelope = TaskEnvelope.model_validate(state["spec"])
        plan = TaskPlan.model_validate(state["plan"])
        _, report, result = service.execute_task_plan(state["task_id"], envelope, plan)
        return {
            "result": {
                "verification": report.model_dump(mode="json"),
                "tool_result": result.model_dump(mode="json"),
            },
            "status": "completed",
        }

    graph = StateGraph(WorkflowState)
    graph.add_node("parse_task", parse_task)
    graph.add_node("plan_task", plan_task)
    graph.add_node("clarification", clarification)
    graph.add_node("approval", approval_v2)
    graph.add_node("execute", execute_v2)
    graph.add_edge(START, "parse_task")
    graph.add_edge("parse_task", "plan_task")
    graph.add_conditional_edges(
        "plan_task", route_after_plan, {"approval": "approval", "clarification": "clarification"}
    )
    graph.add_edge("clarification", "plan_task")
    graph.add_conditional_edges(
        "approval",
        route_after_approval_v2,
        {"execute": "execute", "plan": "plan_task", "stop": END},
    )
    graph.add_edge("execute", END)
    return graph.compile(checkpointer=checkpointer or MemorySaver())


def build_execution_workflow(service, checkpointer=None):
    """Durable execution graph used by the chat UI after plan approval.

    Domain-level checkpoint files and retries live in SheetPilotService; this graph
    is the single orchestration entrypoint and persists the conversation execution
    boundary with the workbook task as its stable thread.
    """
    try:
        from langgraph.checkpoint.memory import MemorySaver
        from langgraph.graph import END, START, StateGraph
    except ImportError as exc:
        raise RuntimeError("Install project dependencies to use the LangGraph workflow.") from exc

    def validate_approved_plan(state: WorkflowState):
        row = service.repository.get_task(state["task_id"])
        if not row or row["status"] != "approved":
            return {"status": "failed", "error": {"message": "计划尚未获得批准。"}}
        plan = TaskPlan.model_validate(state["plan"])
        if row.get("plan_json"):
            stored = TaskPlan.model_validate_json(row["plan_json"])
            if stored.plan_id != plan.plan_id or stored.revision != plan.revision:
                return {"status": "failed", "error": {"message": "批准已因计划修订而失效。"}}
        return {"status": "executing"}

    def route_validation(state: WorkflowState):
        return "execute" if state["status"] == "executing" else "stop"

    def execute_transaction(state: WorkflowState):
        envelope = TaskEnvelope.model_validate(state["spec"])
        plan = TaskPlan.model_validate(state["plan"])
        output, report, result = service.execute_task_plan(
            state["task_id"], envelope, plan, run_id=state.get("run_id")
        )
        return {
            "status": "completed",
            "result": {
                "output_path": str(output),
                "verification": report.model_dump(mode="json"),
                "tool_result": result.model_dump(mode="json"),
            },
        }

    graph = StateGraph(WorkflowState)
    graph.add_node("validate_approved_plan", validate_approved_plan)
    graph.add_node("execute_transaction", execute_transaction)
    graph.add_edge(START, "validate_approved_plan")
    graph.add_conditional_edges(
        "validate_approved_plan", route_validation,
        {"execute": "execute_transaction", "stop": END},
    )
    graph.add_edge("execute_transaction", END)
    return graph.compile(checkpointer=checkpointer or MemorySaver())


@contextmanager
def sqlite_workflow(service) -> Iterator[Any]:
    """Compile the graph with durable SQLite checkpoints for a running app process."""
    try:
        from langgraph.checkpoint.sqlite import SqliteSaver
    except ImportError as exc:
        raise RuntimeError("Install langgraph-checkpoint-sqlite to enable durable checkpoints.") from exc
    checkpoint_path = service.settings.data_dir / "langgraph-checkpoints.sqlite3"
    with SqliteSaver.from_conn_string(str(checkpoint_path)) as checkpointer:
        yield build_workflow(service, checkpointer=checkpointer)


@contextmanager
def sqlite_v2_workflow(service) -> Iterator[Any]:
    """Compile the v0.2 workflow with durable SQLite checkpoints."""
    try:
        from langgraph.checkpoint.sqlite import SqliteSaver
    except ImportError as exc:
        raise RuntimeError("Install langgraph-checkpoint-sqlite to enable durable checkpoints.") from exc
    checkpoint_path = service.settings.data_dir / "langgraph-v2-checkpoints.sqlite3"
    with SqliteSaver.from_conn_string(str(checkpoint_path)) as checkpointer:
        yield build_v2_workflow(service, checkpointer=checkpointer)


@contextmanager
def sqlite_execution_workflow(service) -> Iterator[Any]:
    """Compile the chat execution graph with durable SQLite checkpoints."""
    try:
        from langgraph.checkpoint.sqlite import SqliteSaver
    except ImportError as exc:
        raise RuntimeError("Install langgraph-checkpoint-sqlite to enable durable checkpoints.") from exc
    checkpoint_path = service.settings.data_dir / "langgraph-v3-checkpoints.sqlite3"
    with SqliteSaver.from_conn_string(str(checkpoint_path)) as checkpointer:
        yield build_execution_workflow(service, checkpointer=checkpointer)


# Compatibility aliases for callers and persisted tests that use the historical names.
def build_v1_planning_workflow(service, checkpointer=None):
    return build_planning_workflow(service, checkpointer=checkpointer)


def build_v3_execution_workflow(service, checkpointer=None):
    return build_execution_workflow(service, checkpointer=checkpointer)


@contextmanager
def sqlite_v3_execution_workflow(service) -> Iterator[Any]:
    with sqlite_execution_workflow(service) as workflow:
        yield workflow
