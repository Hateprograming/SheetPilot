from __future__ import annotations

import json
import logging
import sys
from pathlib import Path

import streamlit as st


ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from sheetpilot.config import Settings
from sheetpilot.errors import SheetPilotError
from sheetpilot.schemas import MessageKind
from sheetpilot.service import SheetPilotService
from sheetpilot.ui_utils import display_dataframe


LOGGER = logging.getLogger(__name__)


st.set_page_config(
    page_title="SheetPilot v2.0.10 · 表格 Agent",
    page_icon="▦",
    layout="wide",
    initial_sidebar_state="collapsed",
)
st.markdown(
    """
    <style>
    :root {
      --paper:#F7F9FC; --ink:#172033; --action:#2E63E7;
      --workbook:#16856B; --line:#DCE3EC; --danger:#B5473A; --muted:#667085;
    }
    html, body, [class*="css"] { font-family:"Microsoft YaHei UI","Segoe UI Variable",sans-serif; }
    .stApp { background:var(--paper); color:var(--ink); }
    .block-container { max-width:920px; padding-top:2.2rem; padding-bottom:7rem; }
    [data-testid="stSidebar"] { background:#EEF3F8; border-right:1px solid var(--line); }
    h1,h2,h3 { color:var(--ink); letter-spacing:-.025em; }
    h1 { font-weight:760; font-size:2.05rem!important; }
    .eyebrow { color:var(--action); font:700 .7rem/1.3 "Cascadia Mono",monospace; letter-spacing:.13em; text-transform:uppercase; }
    .context-bar {
      display:flex; align-items:center; justify-content:space-between; gap:1rem;
      border:1px solid #BFD9D2; background:#F1F9F6; border-radius:12px;
      padding:.7rem .9rem; margin:.8rem 0 1.5rem;
    }
    .context-file { font-weight:700; color:var(--ink); overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
    .context-state { color:var(--workbook); font:650 .78rem/1.2 "Cascadia Mono",monospace; white-space:nowrap; }
    [data-testid="stChatMessage"] {
      background:transparent; border:0; padding:.35rem 0 1rem;
    }
    [data-testid="stChatMessageContent"] { color:var(--ink); }
    [data-testid="stChatInput"] { border:1px solid #BCC8D8; border-radius:15px; background:white; box-shadow:0 8px 24px rgba(23,32,51,.08); }
    .plan-card, .result-card {
      background:white; border:1px solid var(--line); border-radius:12px;
      padding:.75rem 1rem; margin:.4rem 0 .8rem;
    }
    .plan-card { border-left:4px solid var(--action); }
    .result-card { border-left:4px solid var(--workbook); }
    .empty-shell { max-width:720px; margin:8vh auto 0; text-align:center; }
    .empty-shell p { color:var(--muted); font-size:1.02rem; }
    .sample-grid { display:grid; grid-template-columns:1fr 1fr; gap:.65rem; margin:1.2rem 0; text-align:left; }
    .sample { background:white; border:1px solid var(--line); border-radius:10px; padding:.75rem .85rem; color:#344054; font-size:.88rem; }
    .version-code { font:650 .76rem/1.2 "Cascadia Mono",monospace; color:var(--workbook); }
    .stButton>button { border-radius:9px; font-weight:650; }
    .stButton>button[kind="primary"] { background:var(--action); border-color:var(--action); }
    :focus-visible { outline:3px solid #8DB0FF!important; outline-offset:2px; }
    @media (max-width:700px) { .sample-grid { grid-template-columns:1fr; } .block-container { padding:1rem .9rem 6rem; } }
    @media (prefers-reduced-motion:reduce) { * { transition:none!important; animation:none!important; } }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_resource
def get_service() -> SheetPilotService:
    return SheetPilotService(Settings.load(ROOT))


service = get_service()
state = st.session_state
state.setdefault("task_id", None)
state.setdefault("editing_plan", False)
state.setdefault("exception_followup", False)
state.setdefault("ui_error", None)


STATUS_LABELS = {
    "uploaded": "准备分析",
    "profiled": "等待你的要求",
    "needs_conversion": "等待转换要求",
    "intent_parsed": "正在理解",
    "analyzing_request": "正在理解",
    "needs_clarification": "等待补充",
    "awaiting_approval": "等待确认",
    "approved": "准备执行",
    "executing": "正在执行",
    "completed": "可以继续修改",
    "failed": "需要调整",
    "verification_failed": "已安全回滚",
    "paused": "等待恢复选择",
}

STEP_LABELS = {
    "pending": "等待", "running": "执行中", "completed": "已核验",
    "retrying": "正在重试", "paused": "已暂停", "failed": "失败",
    "rolled_back": "已回滚",
}


def rerun_after(action) -> None:
    try:
        action()
        state.ui_error = None
    except SheetPilotError as exc:
        state.ui_error = exc.message
    except Exception:
        state.ui_error = "这次操作没有完成，文件未被覆盖。请检查要求后重试。"
    st.rerun()


def open_workspace(task_id: str) -> None:
    state.task_id = task_id
    state.editing_plan = False
    state.ui_error = None


def render_upload() -> None:
    st.markdown(
        """
        <div class="empty-shell">
          <div class="eyebrow">Spreadsheet work agent</div>
          <h1>上传工作簿，然后告诉我你想得到什么</h1>
          <p>SheetPilot 会理解表格结构、必要时向你确认、展示计划，再在副本上执行并核验结果。</p>
          <div class="sample-grid">
            <div class="sample">按客户编号，把客户等级补充到订单明细。</div>
            <div class="sample">只保留已完成订单，按地区汇总销售额。</div>
            <div class="sample">统一客户编号格式，去掉空格并转成大写。</div>
            <div class="sample">按部门对业绩排名，每组保留前 3 名。</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    upload = st.file_uploader("选择 XLSX 工作簿或 CSV 文件", type=["xlsx", "csv"])
    consent = st.checkbox(
        "我同意将任务描述、Sheet 名、表头和结构统计发送到已配置的模型；完整工作簿保留在本机。"
    )
    if upload and st.button("分析这个工作簿", type="primary", width="stretch", disabled=not consent):
        with st.status("正在读取工作簿结构……", expanded=True) as status:
            try:
                profile = service.create_task(upload.name, upload.getvalue())
                open_workspace(service.task_id_from_profile(profile))
                status.update(label="结构分析完成", state="complete")
                st.rerun()
            except Exception as exc:
                status.update(label="无法读取这个工作簿", state="error")
                st.error(str(exc))


def render_sidebar(workspace: dict | None) -> None:
    with st.sidebar:
        st.markdown("## SheetPilot")
        st.caption("v2.0.10 · 可审计的数据准备 Agent")
        if st.button("＋ 处理另一个工作簿", width="stretch"):
            state.task_id = None
            state.editing_plan = False
            st.rerun()

        recent = service.recent_workspaces()
        if recent:
            st.markdown("### 最近工作簿")
            labels = {
                f"{item['name']} · V{item['version']}": item["task_id"] for item in recent
            }
            selected = st.selectbox("最近工作簿", ["", *labels], label_visibility="collapsed")
            if selected and labels[selected] != state.task_id:
                if st.button("打开这段对话", width="stretch"):
                    open_workspace(labels[selected])
                    st.rerun()

        if not workspace:
            return
        task = workspace["task"]
        profile = workspace["profile"]
        st.divider()
        st.markdown(f"### {task['original_name']}")
        st.markdown(f'<span class="version-code">CURRENT · V{task["current_version"]}</span>', unsafe_allow_html=True)
        st.caption(f"{len(profile['sheets'])} 个工作表 · {profile['size_bytes'] / 1024:.1f} KB")

        sheet_names = [item["name"] for item in profile["sheets"]]
        if sheet_names:
            with st.expander("查看工作簿内容"):
                selected_sheet = st.selectbox("工作表", sheet_names, key="sidebar_sheet")
                try:
                    preview = service.get_workbook_preview(
                        state.task_id, int(task["current_version"]), selected_sheet, 20
                    )
                    st.caption(f"共 {preview['total_rows']} 行，当前显示前 {len(preview['rows'])} 行")
                    st.dataframe(display_dataframe(preview["rows"]), width="stretch", hide_index=True)
                except Exception:
                    LOGGER.exception("Failed to render sidebar workbook preview")
                    st.warning("暂时无法预览这个工作表，但文件仍可继续分析或下载。")

        st.markdown("### 版本")
        for version in reversed(workspace["versions"]):
            number = int(version["version"])
            columns = st.columns([2, 1])
            columns[0].caption(f"V{number}" + (" · 当前" if number == int(task["current_version"]) else ""))
            if number != int(task["current_version"]) and columns[1].button("恢复", key=f"restore_{number}"):
                rerun_after(lambda number=number: service.restore_version(state.task_id, number))

        original_name, original_bytes = service.get_download(state.task_id, original=True)
        if task["status"] != "needs_conversion":
            current_name, current_bytes = service.get_download(state.task_id)
            st.download_button("下载当前版本", current_bytes, current_name, width="stretch", type="primary")
        st.download_button("下载原始文件", original_bytes, original_name, width="stretch")


def render_result_preview(message: dict, workspace: dict) -> None:
    metadata = message.get("metadata", {})
    result = metadata.get("result", {})
    counters = result.get("counters", {})
    if counters:
        columns = st.columns(min(4, len(counters)))
        for index, (name, value) in enumerate(counters.items()):
            columns[index % len(columns)].metric(name.replace("_", " "), value)
    formula_steps = [
        step for step in result.get("details", {}).get("steps", [])
        if step.get("tool_name") == "formula"
    ]
    for step in formula_steps:
        calculation = step.get("details", {}).get("calculation") or {}
        if calculation.get("backend") and calculation.get("backend") != "none":
            backend = {
                "excel": "Microsoft Excel",
                "libreoffice": "LibreOffice",
                "test": "测试计算引擎",
            }.get(calculation.get("backend"), calculation.get("backend", "计算引擎"))
            calculation_message = (
                f"公式已通过 {backend} 重算："
                f"{calculation.get('calculated_cells', 0)} / {calculation.get('formula_cells', 0)} 个公式已写入缓存值。"
            )
            if calculation.get("error_cells"):
                st.warning(
                    f"{calculation_message} 仍检测到 {len(calculation['error_cells'])} 个公式错误。"
                )
            else:
                st.success(calculation_message)
        elif step.get("details", {}).get("mode"):
            st.warning("当前环境没有可用的公式重算引擎；文件已设置为在 Excel 打开时自动完整重算。")
    preview_sheets = [item for item in metadata.get("preview_sheets", []) if item]
    if preview_sheets:
        with st.expander("预览处理结果", expanded=True):
            selected = st.selectbox(
                "预览工作表",
                preview_sheets,
                key=f"result_sheet_{message['id']}",
            )
            try:
                preview = service.get_workbook_preview(
                    state.task_id, metadata.get("version"), selected, 50
                )
                st.caption(f"{selected} · 共 {preview['total_rows']} 行，显示前 {len(preview['rows'])} 行")
                st.dataframe(display_dataframe(preview["rows"]), width="stretch", hide_index=True)
            except Exception:
                LOGGER.exception("Failed to render result workbook preview")
                st.warning("结果文件已生成，但当前无法显示表格预览，请直接下载文件检查。")
    report = metadata.get("verification", {})
    if report:
        with st.expander("查看核验结果"):
            for check in report.get("checks", []):
                st.write(f"{'✓' if check['passed'] else '×'} {check['name']} — {check.get('detail', '')}")
    for step in result.get("details", {}).get("steps", []):
        artifact_path = step.get("details", {}).get("artifact_path")
        if artifact_path:
            try:
                name, content = service.get_artifact_download(state.task_id, artifact_path)
                st.download_button(
                    "下载异常 CSV" if step.get("details", {}).get("operation") else "下载 CSV 导出文件",
                    content,
                    file_name=name,
                    mime="text/csv",
                    key=f"artifact_{message['id']}_{name}",
                )
            except Exception:
                st.warning("导出产物已不可用，请重新执行转换。")
    if any(
        step.get("details", {}).get("follow_up_available")
        for step in result.get("details", {}).get("steps", [])
    ):
        if st.button("继续处理异常数据", key=f"followup_{message['id']}"):
            state.editing_plan = True
            state.exception_followup = True
            st.rerun()


def render_exception_review(message: dict) -> None:
    review = message.get("metadata", {}).get("review") or {}
    rows = review.get("rows") or []
    if not rows:
        return
    display_rows = []
    value_columns: list[str] = []
    for item in rows:
        raw_values = item.get("raw_values") or {}
        for name in raw_values:
            if name not in value_columns:
                value_columns.append(name)
        display_rows.append({
            "工作表": item.get("sheet", ""),
            "Excel 行号": item.get("row", ""),
            **raw_values,
            "问题代码": item.get("issue_code", ""),
            "说明": item.get("description", ""),
        })
    total = int(review.get("total") or len(rows))
    st.caption(
        f"以下是需要你决定处理方式的异常记录（共 {total} 条）"
        + ("，当前展示前 50 条" if review.get("truncated") else "")
    )
    ordered = ["工作表", "Excel 行号", *value_columns, "问题代码", "说明"]
    frame = display_dataframe(display_rows)
    st.dataframe(frame[[name for name in ordered if name in frame.columns]], width="stretch", hide_index=True)


def stop_active_work(workspace: dict) -> None:
    planning = workspace.get("planning_run")
    execution = workspace.get("active_run")
    if planning and planning.get("status") in {"queued", "waiting_for_execution", "running"}:
        service.cancel_planning(state.task_id)
    if execution and execution.get("status") == "running":
        service.control_current_run(state.task_id, "cancel")


def render_activity_timeline(workspace: dict) -> None:
    planning = workspace.get("planning_run")
    execution = workspace.get("active_run")
    planning_active = planning and planning.get("status") in {
        "queued", "waiting_for_execution", "running"
    }
    execution_active = execution and execution.get("status") == "running"
    if not planning_active and not execution_active:
        return
    run_ids = set()
    if planning_active:
        run_ids.add(planning["job_id"])
    events = [
        item for item in workspace.get("activity_events", [])
        if item.get("run_id") in run_ids or (execution_active and item.get("run_kind") == "execution")
    ]
    label = (
        "正在完成当前安全步骤后停止，并按新要求规划"
        if planning and planning.get("status") == "waiting_for_execution"
        else "正在处理你的要求"
    )
    with st.status(label, expanded=True, state="running"):
        for event in events[-12:]:
            icon = "✓" if event.get("status") == "completed" else "·"
            st.write(f"{icon} {event.get('label', '')}")
        if execution_active:
            for item in workspace.get("tool_runs", []):
                if item.get("run_id") == execution.get("run_id"):
                    text = "正在执行数据处理步骤" if item.get("status") == "running" else "已完成并核验数据处理步骤"
                    st.write(f"{'·' if item.get('status') == 'running' else '✓'} {text}")
        if st.button("停止当前处理", key="stop_active_work"):
            rerun_after(lambda: stop_active_work(workspace))


def render_completed_activity(message: dict, workspace: dict) -> None:
    run_id = message.get("metadata", {}).get("activity_run_id")
    if not run_id:
        return
    events = [
        item for item in workspace.get("activity_events", [])
        if item.get("run_id") == run_id
    ]
    if not events or not any(item.get("status") in {"completed", "failed", "cancelled"} for item in events):
        return
    with st.expander("查看处理时间线"):
        for event in events:
            icon = "✓" if event.get("status") == "completed" else "×" if event.get("status") == "failed" else "·"
            duration = event.get("metadata", {}).get("duration_ms")
            suffix = f" · {duration} ms" if duration is not None else ""
            st.write(f"{icon} {event.get('label', '')}{suffix}")


def render_chat(workspace: dict) -> None:
    task = workspace["task"]
    status = STATUS_LABELS.get(task["status"], "等待你的要求")
    st.markdown('<div class="eyebrow">Workbook conversation</div>', unsafe_allow_html=True)
    st.title("和你的工作簿对话")
    st.markdown(
        f'<div class="context-bar"><span class="context-file">{task["original_name"]}</span>'
        f'<span class="context-state">V{task["current_version"]} · {status}</span></div>',
        unsafe_allow_html=True,
    )
    legacy_notice_key = f"integrity_notice_{task['task_id']}"
    if task.get("integrity_level") == "legacy_rebased" and not state.get(legacy_notice_key):
        st.info("这是旧任务，系统已基于当前原文件建立新的完整性基线；这不代表能够证明文件自最初上传后从未变化。")
        state[legacy_notice_key] = True
    elif task.get("integrity_level") == "missing":
        st.warning("这个旧任务缺少可验证的原始文件，只能查看或下载已有版本；如需继续处理，请重新上传。")

    messages = workspace["messages"]
    latest_id = messages[-1]["id"] if messages else None
    for message in messages:
        role = "user" if message["role"] == "user" else "assistant"
        with st.chat_message(role, avatar="📊" if role == "assistant" else None):
            kind = message["kind"]
            if kind == MessageKind.PLAN.value:
                with st.container(border=True):
                    st.markdown(message["content"])
                    plan_data = message.get("metadata", {}).get("plan", {})
                    previous_plan = message.get("metadata", {}).get("previous_plan")
                    if previous_plan:
                        st.info(
                            f"计划已从 R{previous_plan.get('revision', 1)} 修订为 "
                            f"R{plan_data.get('revision', 1)}；旧审批已失效。"
                        )
                        with st.expander("查看计划变化"):
                            before = [item.get("description") for item in previous_plan.get("steps", [])]
                            after = [item.get("description") for item in plan_data.get("steps", [])]
                            st.write("原计划：", before)
                            st.write("新计划：", after)
                            st.write(
                                "风险：",
                                f"{previous_plan.get('risk_level', '—')} → {plan_data.get('risk_level', '—')}",
                            )
                    if plan_data.get("steps"):
                        st.caption("执行时间线")
                        completed = {
                            item.get("plan_step_id")
                            for item in workspace.get("tool_runs", [])
                            if item.get("status") == "completed"
                        }
                        for step in plan_data["steps"]:
                            step_status = "completed" if step["id"] in completed else step.get("status", "pending")
                            dependency = f" · 依赖 {', '.join(step.get('depends_on', []))}" if step.get("depends_on") else ""
                            st.markdown(
                                f"`{step['order']:02d}` **{STEP_LABELS.get(step_status, step_status)}** "
                                f"{step['description']}{dependency}"
                            )
                    with st.expander("查看技术详情"):
                        st.json(plan_data)
            elif kind == MessageKind.RESULT.value:
                with st.container(border=True):
                    st.markdown(message["content"])
                    render_result_preview(message, workspace)
            elif kind == MessageKind.ERROR.value:
                st.error(message["content"])
            elif kind == MessageKind.REFUSAL.value:
                st.warning(message["content"])
            else:
                st.markdown(message["content"])

            if role == "assistant":
                render_completed_activity(message, workspace)

            if kind == MessageKind.CLARIFICATION.value:
                render_exception_review(message)

            if (
                kind == MessageKind.CLARIFICATION.value
                and message["id"] == latest_id
                and task["status"] == "needs_clarification"
            ):
                choices = message.get("metadata", {}).get("choices", [])
                if choices:
                    columns = st.columns(min(3, len(choices)))
                    for index, choice in enumerate(choices):
                        if columns[index % len(columns)].button(choice, key=f"choice_{message['id']}_{index}"):
                            if "输入具体" in choice:
                                def enter_rule_mode(choice=choice):
                                    service.select_custom_clarification_input(state.task_id, choice)
                                    state.editing_plan = True
                                    state.exception_followup = True
                                rerun_after(enter_rule_mode)
                            else:
                                rerun_after(lambda choice=choice: service.enqueue_message(state.task_id, choice))

    render_activity_timeline(workspace)

    if state.ui_error:
        st.error(state.ui_error)

    active_run = workspace.get("active_run")
    if active_run and active_run.get("status") in {"paused", "failed"}:
        with st.container(border=True):
            st.markdown("**任务停在安全检查点**")
            error = json.loads(active_run.get("error_json") or "{}")
            st.write(error.get("safe_message", "运行已暂停，上一正式版本没有改变。"))
            if error.get("technical_detail"):
                with st.expander("查看失败原因"):
                    st.write(error["technical_detail"])
            st.caption(f"已完成 {active_run.get('current_step', 0)} 个步骤 · 内部检查点不会作为正式版本交付")
            retry, revise, cancel = st.columns(3)
            if retry.button("从检查点重试", type="primary", width="stretch"):
                with st.status("正在从最近检查点继续……", expanded=True) as progress:
                    try:
                        service.resume_current_run(state.task_id)
                        progress.update(label="恢复执行完成", state="complete")
                        state.ui_error = None
                    except SheetPilotError as exc:
                        progress.update(label="仍需调整", state="error")
                        state.ui_error = exc.message
                    except Exception:
                        LOGGER.exception("Unexpected error while resuming task run")
                        progress.update(label="恢复未完成", state="error")
                        state.ui_error = "恢复执行时发生内部错误，上一正式版本保持不变。"
                st.rerun()
            if revise.button("修改剩余计划", width="stretch"):
                state.editing_plan = True
                st.rerun()
            if cancel.button("取消并保留上一版本", width="stretch"):
                rerun_after(lambda: service.control_current_run(state.task_id, "cancel"))

    if task["status"] == "awaiting_approval" and messages and messages[-1]["kind"] == MessageKind.PLAN.value:
        plan_data = messages[-1].get("metadata", {}).get("plan", {})
        requirements = plan_data.get("confirmation_requirements", [])
        acknowledged = True
        if "overwrite_source_columns" in requirements:
            acknowledged = st.checkbox(
                "我确认该计划会覆盖当前工作表（包括修改已有列或删除未保留列），并理解旧内容将仅保留在历史版本中。",
                key=f"overwrite_ack_{plan_data.get('plan_id', 'current')}_{plan_data.get('revision', 1)}",
            )
        approve, edit, cancel = st.columns([1.2, 1, 1])
        if approve.button(
            "确认并执行", type="primary", width="stretch",
            disabled=not acknowledged,
        ):
            rerun_after(lambda: service.enqueue_current_plan_execution(
                state.task_id,
                ["overwrite_source_columns"] if requirements else [],
            ))
        if edit.button("修改要求", width="stretch"):
            state.editing_plan = True
            st.rerun()
        if cancel.button("取消计划", width="stretch"):
            rerun_after(lambda: service.cancel_current_plan(state.task_id))

    if state.editing_plan:
        if state.get("exception_followup"):
            st.info("请在下方说明哪些异常值应如何处理。系统会创建新计划，不会自动修改文件。")
        else:
            st.info("直接在下方说明要修改的地方，例如：“来源表改成客户主数据，并使用最后一条重复记录”。")

    prompt = st.chat_input(
        "描述你希望怎样处理这个工作簿……"
        if not state.editing_plan
        else (
            "例如：第13行改为21mm×7mm×16.7mm，第16行改为2mm×3mm×4mm"
            if state.get("exception_followup")
            else "告诉我需要怎样修改计划……"
        )
    )
    if prompt:
        state.editing_plan = False
        state.exception_followup = False
        rerun_after(lambda: service.enqueue_message(state.task_id, prompt))


workspace = None
if state.task_id:
    try:
        workspace = service.get_workspace(state.task_id)
    except Exception:
        state.task_id = None

render_sidebar(workspace)
if workspace:
    def workspace_is_active(current: dict) -> bool:
        planning = current.get("planning_run") or {}
        execution = current.get("active_run") or {}
        return (
            planning.get("status") in {"queued", "waiting_for_execution", "running"}
            or execution.get("status") == "running"
            or current.get("task", {}).get("status") in {"approved", "executing"}
        )

    @st.fragment(run_every=1.0)
    def render_live_chat() -> None:
        try:
            fresh = service.get_workspace(state.task_id)
        except Exception:
            LOGGER.exception("Failed to refresh live chat")
            st.error("暂时无法刷新处理状态；工作簿不会因此被覆盖。")
            return
        if not workspace_is_active(fresh):
            st.rerun(scope="app")
        render_chat(fresh)

    if workspace_is_active(workspace):
        render_live_chat()
    else:
        render_chat(workspace)
else:
    render_upload()
