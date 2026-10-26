from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

import streamlit as st

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from sheetpilot.config import Settings
from sheetpilot.errors import SheetPilotError
from sheetpilot.ui_utils import display_dataframe
from sheetpilot.schemas import (
    AggregateFunction,
    AggregateMetric,
    AggregateTaskSpec,
    CleanOperation,
    CleanTaskSpec,
    DuplicatePolicy,
    FilterOperator,
    LookupTaskSpec,
    NormalizationMode,
    NormalizeOperation,
    NormalizeTaskSpec,
    RankMethod,
    RowFilter,
    SortDirection,
    SortKey,
    SortRankTaskSpec,
    TaskEnvelope,
    TaskPlan,
    TaskType,
    WorkbookProfile,
)
from sheetpilot.service import SheetPilotService


st.set_page_config(page_title="SheetPilot v0.2", page_icon="▦", layout="wide")
st.markdown(
    """
    <style>
    :root { --ink:#14213d; --muted:#64748b; --line:#dbe3ed; --blue:#2559d6; --ice:#f5f8fc; }
    .stApp { background:linear-gradient(90deg,#edf2f7 0,#edf2f7 16px,#f8fafc 16px,#f8fafc 100%); color:var(--ink); }
    h1,h2,h3 { color:var(--ink); letter-spacing:-.025em; }
    h1 { font-family:"Arial Narrow","Microsoft YaHei UI",sans-serif; font-weight:800; }
    [data-testid="stSidebar"] { background:#14213d; }
    [data-testid="stSidebar"] * { color:#f8fafc; }
    .pilot-kicker { font:700 .72rem/1.2 ui-monospace,Consolas,monospace; color:var(--blue); letter-spacing:.14em; text-transform:uppercase; }
    .pilot-lead { max-width:850px; color:var(--muted); font-size:1.02rem; }
    .trace { border-left:3px solid var(--blue); padding:.3rem 0 .3rem 1rem; margin:.8rem 0 1.4rem; }
    .trace strong { display:block; color:var(--ink); }
    .trace span { color:var(--muted); font-size:.88rem; }
    .plan-step { border:1px solid var(--line); border-radius:6px; padding:.9rem 1rem; background:white; margin:.6rem 0; }
    .status-chip { display:inline-block; border:1px solid #b9caec; color:#204e9e; padding:.2rem .55rem; border-radius:999px; font-size:.76rem; }
    .stButton>button { border-radius:4px; font-weight:650; }
    .stButton>button[kind="primary"] { background:var(--blue); }
    :focus-visible { outline:3px solid #8bb3ff!important; outline-offset:2px; }
    @media (prefers-reduced-motion:reduce) { * { transition:none!important; } }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_resource
def get_service() -> SheetPilotService:
    return SheetPilotService(Settings.load(ROOT))


service = get_service()
state = st.session_state
for key, value in {
    "stage": "upload",
    "task_id": None,
    "profile": None,
    "envelope": None,
    "plan": None,
    "validation": None,
    "result": None,
}.items():
    state.setdefault(key, value)


TASK_LABELS = {
    TaskType.LOOKUP: "跨表精确匹配",
    TaskType.NORMALIZE: "数据标准化",
    TaskType.AGGREGATE: "聚合与汇总",
    TaskType.SORT_RANK: "排序、排名与 Top N",
    TaskType.CLEAN: "重复与条件清洗",
}


def reset_all() -> None:
    for key in ("task_id", "profile", "envelope", "plan", "validation", "result"):
        state[key] = None
    state.stage = "upload"


def reset_task() -> None:
    for key in ("envelope", "plan", "validation", "result"):
        state[key] = None
    state.stage = "task"


def current_profile() -> WorkbookProfile:
    return WorkbookProfile.model_validate(state.profile)


def sheet_profile(name: str):
    return next(sheet for sheet in current_profile().sheets if sheet.name == name)


def columns_for(name: str) -> list[str]:
    return [column.name for column in sheet_profile(name).columns]


def restore_task(task_id: str) -> None:
    row = service.repository.get_task(task_id)
    if not row or not row["profile_json"]:
        return
    state.task_id = task_id
    state.profile = json.loads(row["profile_json"])
    state.envelope = None
    state.plan = None
    state.validation = None
    state.result = None
    if row["spec_json"]:
        try:
            state.envelope = TaskEnvelope.model_validate_json(row["spec_json"]).model_dump(mode="json")
        except Exception:
            try:
                lookup = LookupTaskSpec.model_validate_json(row["spec_json"])
                state.envelope = TaskEnvelope(goal="历史跨表匹配任务", spec=lookup).model_dump(mode="json")
            except Exception:
                pass
    if row["plan_json"]:
        try:
            state.plan = TaskPlan.model_validate_json(row["plan_json"]).model_dump(mode="json")
        except Exception:
            pass
    if row["status"] == "completed" and row["verification_json"]:
        state.result = service.result_payload(task_id)
        state.stage = "result"
    elif state.plan:
        state.validation = {"valid": row["status"] != "needs_clarification", "errors": [], "warnings": []}
        state.stage = "review"
    else:
        state.stage = "task"


def render_workbook() -> None:
    profile = current_profile()
    st.subheader("当前工作簿")
    task = service.repository.get_task(state.task_id)
    st.caption(f"版本 V{task['current_version']} · {profile.file_name} · {profile.size_bytes / 1024:.1f} KB")
    for sheet in profile.sheets:
        with st.expander(f"{sheet.name} · {sheet.max_row} 行 × {sheet.max_column} 列"):
            st.caption(f"表头行：{sheet.header_row or '待确认'} · 置信度 {sheet.header_confidence:.0%}")
            if sheet.preview_rows:
                st.dataframe(display_dataframe(sheet.preview_rows), width="stretch", hide_index=True)
    if profile.risks:
        st.warning("\n".join(profile.risks))


def filter_editor(sheet: str, key_prefix: str, defaults: list[RowFilter] | None = None) -> list[RowFilter]:
    defaults = defaults or []
    columns = columns_for(sheet)
    enabled = st.checkbox("添加筛选条件", value=bool(defaults), key=f"{key_prefix}_enabled")
    if not enabled:
        return []
    count = st.number_input("条件数量", min_value=1, max_value=2, value=max(1, len(defaults)), key=f"{key_prefix}_count")
    result: list[RowFilter] = []
    operator_labels = {
        FilterOperator.EQ: "等于",
        FilterOperator.NE: "不等于",
        FilterOperator.GT: "大于",
        FilterOperator.GTE: "大于等于",
        FilterOperator.LT: "小于",
        FilterOperator.LTE: "小于等于",
        FilterOperator.CONTAINS: "包含文本",
        FilterOperator.IS_BLANK: "为空",
        FilterOperator.NOT_BLANK: "不为空",
    }
    for index in range(int(count)):
        default = defaults[index] if index < len(defaults) else None
        a, b, c = st.columns([1.2, 1, 1.4])
        column = a.selectbox(
            f"条件 {index + 1} · 列",
            columns,
            index=columns.index(default.column) if default and default.column in columns else 0,
            key=f"{key_prefix}_column_{index}",
        )
        operators = list(FilterOperator)
        operator = b.selectbox(
            f"条件 {index + 1} · 关系",
            operators,
            index=operators.index(default.operator) if default else 0,
            format_func=lambda item: operator_labels[item],
            key=f"{key_prefix}_operator_{index}",
        )
        if operator in {FilterOperator.IS_BLANK, FilterOperator.NOT_BLANK}:
            value: Any = None
            c.caption("此条件不需要比较值")
        else:
            value = c.text_input(
                f"条件 {index + 1} · 值",
                value="" if not default or default.value is None else str(default.value),
                key=f"{key_prefix}_value_{index}",
            )
        result.append(RowFilter(column=column, operator=operator, value=value))
    return result


def lookup_form(current: LookupTaskSpec | None) -> LookupTaskSpec:
    profile = current_profile()
    sheets = [sheet.name for sheet in profile.sheets]
    current = current or LookupTaskSpec()
    main = st.selectbox("主表（将追加列）", sheets, index=sheets.index(current.main_sheet) if current.main_sheet in sheets else 0)
    source_options = [name for name in sheets if name != main]
    source = st.selectbox("来源表", source_options, index=source_options.index(current.source_sheet) if current.source_sheet in source_options else 0)
    main_columns, source_columns = columns_for(main), columns_for(source)
    main_key = st.selectbox("主表关联列", main_columns, index=main_columns.index(current.main_key_column) if current.main_key_column in main_columns else 0)
    source_key = st.selectbox("来源表关联列", source_columns, index=source_columns.index(current.source_key_column) if current.source_key_column in source_columns else 0)
    returns = st.multiselect("需要补充的列（1–3列）", source_columns, default=[item for item in current.return_columns if item in source_columns], max_selections=3)
    outputs = [
        st.text_input(
            f"输出列名 · {column}",
            value=current.output_column_names[index] if index < len(current.output_column_names) else column,
            key=f"lookup_output_{index}_{column}",
        )
        for index, column in enumerate(returns)
    ]
    normalization = st.radio(
        "关联键比较方式",
        list(NormalizationMode),
        index=list(NormalizationMode).index(current.normalization_mode),
        format_func=lambda item: "精确文本（保留前导零）" if item == NormalizationMode.EXACT_TEXT else "数值等价（001与1视为相同）",
        horizontal=True,
    )
    duplicate = st.selectbox(
        "来源键重复时",
        list(DuplicatePolicy),
        index=list(DuplicatePolicy).index(current.duplicate_policy),
        format_func=lambda item: {DuplicatePolicy.BLOCK: "阻止并报告", DuplicatePolicy.FIRST: "使用第一条", DuplicatePolicy.LAST: "使用最后一条"}[item],
    )
    return LookupTaskSpec(
        main_sheet=main,
        source_sheet=source,
        main_key_column=main_key,
        source_key_column=source_key,
        return_columns=returns,
        output_column_names=outputs,
        normalization_mode=normalization,
        duplicate_policy=duplicate,
        confidence=max(current.confidence, 0.95),
    )


def normalize_form(current: NormalizeTaskSpec | None) -> NormalizeTaskSpec:
    profile = current_profile()
    sheets = [sheet.name for sheet in profile.sheets]
    current = current or NormalizeTaskSpec()
    sheet = st.selectbox("工作表", sheets, index=sheets.index(current.sheet) if current.sheet in sheets else 0)
    columns = columns_for(sheet)
    column = st.selectbox("需要标准化的列", columns, index=columns.index(current.column) if current.column in columns else 0)
    labels = {
        NormalizeOperation.TRIM: "去除首尾空格",
        NormalizeOperation.NFKC: "统一全角/半角字符",
        NormalizeOperation.LOWER: "转小写",
        NormalizeOperation.UPPER: "转大写",
        NormalizeOperation.TO_TEXT: "转为文本",
        NormalizeOperation.TO_NUMBER: "转为数值",
        NormalizeOperation.DATE_FORMAT: "统一日期格式",
        NormalizeOperation.FILL_BLANK: "填充空值",
    }
    operations = st.multiselect(
        "标准化操作（按选择顺序执行）",
        list(NormalizeOperation),
        default=current.operations,
        format_func=lambda item: labels[item],
        max_selections=4,
    )
    date_format = st.text_input("日期输出格式", value=current.date_format or "%Y-%m-%d") if NormalizeOperation.DATE_FORMAT in operations else None
    fill_value = st.text_input("空值填充值", value="" if current.fill_value is None else str(current.fill_value)) if NormalizeOperation.FILL_BLANK in operations else None
    overwrite = st.checkbox("覆盖原列（高风险）", value=current.overwrite)
    output = "" if overwrite else st.text_input("输出列名", value=current.output_column or f"{column}_标准化")
    return NormalizeTaskSpec(
        sheet=sheet,
        column=column,
        operations=operations,
        output_column=output,
        date_format=date_format,
        fill_value=fill_value,
        overwrite=overwrite,
        confidence=max(current.confidence, 0.95),
    )


def aggregate_form(current: AggregateTaskSpec | None) -> AggregateTaskSpec:
    profile = current_profile()
    sheets = [sheet.name for sheet in profile.sheets]
    current = current or AggregateTaskSpec()
    sheet = st.selectbox("来源工作表", sheets, index=sheets.index(current.sheet) if current.sheet in sheets else 0)
    columns = columns_for(sheet)
    groups = st.multiselect("分组列（1–3列）", columns, default=[item for item in current.group_by if item in columns], max_selections=3)
    metric_count = st.number_input("指标数量", min_value=1, max_value=3, value=max(1, len(current.metrics)))
    metrics: list[AggregateMetric] = []
    function_labels = {AggregateFunction.SUM: "求和", AggregateFunction.COUNT: "计数", AggregateFunction.MEAN: "平均", AggregateFunction.MIN: "最小", AggregateFunction.MAX: "最大"}
    for index in range(int(metric_count)):
        default = current.metrics[index] if index < len(current.metrics) else None
        a, b, c = st.columns([1.2, 1, 1.2])
        column = a.selectbox(f"指标 {index + 1} · 列", columns, index=columns.index(default.column) if default and default.column in columns else 0, key=f"agg_col_{index}")
        functions = list(AggregateFunction)
        function = b.selectbox(f"指标 {index + 1} · 计算", functions, index=functions.index(default.function) if default else 0, format_func=lambda item: function_labels[item], key=f"agg_fn_{index}")
        output = c.text_input(f"指标 {index + 1} · 输出列", value=default.output_column if default else f"{column}_{function.value}", key=f"agg_out_{index}")
        metrics.append(AggregateMetric(column=column, function=function, output_column=output))
    filters = filter_editor(sheet, "aggregate_filter", current.filters)
    output_sheet = st.text_input("汇总结果 Sheet 名", value=current.output_sheet)
    return AggregateTaskSpec(sheet=sheet, group_by=groups, metrics=metrics, filters=filters, output_sheet=output_sheet, confidence=max(current.confidence, 0.95))


def sort_rank_form(current: SortRankTaskSpec | None) -> SortRankTaskSpec:
    profile = current_profile()
    sheets = [sheet.name for sheet in profile.sheets]
    current = current or SortRankTaskSpec()
    sheet = st.selectbox("来源工作表", sheets, index=sheets.index(current.sheet) if current.sheet in sheets else 0)
    columns = columns_for(sheet)
    sort_columns = st.multiselect("排序列（最多3列）", columns, default=[item.column for item in current.sort_keys if item.column in columns], max_selections=3)
    sort_keys: list[SortKey] = []
    for index, column in enumerate(sort_columns):
        default = next((item for item in current.sort_keys if item.column == column), None)
        direction = st.radio(f"{column} 的方向", list(SortDirection), index=list(SortDirection).index(default.direction) if default else 0, format_func=lambda item: "升序" if item == SortDirection.ASC else "降序", horizontal=True, key=f"sort_dir_{index}")
        sort_keys.append(SortKey(column=column, direction=direction))
    rank_options = ["（不排名）", *columns]
    rank_column = st.selectbox("排名列", rank_options, index=rank_options.index(current.rank_column) if current.rank_column in rank_options else 0)
    actual_rank = None if rank_column == "（不排名）" else rank_column
    group_by: list[str] = []
    rank_direction = current.rank_direction
    rank_method = current.rank_method
    rank_output = current.rank_output_column
    if actual_rank:
        group_by = st.multiselect("分组内排名（可不选）", columns, default=[item for item in current.group_by if item in columns], max_selections=2)
        rank_direction = st.radio("排名方向", list(SortDirection), index=list(SortDirection).index(current.rank_direction), format_func=lambda item: "数值越小排名越前" if item == SortDirection.ASC else "数值越大排名越前", horizontal=True)
        rank_method = st.selectbox("并列策略", list(RankMethod), index=list(RankMethod).index(current.rank_method), format_func=lambda item: {RankMethod.DENSE: "密集排名 1,2,2,3", RankMethod.MIN: "竞赛排名 1,2,2,4", RankMethod.MAX: "并列使用末位"}[item])
        rank_output = st.text_input("排名输出列", value=current.rank_output_column)
    use_top = st.checkbox("只保留 Top N", value=current.top_n is not None)
    top_n = int(st.number_input("Top N", min_value=1, max_value=10000, value=current.top_n or 10)) if use_top else None
    output_sheet = st.text_input("结果 Sheet 名", value=current.output_sheet)
    return SortRankTaskSpec(sheet=sheet, sort_keys=sort_keys, rank_column=actual_rank, rank_output_column=rank_output, rank_direction=rank_direction, rank_method=rank_method, group_by=group_by, top_n=top_n, output_sheet=output_sheet, confidence=max(current.confidence, 0.95))


def clean_form(current: CleanTaskSpec | None) -> CleanTaskSpec:
    profile = current_profile()
    sheets = [sheet.name for sheet in profile.sheets]
    current = current or CleanTaskSpec()
    sheet = st.selectbox("来源工作表", sheets, index=sheets.index(current.sheet) if current.sheet in sheets else 0)
    columns = columns_for(sheet)
    labels = {CleanOperation.MARK_DUPLICATES: "标记重复记录", CleanOperation.REMOVE_DUPLICATES: "移除重复记录（结果写入新Sheet）", CleanOperation.FILTER_ROWS: "按条件筛选"}
    operation = st.radio("清洗方式", list(CleanOperation), index=list(CleanOperation).index(current.operation), format_func=lambda item: labels[item])
    key_columns: list[str] = []
    filters: list[RowFilter] = []
    keep = current.keep
    flag = current.duplicate_flag_column
    if operation in {CleanOperation.MARK_DUPLICATES, CleanOperation.REMOVE_DUPLICATES}:
        key_columns = st.multiselect("判断重复的列（1–5列）", columns, default=[item for item in current.key_columns if item in columns], max_selections=5)
        keep = st.radio("去重时保留", ["first", "last"], index=0 if current.keep == "first" else 1, format_func=lambda item: "第一条" if item == "first" else "最后一条", horizontal=True)
        if operation == CleanOperation.MARK_DUPLICATES:
            flag = st.text_input("重复标记列名", value=current.duplicate_flag_column)
    else:
        filters = filter_editor(sheet, "clean_filter", current.filters)
    output_sheet = st.text_input("结果 Sheet 名", value=current.output_sheet)
    return CleanTaskSpec(sheet=sheet, operation=operation, key_columns=key_columns, keep=keep, filters=filters, output_sheet=output_sheet, duplicate_flag_column=flag, confidence=max(current.confidence, 0.95))


def render_task_form(task_type: TaskType, envelope: TaskEnvelope | None):
    current = envelope.spec if envelope and envelope.spec.task_type == task_type else None
    if task_type == TaskType.LOOKUP:
        return lookup_form(current if isinstance(current, LookupTaskSpec) else None)
    if task_type == TaskType.NORMALIZE:
        return normalize_form(current if isinstance(current, NormalizeTaskSpec) else None)
    if task_type == TaskType.AGGREGATE:
        return aggregate_form(current if isinstance(current, AggregateTaskSpec) else None)
    if task_type == TaskType.SORT_RANK:
        return sort_rank_form(current if isinstance(current, SortRankTaskSpec) else None)
    return clean_form(current if isinstance(current, CleanTaskSpec) else None)


with st.sidebar:
    st.markdown("## SheetPilot")
    st.caption("v0.2 · controlled spreadsheet workflows")
    st.markdown("---")
    for number, label in (("01", "文件"), ("02", "任务"), ("03", "计划"), ("04", "执行与核验")):
        st.markdown(f"**{number}**  {label}")
    st.markdown("---")
    st.caption("原文件只读 · 计划先审阅 · 确定性工具执行")
    recent = service.repository.recent_tasks(8)
    if recent:
        options = {f"{item['original_name']} · V{item['current_version']} · {item['status']}": item["task_id"] for item in recent}
        selected = st.selectbox("恢复本地任务", ["", *options], label_visibility="collapsed")
        if selected and st.button("恢复任务", width="stretch"):
            restore_task(options[selected])
            st.rerun()
    if state.task_id and st.button("上传另一个文件", width="stretch"):
        reset_all()
        st.rerun()


st.markdown('<div class="pilot-kicker">Review-first spreadsheet operations</div>', unsafe_allow_html=True)
st.title("从一个匹配工具，升级为受控表格工作流")
st.markdown('<div class="pilot-lead">SheetPilot v0.2 支持跨表匹配、列标准化、聚合汇总、排序排名和基础清洗。每次修改都基于明确计划生成新版本并经过程序核验。</div>', unsafe_allow_html=True)
stage_labels = {"upload": "上传并分析", "task": "选择与配置任务", "review": "审阅执行计划", "result": "核验、版本与交付"}
st.markdown(f'<div class="trace"><strong>{stage_labels[state.stage]}</strong><span>当前阶段 · 不支持的任务会明确拒绝</span></div>', unsafe_allow_html=True)


try:
    if state.stage == "upload":
        st.subheader("上传工作簿")
        st.caption("支持单个无宏、无密码、无外部链接的 .xlsx；单 Sheet 最多 100,000 行。")
        upload = st.file_uploader("选择 XLSX 文件", type=["xlsx"])
        consent = st.checkbox("我同意将任务描述、Sheet名、表头、统计摘要和最多3行预览发送到Qwen Cloud。")
        if upload and st.button("建立只读任务并分析", type="primary", disabled=not consent):
            with st.status("正在保存原文件并分析工作簿……", expanded=True):
                profile = service.create_task(upload.name, upload.getvalue())
            state.task_id = service.task_id_from_profile(profile)
            profile.file_name = upload.name
            state.profile = profile.model_dump(mode="json")
            state.stage = "task"
            st.rerun()

    elif state.stage == "task":
        left, right = st.columns([1, 1.4], gap="large")
        with left:
            render_workbook()
            with st.expander("表头识别不正确？"):
                overrides = {}
                for sheet in current_profile().sheets:
                    candidates = sheet.header_candidates or [sheet.header_row or 1]
                    default_index = candidates.index(sheet.header_row) if sheet.header_row in candidates else 0
                    overrides[sheet.name] = st.selectbox(f"{sheet.name} 的表头行", candidates, index=default_index, key=f"header_{sheet.name}")
                if st.button("按所选表头重新分析"):
                    updated = service.update_headers(state.task_id, overrides)
                    state.profile = updated.model_dump(mode="json")
                    st.rerun()
        with right:
            st.subheader("描述或配置任务")
            request = st.text_area("自然语言目标", placeholder="例如：按地区汇总订单表中的订单额，计算总和和平均值。", height=100)
            if service.settings.llm_configured:
                if st.button("让 Agent 识别任务", disabled=not request.strip()):
                    with st.spinner("正在选择受控工作流并提取参数……"):
                        envelope = service.parse_task(state.task_id, request)
                    state.envelope = envelope.model_dump(mode="json")
                    st.rerun()
            else:
                st.info("尚未配置Qwen Cloud；可以直接使用下面的结构化配置。")

            parsed = TaskEnvelope.model_validate(state.envelope) if state.envelope else None
            task_types = list(TaskType)
            task_type = st.selectbox("任务类型", task_types, index=task_types.index(parsed.spec.task_type) if parsed else 0, format_func=lambda item: TASK_LABELS[item])
            st.caption("v0.2 每次只执行一个主要业务任务；预处理、动态重规划和公式生成暂不开放。")
            spec = render_task_form(task_type, parsed)
            if parsed and parsed.spec.missing_fields:
                st.warning("Agent 尚未确定：" + "、".join(parsed.spec.missing_fields))
            if parsed and parsed.spec.ambiguities:
                st.warning("需要确认：" + "；".join(parsed.spec.ambiguities))
            if st.button("生成可审阅计划", type="primary"):
                envelope = TaskEnvelope(goal=request.strip() or TASK_LABELS[task_type], spec=spec)
                plan, validation = service.make_task_plan(state.task_id, envelope)
                state.envelope = envelope.model_dump(mode="json")
                state.plan = plan.model_dump(mode="json")
                state.validation = validation.model_dump(mode="json")
                state.stage = "review"
                st.rerun()

    elif state.stage == "review":
        plan = TaskPlan.model_validate(state.plan)
        validation = state.validation
        st.subheader("审阅执行计划")
        a, b, c = st.columns(3)
        a.metric("输入版本", f"V{plan.input_version}")
        b.metric("任务类型", TASK_LABELS[plan.task_type])
        c.metric("风险等级", plan.risk_level.value.upper())
        st.markdown(f"### {plan.summary}")
        for step in plan.steps:
            preview = step.expected_changes
            st.markdown(
                f'<div class="plan-step"><strong>{step.order:02d} · {step.description}</strong><br>'
                f'<span>工具：{step.tool_name} · 预计处理 {preview.estimated_rows} 行 · '
                f'新建 Sheet：{", ".join(preview.created_sheets) or "无"} · '
                f'新增列：{", ".join(preview.new_columns) or "无"}</span></div>',
                unsafe_allow_html=True,
            )
            with st.expander("查看结构化参数"):
                st.json(step.arguments)
        for warning in validation["warnings"]:
            st.warning(warning)
        for error in validation["errors"]:
            st.error(error)
        if plan.risk_level.value == "high":
            st.error("这是覆盖原列的高风险计划。结果仍写入新文件版本，但原列会在新版本中被替换。")
            high_risk_confirmed = st.checkbox("我已确认将在新版本中覆盖原列，并保留原始上传文件。")
        else:
            high_risk_confirmed = True
        blocked = not validation["valid"]
        col1, col2, col3 = st.columns([1, 1, 2])
        if col1.button("批准并执行", type="primary", disabled=blocked or not high_risk_confirmed):
            service.approve(state.task_id, "approve", {"source": "streamlit-v2", "plan": plan.model_dump(mode="json")})
            envelope = TaskEnvelope.model_validate(state.envelope)
            with st.status("正在执行受控工具并核验……", expanded=True) as status:
                status.write("重新检查文件版本和工具参数")
                status.write("执行计划中的确定性表格工具")
                output, report, tool_result = service.execute_task_plan(state.task_id, envelope, plan)
                status.write("重新打开文件并执行通用与任务专用核验")
                status.update(label="执行完成", state="complete")
            state.result = service.result_payload(state.task_id)
            state.profile = json.loads(state.result["task"]["profile_json"])
            state.stage = "result"
            st.rerun()
        if col2.button("返回编辑"):
            service.approve(state.task_id, "edit", {"source": "streamlit-v2"})
            state.stage = "task"
            st.rerun()
        if col3.button("拒绝并取消"):
            service.approve(state.task_id, "reject", {"source": "streamlit-v2"})
            reset_task()
            st.rerun()

    elif state.stage == "result":
        payload = service.result_payload(state.task_id)
        result = payload["changeset"] or {}
        report = payload["verification"] or {}
        task = payload["task"]
        st.subheader(f"版本 V{task['current_version']} 已生成")
        st.markdown(f'<span class="status-chip">{report.get("status", "completed")}</span>', unsafe_allow_html=True)
        counters = result.get("counters") or {}
        metric_columns = st.columns(max(1, min(4, len(counters) or 1)))
        if counters:
            for index, (name, value) in enumerate(counters.items()):
                metric_columns[index % len(metric_columns)].metric(name.replace("_", " "), value)
        else:
            metric_columns[0].metric("处理行", result.get("rows_affected", 0))

        left, right = st.columns([1.1, 1], gap="large")
        with left:
            st.markdown("### 变更")
            st.write("修改工作表：", ", ".join(result.get("modified_sheets", [])) or "无")
            st.write("新建工作表：", ", ".join(result.get("created_sheets", [])) or "无")
            st.write("修改区域：", ", ".join(result.get("modified_ranges", [])) or "未记录")
            for warning in result.get("warnings", []):
                st.warning(warning)
        with right:
            st.markdown("### 核验")
            for check in report.get("checks", []):
                st.write(f"{'✓' if check['passed'] else '×'} {check['name']} — {check.get('detail', '')}")
            for warning in report.get("warnings", []):
                st.warning(warning)

        output_path = Path(task["current_path"])
        original_path = Path(task["original_path"])
        a, b, c = st.columns(3)
        a.download_button("下载当前版本", output_path.read_bytes(), file_name=f"SheetPilot_v{task['current_version']}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
        b.download_button("下载原文件", original_path.read_bytes(), file_name=task["original_name"], mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        c.download_button("下载核验报告", json.dumps(report, ensure_ascii=False, indent=2), file_name=f"verification_v{task['current_version']}.json", mime="application/json")

        st.markdown("### 版本历史")
        versions = payload["versions"]
        for version in reversed(versions):
            cols = st.columns([3, 2, 1])
            cols[0].write(f"V{version['version']} · {Path(version['path']).name}")
            cols[1].caption(version["created_at"])
            if version["version"] != task["current_version"] and cols[2].button("恢复", key=f"restore_v{version['version']}"):
                profile = service.restore_version(state.task_id, int(version["version"]))
                state.profile = profile.model_dump(mode="json")
                reset_task()
                st.rerun()
        if st.button("基于当前版本继续任务", type="primary"):
            state.profile = json.loads(service.repository.get_task(state.task_id)["profile_json"])
            reset_task()
            st.rerun()

except SheetPilotError as exc:
    st.error(f"{exc.code} · {exc.message}")
    if exc.detail:
        with st.expander("查看详情"):
            st.json(exc.detail)
except Exception as exc:
    st.error(f"UNEXPECTED_ERROR · {type(exc).__name__}: {exc}")
