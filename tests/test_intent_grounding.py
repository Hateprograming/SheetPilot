import json
from dataclasses import replace
from pathlib import Path

import httpx
from pydantic import BaseModel

from evals.phase1.runner import semantic
from sheetpilot.config import Settings
from sheetpilot.llm import QwenCompatibleClient
from sheetpilot.policy import PolicyEngine
from sheetpilot.service import SheetPilotService
from sheetpilot.schemas import (
    AggregateFunction,
    AggregateMetric,
    AggregateTaskSpec,
    ColumnProfile,
    ColumnRef,
    CleanOperation,
    CleanTaskSpec,
    DuplicatePolicy,
    FormulaMode,
    FormulaTaskSpec,
    ImputeNumericRule,
    InvalidNumericPolicy,
    LookupTaskSpec,
    MessageKind,
    ConversationMessage,
    ConversationRole,
    NormalizeOperation,
    NormalizeTaskSpec,
    SortDirection,
    SortRankTaskSpec,
    SortKey,
    SheetProfile,
    TaskEnvelope,
    TransformOperation,
    TransformTaskSpec,
    ValueMapRule,
    WorkbookProfile,
)
from sheetpilot.conversation import clarification_for
from sheetpilot.skills import SkillRegistry


def profile() -> WorkbookProfile:
    columns = [
        ColumnProfile(
            name=name, display_header=name, index=index, letter=letter,
            inferred_type="number" if name == "金额" else "text",
            empty_count=0, unique_count=3, formula_count=0,
        )
        for index, (letter, name) in enumerate(
            (("A", "编号"), ("B", "分类"), ("C", "金额"), ("D", "备注")), 1
        )
    ]
    return WorkbookProfile(
        file_name="input.xlsx", size_bytes=1, sha256="a" * 64,
        sheets=[SheetProfile(
            name="Data", state="visible", max_row=4, max_column=4,
            header_row=1, header_confidence=1, columns=columns,
            preview_rows=[{"编号": "DO_NOT_SEND_CELL_VALUE"}],
        )],
    )


def validate(spec, request: str):
    return PolicyEngine(SkillRegistry()).validate_envelope(
        profile(), TaskEnvelope(goal=request, spec=spec), user_source_text=request
    )


def test_policy_blocks_unresolved_ambiguity_and_unauthorized_lookup_choice():
    ambiguous = NormalizeTaskSpec(
        sheet="Data", column="编号", operations=[NormalizeOperation.TRIM],
        output_column="标准编号", ambiguities=["编号还是备注"], confidence=0.9,
    )
    assert any("存在歧义" in item for item in validate(ambiguous, "整理Data").errors)

    lookup = LookupTaskSpec(
        main_sheet="Data", source_sheet="Data", main_key_column="编号",
        source_key_column="编号", return_columns=["金额"],
        output_column_names=["匹配金额"], duplicate_policy=DuplicatePolicy.LAST,
        confidence=1,
    )
    result = validate(lookup, "按编号匹配金额，来源重复时停止。")
    assert any("未经用户授权" in item for item in result.errors)


def test_policy_preserves_explicit_column_output_names_and_mapping_pairs():
    wrong_column = NormalizeTaskSpec(
        sheet="Data", column="备注", operations=[NormalizeOperation.TRIM],
        output_column="标准编号", confidence=1,
    )
    result = validate(wrong_column, "把Data的编号去空格，写入新列标准编号。")
    assert any("编号" in item and "备注" in item for item in result.errors)

    wrong_output = AggregateTaskSpec(
        sheet="Data", group_by=["分类"],
        metrics=[AggregateMetric(column="金额", function="sum", output_column="金额合计")],
        output_sheet="汇总", confidence=1,
    )
    result = validate(wrong_output, "按分类汇总金额，新表叫统计结果，结果列叫统计值。")
    assert any("统计结果" in item for item in result.errors)
    assert any("统计值" in item for item in result.errors)

    mapping = TransformTaskSpec(
        sheet="Data", operation=TransformOperation.MAP_VALUES,
        map_rules=[ValueMapRule(
            source=ColumnRef(sheet="Data", display_header="分类"),
            mapping={"有": 1, "无": 0, "是": 1}, output_column="类别编码",
        )], confidence=1,
    )
    result = validate(
        mapping,
        "把Data的分类中值“有”映射为1、值“无”映射为0，新列类别编码。",
    )
    assert any("不能增加其他映射" in item for item in result.errors)


def test_model_payload_uses_safe_context_codes_and_previous_envelope():
    previous = TaskEnvelope(
        goal="整理编号",
        spec=NormalizeTaskSpec(
            sheet="Data", column="编号", operations=[NormalizeOperation.TRIM],
            output_column="标准编号", confidence=1,
        ),
    )
    captured = {}
    client = object.__new__(QwenCompatibleClient)

    def complete(system, user, model):
        captured["system"] = system
        captured["user"] = user
        return previous

    client._complete_json = complete
    result = client.parse_task(
        "输出列改叫清理编号", profile(), previous_envelope=previous,
        validation_errors=["列不存在：DO_NOT_SEND_ERROR_VALUE"],
        conversation=[
            {"role": "assistant", "content": "DO_NOT_SEND_ASSISTANT_VALUE"},
            {"role": "user", "content": "输出列改叫清理编号"},
        ],
        workbook_context={"raw_values": ["DO_NOT_SEND_CALLER_VALUE"]},
        capability_manifest=SkillRegistry().manifest(),
    )
    payload = json.loads(captured["user"])
    serialized = captured["user"]
    assert result == previous
    assert payload["previous_task"] == previous.model_dump(mode="json")
    assert payload["validation_error_codes"] == ["COLUMN_NOT_FOUND"]
    assert payload["recent_conversation"] == [
        {"role": "user", "content": "输出列改叫清理编号"}
    ]
    assert payload["workbook_context"]["selected_skill"]["id"] == "normalization_skill"
    for secret in (
        "DO_NOT_SEND_CELL_VALUE", "DO_NOT_SEND_ERROR_VALUE",
        "DO_NOT_SEND_ASSISTANT_VALUE", "DO_NOT_SEND_CALLER_VALUE",
    ):
        assert secret not in serialized


def test_clarification_clears_only_the_answered_issue():
    draft = TaskEnvelope(
        goal="填补金额",
        spec=TransformTaskSpec(
            sheet="Data", operation=TransformOperation.IMPUTE_NUMERIC,
            impute_rules=[ImputeNumericRule(
                source=ColumnRef(sheet="Data", display_header="金额"),
                output_column="填补值",
            )],
            missing_fields=["invalid_numeric：错误值怎么处理", "输出列名待确认"],
            ambiguities=["错误值留空还是填补", "是否覆盖原列"], confidence=1,
        ),
    )
    history = [
        ConversationMessage(
            task_id="t", role=ConversationRole.ASSISTANT,
            kind=MessageKind.CLARIFICATION, content="错误值怎么办？",
            metadata={"review": {"issue_code": "invalid_numeric"}},
        ),
        ConversationMessage(
            task_id="t", role=ConversationRole.USER, content="错误值也用均值填补"
        ),
    ]
    service = object.__new__(SheetPilotService)
    resumed = service._resume_clarification_envelope(draft, history[-1].content, history)
    assert resumed is not None
    spec = resumed.spec
    assert isinstance(spec, TransformTaskSpec)
    assert spec.impute_rules[0].invalid_policy == InvalidNumericPolicy.IMPUTE
    assert spec.missing_fields == ["输出列名待确认"]
    assert spec.ambiguities == ["是否覆盖原列"]


def test_transform_output_name_stops_before_conjunction():
    envelope = TaskEnvelope(
        goal="映射状态",
        spec=TransformTaskSpec(
            sheet="Data", operation=TransformOperation.MAP_VALUES,
            map_rules=[ValueMapRule(
                source=ColumnRef(sheet="Data", display_header="分类"),
                mapping={"有": 1, "无": 0}, output_column="状态编码",
            )], confidence=1,
        ),
    )
    request = "把Data的分类中有映射为1、无映射为0，新列叫状态编码并保留原列。"
    prepared = SheetPilotService.prepare_model_envelope(envelope, request, profile())
    assert prepared.spec.map_rules[0].output_column == "状态编码"
    assert validate(prepared.spec, request).valid


def test_unspecified_presentation_names_are_deterministic():
    aggregate = TaskEnvelope(
        goal="按分类汇总金额",
        spec=AggregateTaskSpec(
            sheet="Data", group_by=["分类"],
            metrics=[AggregateMetric(
                column="金额", function=AggregateFunction.SUM, output_column="模型随机构造名"
            )], output_sheet="汇总", confidence=1,
        ),
    )
    prepared = SheetPilotService.prepare_model_envelope(
        aggregate, "按分类汇总金额，结果放到汇总表。", profile()
    )
    assert prepared.spec.metrics[0].output_column == "金额_sum"

    mapping = TaskEnvelope(
        goal="映射分类",
        spec=TransformTaskSpec(
            sheet="Data", operation=TransformOperation.MAP_VALUES,
            map_rules=[ValueMapRule(
                source=ColumnRef(sheet="Data", display_header="分类"),
                mapping={"有": 1, "无": 0}, output_column="模型随机构造名",
            )], confidence=1,
        ),
    )
    prepared = SheetPilotService.prepare_model_envelope(
        mapping, "把分类中的有映射为1、无映射为0并保留原列。", profile()
    )
    assert prepared.spec.map_rules[0].output_column == ""


def test_generic_normalize_uses_default_rules_and_asks_only_for_column():
    envelope = TaskEnvelope(
        goal="标准化客户列",
        spec=NormalizeTaskSpec(
            sheet="Data", column="客户", operations=[NormalizeOperation.TRIM],
            output_column="客户_标准化", ambiguities=["列无法确定"], confidence=.7,
        ),
    )
    prepared = SheetPilotService.prepare_model_envelope(
        envelope, "把Data里客户相关的列标准化。", profile()
    )
    validation = validate(prepared.spec, "把Data里客户相关的列标准化。")
    first = clarification_for(validation, prepared, profile())
    assert prepared.spec.operations == [NormalizeOperation.NFKC, NormalizeOperation.TRIM]
    assert first.field == "column"
    assert first.issue_code == "ambiguous_column"

    history = [
        ConversationMessage(
            task_id="t", role=ConversationRole.ASSISTANT,
            kind=MessageKind.CLARIFICATION, content=first.question,
            metadata={"field": first.field, "issue_code": first.issue_code},
        ),
        ConversationMessage(task_id="t", role=ConversationRole.USER, content="编号"),
    ]
    service = object.__new__(SheetPilotService)
    resumed = service._resume_clarification_envelope(prepared, history[-1].content, history)
    assert resumed is not None
    assert resumed.spec.operations == [NormalizeOperation.NFKC, NormalizeOperation.TRIM]
    assert resumed.spec.column == "编号"


def test_formula_and_rank_clarifications_merge_into_original_goal():
    formula = TaskEnvelope(
        goal="给金额列填公式",
        spec=FormulaTaskSpec(
            sheet="Data", column="金额", mode=FormulaMode.FILL_GAPS,
            confidence=.8,
        ),
    )
    prepared = SheetPilotService.prepare_model_envelope(
        formula, "给Data的金额列填公式，处理整列。", profile()
    )
    assert prepared.spec.mode == FormulaMode.SET_FORMULA
    assert (prepared.spec.start_row, prepared.spec.end_row) == (2, 4)
    validation = validate(prepared.spec, "给Data的金额列填公式，处理整列。")
    prompt = clarification_for(validation, prepared, profile())
    assert prompt.field == "formula"
    history = [
        ConversationMessage(
            task_id="t", role=ConversationRole.ASSISTANT,
            kind=MessageKind.CLARIFICATION, content=prompt.question,
            metadata={"field": prompt.field, "issue_code": prompt.issue_code},
        ),
        ConversationMessage(task_id="t", role=ConversationRole.USER, content="=C2*2"),
    ]
    service = object.__new__(SheetPilotService)
    resumed = service._resume_clarification_envelope(prepared, "=C2*2", history)
    assert resumed.spec.mode == FormulaMode.SET_FORMULA
    assert resumed.spec.formula == "=C2*2"

    ranking = TaskEnvelope(
        goal="给Data排个名",
        spec=SortRankTaskSpec(
            sheet="Data", missing_fields=["rank_column"], confidence=.8,
        ),
    )
    rank_history = [
        ConversationMessage(
            task_id="t", role=ConversationRole.ASSISTANT,
            kind=MessageKind.CLARIFICATION, content="按哪列排名？",
            metadata={"field": "rank_column", "issue_code": "missing_sort_rank_column"},
        ),
        ConversationMessage(task_id="t", role=ConversationRole.USER, content="按金额降序"),
    ]
    resumed_rank = service._resume_clarification_envelope(
        ranking, "按金额降序", rank_history
    )
    assert resumed_rank.spec.sort_keys == [
        SortKey(column="金额", direction=SortDirection.DESC)
    ]
    assert resumed_rank.spec.rank_column == "金额"
    assert resumed_rank.spec.rank_direction == SortDirection.DESC


def test_explicit_formula_gap_request_remains_fill_gaps():
    envelope = TaskEnvelope(
        goal="填补金额公式断层",
        spec=FormulaTaskSpec(
            sheet="Data", column="金额", mode=FormulaMode.FILL_GAPS, confidence=1,
        ),
    )
    prepared = SheetPilotService.prepare_model_envelope(
        envelope, "填补Data金额列的公式断层。", profile()
    )
    assert prepared.spec.mode == FormulaMode.FILL_GAPS
    assert prepared.spec.formula is None


def test_phase1_semantic_comparison_ignores_only_inactive_fields():
    expected_sort = {
        "task_type": "sort_rank", "sheet": "Data",
        "sort_keys": [{"column": "金额", "direction": "desc"}],
        "rank_column": None, "rank_output_column": "名次",
        "rank_direction": "desc", "rank_method": "dense", "group_by": [],
    }
    actual_sort = {**expected_sort, "rank_output_column": "排名", "rank_method": "max"}
    assert semantic(expected_sort, profile()) == semantic(actual_sort, profile())

    expected_filter = {
        "task_type": "clean", "sheet": "Data", "operation": "filter_rows",
        "filters": [{"column": "金额", "operator": "gt", "value": 10}],
        "key_columns": [], "keep": "first", "duplicate_flag_column": "重复标记",
    }
    actual_filter = {**expected_filter, "keep": "last", "duplicate_flag_column": "其他名"}
    assert semantic(expected_filter, profile()) == semantic(actual_filter, profile())

    implicit_blank = {"resolution_policies": {}}
    explicit_blank = {"resolution_policies": {"invalid_input": "leave_blank"}}
    assert semantic(implicit_blank, profile()) == semantic(explicit_blank, profile())


def test_unique_explicit_names_are_bound_without_model_prefixes():
    impute = TaskEnvelope(
        goal="填补评分",
        spec=TransformTaskSpec(
            sheet="Data", operation=TransformOperation.IMPUTE_NUMERIC,
            impute_rules=[ImputeNumericRule(
                source=ColumnRef(sheet="Data", display_header="金额"),
                output_column="金额_填补", indicator_column="金额_缺失标记",
            )], confidence=1,
        ),
    )
    prepared = SheetPilotService.prepare_model_envelope(
        impute, "把Data的金额空值按均值填补，新列输出填补值和缺失标记，保留原列。"
    )
    rule = prepared.spec.impute_rules[0]
    assert rule.output_column == "填补值"
    assert rule.indicator_column == "缺失标记"

    aggregate = TaskEnvelope(
        goal="汇总",
        spec=AggregateTaskSpec(
            sheet="Data", group_by=["分类"],
            metrics=[AggregateMetric(column="金额", function="sum", output_column="合计")],
            output_sheet="汇总表", confidence=1,
        ),
    )
    prepared = SheetPilotService.prepare_model_envelope(
        aggregate, "按分类汇总金额，新建汇总表，结果列叫合计。"
    )
    assert prepared.spec.output_sheet == "汇总"

    wrong_source = TaskEnvelope(
        goal="整理编号",
        spec=NormalizeTaskSpec(
            sheet="Data", column="备注", operations=[NormalizeOperation.TRIM],
            output_column="标准编号", confidence=1,
        ),
    )
    prepared = SheetPilotService.prepare_model_envelope(
        wrong_source, "把Data的编号去空格，写入新列标准编号。", profile()
    )
    assert prepared.spec.column == "编号"


def test_unique_column_location_resolves_sheet_without_extra_clarification():
    formula = TaskEnvelope(
        goal="设置公式",
        spec=FormulaTaskSpec(
            sheet="", column="金额", start_row=2, end_row=4,
            mode=FormulaMode.SET_FORMULA, formula="=C2", recalculate=False,
            missing_fields=["sheet"], ambiguities=["无法确定工作表"], confidence=1,
        ),
    )
    prepared = SheetPilotService.prepare_model_envelope(
        formula, "在金额列填入公式 =C2，不重算。", profile()
    )
    assert prepared.spec.sheet == "Data"
    assert prepared.spec.missing_fields == []
    assert prepared.spec.ambiguities == []


def test_explicit_normalization_operations_are_not_dropped_from_composite():
    envelope = TaskEnvelope.model_validate({
        "goal": "标准化两列",
        "spec": {
            "task_type": "composite",
            "confidence": 1,
            "operations": [
                {"task_type": "normalize", "sheet": "Data", "column": "编号",
                 "operations": ["nfkc"], "output_column": "标准编号", "confidence": 1},
                {"task_type": "normalize", "sheet": "Data", "column": "备注",
                 "operations": ["upper"], "output_column": "标准备注", "confidence": 1},
            ],
        },
    })
    prepared = SheetPilotService.prepare_model_envelope(
        envelope,
        "先把编号去掉首尾空格并做 NFKC；再把备注去掉首尾空格并转成大写。",
        profile(),
    )
    first, second = prepared.spec.operations
    assert [str(item) for item in first.operations] == ["nfkc", "trim"]
    assert [str(item) for item in second.operations] == ["trim", "upper"]


def test_deepseek_payload_uses_official_non_thinking_json_parameters():
    class Payload(BaseModel):
        value: int

    captured = {}

    def respond(request: httpx.Request):
        captured.update(json.loads(request.content))
        return httpx.Response(200, json={"choices": [{"message": {"content": '{"value":1}'}}]})

    project_root = Path(__file__).resolve().parents[1]
    settings = replace(
        Settings.load(project_root),
        llm_provider="deepseek", api_key="test-key",
        base_url="https://api.deepseek.com", model="deepseek-flash",
        data_dir=project_root / ".testdata", database_path=project_root / ".testdata/test.sqlite3",
    )
    client = QwenCompatibleClient(settings)
    client.client.close()
    client.client = httpx.Client(base_url=settings.base_url, transport=httpx.MockTransport(respond))
    try:
        assert client._complete_json("return json", "request", Payload).value == 1
    finally:
        client.close()
    assert captured["thinking"] == {"type": "disabled"}
    assert captured["response_format"] == {"type": "json_object"}
    assert "enable_thinking" not in captured


def test_unauthorized_guesses_are_removed_and_explicit_revision_is_kept():
    lookup = TaskEnvelope(
        goal="匹配信息",
        spec=LookupTaskSpec(
            main_sheet="Data", source_sheet="Data", main_key_column="编号",
            source_key_column="编号", return_columns=["金额"],
            output_column_names=["金额"], confidence=0.8,
        ),
    )
    prepared = SheetPilotService.prepare_model_envelope(
        lookup, "按编号匹配相关信息。", profile()
    )
    assert prepared.spec.return_columns == []
    assert "return_columns" in prepared.spec.missing_fields

    clean = TaskEnvelope(
        goal="去重",
        spec=CleanTaskSpec(
            sheet="Data", operation=CleanOperation.REMOVE_DUPLICATES,
            key_columns=["编号"], output_sheet="清理", confidence=0.8,
        ),
    )
    prepared = SheetPilotService.prepare_model_envelope(
        clean, "Data去重保留第一条，输出清理表。", profile()
    )
    assert prepared.spec.key_columns == []
    assert "key_columns" in prepared.spec.missing_fields
    assert prepared.spec.output_sheet == "清理"

    formula = TaskEnvelope(
        goal="设置公式",
        spec=FormulaTaskSpec(
            sheet="Data", column="金额", start_row=2, end_row=4,
            mode=FormulaMode.SET_FORMULA, formula="=1", recalculate=True,
            confidence=1,
        ),
    )
    prepared = SheetPilotService.prepare_model_envelope(
        formula, "设置公式=1，仅检查文本不重算。", profile()
    )
    assert prepared.spec.recalculate is False

    impute = TaskEnvelope(
        goal="改成中位数",
        spec=TransformTaskSpec(
            sheet="Data", operation=TransformOperation.IMPUTE_NUMERIC,
            impute_rules=[ImputeNumericRule(
                source=ColumnRef(sheet="Data", display_header="金额"),
                output_column="填补值", indicator_column="缺失标记",
                overwrite_existing_outputs=True,
            )], confidence=1,
        ),
    )
    prepared = SheetPilotService.prepare_model_envelope(
        impute, "改成中位数，其他要求不变。", profile()
    )
    assert prepared.spec.impute_rules[0].overwrite_existing_outputs is False
