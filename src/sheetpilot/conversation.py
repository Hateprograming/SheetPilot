from __future__ import annotations

from .schemas import (
    AggregateTaskSpec,
    ClarificationPrompt,
    CleanTaskSpec,
    CompositeTaskSpec,
    ConversionDirection,
    ConversionTaskSpec,
    FormulaTaskSpec,
    LookupTaskSpec,
    NormalizeTaskSpec,
    SortRankTaskSpec,
    TaskEnvelope,
    TaskPlan,
    TaskType,
    TaskValidation,
    TransformTaskSpec,
    WorkbookProfile,
)


TASK_LABELS = {
    TaskType.LOOKUP: "跨表匹配",
    TaskType.NORMALIZE: "数据标准化",
    TaskType.AGGREGATE: "聚合汇总",
    TaskType.SORT_RANK: "排序与排名",
    TaskType.CLEAN: "数据清洗",
    TaskType.FORMULA: "公式诊断与修复",
    TaskType.CONVERSION: "文件转换",
    TaskType.TRANSFORM: "可审计数据准备",
    TaskType.COMPOSITE: "复合任务",
}

RISK_LABELS = {"low": "低", "medium": "中", "high": "高"}


def workbook_greeting(profile: WorkbookProfile) -> str:
    sheets = "、".join(
        f"“{sheet.name}”({max(0, sheet.max_row - (sheet.header_row or 1))} 行 × {sheet.max_column} 列)"
        for sheet in profile.sheets
    )
    risk = f" 我还发现：{'；'.join(profile.risks)}" if profile.risks else ""
    return (
        f"我已经完成本地结构分析。这个工作簿包含 {len(profile.sheets)} 个工作表：{sheets}。"
        f"{risk}\n\n告诉我你希望怎样处理它即可，我会先复述理解并给出执行计划，不会直接修改原文件。"
    )


def clarification_for(
    validation: TaskValidation, envelope: TaskEnvelope, profile: WorkbookProfile
) -> ClarificationPrompt:
    errors = "；".join(validation.errors)
    spec = envelope.spec
    sheet_names = [sheet.name for sheet in profile.sheets]

    def source_labels() -> list[str]:
        refs = []
        if isinstance(spec, TransformTaskSpec):
            if spec.operation.value == "impute_numeric":
                refs = [rule.source for rule in spec.impute_rules]
            elif spec.operation.value == "map_values":
                refs = [rule.source for rule in spec.map_rules]
            elif spec.operation.value == "parse_dimensions" and spec.dimension_rule:
                refs = [spec.dimension_rule.source]
            elif spec.operation.value == "derive_metric" and spec.derive_rule:
                refs = list(spec.derive_rule.sources)
        labels = [
            ref.display_header or ref.column_letter or str(ref.column_index or "")
            for ref in refs
        ]
        return list(dict.fromkeys(label for label in labels if label))

    def source_phrase(default: str) -> str:
        labels = source_labels()
        if not labels:
            return default
        return "、".join(f"“{label}”" for label in labels) + "列"

    if any(
        token in errors
        for token in ("体积公式未在用户原始要求中明确选择", "派生指标必须明确选择注册公式模板")
    ):
        return ClarificationPrompt(
            question="计算体积需要明确公式。你希望使用哪一种？",
            choices=[
                "椭球体积 π/6×L×W×H",
                "半乘积体积 L×W×H÷2",
                "长方体积 L×W×H",
                "三维用 L×W×H，单一直径用 D³",
            ],
            field="formula_template",
            issue_code="derived_formula",
        )

    if "DERIVE_RESOLUTION_REQUIRED:" in errors:
        marker = next(
            item for item in validation.errors
            if item.startswith("DERIVE_RESOLUTION_REQUIRED:")
        )
        _, _, count = marker.split(":", 2)
        return ClarificationPrompt(
            question=(
                f"检查到{source_phrase('派生指标来源列')}中有 {count} 条记录的输入无法用于计算派生指标。"
                "请查看下方记录后选择统一处理方式，也可以直接输入具体的来源修正规则。"
            ),
            choices=[
                "保持结果留空并生成异常清单",
                "将这类无效派生结果统一填 0",
                "我输入具体的来源修正规则",
            ],
            field="resolution_policy",
            issue_code="invalid_input",
        )

    if "DIMENSION_RESOLUTION_REQUIRED:" in errors:
        marker = next(
            item for item in validation.errors
            if item.startswith("DIMENSION_RESOLUTION_REQUIRED:")
        )
        _, status, count = marker.split(":", 2)
        labels = {
            "multiple_candidates": "包含多组候选尺寸，系统无法确定应选择哪一组",
            "unit_conflict": "同时使用了不同单位，需要确认是否分别换算为 mm",
            "invalid_format": "格式无法可靠解析为尺寸",
            "no_numeric_measurement": "没有可识别的尺寸数值",
        }
        choices = {
            "multiple_candidates": [
                "这类记录统一取第一组尺寸",
                "这类记录统一取最后一组尺寸",
                "保持留空并生成异常清单",
            ],
            "unit_conflict": [
                "按每个数值的显式单位分别换算为 mm",
                "保持留空并生成异常清单",
            ],
            "invalid_format": [
                "这类记录统一填 0",
                "保持留空并生成异常清单",
                "我输入具体修正规则",
            ],
            "no_numeric_measurement": [
                "这类记录统一填 0",
                "保持留空并生成异常清单",
            ],
        }
        return ClarificationPrompt(
            question=(
                f"检查到{source_phrase('尺寸列')}中有 {count} 条记录"
                f"{labels.get(status, '属于同一类待确认问题')}。"
                "请查看下方记录后选择统一处理方式，也可以直接输入自己的规则。"
            ),
            choices=choices.get(status, ["保持留空并生成异常清单"]),
            field="resolution_policy",
            issue_code=status,
        )

    if "IMPUTE_RESOLUTION_REQUIRED:" in errors:
        marker = next(
            item for item in validation.errors
            if item.startswith("IMPUTE_RESOLUTION_REQUIRED:")
        )
        _, _, count = marker.split(":", 2)
        return ClarificationPrompt(
            question=(
                f"检查到{source_phrase('目标数值列')}中有 {count} 条非空内容无法解析为数值。"
                "这些内容不属于普通空单元格，请查看下方记录后确认怎样处理。"
            ),
            choices=[
                "将这类错误值也用同一均值填补",
                "保持结果留空并生成异常清单",
                "我输入哪些值应视为缺失",
            ],
            field="invalid_policy",
            issue_code="invalid_numeric",
        )

    if "尺寸异常处理“first_candidate”未在用户要求中明确授权" in errors:
        return ClarificationPrompt(
            question="尺寸列中存在多组候选值。你希望怎样处理这类记录？",
            choices=[
                "这类记录统一取第一组尺寸",
                "这类记录统一取最后一组尺寸",
                "保持留空并生成异常清单",
            ],
            field="resolution_policy",
            issue_code="multiple_candidates",
        )
    if "尺寸异常处理“last_candidate”未在用户要求中明确授权" in errors:
        return ClarificationPrompt(
            question="尺寸列中存在多组候选值。你希望怎样处理这类记录？",
            choices=[
                "这类记录统一取最后一组尺寸",
                "这类记录统一取第一组尺寸",
                "保持留空并生成异常清单",
            ],
            field="resolution_policy",
            issue_code="multiple_candidates",
        )

    if "return_columns" in errors:
        choices: list[str] = []
        if isinstance(spec, LookupTaskSpec):
            source = next((item for item in profile.sheets if item.name == spec.source_sheet), None)
            choices = [
                item.name for item in (source.columns if source else [])
                if item.name != spec.source_key_column
            ][:4]
        return ClarificationPrompt(
            question="你希望从来源表补充哪些列？", choices=choices,
            field="return_columns", issue_code="missing_return_columns",
        )
    if "来源表关联键存在重复" in errors or "重复" in errors and isinstance(spec, LookupTaskSpec):
        return ClarificationPrompt(
            question="来源表中发现同一关联编号对应多条记录。你希望怎样处理？",
            choices=["使用第一条", "使用最后一条", "停止处理"],
            field="duplicate_policy", issue_code="duplicate_source_key",
        )
    if "group_by" in errors or "至少需要一个分组列" in errors:
        return ClarificationPrompt(
            question="你希望按哪一列进行分组汇总？",
            choices=_columns(profile, getattr(spec, "sheet", ""))[:4],
            field="group_by", issue_code="missing_group_by",
        )
    if "metrics" in errors or "至少需要一个指标" in errors:
        numeric = [
            column.name
            for sheet in profile.sheets if sheet.name == getattr(spec, "sheet", "")
            for column in sheet.columns if column.inferred_type == "number"
        ]
        return ClarificationPrompt(
            question="需要计算哪个数值列？也请告诉我是求和、平均、计数、最大值还是最小值。",
            choices=[f"汇总{item}" for item in numeric[:4]],
            field="metrics", issue_code="missing_metrics",
        )
    if "operations" in errors or "至少选择一种标准化操作" in errors:
        return ClarificationPrompt(
            question="你希望怎样统一这一列？",
            choices=["去除首尾空格", "统一全角和半角", "转为数值", "转为文本"],
            field="operations", issue_code="missing_normalize_operations",
        )
    if isinstance(spec, NormalizeTaskSpec) and (
        ("列" in errors and any(token in errors for token in ("歧义", "无法确定", "不唯一")))
        or "column" in errors
    ):
        candidates = [
            column.name
            for sheet in profile.sheets
            for column in sheet.columns
            if not spec.sheet or sheet.name == spec.sheet
        ]
        stem = spec.column.replace("客户", "").replace("文本", "") if spec.column else ""
        filtered = [name for name in candidates if stem and stem in name]
        return ClarificationPrompt(
            question="需要标准化哪一列？请从候选列中选择。",
            choices=list(dict.fromkeys(filtered or candidates))[:6],
            field="column", issue_code="ambiguous_column",
        )
    if isinstance(spec, FormulaTaskSpec) and "多个同名表头" in errors:
        return ClarificationPrompt(
            question=f"“{spec.column}”对应多个同名表头，请选择目标列的列字母。",
            choices=[f"{col.letter} 列（{col.display_header or col.name}）" for sheet in profile.sheets if sheet.name == spec.sheet for col in sheet.columns if (col.display_header or col.name) == spec.column],
            field="column", issue_code="ambiguous_formula_column",
        )
    if isinstance(spec, FormulaTaskSpec) and any(
        token in errors for token in ("设置公式时必须提供", "缺少信息：formula", "公式文本")
    ):
        return ClarificationPrompt(
            question=f"请说明“{spec.column or '目标'}”列的计算规则，例如“数量×单价”，也可以提供以 = 开头的 Excel 公式；列名需能唯一确定。",
            choices=[], field="formula", issue_code="missing_formula",
        )
    if "至少指定一个排序列或排名列" in errors:
        return ClarificationPrompt(
            question="你希望按哪一列排序或排名？",
            choices=_columns(profile, getattr(spec, "sheet", ""))[:4],
            field="rank_column", issue_code="missing_sort_rank_column",
        )
    if "重复记录操作至少需要一个判断列" in errors:
        return ClarificationPrompt(
            question="应根据哪些列判断两条记录重复？",
            choices=_columns(profile, getattr(spec, "sheet", ""))[:4],
            field="key_columns", issue_code="missing_duplicate_keys",
        )
    if "筛选任务至少需要一个条件" in errors:
        return ClarificationPrompt(
            question="需要保留哪些记录？请告诉我筛选列、条件和值。",
            choices=[],
            field="filters", issue_code="missing_filters",
        )
    if "已存在" in errors or "冲突" in errors:
        return ClarificationPrompt(
            question="目标位置已经存在同名内容。请告诉我一个新的输出列名或工作表名。",
            choices=["换一个名称", "停止处理"],
            field="output_name", issue_code="output_conflict",
        )
    if isinstance(spec, ConversionTaskSpec):
        if spec.direction == ConversionDirection.XLSX_TO_CSV and not spec.source_sheet:
            return ClarificationPrompt(
                question="你希望将哪个工作表导出为 CSV？",
                choices=sheet_names[:4],
                field="source_sheet", issue_code="missing_source_sheet",
            )
        if any("分隔符" in item for item in validation.errors):
            return ClarificationPrompt(
                question="我无法可靠判断 CSV 分隔符，请选择实际使用的分隔符。",
                choices=["逗号", "制表符", "分号", "竖线"],
                field="delimiter", issue_code="ambiguous_delimiter",
            )
    if any(token in errors for token in (
        "main_sheet", "source_sheet", "未指定工作表", "工作表不能为空", "工作表“"
    )) and any(token in errors for token in ("不存在", "缺少", "未指定", "为空")):
        return ClarificationPrompt(
            question="我还不能确定要处理哪个工作表，请选择或直接告诉我名称。",
            choices=sheet_names[:4],
            field="sheet", issue_code="missing_sheet",
        )
    if any(token in errors for token in ("column", "关联列", "不存在或不唯一")):
        return ClarificationPrompt(
            question="我还不能确定要使用哪一列，请直接告诉我列名和用途。",
            choices=list(dict.fromkeys(
                column.name for sheet in profile.sheets for column in sheet.columns
            ))[:6],
            field="column", issue_code="missing_column",
        )
    unresolved = next(
        (item for item in validation.errors if "缺少信息：" in item or "存在歧义：" in item),
        validation.errors[0] if validation.errors else "需要补充任务规则",
    )
    return ClarificationPrompt(
        question=f"我还不能执行，因为：{unresolved}。请补充这一项。",
        choices=[],
        field="unknown", issue_code="unresolved_validation",
    )


def _columns(profile: WorkbookProfile, sheet_name: str) -> list[str]:
    sheet = next((item for item in profile.sheets if item.name == sheet_name), None)
    return [column.name for column in sheet.columns] if sheet else []


def plan_text(envelope: TaskEnvelope, plan: TaskPlan, validation: TaskValidation) -> str:
    spec = envelope.spec
    skills = "、".join(plan.skill_ids) or "未标记"
    lines = [
        f"我理解你要完成的是：**{envelope.goal or plan.summary}**",
        f"使用 Skill：`{skills}`",
        "",
        "执行前请确认：",
    ]
    for step in plan.steps:
        lines.append(f"{step.order}. {step.description}")
        if step.expected_changes.summary_counts:
            preview = "、".join(
                f"{_counter_label(key)} {value}"
                for key, value in step.expected_changes.summary_counts.items()
            )
            lines.append(f"   - 预检结果：{preview}")
    modified = list(dict.fromkeys(name for step in plan.steps for name in step.expected_changes.modified_sheets))
    created = list(dict.fromkeys(name for step in plan.steps for name in step.expected_changes.created_sheets))
    columns = list(dict.fromkeys(name for step in plan.steps for name in step.expected_changes.new_columns))
    estimated_rows = sum(step.expected_changes.estimated_rows for step in plan.steps)
    if modified:
        lines.append(f"- 将修改：{'、'.join(modified)}")
    if created:
        lines.append(f"- 将新建：{'、'.join(created)}")
    if columns:
        lines.append(f"- 将新增列：{'、'.join(columns)}")
    lines.extend(
        [
            f"- 预计处理量：约 {estimated_rows} 行次",
            f"- 风险等级：{RISK_LABELS[plan.risk_level.value]}",
            f"- 输入版本：V{plan.base_version} · 计划修订 R{plan.revision}",
            "- 中间步骤只保存为内部检查点，全部核验通过后才生成一个新版本",
            "- 原始上传文件不会被覆盖",
        ]
    )
    if validation.warnings:
        lines.append(f"\n需要留意：{'；'.join(validation.warnings)}")
    if plan.plan_assumptions:
        lines.append(f"\n采用的安全假设：{'；'.join(plan.plan_assumptions)}")
    if plan.confirmation_requirements:
        lines.append("\n**额外确认：本计划会覆盖当前工作表中的已有内容（包括修改已有列或删除未保留列），只有勾选覆盖风险确认后才能执行。**")
    return "\n".join(lines)


def result_text(version: int, result, report) -> str:
    changed = "、".join(result.modified_sheets) or "无"
    created = "、".join(result.created_sheets) or "无"
    counters = "，".join(f"{_counter_label(key)} {value}" for key, value in result.counters.items())
    status = "已完成既定检查" if report.status == "passed" else "已完成，存在需要留意的项目"
    text = (
        f"**{status}，已生成 V{version}。**\n\n"
        f"- 修改工作表：{changed}\n"
        f"- 新建工作表：{created}\n"
        f"- 处理范围：{counters or f'{result.rows_affected} 行'}"
    )
    warnings = [*result.warnings, *report.warnings]
    if warnings:
        text += f"\n- 提醒：{'；'.join(dict.fromkeys(warnings))}"
    if result.tool_name == "transaction":
        text += "\n- 各步骤分别统计影响范围，同一行可能重复参与，不代表去重记录数。"
    text += "\n- 业务规则含义及上游数据是否最新仍需人工确认。"
    return text


def _counter_label(name: str) -> str:
    return {
        "matched": "匹配",
        "unmatched": "未匹配",
        "empty_key": "空键",
        "duplicate_key": "重复键",
        "input_rows": "输入",
        "output_rows": "输出",
        "groups": "分组",
        "non_numeric": "非数值",
        "duplicates_removed": "移除重复",
        "duplicate_rows": "重复记录",
        "input_columns": "原有列",
        "output_columns": "保留列",
        "kept_columns": "保留列",
        "deleted_columns": "删除列",
        "formula_count": "公式",
        "gaps": "断层",
        "outliers": "异常公式",
        "broken_refs": "损坏引用",
        "repaired": "已修复",
        "generated": "已生成公式",
        "copied": "已复制公式",
        "calculated": "已重算",
        "calculation_errors": "计算错误",
        "success": "成功",
        "valid": "有效",
        "invalid": "格式错误",
        "invalid_imputed": "错误值经授权填补",
        "missing": "缺失",
        "errors": "异常",
        "exceptions": "异常",
        "imputed": "已填补",
        "mapped": "已映射",
        "mapped_blank": "空值按明确规则映射",
        "mapped_default": "其他值按明确规则映射",
        "mapping_coverage_pct": "映射覆盖率(%)",
        "unknown": "未知值",
        "blank": "空值",
        "observed_basis": "有效样本",
        "planned_fill_value": "计划填补值",
        "parsed_3d": "三维解析",
        "parsed_2d": "二维解析",
        "diameter_only": "单径解析",
        "multiple_candidates": "多组候选",
        "unit_conflict": "单位冲突",
        "invalid_format": "格式错误",
        "no_numeric_measurement": "无数值尺寸",
        "resolved_multiple_first": "多候选取第一组",
        "resolved_multiple_last": "多候选取最后一组",
        "resolved_unit_conflict": "混合单位已换算",
        "resolved_zero": "异常经授权置零",
        "corrected_value": "按明确规则修正",
        "derived": "已派生",
        "input_missing": "输入缺失",
        "invalid_input": "无效输入",
        "resolved_invalid_zero": "无效派生结果经授权置零",
        "resolved_missing_zero": "缺失派生结果经授权置零",
        "deferred_to_dependency": "将在前置步骤完成后预检",
    }.get(name, name.replace("_", " "))
