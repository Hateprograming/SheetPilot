from __future__ import annotations

from .formula_intent import column_formula

import re
import unicodedata

from .schemas import (
    AggregateTaskSpec,
    CleanTaskSpec,
    DuplicatePolicy,
    CompositeTaskSpec,
    DimensionResolutionAction,
    FormulaMode,
    FormulaTaskSpec,
    LookupTaskSpec,
    NormalizeTaskSpec,
    SortRankTaskSpec,
    DerivedFormulaTemplate,
    DerivedResolutionAction,
    TaskEnvelope,
    TaskType,
    TaskValidation,
    TransformOperation,
    TransformTaskSpec,
    InvalidNumericPolicy,
    WorkbookProfile,
)
from .skills import SkillRegistry


class PolicyEngine:
    """Deterministic policy checks that do not rely on the model's instructions."""

    def __init__(self, skills: SkillRegistry):
        self.skills = skills

    def validate_envelope(
        self,
        profile: WorkbookProfile,
        envelope: TaskEnvelope,
        *,
        user_source_text: str | None = None,
    ) -> TaskValidation:
        errors: list[str] = []
        warnings: list[str] = []
        specs = (
            list(envelope.spec.operations)
            if isinstance(envelope.spec, CompositeTaskSpec)
            else [envelope.spec]
        )
        if isinstance(envelope.spec, CompositeTaskSpec):
            errors.extend(f"复合任务缺少信息：{item}" for item in envelope.spec.missing_fields)
            errors.extend(f"复合任务存在歧义：{item}" for item in envelope.spec.ambiguities)
            if envelope.spec.confidence < 0.6:
                errors.append("复合任务整体置信度过低，需要进一步澄清。")
        for index, spec in enumerate(specs, 1):
            errors.extend(f"步骤 {index} 缺少信息：{item}" for item in spec.missing_fields)
            errors.extend(f"步骤 {index} 存在歧义：{item}" for item in spec.ambiguities)
            try:
                self.skills.get(spec.task_type)
            except KeyError:
                errors.append(f"步骤 {index} 使用了未注册的任务类型 {spec.task_type}。")
                continue
            if spec.task_type == TaskType.CONVERSION and len(specs) > 1:
                if spec.direction.value != "xlsx_to_csv" or index != len(specs):
                    errors.append("XLSX 导出 CSV 必须是最后一步；CSV 转 XLSX 必须单独执行。")
            if user_source_text is not None:
                errors.extend(
                    self._ground_explicit_intent(
                        profile, spec, user_source_text, index, len(specs)
                    )
                )
            if (
                isinstance(spec, FormulaTaskSpec)
                and spec.mode == FormulaMode.SET_FORMULA
                and spec.formula
                and user_source_text is not None
                and spec.formula not in user_source_text
                and spec.formula != column_formula(profile, spec.sheet, spec.column, user_source_text, spec.start_row)
            ):
                errors.append(f"步骤 {index} 的设置公式未出现在用户原始要求中，请明确提供公式文本。")
            if (
                isinstance(spec, TransformTaskSpec)
                and spec.operation == TransformOperation.DERIVE_METRIC
                and spec.derive_rule
                and spec.derive_rule.formula_template
                and user_source_text is not None
            ):
                cues = {
                    DerivedFormulaTemplate.ELLIPSOID_PI_OVER_6: ("π/6", "pi/6", "椭球"),
                    DerivedFormulaTemplate.HALF_PRODUCT: ("÷2", "/2", "半乘积"),
                    DerivedFormulaTemplate.RECTANGULAR_PRODUCT: ("长方体", "l×w×h", "l*w*h"),
                    DerivedFormulaTemplate.RECTANGULAR_OR_DIAMETER_CUBED: (
                        "直径^3", "直径³", "直径的三次方", "diameter^3"
                    ),
                }[spec.derive_rule.formula_template]
                normalized = user_source_text.casefold().replace(" ", "")
                if not any(cue.casefold().replace(" ", "") in normalized for cue in cues):
                    errors.append(
                        f"步骤 {index} 的体积公式未在用户原始要求中明确选择；"
                        "请指定椭球体积、半乘积体积、长方体积或三维乘积/单径立方。"
                    )
            if isinstance(spec, TransformTaskSpec) and user_source_text is not None:
                normalized = unicodedata.normalize("NFKC", user_source_text).casefold()
                compact = re.sub(r"\s+", "", normalized)
                if spec.operation == TransformOperation.IMPUTE_NUMERIC:
                    for rule in spec.impute_rules:
                        if (
                            rule.invalid_policy == InvalidNumericPolicy.IMPUTE
                            and not (
                                any(cue in compact for cue in ("错误", "invalid"))
                                and any(cue in compact for cue in ("填补", "填充", "平均", "中位", "常数"))
                            )
                        ):
                            errors.append(
                                f"步骤 {index} 未获得将格式错误值作为填补目标的明确授权。"
                            )
                if spec.operation == TransformOperation.MAP_VALUES:
                    explicit_one_hot = (
                        "indicator" in compact
                        and any(cue in compact for cue in ("分别", "各", "每个"))
                    )
                    for rule in spec.map_rules:
                        positive_count = sum(
                            value in (1, True) for value in rule.mapping.values()
                        )
                        if (
                            explicit_one_hot
                            and positive_count > 1
                            and rule.unmapped_value in (0, False)
                        ):
                            errors.append(
                                f"步骤 {index} 将多个用户要求分开的 indicator "
                                "合并到了同一列。"
                            )
                        if rule.blank_value is not None and not (
                            any(cue in compact for cue in ("没有记录", "无记录", "空值", "缺失"))
                            and str(rule.blank_value).casefold() in normalized
                            or (rule.blank_value == 0 and "算作无" in compact)
                        ):
                            errors.append(
                                f"步骤 {index} 的空值映射未在用户要求中明确授权。"
                            )
                        if rule.unmapped_value is not None and not (
                            any(cue in compact for cue in ("其他", "其余", "除此以外"))
                            and str(rule.unmapped_value).casefold() in normalized
                        ):
                            errors.append(
                                f"步骤 {index} 的其他值默认映射未在用户要求中明确授权。"
                            )
                if (
                    spec.operation == TransformOperation.PARSE_DIMENSIONS
                    and spec.dimension_rule is not None
                ):
                    cue_groups = {
                        DimensionResolutionAction.FIRST_CANDIDATE: ("取前", "第一组", "前者"),
                        DimensionResolutionAction.LAST_CANDIDATE: ("取后", "最后一组", "后者"),
                        DimensionResolutionAction.CONVERT_EACH_TO_MM: ("换算为mm", "统一为mm", "转为mm"),
                        DimensionResolutionAction.ZERO: ("填0", "为0", "统一为0", "置零"),
                    }
                    for action in spec.dimension_rule.resolution_policies.values():
                        if action == DimensionResolutionAction.LEAVE_BLANK:
                            continue
                        if not any(cue in compact for cue in cue_groups[action]):
                            errors.append(
                                f"步骤 {index} 的尺寸异常处理“{action.value}”未在用户要求中明确授权。"
                            )
                    if spec.dimension_rule.value_corrections and not any(
                        cue in compact for cue in ("改为", "修正为", "更正为", "应为", "替换为", "为")
                    ):
                        errors.append(f"步骤 {index} 的尺寸值修正规则未在用户要求中明确提供。")
                if (
                    spec.operation == TransformOperation.DERIVE_METRIC
                    and spec.derive_rule is not None
                ):
                    for action in spec.derive_rule.resolution_policies.values():
                        if action == DerivedResolutionAction.LEAVE_BLANK:
                            continue
                        if action == DerivedResolutionAction.ZERO and not any(
                            cue in compact for cue in ("填0", "为0", "统一为0", "置零")
                        ):
                            errors.append(
                                f"步骤 {index} 的无效派生结果置零未在用户要求中明确授权。"
                            )
        return TaskValidation(valid=not errors, errors=errors, warnings=warnings)
    LookupTaskSpec,
    NormalizeTaskSpec,
    SortRankTaskSpec,

    @staticmethod
    def _ground_explicit_intent(
        profile, spec, source: str, index: int, operation_count: int
    ) -> list[str]:
        """Reject model fields that contradict values stated plainly by the user.

        Extraction is deliberately conservative: if a phrase can refer to several
        fields, Policy leaves it to clarification instead of rewriting the plan.
        """
        errors: list[str] = []
        text = unicodedata.normalize("NFKC", source)
        latest = next((line.strip() for line in reversed(text.splitlines()) if line.strip()), text)
        compact = re.sub(r"\s+", "", text).casefold()

        # Explicit output names are execution-visible and must be copied literally.
        sheet_pattern = (
            r"(?:写入|输出到|筛到)?新(?:工作)?表(?:叫|名为)?\s*[‘'\"“]?"
            r"(?P<name>[^\s，,；;。的‘’'\"“”]+)"
        )
        expected_sheets = [
            match.group("name").strip(" \t'\"“”")
            for match in re.finditer(sheet_pattern, text, re.IGNORECASE)
        ]
        actual_sheet = getattr(spec, "output_sheet", None)
        if expected_sheets and actual_sheet and actual_sheet not in expected_sheets:
            errors.append(
                f"步骤 {index} 的输出工作表名应采用“{'、'.join(expected_sheets)}”中的对应名称，"
                f"模型却给出“{actual_sheet}”。"
            )

        field_specific_name = isinstance(spec, (AggregateTaskSpec, SortRankTaskSpec))
        column_cue = (
            "结果列" if isinstance(spec, AggregateTaskSpec)
            else "排名列" if isinstance(spec, SortRankTaskSpec)
            else r"(?:新列|新增(?:列)?)"
        )
        named_column_pattern = (
            rf"(?:{column_cue})(?:叫|名为)?\s*[‘'\"“]?"
            r"(?P<name>[^\s，,；;。、和‘’'\"“”]+?)"
            r"(?=\s*(?:并|且|同时|保留|，|,|；|;|。|、|$))"
        )
        expected_columns = []
        for match in re.finditer(named_column_pattern, text, re.IGNORECASE):
            name = match.group("name").strip(" \t'\"“”")
            if match.group(0).startswith("新增") and name.endswith("列"):
                name = name[:-1]
            expected_columns.append(name)
        output_columns: list[str] = []
        if isinstance(spec, LookupTaskSpec):
            output_columns = list(spec.output_column_names)
        elif isinstance(spec, NormalizeTaskSpec):
            output_columns = [spec.output_column]
        elif isinstance(spec, AggregateTaskSpec):
            output_columns = [metric.output_column for metric in spec.metrics]
        elif isinstance(spec, SortRankTaskSpec) and spec.rank_column:
            output_columns = [spec.rank_output_column]
        elif isinstance(spec, TransformTaskSpec):
            if spec.operation == TransformOperation.IMPUTE_NUMERIC:
                output_columns = [
                    name for rule in spec.impute_rules
                    for name in (rule.output_column, rule.indicator_column) if name
                ]
            elif spec.operation == TransformOperation.MAP_VALUES:
                output_columns = [
                    name for rule in spec.map_rules
                    for name in (rule.output_column, rule.status_column) if name
                ]
            elif spec.operation == TransformOperation.PARSE_DIMENSIONS and spec.dimension_rule:
                output_columns = [
                    name for name in (
                        spec.dimension_rule.length_column,
                        spec.dimension_rule.width_column,
                        spec.dimension_rule.height_column,
                        spec.dimension_rule.max_dimension_column,
                        spec.dimension_rule.status_column,
                    ) if name
                ]
            elif spec.operation == TransformOperation.DERIVE_METRIC and spec.derive_rule:
                output_columns = [
                    name for name in (
                        spec.derive_rule.output_column, spec.derive_rule.status_column
                    ) if name
                ]
        if (
            expected_columns and output_columns
            and not set(expected_columns).intersection(output_columns)
            and (operation_count == 1 or field_specific_name)
        ):
            errors.append(
                f"步骤 {index} 的输出列名应包含“{expected_columns[-1]}”，模型计划未保留该名称。"
            )

        # A first/last choice changes lookup results. It needs direct authorization.
        if isinstance(spec, LookupTaskSpec) and spec.duplicate_policy != DuplicatePolicy.BLOCK:
            authorized = {
                DuplicatePolicy.FIRST: ("取第一", "取最先", "首条", "第一条"),
                DuplicatePolicy.LAST: ("取最后", "末条", "最后一条"),
            }[spec.duplicate_policy]
            if not any(cue in compact for cue in authorized):
                errors.append(
                    f"步骤 {index} 的来源重复处理策略“{spec.duplicate_policy.value}”未经用户授权。"
                )

        if isinstance(spec, AggregateTaskSpec) and spec.filters and not any(
            cue in compact for cue in ("筛选", "过滤", "只统计", "仅统计")
        ):
            errors.append(f"步骤 {index} 增加了用户没有要求的行筛选条件。")

        # Only compare mappings when the wording has an unambiguous key/value pair.
        explicit_pairs: dict[str, str] = {}
        pair_pattern = re.compile(
            r"值\s*(?:[‘'\"“](?P<quoted_key>[^‘’'\"“”]+)[’'\"”]|"
            r"(?P<plain_key>[^\s，,；;。、‘’'\"“”]+))"
            r"\s*映射为\s*[‘'\"“]?(?P<value>[^\s，,；;。、‘’'\"“”]+)[’'\"”]?"
        )
        for match in pair_pattern.finditer(latest):
            explicit_pairs[match.group("quoted_key") or match.group("plain_key")] = match.group("value")
        if explicit_pairs and isinstance(spec, TransformTaskSpec) and spec.operation == TransformOperation.MAP_VALUES:
            planned: dict[str, str] = {}
            for rule in spec.map_rules:
                planned.update({str(key): str(value) for key, value in rule.mapping.items()})
            if planned != explicit_pairs:
                errors.append(
                    f"步骤 {index} 的映射必须逐字采用用户给出的键和值，且不能增加其他映射。"
                )

        # For a single-column operation, one uniquely named workbook column is a
        # strong constraint. Multiple mentioned columns are left for clarification.
        target: str | None = None
        if isinstance(spec, NormalizeTaskSpec):
            target = spec.column
        elif isinstance(spec, TransformTaskSpec):
            refs = []
            if spec.operation == TransformOperation.IMPUTE_NUMERIC:
                refs = [rule.source for rule in spec.impute_rules]
            elif spec.operation == TransformOperation.MAP_VALUES:
                refs = [rule.source for rule in spec.map_rules]
            elif spec.operation == TransformOperation.PARSE_DIMENSIONS and spec.dimension_rule:
                refs = [spec.dimension_rule.source]
            if len(refs) == 1:
                target = refs[0].display_header or refs[0].column_letter
        if target:
            known = {
                column.display_header or column.name
                for sheet in profile.sheets for column in sheet.columns
                if (column.display_header or column.name)
            }
            mentioned = [name for name in known if name and name in latest]
            if len(mentioned) == 1 and mentioned[0] != target:
                errors.append(
                    f"步骤 {index} 应处理用户明确点名的列“{mentioned[0]}”，模型却选择了“{target}”。"
                )
        return errors
