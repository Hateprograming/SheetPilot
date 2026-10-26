from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

from pydantic import BaseModel

from .context import build_context
from .schemas import (
    AggregateTaskSpec,
    CapabilityRefusal,
    CleanTaskSpec,
    CompositeTaskSpec,
    ConversionTaskSpec,
    FormulaTaskSpec,
    LookupTaskSpec,
    NormalizeTaskSpec,
    SkillSelection,
    SortRankTaskSpec,
    TaskEnvelope,
    TaskType,
    TransformTaskSpec,
    WorkbookProfile,
)


@dataclass(frozen=True)
class SkillDefinition:
    id: str
    version: str
    task_type: TaskType
    label: str
    description: str
    spec_model: type[BaseModel]
    tool_name: str
    context_fields: tuple[str, ...]
    boundary_notes: tuple[str, ...]
    verification_rules: tuple[str, ...]
    recovery_policy: str
    composable: bool = True

    def manifest(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "version": self.version,
            "task_type": self.task_type.value,
            "label": self.label,
            "description": self.description,
            "input_schema": self.spec_model.model_json_schema(),
            "allowed_tools": [self.tool_name],
            "boundary_notes": list(self.boundary_notes),
        }

    def build_context(self, profile: WorkbookProfile, request: str) -> dict[str, Any]:
        context = build_context(profile, request, self.task_type)
        base_sheet_fields = {"name", "rows", "columns_count", "header_row", "columns"}
        allowed = base_sheet_fields | set(self.context_fields)
        context["sheets"] = [
            {key: value for key, value in sheet.items() if key in allowed}
            for sheet in context.get("sheets", [])
        ]
        context["selected_skill"] = {
            "id": self.id,
            "version": self.version,
            "context_fields": list(self.context_fields),
        }
        return context


class SkillRegistry:
    def __init__(self, skills: list[SkillDefinition] | None = None):
        self._skills = {skill.task_type: skill for skill in (skills or self._defaults())}

    @staticmethod
    def _defaults() -> list[SkillDefinition]:
        common = ("original_file_unchanged", "input_version_unchanged", "file_reopen")
        return [
            SkillDefinition("lookup_skill", "1.0", TaskType.LOOKUP, "跨表匹配", "单工作簿内的单列精确左连接。", LookupTaskSpec, "lookup", ("sheets", "columns", "key_candidates"), ("不支持模糊匹配、联合键、一对多展开。",), (*common, "row_count_preserved"), "数据歧义需用户确认；瞬时错误仅重试一次。"),
            SkillDefinition("normalization_skill", "1.3", TaskType.NORMALIZE, "数据标准化", "对单列执行受控类型与文本标准化。", NormalizeTaskSpec, "normalize_column", ("sheets", "columns", "normalization_candidates"), ("只处理一个工作表中的一列。",), (*common, "output_column_verified"), "覆盖原列属于高风险，失败不提交。"),
            SkillDefinition("aggregation_skill", "1.0", TaskType.AGGREGATE, "聚合汇总", "结构化筛选、分组及统计。", AggregateTaskSpec, "aggregate", ("sheets", "dimension_candidates", "metric_candidates"), ("仅支持 sum/count/mean/min/max。",), (*common, "aggregate_totals_verified"), "校验失败不提交。"),
            SkillDefinition("sort_rank_skill", "1.0", TaskType.SORT_RANK, "排序排名", "稳定排序、分组排名和 Top N。", SortRankTaskSpec, "sort_rank", ("sheets", "sortable_columns", "numeric_rank_candidates"), ("结果写入新工作表。",), (*common, "row_count_verified"), "校验失败不提交。"),
            SkillDefinition("cleaning_skill", "1.0", TaskType.CLEAN, "数据清洗", "精确去重、标记重复、结构化筛选或受控列保留。", CleanTaskSpec, "clean", ("sheets", "columns", "duplicate_key_candidates"), ("不执行模糊去重或自动推断业务重复；列保留只接受明确列引用。",), (*common, "cleaning_counts_verified"), "覆盖删除其他列属于高风险，必须额外确认并逐列核验。"),
            SkillDefinition("formula_skill", "1.1", TaskType.FORMULA, "公式处理", "诊断、保守修复、显式设置、复制与重算公式。", FormulaTaskSpec, "formula", ("sheets", "formula_candidates"), ("不发明业务公式；业务正确性仍需人工确认。",), (*common, "formula_structure_verified"), "候选不唯一时跳过并报告。"),
            SkillDefinition("conversion_skill", "1.1", TaskType.CONVERSION, "文件转换", "CSV 与 XLSX 的受控转换和指定工作表导出。", ConversionTaskSpec, "conversion", ("file_name", "sheets"), ("CSV 转换完成前不能调用其他 Skill。", "XLSX 导出 CSV 必须明确工作表。"), (*common, "conversion_output_verified"), "编码或分隔符不确定时要求用户确认。", composable=False),
            SkillDefinition("transform_skill", "2.5", TaskType.TRANSFORM, "可审计数据准备", "执行数值填补、显式值映射、尺寸解析或受控派生指标。", TransformTaskSpec, "transform_column", ("sheets", "columns", "transform_candidates"), ("错误值默认不会当作缺失；派生指标必须选择注册公式；异常处理仅接受用户明确授权的受控策略。",), (*common, "transform_counts_reconciled", "output_columns_verified"), "可确定行正常交付；异常行默认留空并报告，也可按用户确认的分类规则或Excel行号修正规则继续同一计划。"),
        ]

    def get(self, task_type: TaskType) -> SkillDefinition:
        try:
            return self._skills[task_type]
        except KeyError as exc:
            raise KeyError(f"unknown skill for task type: {task_type}") from exc

    def get_by_id(self, skill_id: str) -> SkillDefinition:
        for skill in self._skills.values():
            if skill.id == skill_id:
                return skill
        raise KeyError(f"unknown skill: {skill_id}")

    def manifest(self) -> list[dict[str, Any]]:
        return [skill.manifest() for skill in self._skills.values()]

    def select(self, envelope: TaskEnvelope) -> list[SkillSelection]:
        specs = envelope.spec.operations if isinstance(envelope.spec, CompositeTaskSpec) else [envelope.spec]
        return [
            SkillSelection(
                skill_id=self.get(spec.task_type).id,
                reason=f"任务类型 {spec.task_type.value} 由该 Skill 的能力契约处理。",
                confidence=spec.confidence,
                boundary_notes=list(self.get(spec.task_type).boundary_notes),
            )
            for spec in specs
        ]

    def validate_composition(self, envelope: TaskEnvelope) -> list[str]:
        if not isinstance(envelope.spec, CompositeTaskSpec):
            return []
        errors: list[str] = []
        for spec in envelope.spec.operations:
            skill = self.get(spec.task_type)
            if not skill.composable and not (isinstance(spec, ConversionTaskSpec) and spec.direction.value == "xlsx_to_csv" and spec is envelope.spec.operations[-1]):
                errors.append(f"{skill.label}不能放入复合任务，请单独执行。")
        return errors


UNSUPPORTED_RULES: tuple[tuple[tuple[str, ...], str, list[str]], ...] = (
    ((r"vba", r"(?:运行|执行|编写|生成|调用|使用)?(?:javascript)?宏", r"macro"), "v1.0 不执行或生成 VBA/宏。", ["可以先使用公式、清洗或聚合 Skill 完成可确定性验证的部分。"]),
    ((r"powerquery", r"dax"), "v1.0 不操作 Power Query 或 DAX。", ["可将数据保存为普通工作表后再执行清洗或聚合。"]),
    ((r"模糊匹配", r"近似(?:拼写)?匹配", r"语义相似度", r"语义匹配", r"语义关联"), "v1.0 只支持精确键匹配。", ["可以先标准化关联键，再执行精确匹配。"]),
    ((r"(?:两个不同|两个|多个|多|不同|跨)工作簿", r"跨文件", r"多个文件夹.*工作簿"), "v1.0 不执行多工作簿关联。", ["请先将数据整理到同一 XLSX 的不同工作表。"]),
    ((r"python代码", r"执行代码", r"任意代码", r"(?:运行|执行|编写|使用).*?(?:python|powershell|javascript).*(?:脚本|代码)"), "v1.0 不开放任意代码执行。", ["请描述目标，Agent 会选择受控 Skill。"]),
    ((r"(?:联网|上网).*(?:抓取|采集|下载|同步)", r"(?:打开|访问)网页.*(?:采集|抓取)", r"(?:下载|获取).*远程文件"), "当前不联网抓取或下载外部数据。", ["请先将数据保存到本地工作簿。"]),
    ((r"(?:修改|编辑|打开).*(?:密码保护|加密).*工作簿",), "当前不修改受密码保护的工作簿。", ["请先提供可编辑的工作簿。"]),
    ((r"(?:运行|执行|使用|调用).*(?:插件|加载项)",), "当前不运行 Excel 插件或加载项。", ["请描述所需的表格结果，改用现有功能处理。"]),
    ((r"(?:连接|同步).*数据库",), "当前不连接外部数据库。", ["请先将所需数据导出到当前工作簿。"]),
    ((r"(?:修改|编辑|改写).*外部链接.*文件", r"(?:修改|编辑).*?(?:其他用户|另一个用户).*在线表格"), "当前不修改外部文件或在线表格。", ["请上传有权处理的本地工作簿。"]),
    ((r"ocr", r"(?:照片|图片|图像).*识别"), "当前不识别图片中的表格内容。", ["请先提供可读取的 CSV 或 XLSX 文件。"]),
)


def capability_refusal(request: str) -> CapabilityRefusal | None:
    for clause in re.split(r"[，,。；;\n]", request.casefold()):
        compact = re.sub(r"\s+", "", clause)
        for patterns, rule, alternatives in UNSUPPORTED_RULES:
            for pattern in patterns:
                match = re.search(pattern, compact)
                if not match:
                    continue
                prefix = compact[:match.start()]
                if re.search(r"(?:不|不要|无需|不用|别|禁止|避免)(?:再|去|使用|用)?$", prefix):
                    continue
                return CapabilityRefusal(
                    reason=rule,
                    boundary_rule="Only registered, deterministic spreadsheet skills may be planned.",
                    alternatives=alternatives,
                    suggested_request=alternatives[0] if alternatives else None,
                )
    return None
