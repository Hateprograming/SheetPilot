from __future__ import annotations

import json
import logging
import re
from typing import Any, Protocol, TypeVar

from pydantic import BaseModel, ValidationError

from .config import Settings
from .context import build_context, guess_task_type
from .errors import LLMError
from .schemas import ExecutionPlan, LookupTaskSpec, TaskEnvelope, WorkbookProfile
from .skills import SkillRegistry
from .p0 import authorized_model_rows


logger = logging.getLogger(__name__)


T = TypeVar("T", bound=BaseModel)


class LLMClient(Protocol):
    def parse_lookup(self, request: str, profile: WorkbookProfile) -> LookupTaskSpec: ...
    def parse_task(
        self,
        request: str,
        profile: WorkbookProfile,
        *,
        previous_envelope: TaskEnvelope | None = None,
        validation_errors: list[str] | None = None,
        conversation: list[dict[str, Any]] | None = None,
        capability_manifest: list[dict[str, Any]] | None = None,
        workbook_context: dict[str, Any] | None = None,
    ) -> TaskEnvelope: ...
    def explain_plan(self, plan: ExecutionPlan) -> str: ...


def _load_json(text: str):
    content = text.strip()
    if "</think>" in content:
        content = content.split("</think>", 1)[1].strip()
    if content.startswith("```"):
        lines = content.splitlines()[1:]
        if lines and lines[-1].strip() == "```":
            lines.pop()
        content = "\n".join(lines)
    return json.loads(content)


def _safe_validation_codes(errors: list[str] | None) -> list[str]:
    """Reduce local validation details to non-sensitive correction codes."""
    codes: list[str] = []
    for error in errors or []:
        folded = error.casefold()
        if any(token in folded for token in ("缺少", "missing", "不能为空", "至少需要")):
            codes.append("MISSING_FIELD")
        elif any(token in folded for token in ("歧义", "ambigu", "无法确定", "不唯一")):
            codes.append("AMBIGUOUS_RULE")
        elif any(token in folded for token in ("列", "column")) and any(
            token in folded for token in ("不存在", "not found", "不唯一")
        ):
            codes.append("COLUMN_NOT_FOUND")
        elif any(token in folded for token in ("未授权", "用户要求中明确", "原始要求")):
            codes.append("UNAUTHORIZED_RULE")
        else:
            codes.append("PLAN_VALIDATION_FAILED")
    return list(dict.fromkeys(codes))


class OpenAICompatibleClient:
    def __init__(self, settings: Settings):
        try:
            import httpx
        except ImportError as exc:
            raise LLMError("缺少 httpx 依赖，请先安装项目依赖。") from exc
        if not settings.llm_configured:
            raise LLMError("尚未配置模型服务。请在 .env 中填写API地址、模型和密钥。")
        self.settings = settings
        self.client = httpx.Client(
            base_url=settings.base_url,
            timeout=90,
            headers={"Authorization": f"Bearer {settings.api_key}", "Content-Type": "application/json"},
        )

    def close(self) -> None:
        self.client.close()

    def _complete_json(self, system: str, user: str, model: type[T]) -> T:
        messages = [{"role": "system", "content": system}, {"role": "user", "content": user}]
        last_error = ""
        for attempt in range(2):
            payload: dict[str, Any] = {
                "model": self.settings.model,
                "messages": messages,
                "temperature": 0.1,
                "response_format": {"type": "json_object"},
            }
            if self.settings.llm_provider.casefold() == "deepseek":
                # DeepSeek enables thinking by default.  Planning needs only the
                # final JSON object, so disable reasoning explicitly to avoid
                # extra latency and reasoning-token cost.
                payload["thinking"] = {"type": "disabled"}
            else:
                # Kept for the previous Qwen-compatible endpoint.
                payload["enable_thinking"] = False
            try:
                response = self.client.post("/chat/completions", json=payload)
                response.raise_for_status()
                content = response.json()["choices"][0]["message"]["content"]
                return model.model_validate(_load_json(content))
            except Exception as exc:
                last_error = f"{type(exc).__name__}: {exc}"
                if attempt == 0:
                    messages.extend(
                        [
                            {"role": "assistant", "content": content if 'content' in locals() else "{}"},
                            {
                                "role": "user",
                                "content": "上一个响应未通过JSON Schema验证。只返回修复后的JSON，不要解释。",
                            },
                        ]
                    )
        safe_error = re.sub(r"(?i)(bearer|api[_ -]?key)\s*[:=]?\s*\S+", r"\1=[REDACTED]", last_error)
        raise LLMError("模型两次未能返回有效的结构化结果。", detail={"reason": safe_error[:500]})

    def parse_lookup(self, request: str, profile: WorkbookProfile) -> LookupTaskSpec:
        safe_profile = build_context(profile, request)
        system = """你是SheetPilot v0.1的意图解析器。你只能解析同一XLSX内、单个精确关联键、1到3个返回列的跨Sheet静态值匹配任务。不要执行任务，不要编造Sheet或列。无法确定的字段使用空字符串，并列入missing_fields；多种可能写入ambiguities。只返回符合LookupTaskSpec的JSON。默认normalization_mode为exact_text，unmatched_policy为blank_and_report，duplicate_policy为block，constraints固定包含preserve_row_order、append_only、static_values。"""
        user = json.dumps(
            {"request": request, "workbook_profile": safe_profile}, ensure_ascii=False, default=str
        )
        return self._complete_json(system, user, LookupTaskSpec)

    def parse_task(
        self,
        request: str,
        profile: WorkbookProfile,
        *,
        previous_envelope: TaskEnvelope | None = None,
        validation_errors: list[str] | None = None,
        conversation: list[dict[str, Any]] | None = None,
        capability_manifest: list[dict[str, Any]] | None = None,
        workbook_context: dict[str, Any] | None = None,
    ) -> TaskEnvelope:
        system = """你是SheetPilot v2.0的 Skill 路由器和意图解析器。只能从 capability_manifest 中选择能力，并严格输出对应输入 Schema。不得发明 Skill、任务类型、工具或参数。
用户明确写出的工作表名、来源列名、输出工作表名、输出列名、映射键和值必须逐字保留，不得换成默认名、近义词或更长的名字。只有用户没有命名时才能使用Schema默认值。用户提到的工作簿列存在时必须选择该列，不得选择名称相近的其他列。
“新建汇总表”表示输出Sheet名是“汇总”，不得改成“汇总表”；“新列输出填补值和缺失标记”表示output_column="填补值"、indicator_column="缺失标记"，不得添加来源列名前缀。aggregate本身支持filters；“先筛选状态等于完成，再按部门分组统计”必须是一个aggregate及其filters，不能拆成clean临时表再aggregate。count已经表示非空计数，不得另外添加not_blank筛选。
lookup中，用户没有明确说来源键重复时取第一条或最后一条，duplicate_policy必须为block。map_values只能包含用户明确给出的映射，禁止补充同义类别。例如“把分类中的值‘有’映射为1、值‘无’映射为0”必须输出mapping={"有":1,"无":0}，不得输出“有映射/无映射”，也不得增加“是/否”。
如果用户原话内部冲突，例如排序列和排名列指向不同但关系不清楚，必须把冲突写入ambiguities，不得自行假设。previous_task是上一轮完整草稿；新消息优先于旧值，未被新消息修改的字段必须保留。
你可以选择以下任务：
lookup：同一XLSX内两个Sheet之间的单列精确左连接；
normalize：泛化“标准化”默认按顺序 nfkc、trim，保留前导零和内部空格；明确指定操作时只执行对应操作。对一个Sheet的一列执行trim、nfkc、lower、upper、to_text、to_number、date_format、fill_blank或strip_suffix；column优先填写工作簿上下文中的准确display_header（也兼容列字母）；strip_suffix只能移除用户明确给出的strip_suffixes，例如年龄“60岁”使用operations=[strip_suffix,to_number]和strip_suffixes=["岁"]；用户明确说“原列全部改为/去掉后改为”时overwrite=true；
aggregate：按1到3列分组并执行sum、count、mean、min或max，写入新Sheet；
sort_rank：对完整矩形数据区域排序、排名或Top N，写入新Sheet；
clean：按指定列精确识别重复、标记/移除重复，按结构化条件筛选，或使用 keep_columns 只保留用户明确指定的列。keep_columns 必须使用 ColumnRef；用户明确要求删除当前表中的其他列时设置 overwrite=true，这属于高风险覆盖；不得自行保留用户未提及的编号、姓名或其他标识列。非覆盖模式写入新Sheet。
formula：读取/诊断公式、填补断层、替换异常、设置公式、复制公式或重算缓存；mode 可选择 diagnose、fill_gaps、replace_outliers、set_formula、copy_formula、recalculate；
conversion：CSV 转 XLSX 必须单独执行；XLSX 导出 CSV 可以作为 composite 最后一步，使用此前步骤的最终数据；单表可直接转换，多表必须由用户选择来源；
transform：数值填补(impute_numeric)、显式分类映射(map_values)、尺寸解析(parse_dimensions)或受控派生指标(derive_metric)。重复表头必须用ColumnRef的列字母或序号消歧；空表头有数据的列也用列字母选择。除非用户明确要求覆盖，默认overwrite=false；
impute_numeric只把空单元格和用户明确给出的missing_tokens当作缺失；非空格式错误默认不填补。只有用户明确说“错误值也填补/错误值也用均值”等时，才可把对应规则的invalid_policy设为impute，并设置invalid_policy_confirmed=true；用户明确要求错误值留空时也设置invalid_policy_confirmed=true。此时统计量仍只由原始有效数值计算，错误值不进入统计量。若用户是在上一版本已经存在的填补结果列中继续修正None，应保留原始来源列并设置overwrite_existing_outputs=true，重算既有结果列和指示列；不要把overwrite设为true而误覆盖原始来源列；
parse_dimensions默认对不确定记录留空。若用户明确说单个或两个数必须带文字说明，否则视为错误，必须设置require_context_for_partial=true。用户可以明确授权dimension_rule.resolution_policies：multiple_candidates可设first_candidate或last_candidate，unit_conflict可设convert_each_to_mm，invalid_format/no_numeric_measurement/blank可设zero；用户明确要求保持留空时设leave_blank。用户对具体原始值给出精确修正时，写入value_corrections，例如“21.*7*16.7改为21*7*16.7”；用户按Excel行号给出修正时写入row_corrections，例如“第13行改为21mm7mm16.7mm”应写为{13:"21mm7mm16.7mm"}。recent_conversation中的“安全的结构化澄清上下文”只包含当前目标列，可用于将行号回答与异常记录对应。不得自行添加这些异常处理策略；
当previous_task是尺寸解析、数值填补或派生指标，而用户说“这些错误值/显示为None的错误值/这类异常”时，应修订原Transform规则而不是新建无关任务，并返回合并后的完整TaskEnvelope。map_values必须由用户给出显式mapping，未知值不猜测；仅当用户明确说“没有记录/空值算作某值”时设置blank_value，仅当用户明确说“其他类别均为某值”时设置unmapped_value。map_values的status_column默认必须为空，mapped、mapped_default等状态只保留在内部计数中；只有用户明确要求在工作簿中输出映射状态列时才填写status_column。多个同Sheet映射必须合并进一个map_values步骤，多个数值填补必须合并进一个impute_numeric步骤。派生体积只允许ellipsoid_pi_over_6、half_product、rectangular_product、rectangular_or_diameter_cubed；当用户明确要求三维尺寸用L×W×H、单一直径用D³时选择rectangular_or_diameter_cubed，并依次提供长、宽、高、最大径四个sources。用户未明确公式时formula_template必须为null并列入missing_fields；派生输入无效或缺失时默认留空，只有用户明确要求统一填0时才能把derive_rule.resolution_policies中对应的invalid_input或input_missing设为zero，明确选择保持留空时设为leave_blank；
set_formula 可来自用户明确给出的 Excel 公式或明确的列名算式（如“在总额列填入数量×单价”），应选择 formula 任务并填写目标列；列名算式由本地安全转换，不要求用户手写 Excel 公式，不得自行编造业务规则；copy_formula 必须填写 source_row；recalculate 可对工作簿现有公式重算；
公式任务默认保持 recalculate=true；只有用户明确要求仅检查公式文本时才可设为 false；
fill_gaps、replace_outliers 和 diagnose 不要求用户提供具体公式文本：确定性工具会读取已有相邻公式并判断候选。不得因为“用户未提供公式文本”写入 missing_fields 或 ambiguities；候选不唯一是允许的运行结果，工具会跳过该单元格并报告，不需要事前澄清。只要工作表、列和范围已经明确，这三种模式的 missing_fields 与 ambiguities 应为空；
composite：当用户明确提出多个处理要求时，用operations按依赖顺序组合2到5个上述原子任务，禁止嵌套composite；同一Sheet的映射规则和填补规则必须分别合并，避免超过五步。每个TransformTaskSpec只能配置其operation对应的一种payload：impute_numeric只填impute_rules，map_values只填map_rules，parse_dimensions只填dimension_rule，derive_metric只填derive_rule，绝不能把多类payload塞进一个Transform步骤。字段已明确时，每个原子任务和composite的confidence必须不低于0.6；
不要执行任务，不要编造Sheet或列，不要进行模糊匹配、联合键、一对多展开、宏或任意代码。公式任务不得编造具体公式，候选公式由确定性工具推断。无法确定的关键信息写入对应spec.missing_fields；多种可能写入spec.ambiguities。如果提供了previous_task，用户的新消息是在修订或补充该任务；必须合并上下文并返回完整的新TaskEnvelope，不能只返回变化字段。所有输出必须符合提供的JSON Schema，只返回JSON。"""
        # Rebuild from the verified profile. Never trust caller-provided context,
        # which could otherwise smuggle cell values past the egress allowlist.
        hint = guess_task_type(request)
        if hint is None and previous_envelope is not None:
            hint = previous_envelope.spec.task_type
        skill_context = build_context(profile, request)
        if hint and hint.value != "composite":
            try:
                skill_context = SkillRegistry().get(hint).build_context(profile, request)
            except KeyError:
                pass
        user = json.dumps(
            {
                "request": request,
                "workbook_context": skill_context,
                "previous_task_type": previous_envelope.spec.task_type.value if previous_envelope else None,
                "previous_task": previous_envelope.model_dump(mode="json") if previous_envelope else None,
                "validation_error_codes": _safe_validation_codes(validation_errors),
                "recent_conversation": [m for m in (conversation or []) if m.get("role") == "user"][-8:],
                "authorized_exception_rows": authorized_model_rows.get(),
                "capability_manifest": capability_manifest or [],
                "output_schema": TaskEnvelope.model_json_schema(),
            },
            ensure_ascii=False,
            default=str,
        )
        return self._complete_json(system, user, TaskEnvelope)

    def explain_plan(self, plan: ExecutionPlan) -> str:
        return (
            f"我会以 {plan.key_mapping} 为关联键，从“{plan.source_sheet}”读取所需字段，"
            f"并在“{plan.main_sheet}”末尾新增 {', '.join(plan.expected_new_columns)}。"
            "原有行序和内容保持不变；未匹配记录将单独列出。"
        )


class FakeLLMClient:
    def __init__(self, spec: LookupTaskSpec | TaskEnvelope):
        self.spec = spec

    def parse_lookup(self, request: str, profile: WorkbookProfile) -> LookupTaskSpec:
        if isinstance(self.spec, TaskEnvelope):
            if not isinstance(self.spec.spec, LookupTaskSpec):
                raise TypeError("fake task is not a lookup")
            return self.spec.spec.model_copy(deep=True)
        return self.spec.model_copy(deep=True)

    def parse_task(
        self,
        request: str,
        profile: WorkbookProfile,
        *,
        previous_envelope: TaskEnvelope | None = None,
        validation_errors: list[str] | None = None,
        conversation: list[dict[str, Any]] | None = None,
        capability_manifest: list[dict[str, Any]] | None = None,
        workbook_context: dict[str, Any] | None = None,
    ) -> TaskEnvelope:
        if isinstance(self.spec, TaskEnvelope):
            return self.spec.model_copy(deep=True)
        return TaskEnvelope(goal=request, spec=self.spec.model_copy(deep=True))

    def explain_plan(self, plan: ExecutionPlan) -> str:
        return plan.summary


# Compatibility alias for persisted imports and older tests.  The implementation
# is provider-neutral even though the historical class name mentions Qwen.
QwenCompatibleClient = OpenAICompatibleClient
