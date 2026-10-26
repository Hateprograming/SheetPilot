from __future__ import annotations

from enum import StrEnum
from typing import Annotated, Any, Literal, Union
from uuid import uuid4

from pydantic import BaseModel, ConfigDict, Field, model_validator


class StrictModel(BaseModel):
    model_config = ConfigDict(extra="forbid")


class NormalizationMode(StrEnum):
    EXACT_TEXT = "exact_text"
    NUMERIC_EQUIVALENCE = "numeric_equivalence"


class DuplicatePolicy(StrEnum):
    BLOCK = "block"
    FIRST = "first"
    LAST = "last"


class TaskType(StrEnum):
    LOOKUP = "lookup"
    NORMALIZE = "normalize"
    AGGREGATE = "aggregate"
    SORT_RANK = "sort_rank"
    CLEAN = "clean"
    FORMULA = "formula"
    CONVERSION = "conversion"
    TRANSFORM = "transform"
    COMPOSITE = "composite"


class ConversationState(StrEnum):
    EMPTY = "empty"
    WORKBOOK_READY = "workbook_ready"
    ANALYZING_REQUEST = "analyzing_request"
    NEEDS_CLARIFICATION = "needs_clarification"
    PLAN_READY = "plan_ready"
    EXECUTING = "executing"
    RESULT_READY = "result_ready"
    PAUSED = "paused"


class ConversationRole(StrEnum):
    USER = "user"
    ASSISTANT = "assistant"
    SYSTEM = "system"


class MessageKind(StrEnum):
    TEXT = "text"
    WORKBOOK = "workbook"
    CLARIFICATION = "clarification"
    PLAN = "plan"
    PROGRESS = "progress"
    RESULT = "result"
    ERROR = "error"
    RECOVERY = "recovery"
    REFUSAL = "refusal"


class RiskLevel(StrEnum):
    LOW = "low"
    MEDIUM = "medium"
    HIGH = "high"


class NormalizeOperation(StrEnum):
    TRIM = "trim"
    NFKC = "nfkc"
    LOWER = "lower"
    UPPER = "upper"
    TO_TEXT = "to_text"
    TO_NUMBER = "to_number"
    DATE_FORMAT = "date_format"
    FILL_BLANK = "fill_blank"
    STRIP_SUFFIX = "strip_suffix"


class AggregateFunction(StrEnum):
    SUM = "sum"
    COUNT = "count"
    MEAN = "mean"
    MIN = "min"
    MAX = "max"


class SortDirection(StrEnum):
    ASC = "asc"
    DESC = "desc"


class RankMethod(StrEnum):
    DENSE = "dense"
    MIN = "min"
    MAX = "max"


class CleanOperation(StrEnum):
    MARK_DUPLICATES = "mark_duplicates"
    REMOVE_DUPLICATES = "remove_duplicates"
    FILTER_ROWS = "filter_rows"
    KEEP_COLUMNS = "keep_columns"


class FilterOperator(StrEnum):
    EQ = "eq"
    NE = "ne"
    GT = "gt"
    GTE = "gte"
    LT = "lt"
    LTE = "lte"
    CONTAINS = "contains"
    IS_BLANK = "is_blank"
    NOT_BLANK = "not_blank"


class FormulaMode(StrEnum):
    DIAGNOSE = "diagnose"
    FILL_GAPS = "fill_gaps"
    REPLACE_OUTLIERS = "replace_outliers"
    SET_FORMULA = "set_formula"
    COPY_FORMULA = "copy_formula"
    RECALCULATE = "recalculate"


class ConversionDirection(StrEnum):
    CSV_TO_XLSX = "csv_to_xlsx"
    XLSX_TO_CSV = "xlsx_to_csv"


class TransformOperation(StrEnum):
    IMPUTE_NUMERIC = "impute_numeric"
    MAP_VALUES = "map_values"
    PARSE_DIMENSIONS = "parse_dimensions"
    DERIVE_METRIC = "derive_metric"


class ImputationMethod(StrEnum):
    MEAN = "mean"
    MEDIAN = "median"
    CONSTANT = "constant"


class InvalidNumericPolicy(StrEnum):
    LEAVE_BLANK = "leave_blank"
    IMPUTE = "impute"


class DimensionResolutionAction(StrEnum):
    LEAVE_BLANK = "leave_blank"
    FIRST_CANDIDATE = "first_candidate"
    LAST_CANDIDATE = "last_candidate"
    CONVERT_EACH_TO_MM = "convert_each_to_mm"
    ZERO = "zero"


class DerivedResolutionAction(StrEnum):
    LEAVE_BLANK = "leave_blank"
    ZERO = "zero"


class DerivedFormulaTemplate(StrEnum):
    ELLIPSOID_PI_OVER_6 = "ellipsoid_pi_over_6"
    HALF_PRODUCT = "half_product"
    RECTANGULAR_PRODUCT = "rectangular_product"
    RECTANGULAR_OR_DIAMETER_CUBED = "rectangular_or_diameter_cubed"


class ColumnProfile(StrictModel):
    name: str
    index: int
    inferred_type: str
    empty_count: int
    unique_count: int
    formula_count: int
    letter: str = ""
    display_header: str = ""
    header_occurrence: int = 1
    nonempty_count: int = 0
    active: bool = True


class ColumnRef(StrictModel):
    sheet: str
    column_index: int | None = Field(default=None, ge=1)
    column_letter: str = ""
    display_header: str = ""
    header_occurrence: int = Field(default=1, ge=1)

    @model_validator(mode="after")
    def require_locator(self) -> "ColumnRef":
        if not self.column_index and not self.column_letter and not self.display_header:
            raise ValueError("column reference needs an index, letter, or display header")
        return self


class SheetProfile(StrictModel):
    name: str
    state: str
    max_row: int
    max_column: int
    header_row: int | None
    header_confidence: float
    header_candidates: list[int] = Field(default_factory=list)
    columns: list[ColumnProfile] = Field(default_factory=list)
    preview_rows: list[dict[str, Any]] = Field(default_factory=list)
    merged_ranges: list[str] = Field(default_factory=list)
    hidden_rows: int = 0
    hidden_columns: int = 0


class WorkbookProfile(StrictModel):
    file_name: str
    size_bytes: int
    sha256: str
    sheets: list[SheetProfile]
    external_links: list[str] = Field(default_factory=list)
    risks: list[str] = Field(default_factory=list)


class LookupTaskSpec(StrictModel):
    task_type: Literal[TaskType.LOOKUP] = TaskType.LOOKUP
    main_sheet: str = ""
    source_sheet: str = ""
    main_key_column: str = ""
    source_key_column: str = ""
    return_columns: list[str] = Field(default_factory=list, max_length=3)
    output_column_names: list[str] = Field(default_factory=list, max_length=3)
    normalization_mode: NormalizationMode = NormalizationMode.EXACT_TEXT
    unmatched_policy: Literal["blank_and_report"] = "blank_and_report"
    duplicate_policy: DuplicatePolicy = DuplicatePolicy.BLOCK
    constraints: list[str] = Field(
        default_factory=lambda: ["preserve_row_order", "append_only", "static_values"]
    )
    missing_fields: list[str] = Field(default_factory=list)
    ambiguities: list[str] = Field(default_factory=list)
    confidence: float = Field(default=0.0, ge=0.0, le=1.0)

    @model_validator(mode="after")
    def align_output_columns(self) -> "LookupTaskSpec":
        if self.output_column_names and len(self.output_column_names) != len(self.return_columns):
            raise ValueError("output_column_names must align with return_columns")
        return self


class NormalizeTaskSpec(StrictModel):
    task_type: Literal[TaskType.NORMALIZE] = TaskType.NORMALIZE
    sheet: str = ""
    column: str = ""
    operations: list[NormalizeOperation] = Field(default_factory=list, max_length=4)
    output_column: str = ""
    date_format: str | None = None
    fill_value: str | int | float | None = None
    strip_suffixes: list[str] = Field(default_factory=list, max_length=10)
    overwrite: bool = False
    missing_fields: list[str] = Field(default_factory=list)
    ambiguities: list[str] = Field(default_factory=list)
    confidence: float = Field(default=0.0, ge=0.0, le=1.0)

    @model_validator(mode="after")
    def validate_operation_parameters(self) -> "NormalizeTaskSpec":
        if NormalizeOperation.STRIP_SUFFIX in self.operations and not self.strip_suffixes:
            raise ValueError("strip_suffix needs at least one explicit suffix")
        if any(not suffix for suffix in self.strip_suffixes):
            raise ValueError("strip suffixes cannot be empty")
        return self

class AggregateMetric(StrictModel):
    column: str
    function: AggregateFunction
    output_column: str = ""

    @model_validator(mode="after")
    def default_output_name(self) -> "AggregateMetric":
        if not self.output_column:
            self.output_column = f"{self.column}_{self.function.value}"
        return self


class RowFilter(StrictModel):
    column: str
    operator: FilterOperator
    value: Any = None


class AggregateTaskSpec(StrictModel):
    task_type: Literal[TaskType.AGGREGATE] = TaskType.AGGREGATE
    sheet: str = ""
    group_by: list[str] = Field(default_factory=list, max_length=3)
    metrics: list[AggregateMetric] = Field(default_factory=list, max_length=5)
    filters: list[RowFilter] = Field(default_factory=list, max_length=5)
    output_sheet: str = "SheetPilot_Summary"
    missing_fields: list[str] = Field(default_factory=list)
    ambiguities: list[str] = Field(default_factory=list)
    confidence: float = Field(default=0.0, ge=0.0, le=1.0)


class SortKey(StrictModel):
    column: str
    direction: SortDirection = SortDirection.ASC


class SortRankTaskSpec(StrictModel):
    task_type: Literal[TaskType.SORT_RANK] = TaskType.SORT_RANK
    sheet: str = ""
    sort_keys: list[SortKey] = Field(default_factory=list, max_length=3)
    rank_column: str | None = None
    rank_output_column: str = "排名"
    rank_direction: SortDirection = SortDirection.DESC
    rank_method: RankMethod = RankMethod.DENSE
    group_by: list[str] = Field(default_factory=list, max_length=2)
    top_n: int | None = Field(default=None, ge=1, le=10_000)
    output_sheet: str = "SheetPilot_Sorted"
    missing_fields: list[str] = Field(default_factory=list)
    ambiguities: list[str] = Field(default_factory=list)
    confidence: float = Field(default=0.0, ge=0.0, le=1.0)

class CleanTaskSpec(StrictModel):
    task_type: Literal[TaskType.CLEAN] = TaskType.CLEAN
    sheet: str = ""
    operation: CleanOperation = CleanOperation.MARK_DUPLICATES
    key_columns: list[str] = Field(default_factory=list, max_length=5)
    keep: Literal["first", "last"] = "first"
    filters: list[RowFilter] = Field(default_factory=list, max_length=5)
    keep_columns: list[ColumnRef] = Field(default_factory=list, max_length=200)
    overwrite: bool = False
    output_sheet: str = "SheetPilot_Cleaned"
    duplicate_flag_column: str = "SheetPilot_重复标记"
    missing_fields: list[str] = Field(default_factory=list)
    ambiguities: list[str] = Field(default_factory=list)
    confidence: float = Field(default=0.0, ge=0.0, le=1.0)


class FormulaTaskSpec(StrictModel):
    task_type: Literal[TaskType.FORMULA] = TaskType.FORMULA
    sheet: str = ""
    column: str = ""
    start_row: int | None = Field(default=None, ge=1)
    end_row: int | None = Field(default=None, ge=1)
    mode: FormulaMode = FormulaMode.DIAGNOSE
    formula: str | None = None
    source_row: int | None = Field(default=None, ge=1)
    overwrite_existing: bool = False
    recalculate: bool = True
    missing_fields: list[str] = Field(default_factory=list)
    ambiguities: list[str] = Field(default_factory=list)
    confidence: float = Field(default=0.0, ge=0.0, le=1.0)

    @model_validator(mode="after")
    def validate_row_range(self) -> "FormulaTaskSpec":
        if self.start_row and self.end_row and self.start_row > self.end_row:
            raise ValueError("start_row must not exceed end_row")
        return self


class ConversionTaskSpec(StrictModel):
    task_type: Literal[TaskType.CONVERSION] = TaskType.CONVERSION
    direction: ConversionDirection
    source_sheet: str = ""
    output_name: str = ""
    encoding: Literal["auto", "utf-8", "utf-8-sig", "gb18030"] = "auto"
    delimiter: Literal["auto", ",", "\t", ";", "|"] = "auto"
    preserve_text: bool = True
    infer_types: bool = False
    missing_fields: list[str] = Field(default_factory=list)
    ambiguities: list[str] = Field(default_factory=list)
    confidence: float = Field(default=0.0, ge=0.0, le=1.0)


class ImputeNumericRule(StrictModel):
    source: ColumnRef
    method: ImputationMethod = ImputationMethod.MEAN
    constant_value: float | int | None = None
    missing_tokens: list[str] = Field(default_factory=list, max_length=20)
    output_column: str = ""
    indicator_column: str = ""
    overwrite: bool = False
    overwrite_existing_outputs: bool = False
    invalid_policy: InvalidNumericPolicy = InvalidNumericPolicy.LEAVE_BLANK
    invalid_policy_confirmed: bool = False

    @model_validator(mode="after")
    def validate_write_mode(self) -> "ImputeNumericRule":
        if self.overwrite and self.overwrite_existing_outputs:
            raise ValueError("cannot overwrite the source and existing outputs in one imputation rule")
        return self

    @model_validator(mode="after")
    def validate_constant(self) -> "ImputeNumericRule":
        if self.method == ImputationMethod.CONSTANT and self.constant_value is None:
            raise ValueError("constant imputation needs constant_value")
        return self


class ValueMapRule(StrictModel):
    source: ColumnRef
    mapping: dict[str, bool | int | float | str] = Field(default_factory=dict)
    output_column: str = ""
    status_column: str = ""
    missing_tokens: list[str] = Field(default_factory=list, max_length=20)
    casefold: bool = False
    overwrite: bool = False
    blank_value: bool | int | float | str | None = None
    unmapped_value: bool | int | float | str | None = None


class ParseDimensionsRule(StrictModel):
    source: ColumnRef
    default_unit: Literal["mm", "cm"] = "mm"
    length_column: str = ""
    width_column: str = ""
    height_column: str = ""
    max_dimension_column: str = ""
    status_column: str = ""
    overwrite_existing_outputs: bool = False
    require_context_for_partial: bool = False
    resolution_policies: dict[str, DimensionResolutionAction] = Field(
        default_factory=dict, max_length=8
    )
    value_corrections: dict[str, str] = Field(default_factory=dict, max_length=50)
    row_corrections: dict[int, str] = Field(default_factory=dict, max_length=50)

    @model_validator(mode="after")
    def validate_resolution_policies(self) -> "ParseDimensionsRule":
        allowed = {
            "multiple_candidates": {
                DimensionResolutionAction.LEAVE_BLANK,
                DimensionResolutionAction.FIRST_CANDIDATE,
                DimensionResolutionAction.LAST_CANDIDATE,
                DimensionResolutionAction.ZERO,
            },
            "unit_conflict": {
                DimensionResolutionAction.LEAVE_BLANK,
                DimensionResolutionAction.CONVERT_EACH_TO_MM,
                DimensionResolutionAction.ZERO,
            },
            "invalid_format": {
                DimensionResolutionAction.LEAVE_BLANK,
                DimensionResolutionAction.ZERO,
            },
            "no_numeric_measurement": {
                DimensionResolutionAction.LEAVE_BLANK,
                DimensionResolutionAction.ZERO,
            },
            "blank": {
                DimensionResolutionAction.LEAVE_BLANK,
                DimensionResolutionAction.ZERO,
            },
        }
        for status, action in self.resolution_policies.items():
            if status not in allowed:
                raise ValueError(f"unsupported dimension status policy: {status}")
            if action not in allowed[status]:
                raise ValueError(f"{action.value} cannot resolve {status}")
        if any(not key.strip() or not value.strip() for key, value in self.value_corrections.items()):
            raise ValueError("dimension value corrections need non-empty source and replacement text")
        if any(row < 1 or not value.strip() for row, value in self.row_corrections.items()):
            raise ValueError("dimension row corrections need positive Excel rows and replacement text")
        return self


class DeriveMetricRule(StrictModel):
    sources: list[ColumnRef] = Field(min_length=3, max_length=4)
    formula_template: DerivedFormulaTemplate | None = None
    output_column: str = ""
    status_column: str = ""
    output_unit: Literal["cm3"] = "cm3"
    overwrite_existing_outputs: bool = False
    resolution_policies: dict[str, DerivedResolutionAction] = Field(
        default_factory=dict, max_length=2
    )

    @model_validator(mode="after")
    def validate_resolution_policies(self) -> "DeriveMetricRule":
        allowed = {"invalid_input", "input_missing"}
        if any(status not in allowed for status in self.resolution_policies):
            raise ValueError("unsupported derived metric resolution status")
        expected_sources = (
            4
            if self.formula_template == DerivedFormulaTemplate.RECTANGULAR_OR_DIAMETER_CUBED
            else 3
        )
        if len(self.sources) != expected_sources:
            raise ValueError(
                f"{self.formula_template or 'derived formula'} needs {expected_sources} source columns"
            )
        return self


class TransformTaskSpec(StrictModel):
    task_type: Literal[TaskType.TRANSFORM] = TaskType.TRANSFORM
    operation: TransformOperation
    sheet: str = ""
    impute_rules: list[ImputeNumericRule] = Field(default_factory=list, max_length=20)
    map_rules: list[ValueMapRule] = Field(default_factory=list, max_length=20)
    dimension_rule: ParseDimensionsRule | None = None
    derive_rule: DeriveMetricRule | None = None
    missing_fields: list[str] = Field(default_factory=list)
    ambiguities: list[str] = Field(default_factory=list)
    confidence: float = Field(default=0.0, ge=0.0, le=1.0)

    @model_validator(mode="after")
    def validate_operation_payload(self) -> "TransformTaskSpec":
        payloads = {
            TransformOperation.IMPUTE_NUMERIC: bool(self.impute_rules),
            TransformOperation.MAP_VALUES: bool(self.map_rules),
            TransformOperation.PARSE_DIMENSIONS: self.dimension_rule is not None,
            TransformOperation.DERIVE_METRIC: self.derive_rule is not None,
        }
        if not payloads[self.operation]:
            raise ValueError(f"{self.operation.value} needs its matching rule payload")
        unexpected = [
            operation.value
            for operation, present in payloads.items()
            if operation != self.operation and present
        ]
        if unexpected:
            raise ValueError(
                f"{self.operation.value} cannot include payloads for: {', '.join(unexpected)}"
            )
        if self.operation == TransformOperation.DERIVE_METRIC:
            assert self.derive_rule is not None
            if self.derive_rule.formula_template is None:
                requirement = "必须明确选择派生公式模板"
                if requirement not in self.missing_fields:
                    self.missing_fields.append(requirement)
        return self


AtomicTaskSpec = Annotated[
    Union[
        LookupTaskSpec,
        NormalizeTaskSpec,
        AggregateTaskSpec,
        SortRankTaskSpec,
        CleanTaskSpec,
        FormulaTaskSpec,
        ConversionTaskSpec,
        TransformTaskSpec,
    ],
    Field(discriminator="task_type"),
]


class CompositeTaskSpec(StrictModel):
    task_type: Literal[TaskType.COMPOSITE] = TaskType.COMPOSITE
    operations: list[AtomicTaskSpec] = Field(min_length=2, max_length=5)
    missing_fields: list[str] = Field(default_factory=list)
    ambiguities: list[str] = Field(default_factory=list)
    confidence: float = Field(default=0.0, ge=0.0, le=1.0)


TaskSpec = Annotated[
    Union[
        LookupTaskSpec,
        NormalizeTaskSpec,
        AggregateTaskSpec,
        SortRankTaskSpec,
        CleanTaskSpec,
        FormulaTaskSpec,
        ConversionTaskSpec,
        TransformTaskSpec,
        CompositeTaskSpec,
    ],
    Field(discriminator="task_type"),
]


class TaskEnvelope(StrictModel):
    goal: str
    spec: TaskSpec
    constraints: list[str] = Field(default_factory=list)
    assumptions: list[str] = Field(default_factory=list)


class ExpectedChanges(StrictModel):
    modified_sheets: list[str] = Field(default_factory=list)
    created_sheets: list[str] = Field(default_factory=list)
    new_columns: list[str] = Field(default_factory=list)
    estimated_rows: int = 0
    destructive: bool = False
    allowed_ranges: list[str] = Field(default_factory=list)
    allowed_created_sheets: list[str] = Field(default_factory=list)
    allowed_artifacts: list[str] = Field(default_factory=list)
    allow_formula_recalculation: bool = False
    summary_counts: dict[str, int | float | str] = Field(default_factory=dict)
    review_rows: list[dict[str, Any]] = Field(default_factory=list)
    review_totals: dict[str, int] = Field(default_factory=dict)
    review_truncated: dict[str, bool] = Field(default_factory=dict)


class ToolStep(StrictModel):
    id: str
    order: int
    tool_name: str
    arguments: dict[str, Any]
    description: str
    risk_level: RiskLevel
    expected_changes: ExpectedChanges
    depends_on: list[str] = Field(default_factory=list)
    preconditions: list[str] = Field(default_factory=list)
    verification_rules: list[str] = Field(default_factory=list)
    retry_policy: Literal["transient_once", "never"] = "transient_once"
    required: bool = True
    status: Literal[
        "pending", "running", "completed", "retrying", "paused", "failed", "rolled_back"
    ] = "pending"


class TaskPlan(StrictModel):
    plan_id: str = Field(default_factory=lambda: uuid4().hex)
    revision: int = Field(default=1, ge=1)
    task_id: str
    task_type: TaskType
    input_version: int
    input_sha256: str
    base_version: int | None = None
    base_sha256: str = ""
    summary: str
    steps: list[ToolStep] = Field(min_length=1, max_length=5)
    invariants: list[str] = Field(
        default_factory=lambda: ["original_file_unchanged", "planned_changes_only"]
    )
    risk_level: RiskLevel
    requires_approval: bool = True
    skill_ids: list[str] = Field(default_factory=list)
    skill_versions: dict[str, str] = Field(default_factory=dict)
    plan_assumptions: list[str] = Field(default_factory=list)
    confirmation_requirements: list[str] = Field(default_factory=list)

    @model_validator(mode="after")
    def populate_legacy_base_fields(self) -> "TaskPlan":
        if self.base_version is None:
            self.base_version = self.input_version
        if not self.base_sha256:
            self.base_sha256 = self.input_sha256
        return self


class TaskValidation(StrictModel):
    valid: bool
    errors: list[str] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)


class ConversationMessage(StrictModel):
    id: int | None = None
    task_id: str
    role: ConversationRole
    kind: MessageKind = MessageKind.TEXT
    content: str
    metadata: dict[str, Any] = Field(default_factory=dict)
    created_at: str | None = None


class ClarificationPrompt(StrictModel):
    question: str
    choices: list[str] = Field(default_factory=list)
    field: str | None = None
    issue_code: str | None = None


class SkillSelection(StrictModel):
    skill_id: str
    reason: str
    confidence: float = Field(ge=0.0, le=1.0)
    boundary_notes: list[str] = Field(default_factory=list)


class CapabilityRefusal(StrictModel):
    reason: str
    boundary_rule: str
    alternatives: list[str] = Field(default_factory=list)
    suggested_request: str | None = None


class AgentTurnResult(StrictModel):
    state: ConversationState
    reply: ConversationMessage
    needs_clarification: bool = False
    envelope: TaskEnvelope | None = None
    plan: TaskPlan | None = None
    clarification: ClarificationPrompt | None = None
    actions: list[Literal["approve", "edit", "cancel"]] = Field(default_factory=list)
    selection: SkillSelection | None = None
    refusal: CapabilityRefusal | None = None


class ToolResult(StrictModel):
    tool_name: str
    modified_sheets: list[str] = Field(default_factory=list)
    modified_ranges: list[str] = Field(default_factory=list)
    created_sheets: list[str] = Field(default_factory=list)
    rows_affected: int = 0
    counters: dict[str, int | float] = Field(default_factory=dict)
    warnings: list[str] = Field(default_factory=list)
    details: dict[str, Any] = Field(default_factory=dict)


class PlanStep(StrictModel):
    order: int
    action: str
    description: str


class ExecutionPlan(StrictModel):
    task_id: str
    summary: str
    steps: list[PlanStep]
    main_sheet: str
    source_sheet: str
    key_mapping: str
    return_mapping: list[str]
    expected_new_columns: list[str]
    estimated_rows: int
    duplicate_key_count: int
    output_conflicts: list[str] = Field(default_factory=list)
    risk_level: Literal["low", "medium", "high"]
    requires_approval: bool = True


class ChangeSet(StrictModel):
    task_id: str
    version_before: int
    version_after: int
    modified_sheets: list[str]
    modified_ranges: list[str]
    created_sheets: list[str] = Field(default_factory=list)
    rows_affected: int
    matched_count: int
    unmatched_count: int
    empty_key_count: int
    duplicate_key_count: int
    warnings: list[str] = Field(default_factory=list)


class VerificationCheck(StrictModel):
    name: str
    passed: bool
    detail: str = ""


class VerificationReport(StrictModel):
    task_id: str
    status: Literal["passed", "passed_with_warnings", "failed_rolled_back"]
    checks: list[VerificationCheck]
    warnings: list[str] = Field(default_factory=list)
    manual_review_required: bool = False
    output_path: str | None = None
    unexpected_changes: list[str] = Field(default_factory=list)
    style_warnings: list[str] = Field(default_factory=list)
    final_sha256: str | None = None
    commit_id: str | None = None
    verification_scope: str = "step"


class WorkbookChangeSet(StrictModel):
    changed_ranges: list[str] = Field(default_factory=list)
    created_sheets: list[str] = Field(default_factory=list)
    deleted_sheets: list[str] = Field(default_factory=list)
    structural_changes: list[str] = Field(default_factory=list)
    unexpected_changes: list[str] = Field(default_factory=list)
    style_warnings: list[str] = Field(default_factory=list)


class FinalVerificationContext(StrictModel):
    task_id: str
    run_id: str
    fencing_token: str
    original_sha256: str
    base_version: int
    base_sha256: str
    candidate_path: str


class ExecutionLease(StrictModel):
    run_id: str
    task_id: str
    fencing_token: str
    plan_id: str
    plan_revision: int
    base_version: int
    base_sha256: str


class CommitRecord(StrictModel):
    commit_id: str
    run_id: str
    task_id: str
    target_version: int
    candidate_path: str
    final_path: str
    candidate_sha256: str
    fencing_token: str
    status: Literal["prepared", "file_committed", "completed", "failed"]
    error_json: str | None = None


class ExecutionResult(StrictModel):
    state: ConversationState
    version: int
    output_path: str
    report: VerificationReport
    result: ToolResult
    reply: ConversationMessage


class FailureCategory(StrEnum):
    TRANSIENT_IO = "transient_io"
    USER_INPUT_REQUIRED = "user_input_required"
    PLAN_REPAIRABLE = "plan_repairable"
    VERIFICATION_FAILED = "verification_failed"
    SAFETY_VIOLATION = "safety_violation"
    FATAL = "fatal"


class FailureRecord(StrictModel):
    category: FailureCategory
    step_id: str
    phase: Literal["prepare", "execute", "verify", "commit"]
    retryable: bool = False
    repairable: bool = False
    requires_user_input: bool = False
    rollback_required: bool = False
    safe_message: str
    technical_detail: str = ""
    attempt: int = 1


class ValidationResult(StrictModel):
    valid: bool
    missing_fields: list[str] = Field(default_factory=list)
    ambiguities: list[str] = Field(default_factory=list)
    duplicate_key_count: int = 0
    duplicate_examples: list[dict[str, Any]] = Field(default_factory=list)
    output_conflicts: list[str] = Field(default_factory=list)
    type_warning: str | None = None
