from __future__ import annotations

import copy
import gzip
import hashlib
import json
from pathlib import Path
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator

APP = Path(__file__).resolve().parents[2]
ROOT = APP.parent
PHASE2_ROOT = APP / ".phase2"
DEFAULT_OUT = PHASE2_ROOT / "deepseek-flash-v2.1"
VERSION = "2.1"
SEED = 20261022


class ClarificationStep(BaseModel):
    model_config = ConfigDict(extra="forbid")
    field: str
    answer: str
    spec_after: dict[str, Any]


class Case(BaseModel):
    model_config = ConfigDict(extra="forbid")
    version: Literal["2.0", "2.1"] = VERSION
    id: str = Field(pattern=r"^P2-[A-Z_]+-\d{3}$")
    journey: Literal["direct", "multi_step", "clarification", "boundary"]
    primary_capability: Literal["lookup", "normalize", "aggregate", "sort_rank", "clean", "transform", "formula", "conversion"]
    capabilities: list[str]
    source: dict[str, Any]
    seed: int = SEED
    family_id: str
    size_class: Literal["small", "medium", "large"]
    row_counts: dict[str, int]
    input: str
    input_sha256: str
    requests: dict[Literal["concise", "colloquial", "contextual"], str]
    expected_action: Literal["execute", "clarify_then_execute", "refuse"]
    spec: dict[str, Any] | None = None
    initial_spec: dict[str, Any] | None = None
    clarification_script: list[ClarificationStep] = Field(default_factory=list)
    allowed_changes: list[str] = Field(default_factory=list)
    forbidden_changes: list[str] = Field(default_factory=list)
    oracle: str | None = None
    oracle_sha256: str | None = None
    max_seconds: int = Field(default=300, ge=1, le=300)
    max_calls: int = Field(default=6, ge=1, le=6)
    max_interactions: int = Field(default=3, ge=0, le=3)
    stability_selected: bool = False

    @model_validator(mode="after")
    def coherent(self):
        for name in (self.input, self.oracle):
            if name and (Path(name).is_absolute() or ".." in Path(name).parts):
                raise ValueError("case files must remain inside the case directory")
        if set(self.requests) != {"concise", "colloquial", "contextual"} or any(not v.strip() for v in self.requests.values()):
            raise ValueError("three non-empty request variants are required")
        if self.journey == "boundary":
            if self.expected_action != "refuse" or self.spec is not None or self.oracle is not None:
                raise ValueError("boundary cases must refuse and have no executable oracle")
        elif not self.spec or not self.oracle or not self.oracle_sha256:
            raise ValueError("executable cases need a spec and oracle")
        if self.journey == "clarification":
            if self.expected_action != "clarify_then_execute" or not self.initial_spec or not self.clarification_script:
                raise ValueError("clarification cases need an initial spec and answer script")
        elif self.clarification_script or self.initial_spec:
            raise ValueError("only clarification cases may define a script")
        return self

    @property
    def request(self) -> str:
        return self.requests["concise"]


def canonical_bytes(value: Any) -> bytes:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"), default=str).encode("utf-8")


def resolve_oracle_defaults(spec: dict[str, Any]) -> dict[str, Any]:
    """Resolve product-documented output defaults in the independent oracle input."""
    resolved = copy.deepcopy(spec)
    operations = resolved.get("operations", []) if resolved.get("task_type") == "composite" else [resolved]
    for operation in operations:
        task_type = operation.get("task_type")
        if task_type == "aggregate":
            for metric in operation.get("metrics", []):
                metric["output_column"] = metric.get("output_column") or f"{metric['column']}_{metric['function']}"
        elif task_type == "normalize" and not operation.get("overwrite"):
            operation["output_column"] = operation.get("output_column") or f"{operation['column']}_标准化"
        elif task_type == "transform" and operation.get("operation") == "map_values":
            for rule in operation.get("map_rules", []):
                label = rule["source"].get("display_header") or rule["source"].get("column_letter")
                rule["output_column"] = rule.get("output_column") or f"{label}_映射"
        elif task_type == "transform" and operation.get("operation") == "impute_numeric":
            for rule in operation.get("impute_rules", []):
                label = rule["source"].get("display_header") or rule["source"].get("column_letter")
                value_name = rule.get("output_column") or f"{label}_填补"
                rule["output_column"] = value_name
                rule["indicator_column"] = rule.get("indicator_column") or f"{value_name}_是否填补"
    return resolved


def dump(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2, sort_keys=True, default=str) + "\n", encoding="utf-8")


def dump_gzip(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("wb") as raw:
        with gzip.GzipFile(filename="", mode="wb", fileobj=raw, mtime=0) as stream:
            stream.write(canonical_bytes(value))


def load_gzip(path: Path) -> Any:
    with gzip.open(path, "rt", encoding="utf-8") as stream:
        return json.load(stream)


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def protected_output(path: Path) -> Path:
    resolved = path.resolve()
    base = PHASE2_ROOT.resolve()
    if resolved != base and base not in resolved.parents:
        raise ValueError(f"output must be inside {PHASE2_ROOT}")
    return resolved


def load_cases(directory: Path) -> list[Case]:
    cases = [Case.model_validate_json(path.read_text(encoding="utf-8")) for path in sorted(directory.glob("*/case.json"))]
    if len({case.id for case in cases}) != len(cases):
        raise ValueError("duplicate case IDs")
    return cases

