from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator

APP = Path(__file__).resolve().parents[2]
ROOT = APP.parent
DEFAULT_OUT = APP / '.phase1'
VERSION = '1.0'


class Case(BaseModel):
    model_config = ConfigDict(extra='forbid')
    version: Literal['1.0'] = VERSION
    id: str = Field(pattern=r'^[A-Z][A-Z0-9_-]+$')
    kind: Literal['normal', 'clarification', 'refusal', 'fault']
    modes: list[Literal['offline', 'live']] = Field(default_factory=lambda: ['offline', 'live'])
    capability: str
    source: dict[str, Any]
    seed: int = 20261018
    group: str
    input: str
    input_sha256: str
    request: str
    spec: dict[str, Any]
    initial_spec: dict[str, Any] | None = None
    reply: dict[str, Any] | None = None
    revision: str | None = None
    fault: Literal['timeout', 'invalid_response', 'step_failure', 'stale_plan'] | None = None
    checks: dict[str, Any]
    max_seconds: int = Field(default=300, ge=1, le=300)
    max_calls: int = Field(default=6, ge=1, le=6)
    max_clarifications: int = Field(default=3, ge=0, le=3)
    acknowledgements: list[str] = Field(default_factory=list)

    @model_validator(mode='after')
    def check_relative(self):
        if Path(self.input).is_absolute() or '..' in Path(self.input).parts:
            raise ValueError('input must remain inside the case directory')
        if self.kind == 'fault' and not self.fault:
            raise ValueError('fault case requires an injection')
        if self.kind != 'fault' and self.fault:
            raise ValueError('fault injection belongs to fault cases only')
        if self.kind == 'fault':
            self.modes = ['offline']
        if not self.modes or len(set(self.modes)) != len(self.modes):
            raise ValueError('each case needs distinct supported modes')
        return self


def dump(path: Path, obj):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=2, default=str), encoding='utf-8')


def sha(path: Path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def protected_output(path: Path) -> Path:
    path = path.resolve()
    # All generated data stays under a dedicated directory; never accept data/data_web.
    if path != DEFAULT_OUT.resolve() and DEFAULT_OUT.resolve() not in path.parents:
        raise ValueError(f'output must be inside {DEFAULT_OUT}')
    return path


def load_cases(directory: Path):
    cases = [Case.model_validate_json(p.read_text(encoding='utf-8')) for p in sorted(directory.glob('*/case.json'))]
    if len({c.id for c in cases}) != len(cases):
        raise ValueError('duplicate case IDs')
    return cases
