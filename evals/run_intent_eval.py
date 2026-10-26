from __future__ import annotations

import json
from pathlib import Path

from sheetpilot.config import Settings
from sheetpilot.llm import OpenAICompatibleClient
from sheetpilot.schemas import ColumnProfile, SheetProfile, WorkbookProfile
from sheetpilot.skills import SkillRegistry, capability_refusal


def sample_profile() -> WorkbookProfile:
    def sheet(name: str, columns: list[tuple[str, str]]) -> SheetProfile:
        return SheetProfile(
            name=name, state="visible", max_row=20, max_column=len(columns),
            header_row=1, header_confidence=1,
            columns=[
                ColumnProfile(
                    name=column, index=index, inferred_type=kind,
                    empty_count=0, unique_count=10, formula_count=0,
                )
                for index, (column, kind) in enumerate(columns, 1)
            ],
        )

    return WorkbookProfile(
        file_name="eval.xlsx", size_bytes=1, sha256="eval",
        sheets=[
            sheet("订单表", [("客户编号", "text"), ("地区", "text"), ("订单额", "number"), ("税后金额", "number")]),
            sheet("客户表", [("客户编号", "text"), ("客户等级", "text")]),
        ],
    )


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    cases = [
        json.loads(line) for line in (root / "evals" / "intent_cases.jsonl").read_text(encoding="utf-8").splitlines()
        if line.strip()
    ]
    client = OpenAICompatibleClient(Settings.load(root))
    skills = SkillRegistry()
    profile = sample_profile()
    passed = 0
    try:
        for case in cases:
            refusal = capability_refusal(case["request"])
            actual_refusal = refusal is not None
            actual_type = None
            clarification = False
            if not refusal:
                envelope = client.parse_task(
                    case["request"], profile, capability_manifest=skills.manifest()
                )
                actual_type = envelope.spec.task_type.value
                specs = getattr(envelope.spec, "operations", [envelope.spec])
                clarification = any(
                    item.missing_fields or item.ambiguities or item.confidence < 0.6
                    for item in specs
                )
            ok = (
                actual_refusal == case["expected_refusal"]
                and actual_type == case["expected_task_type"]
                and clarification == case["expected_clarification"]
            )
            passed += int(ok)
            print(json.dumps({
                "id": case["id"], "passed": ok, "task_type": actual_type,
                "refusal": actual_refusal, "clarification": clarification,
            }, ensure_ascii=False))
    finally:
        client.close()
    print(f"passed={passed}/{len(cases)}")
    return 0 if passed == len(cases) else 1


if __name__ == "__main__":
    raise SystemExit(main())
