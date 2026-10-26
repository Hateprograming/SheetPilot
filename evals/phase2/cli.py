from __future__ import annotations

import argparse
import json
import tempfile
from collections import Counter
from pathlib import Path

from openpyxl import load_workbook

from sheetpilot.schemas import TaskEnvelope

from .common import APP, ROOT, DEFAULT_OUT, Case, dump, load_cases, load_gzip, protected_output, resolve_oracle_defaults, sha


EXPECTED_CAP={"lookup":22,"normalize":18,"aggregate":22,"sort_rank":18,"clean":22,"transform":26,"formula":16,"conversion":16}


def validate(out: Path) -> list[str]:
    from evals.phase1.oracle import reference
    from sheetpilot.skills import capability_refusal
    errors=[];cases=load_cases(out/"cases")
    if len(cases)!=160:errors.append(f"expected 160 cases, found {len(cases)}")
    checks=[(Counter(c.journey for c in cases),{"direct":80,"multi_step":32,"clarification":32,"boundary":16},"journeys"),
            (Counter(c.primary_capability for c in cases),EXPECTED_CAP,"capabilities"),
            (Counter(c.size_class for c in cases),{"small":80,"medium":56,"large":24},"sizes"),
            (Counter(c.source["type"] for c in cases),{"research_reconstructed":64,"spreadsheetbench_adapted":32,"synthetic":64},"sources")]
    for actual,expected,label in checks:
        if actual!=Counter(expected):errors.append(f"unexpected {label}: {dict(actual)}")
    stable=[c for c in cases if c.stability_selected]
    if len(stable)!=40 or Counter(c.primary_capability for c in stable)!=Counter({k:5 for k in EXPECTED_CAP}):errors.append("invalid stability subset")
    clarifications=[c for c in cases if c.journey=="clarification"]
    if Counter(len(c.clarification_script) for c in clarifications)!=Counter({1:24,2:8}):errors.append("invalid clarification rounds")
    phase1_requests=set();phase1_inputs=set();phase1_sources=set();phase1_groups=set()
    for path in (APP/".phase1/cases").glob("*/case.json"):
        old=json.loads(path.read_text(encoding="utf-8"));phase1_requests.add(old.get("request"));phase1_inputs.add(old.get("input_sha256"));s=old.get("source",{})
        phase1_sources.add(str(s.get("source_item_id")));phase1_groups.add(str(s.get("task_id")))
    for case in cases:
        folder=out/"cases"/case.id;input_path=folder/case.input
        if not input_path.is_file() or sha(input_path)!=case.input_sha256:errors.append(f"{case.id}: input checksum")
        if case.input_sha256 in phase1_inputs or any(v in phase1_requests for v in case.requests.values()):errors.append(f"{case.id}: overlaps phase one")
        if input_path.suffix==".xlsx":
            try:
                book=load_workbook(input_path,read_only=True);book.close()
            except Exception as exc:errors.append(f"{case.id}: workbook cannot open: {exc}")
        if case.spec:
            try:TaskEnvelope(goal=case.request,spec=case.spec)
            except Exception as exc:errors.append(f"{case.id}: invalid final spec: {exc}")
            oracle_path=folder/case.oracle
            if not oracle_path.is_file() or sha(oracle_path)!=case.oracle_sha256:errors.append(f"{case.id}: oracle checksum")
            elif reference(input_path,resolve_oracle_defaults(case.spec))!=load_gzip(oracle_path):errors.append(f"{case.id}: oracle changed")
        elif any(capability_refusal(text) is None for text in case.requests.values()):
            errors.append(f"{case.id}: boundary request is not deterministically refused")
        if case.initial_spec:
            try:TaskEnvelope(goal=case.request,spec=case.initial_spec)
            except Exception as exc:errors.append(f"{case.id}: invalid initial spec: {exc}")
            if case.clarification_script[-1].spec_after!=case.spec:errors.append(f"{case.id}: clarification does not reach final spec")
        source=case.source
        if source["type"]=="research_reconstructed" and (not source.get("source_url") or str(source.get("source_item_id")) in phase1_sources):errors.append(f"{case.id}: invalid research source")
        if source["type"]=="spreadsheetbench_adapted":
            original=ROOT/source["original_file"]
            if str(source.get("task_id")) in phase1_groups or not original.is_file() or sha(original)!=source["original_sha256"]:errors.append(f"{case.id}: invalid benchmark source")
        if any("{" in text or "}" in text for text in case.requests.values()):errors.append(f"{case.id}: template residue")
    dump(out/"validation.json",{"count":len(cases),"errors":errors})
    return errors


def main():
    parser=argparse.ArgumentParser(description="SheetPilot phase-two locked evaluation")
    sub=parser.add_subparsers(dest="command",required=True)
    for name in ("prepare","validate","suite-report"):
        p=sub.add_parser(name);p.add_argument("--out",type=Path,default=DEFAULT_OUT)
    p=sub.add_parser("run");p.add_argument("--out",type=Path,default=DEFAULT_OUT);p.add_argument("--suite",choices=["locked","stability"],required=True)
    p.add_argument("--ids",nargs="+");p.add_argument("--max-tasks",type=int,default=20);p.add_argument("--max-calls",type=int,default=60)
    p.add_argument("--batch-seconds",type=int,default=1800);p.add_argument("--case-seconds",type=int,default=300);p.add_argument("--dry-run",action="store_true");p.add_argument("--diagnostic",action="store_true")
    p=sub.add_parser("worker");p.add_argument("--case",type=Path,required=True);p.add_argument("--folder",type=Path,required=True)
    p.add_argument("--mode",choices=["offline","live"],required=True);p.add_argument("--call-limit",type=int,required=True);p.add_argument("--request",required=True);p.add_argument("--instance-id",required=True)
    args=parser.parse_args()
    if args.command=="worker":
        from .runner import worker
        worker(args.case,protected_output(args.folder),args.mode,args.call_limit,args.request,args.instance_id);return 0
    out=protected_output(args.out)
    if args.command=="prepare":
        from .generator import prepare
        prepare(out);errors=validate(out);print(json.dumps({"cases":160,"errors":errors},ensure_ascii=False));return 1 if errors else 0
    if args.command=="validate":
        errors=validate(out);print(json.dumps({"cases":len(load_cases(out/"cases")),"errors":errors},ensure_ascii=False));return 1 if errors else 0
    if args.command=="suite-report":
        from .runner import summarize
        result=summarize(out);print(json.dumps(result["locked"],ensure_ascii=False,indent=2));return 0
    from .runner import run_suite
    run=run_suite(out,args.suite,args.ids,args.max_tasks,args.max_calls,args.batch_seconds,args.case_seconds,args.dry_run,args.diagnostic)
    if run:print(run)
    elif not args.dry_run:print("该套件没有待运行实例。")
    return 0

