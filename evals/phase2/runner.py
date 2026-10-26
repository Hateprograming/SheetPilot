from __future__ import annotations

import copy
import csv
import hashlib
import json
import os
import shutil
import statistics
import subprocess
import sys
import time
import traceback
from collections import Counter, defaultdict
from dataclasses import replace
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from zoneinfo import ZoneInfo

from evals.phase1.oracle import snapshot
from evals.phase1.runner import Calls, case_semantic, check_delivery, guard_plan

from .common import APP, Case, canonical_bytes, dump, load_cases, load_gzip, protected_output, sha


class SequenceClient:
    """Explicit deterministic substitute used only by offline tests."""
    def __init__(self, case: Case): self.case=case; self.index=0
    def parse_task(self, request, profile, **kwargs):
        from sheetpilot.schemas import TaskEnvelope
        if self.case.journey == "clarification":
            specs=[self.case.initial_spec,*[s.spec_after for s in self.case.clarification_script]]
            spec=specs[min(self.index,len(specs)-1)];self.index+=1
        elif self.case.spec is not None: spec=self.case.spec
        else:
            # Refusal is decided by the product's capability boundary.  This
            # valid fallback prevents the offline double from failing first.
            spec={"task_type":"lookup","main_sheet":"交易","source_sheet":"客户","main_key_column":"客户编号",
                  "source_key_column":"客户编号","return_columns":["客户等级"],"output_column_names":["客户等级"],"confidence":1}
        return TaskEnvelope(goal=request,spec=copy.deepcopy(spec))
    def parse_lookup(self,*args,**kwargs): return self.parse_task(*args,**kwargs).spec
    def explain_plan(self,plan): return plan.summary


def adapter(case: Case, request: str, case_dir: Path):
    from sheetpilot.schemas import TaskEnvelope
    checks=load_gzip(case_dir/case.oracle) if case.oracle else {}
    normalized_spec=TaskEnvelope(goal=request,spec=case.spec).spec.model_dump(mode="json") if case.spec else None
    return SimpleNamespace(id=case.id,kind="refusal" if case.journey=="boundary" else "clarification" if case.journey=="clarification" else "normal",
        capability=case.primary_capability,spec=normalized_spec,initial_spec=case.initial_spec,reply=None,revision=None,fault=None,
        request=request,checks=checks,max_calls=case.max_calls,max_clarifications=case.max_interactions,acknowledgements=[])


def _usage(events):
    totals={"input_tokens":0,"output_tokens":0,"cached_tokens":0}
    for event in events:
        usage=event.get("usage") or {}
        totals["input_tokens"]+=int(usage.get("prompt_tokens",usage.get("input_tokens",0)) or 0)
        totals["output_tokens"]+=int(usage.get("completion_tokens",usage.get("output_tokens",0)) or 0)
        details=usage.get("prompt_tokens_details") or usage.get("input_tokens_details") or {}
        totals["cached_tokens"]+=int(
            usage.get("prompt_cache_hit_tokens")
            or details.get("cached_tokens",details.get("cache_read_input_tokens",0))
            or 0
        )
    return totals


def _event_cost(events, pricing):
    def cost(usage, rates):
        input_tokens=int(usage.get("prompt_tokens",usage.get("input_tokens",0)) or 0)
        output_tokens=int(usage.get("completion_tokens",usage.get("output_tokens",0)) or 0)
        details=usage.get("prompt_tokens_details") or usage.get("input_tokens_details") or {}
        cached=int(
            usage.get("prompt_cache_hit_tokens")
            or details.get("cached_tokens",details.get("cache_read_input_tokens",0))
            or 0
        )
        return (
            (input_tokens-cached)*rates["input_per_million"]/1e6
            +cached*rates.get("cached_input_per_million",rates["input_per_million"])/1e6
            +output_tokens*rates["output_per_million"]/1e6
        )

    if "peak" not in pricing or "off_peak" not in pricing:
        return sum(cost(event.get("usage") or {},pricing) for event in events),{}
    timezone=ZoneInfo(pricing.get("timezone","Asia/Shanghai"));counts=Counter()
    windows=pricing.get("peak_windows",[]);total=0.0
    for event in events:
        local=datetime.fromtimestamp(float(event.get("started",0)),tz=timezone)
        minute=local.hour*60+local.minute
        peak=any(
            local.weekday() in window.get("weekdays",range(5))
            and int(window["start"].split(":")[0])*60+int(window["start"].split(":")[1]) <= minute
            < int(window["end"].split(":")[0])*60+int(window["end"].split(":")[1])
            for window in windows
        )
        tier="peak" if peak else "off_peak";counts[tier]+=1
        total+=cost(event.get("usage") or {},pricing[tier])
    return total,dict(counts)


def _environment_failure(events):
    statuses = [event.get("status") for event in events]
    if 401 in statuses or 403 in statuses:
        return "authentication_error", "model API authentication failed"
    if 429 in statuses:
        return "provider_rate_limit", "model API returned HTTP 429"
    if any(isinstance(status, int) and status >= 500 for status in statuses):
        return "provider_unavailable", "model API returned a server error"
    if "transport_error" in statuses:
        return "provider_unavailable", "model API transport failed"
    return None


def _clarification_answer(case: Case, turn, used_fields: set[str]):
    if not turn.envelope:
        return None
    payload=turn.envelope.spec.model_dump(mode="json")
    items=payload.get("operations",[]) if payload.get("task_type")=="composite" else [payload]
    missing={str(field) for item in items for field in item.get("missing_fields",[])}
    remaining=[step for step in case.clarification_script if step.field not in used_fields]
    requested = getattr(getattr(turn, "clarification", None), "field", None)
    if requested:
        exact = next((step for step in remaining if step.field == requested), None)
        if exact:
            return exact
        aliases = {
            "rank_column": {"sort_keys", "rank_column"},
            "sort_keys": {"sort_keys", "rank_column"},
            "sheet": {"sheet", "main_sheet", "source_sheet"},
            "column": {"column"},
        }
        matched = [step for step in remaining if step.field in aliases.get(requested, set())]
        if len(matched) == 1:
            return matched[0]
    for step in remaining:
        if step.field in missing or any(step.field in field for field in missing):
            return step
    return remaining[0] if len(remaining)==1 else None


def _semantic_plan_sha(plan, profile, request: str) -> str:
    payload = [
        {
            "tool_name": step.tool_name,
            "arguments": case_semantic(step.arguments, profile, request),
        }
        for step in plan.steps
    ]
    return hashlib.sha256(canonical_bytes(payload)).hexdigest()


def _semantic_output_sha(output: Path, artifact_dir: Path) -> str:
    artifacts = []
    for path in sorted(artifact_dir.glob("*.csv")) if artifact_dir.exists() else []:
        with path.open(encoding="utf-8-sig", newline="") as stream:
            artifacts.append(list(csv.reader(stream)))
    artifacts.sort(key=canonical_bytes)
    payload = {"workbook": snapshot(output), "csv_artifacts": artifacts}
    return hashlib.sha256(canonical_bytes(payload)).hexdigest()


def _progress(folder: Path, stage: str, timings: dict | None = None) -> None:
    dump(folder / "progress.json", {"stage": stage, "timings": timings or {}})


def _restore_call_metrics(result: dict, calls_path: Path) -> dict:
    """Use the append-only provider log as the source of truth after a timeout."""
    if calls_path.exists():
        events = json.loads(calls_path.read_text(encoding="utf-8"))
        result["calls"] = len(events)
        result["token_usage"] = _usage(events)
    return result


def worker(case_path: Path, folder: Path, mode: str, call_limit: int, request: str, instance_id: str):
    from sheetpilot.config import Settings
    from sheetpilot.llm import OpenAICompatibleClient
    from sheetpilot.schemas import MessageKind
    from sheetpilot.service import SheetPilotService

    case=Case.model_validate_json(case_path.read_text(encoding="utf-8"))
    compat=adapter(case,request,case_path.parent);folder.mkdir(parents=True,exist_ok=True)
    result={"case_id":case.id,"instance_id":instance_id,"journey":case.journey,"capability":case.primary_capability,
            "mode":mode,"status":"tool_error","stage":"setup","calls":0,"clarification_rounds":0,"user_turns":1,
            "timings":{},"token_usage":{"input_tokens":0,"output_tokens":0,"cached_tokens":0},"differences":[]}
    started=time.monotonic();service=client=task_id=calls=None
    try:
        _progress(folder,"setup")
        input_path=case_path.parent/case.input
        if sha(input_path)!=case.input_sha256: raise ValueError("case input checksum mismatch")
        settings=Settings.load(APP)
        if mode=="live" and not settings.llm_configured:
            result.update(status="blocked",stage="configuration",error_category="configuration",reason="model configuration missing");return result
        settings=replace(settings,data_dir=folder/"workspace",database_path=folder/"workspace/eval.sqlite3")
        result["provider"]=settings.llm_provider;result["model"]=settings.model
        calls=Calls(folder/"calls.json",min(call_limit,case.max_calls))
        if mode=="offline": client=SequenceClient(case)
        else:
            client=OpenAICompatibleClient(settings);client.client.event_hooks={"request":[calls.before],"response":[calls.after]}
        service=SheetPilotService(settings,client)
        t=time.monotonic();profile=service.create_task(input_path.name,input_path.read_bytes());task_id=service.task_id_from_profile(profile)
        result["timings"]["file_load_seconds"]=time.monotonic()-t
        _progress(folder,"planning",result["timings"])
        baseline=len(service.repository.list_versions(task_id));t=time.monotonic();turn=service.submit_message(task_id,request)
        result["timings"]["planning_seconds"]=time.monotonic()-t
        if turn.reply.kind==MessageKind.ERROR:
            environment=_environment_failure(calls.events if calls else [])
            if environment:
                category,reason=environment
                result.update(status="blocked",stage="environment",error_category=category,reason=reason);return result
            result.update(status="failed",stage="planning",error_category="planning_error",reason="product returned planning error");return result
        if case.journey=="boundary":
            safe=turn.refusal is not None and turn.plan is None and len(service.repository.list_versions(task_id))==baseline
            result.update(status="passed" if safe else "failed",stage="refusal",error_category=None if safe else "missed_refusal",
                          reason="explicit refusal" if safe else "unsupported request not refused");return result
        clarification_seconds=0.0
        used_fields:set[str]=set()
        while turn.needs_clarification and result["clarification_rounds"] < case.max_interactions:
            step=_clarification_answer(case,turn,used_fields)
            if step is None:
                result.update(status="failed",stage="clarification",error_category="clarification_error",reason="product requested information outside the case script");return result
            t=time.monotonic();turn=service.submit_message(task_id,step.answer);clarification_seconds+=time.monotonic()-t
            used_fields.add(step.field)
            result["clarification_rounds"]+=1;result["user_turns"]+=1
        result["timings"]["clarification_seconds"]=clarification_seconds
        if turn.needs_clarification:
            result.update(status="failed",stage="clarification",error_category="clarification_error",reason="unexpected extra clarification");return result
        if turn.reply.kind==MessageKind.ERROR or not turn.plan:
            result.update(status="failed",stage="planning",error_category="planning_error",reason="no executable plan");return result
        diffs=guard_plan(compat,turn,profile)
        if diffs:
            result.update(status="failed",stage="planning",error_category="planning_error",reason="plan does not match locked task",differences=diffs);return result
        result["plan_sha256"]=hashlib.sha256(json.dumps(turn.plan.model_dump(mode="json"),ensure_ascii=False,sort_keys=True).encode()).hexdigest()
        result["semantic_plan_sha256"]=_semantic_plan_sha(turn.plan,profile,request)
        _progress(folder,"execution",result["timings"])
        t=time.monotonic();executed=service.approve_current_plan(task_id,[]);result["timings"]["execution_seconds"]=time.monotonic()-t
        _progress(folder,"checking",result["timings"])
        t=time.monotonic();output=Path(executed.output_path);diffs=check_delivery(compat,input_path,output,service,task_id,executed)
        result["timings"]["checking_seconds"]=time.monotonic()-t
        result.update(status="failed" if diffs else "passed",stage="checking",error_category="result_error" if diffs else None,
                      reason="independent output checks",differences=diffs,output=str(output),output_sha256=sha(output),
                      semantic_output_sha256=_semantic_output_sha(output,service.files.task_dir(task_id)/"artifacts"))
        _progress(folder,"completed",result["timings"])
        return result
    except Exception as exc:
        result.update(status="tool_error",stage=result.get("stage","worker"),error_category="execution_error",reason=str(exc),
                      error_type=type(exc).__name__,traceback=traceback.format_exc());return result
    finally:
        result["timings"]["total_seconds"]=time.monotonic()-started
        events=calls.events if calls else []
        result["calls"]=len(events);result["token_usage"]=_usage(events)
        if service and task_id:
            try:
                dump(folder/"conversation.json",[m.model_dump(mode="json") for m in service.get_conversation(task_id)])
                task=service.repository.get_task(task_id);dump(folder/"plan.json",json.loads(task["plan_json"]) if task.get("plan_json") else None)
            except Exception: pass
            service._background_executor.shutdown(wait=False)
        if client and mode=="live":
            try: client.close()
            except Exception: pass
        dump(folder/"result.json",result)


def instances(cases: list[Case], suite: str):
    if suite=="locked": return [(c,"concise",1,c.requests["concise"],c.id) for c in cases]
    return [(c,variant,repeat,c.requests[variant],f"{c.id}--{variant}--r{repeat}")
            for c in cases if c.stability_selected for variant in ("concise","colloquial","contextual") for repeat in range(1,4)]


def completed(out: Path, suite: str) -> set[str]:
    done=set();root=out/"official-runs"/suite
    for path in root.glob("*/*/result.json") if root.exists() else []:
        result=json.loads(path.read_text(encoding="utf-8"))
        if result.get("stage") not in ("batch_limit","configuration","environment") and result.get("status")!="blocked":
            done.add(result["instance_id"])
    return done


def run_suite(out: Path,suite: str,ids=None,max_tasks=20,max_calls=60,batch_seconds=1800,case_seconds=300,dry_run=False,diagnostic=False):
    out=protected_output(out);all_cases=load_cases(out/"cases")
    if len(all_cases) != 160:
        raise RuntimeError(
            f"phase-two suite is missing or incomplete: expected 160 cases, found {len(all_cases)}; "
            "run `python -m evals.phase2 prepare` first"
        )
    if ids:
        unknown=set(ids)-{c.id for c in all_cases}
        if unknown: raise ValueError(f"unknown case IDs: {sorted(unknown)}")
        all_cases=[c for c in all_cases if c.id in ids]
    work=instances(all_cases,suite)
    if not diagnostic: work=[x for x in work if x[4] not in completed(out,suite)]
    # A partial batch should cover different capabilities and journeys instead
    # of consuming one alphabetic family first.
    work.sort(key=lambda item: hashlib.sha256(f"20261022:{item[4]}".encode()).hexdigest())
    work=work[:min(max_tasks,20)]
    if dry_run:
        print(json.dumps({"suite":suite,"selected_count":len(work),"selected_instances":[x[4] for x in work],"model_calls":0},ensure_ascii=False,indent=2));return None
    if not work: return None
    category="diagnostic-runs" if diagnostic else "official-runs";run_dir=out/category/suite/(datetime.now().strftime("%Y%m%d-%H%M%S-%f"));run_dir.mkdir(parents=True)
    dump(run_dir/"run.json",{"suite":suite,"selected_instances":[x[4] for x in work],"max_tasks":min(max_tasks,20),"max_calls":min(max_calls,60),"official":not diagnostic})
    used=0;batch_start=time.monotonic()
    for case,variant,repeat,request,instance_id in work:
        if used>=min(max_calls,60) or time.monotonic()-batch_start>=min(batch_seconds,1800): break
        folder=run_dir/instance_id;frozen=folder/"case";frozen.mkdir(parents=True)
        source=out/"cases"/case.id/case.input;shutil.copyfile(source,frozen/case.input)
        if case.oracle: shutil.copyfile(out/"cases"/case.id/case.oracle,frozen/case.oracle)
        dump(frozen/"case.json",case.model_dump(mode="json"))
        cmd=[sys.executable,"-m","evals.phase2","worker","--case",str(frozen/"case.json"),"--folder",str(folder),
             "--mode","live","--call-limit",str(min(case.max_calls,60-used)),"--request",request,"--instance-id",instance_id]
        with (folder/"process.log").open("w",encoding="utf-8") as log:
            proc=subprocess.Popen(cmd,cwd=APP,env=dict(os.environ,PYTHONIOENCODING="utf-8"),stdout=log,stderr=subprocess.STDOUT)
            try: proc.wait(timeout=min(case_seconds,case.max_seconds))
            except subprocess.TimeoutExpired:
                proc.kill();proc.wait()
                progress_path=folder/"progress.json"
                progress=json.loads(progress_path.read_text(encoding="utf-8")) if progress_path.exists() else {}
                calls_path=folder/"calls.json"
                timeout_events=json.loads(calls_path.read_text(encoding="utf-8")) if calls_path.exists() else []
                timings=dict(progress.get("timings") or {})
                timings["total_seconds"]=case_seconds
                dump(folder/"result.json",{"case_id":case.id,"instance_id":instance_id,"journey":case.journey,
                    "capability":case.primary_capability,"status":"failed","stage":"runner_timeout","timeout_stage":progress.get("stage","unknown"),
                    "error_category":"timeout","calls":len(timeout_events),"clarification_rounds":0,"user_turns":1,
                    "timings":timings,"token_usage":_usage(timeout_events)})
        events=json.loads((folder/"calls.json").read_text(encoding="utf-8")) if (folder/"calls.json").exists() else [];used+=len(events)
        result_path=folder/"result.json"
        if result_path.exists() and json.loads(result_path.read_text(encoding="utf-8")).get("status")=="blocked":
            break
    summarize(out);return run_dir


def _percentile(values,p):
    if not values:return None
    ordered=sorted(values);index=(len(ordered)-1)*p;low=int(index);high=min(low+1,len(ordered)-1)
    return ordered[low]+(ordered[high]-ordered[low])*(index-low)


def summarize(out: Path):
    out=protected_output(out);results=[]
    for path in (out/"official-runs").glob("*/*/*/result.json") if (out/"official-runs").exists() else []:
        item=json.loads(path.read_text(encoding="utf-8"));item["result_path"]=str(path)
        calls_path=path.with_name("calls.json")
        _restore_call_metrics(item,calls_path)
        results.append(item)
    pricing_path=out/"pricing.json";pricing=json.loads(pricing_path.read_text(encoding="utf-8")) if pricing_path.exists() else None
    for result in results:
        usage=result.get("token_usage",{});result["cost_status"]="unpriced"
        if pricing:
            calls_path=Path(result["result_path"]).with_name("calls.json")
            events=json.loads(calls_path.read_text(encoding="utf-8")) if calls_path.exists() else []
            result["cost"],result["cost_tiers"]=_event_cost(events,pricing)
            result["cost_status"]="priced"
    locked=[r for r in results if "--" not in r["instance_id"]];stable=[r for r in results if "--" in r["instance_id"]]
    valid_locked=[r for r in locked if r.get("status")!="blocked" and r.get("stage")!="environment"]
    valid_stable=[r for r in stable if r.get("status")!="blocked" and r.get("stage")!="environment"]
    times=[r.get("timings",{}).get("total_seconds") for r in valid_locked if r.get("timings",{}).get("total_seconds") is not None]
    payload={"generated_at":datetime.now().isoformat(),"locked":{"completed":len(valid_locked),"attempted":len(locked),
              "environment_blocked":len(locked)-len(valid_locked),"passed":sum(r["status"]=="passed" for r in valid_locked),
              "success_rate":sum(r["status"]=="passed" for r in valid_locked)/len(valid_locked) if valid_locked else None,
              "errors":dict(Counter(r.get("error_category") or "none" for r in valid_locked if r["status"]!="passed")),
              "latency_seconds":{"p50":_percentile(times,.5),"p95":_percentile(times,.95),"mean":statistics.mean(times) if times else None},
              "mean_clarification_rounds":statistics.mean(r.get("clarification_rounds",0) for r in valid_locked) if valid_locked else None,
              "tokens":{k:sum(r.get("token_usage",{}).get(k,0) for r in valid_locked) for k in ("input_tokens","output_tokens","cached_tokens")},
              "cost":sum(r.get("cost",0) for r in valid_locked) if pricing else None,"cost_status":"priced" if pricing else "unpriced"},
             "stability":{"completed":len(valid_stable),"attempted":len(stable),"expected":360,
               "environment_blocked":len(stable)-len(valid_stable),
               "passed":sum(r.get("status")=="passed" for r in valid_stable),
               "success_rate":sum(r.get("status")=="passed" for r in valid_stable)/len(valid_stable) if valid_stable else None,
               "errors":dict(Counter(r.get("error_category") or "none" for r in valid_stable if r.get("status")!="passed")),
               "tokens":{k:sum(r.get("token_usage",{}).get(k,0) for r in valid_stable) for k in ("input_tokens","output_tokens","cached_tokens")},
               "cost":sum(r.get("cost",0) for r in valid_stable) if pricing else None,
               "cost_status":"priced" if pricing else "unpriced"},"results":results}
    case_meta={case.id:case for case in load_cases(out/"cases")}
    def breakdown(rows,key):
        grouped=defaultdict(list)
        for row in rows: grouped[str(key(row))].append(row)
        return {
            name:{"runs":len(group),"passed":sum(item.get("status")=="passed" for item in group),
                  "success_rate":sum(item.get("status")=="passed" for item in group)/len(group)}
            for name,group in sorted(grouped.items())
        }
    payload["locked"]["by_capability"]=breakdown(valid_locked,lambda row:row["capability"])
    payload["locked"]["by_journey"]=breakdown(valid_locked,lambda row:row["journey"])
    payload["locked"]["by_size"]=breakdown(valid_locked,lambda row:case_meta[row["case_id"]].size_class)
    payload["locked"]["by_source"]=breakdown(valid_locked,lambda row:case_meta[row["case_id"]].source["type"])
    payload["stability"]["by_variant"]=breakdown(valid_stable,lambda row:row["instance_id"].split("--")[1]) if valid_stable else {}
    payload["stability"]["by_repeat"]=breakdown(valid_stable,lambda row:row["instance_id"].split("--")[2]) if valid_stable else {}
    groups=defaultdict(list)
    for r in valid_stable: groups[r["case_id"]].append(r)
    stability_cases={}
    for case_id,group in groups.items():
        result_hashes=[r.get("semantic_output_sha256") for r in group]
        plan_hashes=[r.get("semantic_plan_sha256") for r in group]
        stability_cases[case_id]={
            "runs":len(group),"success_rate":sum(r.get("status")=="passed" for r in group)/len(group),
            "result_consistency":all(result_hashes) and len(set(result_hashes))==1,
            "plan_consistency":all(plan_hashes) and len(set(plan_hashes))==1,
        }
    payload["stability"]["cases"]=stability_cases
    complete_cases=[item for item in stability_cases.values() if item["runs"]==9]
    payload["stability"]["plan_consistency_rate"]=(
        sum(item["plan_consistency"] for item in complete_cases)/len(complete_cases)
        if complete_cases else None
    )
    payload["stability"]["result_consistency_rate"]=(
        sum(item["result_consistency"] for item in complete_cases)/len(complete_cases)
        if complete_cases else None
    )
    dump(out/"suite-summary.json",payload)
    with (out/"suite-summary.csv").open("w",encoding="utf-8-sig",newline="") as stream:
        fields=["instance_id","case_id","journey","capability","status","stage","error_category","calls","clarification_rounds","user_turns","cost_status","cost","result_path"]
        writer=csv.DictWriter(stream,fieldnames=fields,extrasaction="ignore");writer.writeheader();writer.writerows(results)
    md=["# 第二阶段评测结果","",f"正式集已完成 {len(valid_locked)}/160，稳定性测试已完成 {len(valid_stable)}/360。","",
        f"正式集成功率：{payload['locked']['success_rate'] if locked else '尚无结果'}。",
        f"延时 P50/P95：{payload['locked']['latency_seconds']['p50']} / {payload['locked']['latency_seconds']['p95']} 秒。",
        f"正式集 Token：{payload['locked']['tokens']}；成本：{payload['locked']['cost']}（{payload['locked']['cost_status']}）。","",
        f"稳定性成功率：{payload['stability']['success_rate'] if valid_stable else '尚无结果'}。",
        f"计划一致率：{payload['stability']['plan_consistency_rate']}；结果一致率：{payload['stability']['result_consistency_rate']}。","",
        "## 分组结果","",
        "| 维度 | 分组 | 运行数 | 成功数 | 成功率 |","|---|---|---:|---:|---:|"]
    for dimension in ("by_capability","by_journey","by_size","by_source"):
        for name,values in payload["locked"][dimension].items():
            md.append(f"| {dimension.removeprefix('by_')} | {name} | {values['runs']} | {values['passed']} | {values['success_rate']:.2%} |")
    for dimension in ("by_variant","by_repeat"):
        for name,values in payload["stability"][dimension].items():
            md.append(f"| {dimension.removeprefix('by_')} | {name} | {values['runs']} | {values['passed']} | {values['success_rate']:.2%} |")
    (out/"suite-report.md").write_text("\n".join(md)+"\n",encoding="utf-8")
    return payload

