from __future__ import annotations

import argparse
import json
from pathlib import Path

from .common import DEFAULT_OUT, Case, dump, load_cases, protected_output, sha


def validate(out):
    from sheetpilot.schemas import TaskEnvelope
    from .oracle import reference
    from collections import Counter
    cases=load_cases(out/'cases');errors=[]
    if len(cases)!=280:errors.append(f'expected 280 cases, found {len(cases)}')
    if Counter(c.kind for c in cases)!=Counter(normal=184,clarification=48,refusal=24,fault=24):errors.append('unexpected case distribution')
    new=[c for c in cases if c.id.startswith('P-')]
    if len(new)!=240:errors.append('expected 240 new cases')
    if Counter(c.kind for c in new)!=Counter(normal=160,clarification=40,refusal=20,fault=20):errors.append('unexpected new-case distribution')
    if Counter(c.source['type'] for c in new)!=Counter(synthetic=168,research_reconstructed=48,spreadsheetbench_adapted=24):errors.append('unexpected source distribution')
    if len({c.source['task_id'] for c in new if c.source['type']=='spreadsheetbench_adapted'})<12:errors.append('too few public-file task groups')
    if len({c.source['source_item_id'] for c in new if c.source['type']=='research_reconstructed'})!=48:errors.append('research sources are not distinct')
    for c in cases:
        path=out/'cases'/c.id/c.input
        if not path.is_file() or sha(path)!=c.input_sha256:errors.append(f'{c.id}: input checksum')
        TaskEnvelope(goal=c.request,spec=c.spec)
        if c.kind not in ('refusal','fault') and reference(path,c.spec)!=c.checks:errors.append(f'{c.id}: reference changed')
        if c.source['type']=='spreadsheetbench_adapted':
            from .common import ROOT
            if sha(ROOT/c.source['original_file'])!=c.source['original_sha256']:errors.append(f'{c.id}: original source changed')
        if c.source['type']=='research_reconstructed' and not c.source.get('source_url'):errors.append(f'{c.id}: research source missing')
    dump(out/'validation.json',{'count':len(cases),'errors':errors})
    print(json.dumps({'count':len(cases),'errors':errors},ensure_ascii=False))
    return 1 if errors else 0


def main():
    parser=argparse.ArgumentParser(description='SheetPilot phase-one cases, runner and independent checks')
    sub=parser.add_subparsers(dest='command',required=True)
    for name in ['prepare','validate','run','browser']:
        p=sub.add_parser(name);p.add_argument('--out',type=Path,default=DEFAULT_OUT)
        if name=='run':
            p.add_argument('--mode',choices=['offline','live'],default='offline')
            p.add_argument('--ids',nargs='+');p.add_argument('--kinds',nargs='+',choices=['normal','clarification','refusal','fault'])
            p.add_argument('--rerun',action='store_true',help='repeat completed live cases intentionally')
            p.add_argument('--max-tasks',type=int,default=20);p.add_argument('--max-calls',type=int,default=60)
            p.add_argument('--batch-seconds',type=int,default=1800);p.add_argument('--case-seconds',type=int,default=300)
    p=sub.add_parser('worker');p.add_argument('--case',type=Path,required=True);p.add_argument('--folder',type=Path,required=True)
    p.add_argument('--mode',choices=['offline','live'],required=True);p.add_argument('--call-limit',type=int,required=True)
    p=sub.add_parser('report');p.add_argument('run_dir',type=Path)
    p=sub.add_parser('suite-report');p.add_argument('--out',type=Path,default=DEFAULT_OUT)
    args=parser.parse_args()
    if args.command=='worker':
        from .runner import worker
        protected_output(args.folder);worker(args.case,args.folder,args.mode,args.call_limit);return 0
    if args.command=='report':
        from .runner import summarize
        summarize(protected_output(args.run_dir));return 0
    if args.command=='suite-report':
        from .runner import summarize_suite
        summary=summarize_suite(protected_output(args.out))
        offline=summary['by_mode'].get('offline',{})
        live=summary['by_mode'].get('live',{})
        print(f"案例总数{summary['total_cases']}；表格工具测试已记录{sum(offline.values())}例，其中通过{offline.get('passed',0)}例、失败{offline.get('failed',0)}例；完整流程测试已记录{sum(live.values())}例，还可运行{summary['live_remaining']}例。")
        print(protected_output(args.out)/'suite-report.md');return 0
    out=protected_output(args.out)
    if args.command=='prepare':
        from .cases import prepare
        print(prepare(out));return validate(out)
    if args.command=='validate':return validate(out)
    if args.command=='browser':
        from .browser import run_browser
        return run_browser(out)
    if min(args.max_tasks,args.max_calls,args.batch_seconds,args.case_seconds)<1:parser.error('limits must be positive')
    from .runner import run_suite,summarize
    run=run_suite(out,args.mode,args.ids,args.kinds,args.max_tasks,args.max_calls,args.batch_seconds,args.case_seconds,args.rerun)
    if run is None:
        print('所有符合条件的案例已有运行记录；用 --rerun 明确重跑。');return 0
    print(run)
    results=summarize(run)
    return 0 if results and all(r['status']=='passed' for r in results) else 1
