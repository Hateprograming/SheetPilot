from __future__ import annotations

import copy
import csv
import json
import os
import re
import subprocess
import sys
import time
import traceback
from collections import Counter
from dataclasses import replace
from pathlib import Path

from .common import APP, Case, dump, load_cases, protected_output, sha
from .oracle import compare, snapshot, unchanged_metadata


class ScriptedClient:
    """Explicit model substitute. Never used as fallback for live runs."""
    def __init__(self,case):self.case=case;self.count=0
    def parse_task(self,request,profile,**kwargs):
        from sheetpilot.schemas import TaskEnvelope
        self.count+=1
        op=self.case.initial_spec if self.count==1 and self.case.initial_spec else self.case.spec
        return TaskEnvelope(goal=request,spec=copy.deepcopy(op))
    def parse_lookup(self,*args,**kwargs):return self.parse_task(*args,**kwargs).spec
    def explain_plan(self,plan):return plan.summary


class Calls:
    def __init__(self,path,limit):self.path=path;self.limit=limit;self.events=[]
    def before(self,request):
        if len(self.events)>=self.limit:raise RuntimeError('PHASE1_CALL_BUDGET_EXCEEDED')
        payload=json.loads(request.content)
        # No requests or cell data are persisted in call telemetry.
        self.events.append({'index':len(self.events)+1,'model':payload.get('model'),
                            'started':time.time(),'usage':None,'status':'started'})
        dump(self.path,self.events)
    def after(self,response):
        response.read()
        last=self.events[-1];last.update(status=response.status_code,elapsed_seconds=time.time()-last['started'])
        try:
            payload=response.json();last['usage']=payload.get('usage')
            dump(self.path.parent/f"response-{last['index']}.json",payload)
        except ValueError:pass
        dump(self.path,self.events)


def attach_client(settings,case,mode,folder,call_limit):
    from sheetpilot.llm import OpenAICompatibleClient
    import httpx
    calls=Calls(folder/'calls.json',call_limit)
    if case.fault in ('timeout','invalid_response'):
        settings=replace(settings,api_key='offline-fault',base_url='https://phase1.invalid',model='fault-double')
        client=OpenAICompatibleClient(settings)
        client.client.close()
        def respond(request):
            if case.fault=='timeout':raise httpx.ReadTimeout('injected model timeout',request=request)
            return httpx.Response(200,json={'choices':[{'message':{'content':'not valid JSON'}}]})
        client.client=httpx.Client(base_url=settings.base_url,transport=httpx.MockTransport(respond),
                                 event_hooks={'request':[calls.before],'response':[calls.after]})
        return client,calls
    if mode=='offline':return ScriptedClient(case),calls
    client=OpenAICompatibleClient(settings)
    client.client.event_hooks={'request':[calls.before],'response':[calls.after]}
    original_send=client.client.send
    def observed_send(*args,**kwargs):
        started=len(calls.events)
        try:return original_send(*args,**kwargs)
        except Exception as exc:
            if len(calls.events)>started and calls.events[-1]['status']=='started':
                message=str(exc).replace(settings.api_key,'[redacted]')
                calls.events[-1].update(status='transport_error',error_type=type(exc).__name__,error=message[:500])
                dump(calls.path,calls.events)
            raise
    client.client.send=observed_send
    return client,calls


def semantic(spec,profile):
    """Compare declared parameters after resolving equivalent column and export defaults."""
    if isinstance(spec,list):return [semantic(v,profile) for v in spec]
    if not isinstance(spec,dict):return spec
    spec=dict(spec)
    if spec.get('task_type')=='conversion' and spec.get('direction')=='xlsx_to_csv':
        # The export tool resolves an unspecified delimiter to a comma.
        # Compare the resulting CSV format, not the two equivalent spellings.
        spec={**spec,'delimiter':',' if spec.get('delimiter','auto')=='auto' else spec['delimiter']}
    if spec.get('task_type')=='sort_rank' and not spec.get('rank_column'):
        for key in ('rank_output_column','rank_direction','rank_method','group_by'):
            spec.pop(key,None)
    if spec.get('task_type')=='clean':
        operation=spec.get('operation')
        if operation=='filter_rows':
            for key in ('key_columns','keep','keep_columns','duplicate_flag_column'):
                spec.pop(key,None)
        elif operation=='keep_columns':
            for key in ('key_columns','keep','filters','duplicate_flag_column'):
                spec.pop(key,None)
        else:
            spec.pop('filters',None);spec.pop('keep_columns',None)
            if operation!='mark_duplicates':spec.pop('duplicate_flag_column',None)
    if 'resolution_policies' in spec:
        # Omitting the policy and spelling out leave_blank execute identically.
        spec['resolution_policies']={
            key:value for key,value in spec['resolution_policies'].items()
            if value!='leave_blank'
        }
    if 'sheet' in spec and any(k in spec for k in ['display_header','column_letter','column_index']):
        sheet=next((s for s in profile.sheets if s.name==spec['sheet']),None)
        if sheet:
            cols=[c for c in sheet.columns if
                  (spec.get('column_index') and c.index==spec['column_index']) or
                  (spec.get('column_letter') and c.letter==spec['column_letter']) or
                  (spec.get('display_header') and c.display_header==spec['display_header'] and c.header_occurrence==spec.get('header_occurrence',1))]
            if len(cols)==1:return {'sheet':sheet.name,'column_index':cols[0].index}
    # constraints is non-executable free text (e.g. 保留行序 versus
    # preserve_row_order); actual order/ranges are checked independently.
    return {k:semantic(v,profile) for k,v in spec.items() if k not in ('confidence','missing_fields','ambiguities','constraints')}


def case_semantic(spec, profile, request):
    """Apply request-aware equivalence before comparing executable semantics."""
    prepared = copy.deepcopy(spec)
    operations = (
        prepared.get('operations', [])
        if isinstance(prepared, dict) and prepared.get('task_type') == 'composite'
        else []
    )
    items = operations or ([prepared] if isinstance(prepared, dict) else [])
    explicit_flag_name = any(
        cue in request for cue in ('标记列叫', '重复标记列叫', '标记列名', '重复标记列名')
    )
    if not explicit_flag_name:
        for item in items:
            if item.get('task_type') == 'clean' and item.get('operation') == 'mark_duplicates':
                item.pop('duplicate_flag_column', None)
    explicit_result_column = bool(re.search(r"结果列(?:叫|名为)", request))
    explicit_new_column = bool(re.search(r"(?:新列|新增列)(?:叫|名为)", request))
    explicit_impute_names = bool(re.search(r"新列(?:输出)?[^，。；;]+和[^，。；;]+", request))
    for item in items:
        if item.get('task_type') == 'aggregate' and not explicit_result_column:
            for metric in item.get('metrics', []):
                metric.pop('output_column', None)
        if item.get('task_type') != 'transform':
            continue
        if item.get('operation') == 'map_values' and not explicit_new_column:
            for rule in item.get('map_rules', []):
                rule.pop('output_column', None)
                rule.pop('status_column', None)
        if item.get('operation') == 'impute_numeric' and not explicit_impute_names:
            for rule in item.get('impute_rules', []):
                rule.pop('output_column', None)
                rule.pop('indicator_column', None)
    return semantic(prepared, profile)


def guard_plan(case,turn,profile):
    if not turn.plan or not turn.envelope:return [{'scope':'plan','reason':'missing plan'}]
    expected=case_semantic(case.spec,profile,case.request)
    actual=case_semantic(turn.envelope.spec.model_dump(mode='json'),profile,case.request)
    differences=[]
    if expected!=actual:differences.append({'scope':'plan_parameters','expected':expected,'actual':actual})
    expected_ops=case.spec['operations'] if case.spec['task_type']=='composite' else [case.spec]
    if len(turn.plan.steps)!=len(expected_ops):differences.append({'scope':'step_count'})
    for step,op in zip(turn.plan.steps,expected_ops):
        tool_name={'normalize':'normalize_column','transform':'transform_column'}.get(op['task_type'],op['task_type'])
        if step.tool_name!=tool_name or case_semantic(step.arguments,profile,case.request)!=case_semantic(op,profile,case.request):
            differences.append({'scope':'executable_arguments','step':step.order,
                                'expected':case_semantic(op,profile,case.request),
                                'actual':case_semantic(step.arguments,profile,case.request)})
    return differences


def answer_clarification(case,turn,already):
    if not case.reply or already:return None
    field=case.reply['field'];op=turn.envelope.spec.model_dump(mode='json') if turn.envelope else {}
    operations=op['operations'] if op.get('task_type')=='composite' else [op]
    missing=any(field in x.get('missing_fields',[]) or not x.get(field) for x in operations)
    if not missing:return None
    return case.reply['answer']


def check_delivery(case,input_path,output,service,task_id,executed):
    expected_book=copy.deepcopy(case.checks['workbook']);actual_book=snapshot(output)
    explicit_flag_name=any(
        cue in case.request for cue in ('标记列叫','重复标记列叫','标记列名','重复标记列名')
    )
    declared=case.spec.get('operations',[]) if case.spec.get('task_type')=='composite' else [case.spec]
    if not explicit_flag_name:
        for op in declared:
            if op.get('task_type')!='clean' or op.get('operation')!='mark_duplicates':continue
            sheet=op.get('output_sheet')
            if sheet not in expected_book or sheet not in actual_book:continue
            header=expected_book[sheet][0];flag=op.get('duplicate_flag_column')
            if flag in header:
                position=header.index(flag)
                expected_book[sheet][0][position]='__DUPLICATE_FLAG__'
                if position<len(actual_book[sheet][0]):
                    actual_book[sheet][0][position]='__DUPLICATE_FLAG__'
    diffs=compare(expected_book,actual_book,case.checks['tolerance'])
    diffs+=unchanged_metadata(input_path,output,case.checks['workbook'])
    task=service.repository.get_task(task_id)
    versions=service.repository.list_versions(task_id)
    version=next((v for v in versions if v['version']==task['current_version']),None)
    if not version or sha(Path(version['path']))!=sha(output) or str(task['current_path'])!=str(output):
        diffs.append({'scope':'delivery_version'})
    expected_csv=case.checks['csv_artifacts']
    artifacts=list((service.files.task_dir(task_id)/'artifacts').glob('*.csv'))
    if expected_csv:
        if len(artifacts)!=len(expected_csv):diffs.append({'scope':'csv_artifact_count','expected':len(expected_csv),'actual':len(artifacts)})
        for expected,p in zip(expected_csv,sorted(artifacts)):
            if not p.read_bytes().startswith(b'\xef\xbb\xbf'):diffs.append({'scope':'csv_bom'})
            with p.open(encoding='utf-8-sig',newline='') as f:actual=list(csv.reader(f))
            # CSV output is textual; no missing/zero coercion on workbook comparisons.
            text=[[str(v) if v is not None else '' for v in row] for row in expected]
            diffs+=compare({'CSV':text},{'CSV':actual})
    return diffs


def worker(case_path: Path,folder: Path,mode: str,call_limit: int):
    from sheetpilot.config import Settings
    from sheetpilot.service import SheetPilotService
    from sheetpilot.schemas import MessageKind
    case=Case.model_validate_json(case_path.read_text(encoding='utf-8'));folder.mkdir(parents=True,exist_ok=True)
    result={'case_id':case.id,'kind':case.kind,'capability':case.capability,'mode':mode,
            'status':'tool_error','stage':'setup','severity':'normal','calls':0,'differences':[],
            'formula_text':'not_checked' if case.capability=='formula' else None,
            'formula_calculation':'not_verified' if case.capability=='formula' else None,
            'reproduce':f'{sys.executable} -m evals.phase1 run --mode {mode} --ids {case.id}'}
    result['input_sha256']=case.input_sha256
    start=time.monotonic();service=None;client=None;task_id=None;original=None;calls=None
    try:
        input_path=case_path.parent/case.input
        if sha(input_path)!=case.input_sha256:raise ValueError('case input checksum mismatch')
        settings=Settings.load(APP)
        if mode=='live' and not settings.llm_configured:
            result.update(status='blocked',stage='configuration',reason='model configuration missing');return result
        settings=replace(settings,data_dir=folder/'workspace',database_path=folder/'workspace/eval.sqlite3')
        result['model']=settings.model if mode=='live' else 'explicit-offline-double'
        client,calls=attach_client(settings,case,mode,folder,min(call_limit,case.max_calls))
        service=SheetPilotService(settings,client)
        profile=service.create_task(input_path.name,input_path.read_bytes());task_id=service.task_id_from_profile(profile)
        original=Path(service.repository.get_task(task_id)['original_path'])
        original_hash=sha(original);baseline_versions=len(service.repository.list_versions(task_id))
        result['stage']='planning'
        turn=service.submit_message(task_id,case.request)
        if case.fault in ('timeout','invalid_response'):
            task=service.repository.get_task(task_id)
            safe=turn.reply.kind==MessageKind.ERROR and task['status']=='failed' and len(service.repository.list_versions(task_id))==baseline_versions
            result.update(status='passed' if safe else 'failed',reason='injected transport failure safely handled' if safe else 'failure not safely handled')
            return result
        if turn.reply.kind==MessageKind.ERROR:
            unavailable=calls and calls.events and all(e['status']=='transport_error' for e in calls.events)
            result.update(status='blocked' if unavailable else 'failed',stage='environment' if unavailable else 'planning',
                          reason='model transport unavailable' if unavailable else 'product returned planning error');return result
        if case.kind=='refusal':
            safe=turn.refusal is not None and turn.plan is None and len(service.repository.list_versions(task_id))==baseline_versions
            result.update(status='passed' if safe else 'failed',reason='explicit refusal' if safe else 'unsupported request not refused')
            return result
        if case.revision:
            if not turn.plan:
                result.update(status='failed',reason='revision case did not reach initial plan');return result
            old_revision=turn.plan.revision;turn=service.submit_message(task_id,case.revision)
            result['initial_revision']=old_revision
            if turn.plan and turn.plan.revision<=old_revision:
                result.update(status='failed',reason='revision did not invalidate old plan');return result
        answered=False;seen=set();turns=0
        if case.kind=='clarification' and not case.revision and not turn.needs_clarification:
            result.update(status='failed',reason='required clarification was skipped');return result
        while turn.needs_clarification:
            signature=json.dumps(turn.envelope.spec.model_dump(mode='json'),sort_keys=True)
            if signature in seen or turns>=case.max_clarifications:
                result.update(status='failed',reason='repeated clarification or round limit');return result
            seen.add(signature);answer=answer_clarification(case,turn,answered)
            if answer is None:
                result.update(status='needs_review',reason='no declared answer for this clarification');return result
            turn=service.submit_message(task_id,answer);answered=True;turns+=1
        result['clarification_rounds']=turns
        if turn.reply.kind==MessageKind.ERROR:
            unavailable=calls and calls.events and all(e['status']=='transport_error' for e in calls.events)
            result.update(status='blocked' if unavailable else 'failed',stage='environment' if unavailable else 'planning',
                          reason='model transport unavailable' if unavailable else 'product returned planning error');return result
        result['stage']='approval'
        diffs=guard_plan(case,turn,profile)
        if diffs:
            result.update(status='failed',reason='plan does not match authorized task; not approved',differences=diffs);return result
        # Independently reconstruct the executable plan from the actual envelope
        # and verify no tool arguments were substituted after planning.
        expected_plan,validation=service.make_task_plan(task_id,turn.envelope,policy_source_text=case.request+'\n'+(case.reply or {}).get('answer','')+'\n'+(case.revision or ''))
        def executable(p):return [(s.tool_name,s.arguments,s.expected_changes.model_dump(mode='json')) for s in p.steps]
        if not validation.valid or executable(expected_plan)!=executable(turn.plan):
            result.update(status='failed',reason='executable plan differs from declared rules');return result
        if case.fault=='stale_plan':service.repository.update_task(task_id,current_version=1)
        if case.fault=='step_failure':
            tool=service.tools.get(turn.plan.steps[1].tool_name)
            def fail(*args,**kwargs):raise RuntimeError('PHASE1_INJECTED_SECOND_STEP_FAILURE')
            tool.execute=fail
        result['stage']='execution'
        try:executed=service.approve_current_plan(task_id,case.acknowledgements)
        except Exception as exc:
            if case.fault in ('step_failure','stale_plan'):
                run=service.repository.latest_plan_run(task_id)
                safe=len(service.repository.list_versions(task_id))==baseline_versions and sha(original)==original_hash
                if case.fault=='step_failure':safe=safe and bool(run and run['current_step']>=1)
                result.update(status='passed' if safe else 'failed',reason='injected failure blocked publication',error_type=type(exc).__name__,run=run)
                return result
            result.update(status='failed',reason=str(exc),error_type=type(exc).__name__);return result
        if case.kind=='fault':
            result.update(status='failed',reason='injected protection did not stop execution');return result
        result['stage']='checking'
        output=Path(executed.output_path)
        diffs=check_delivery(case,input_path,output,service,task_id,executed)
        result.update(status='failed' if diffs else 'passed',differences=diffs,output=str(output),
                      reason='independent output checks',result_version=executed.version)
        if case.capability=='formula':result['formula_text']='checked'
        return result
    except Exception as exc:
        result.update(status='tool_error',error_type=type(exc).__name__,reason=str(exc),traceback=traceback.format_exc())
        return result
    finally:
        if service and task_id:
            try:
                dump(folder/'conversation.json',[x.model_dump(mode='json') for x in service.get_conversation(task_id)])
                task=service.repository.get_task(task_id)
                dump(folder/'plan.json',json.loads(task['plan_json']) if task.get('plan_json') else None)
                dump(folder/'state.json',{'task':task,'versions':service.repository.list_versions(task_id)})
                if original and sha(original)!=case.input_sha256:
                    result.update(status='failed',severity='critical',reason='original input modified')
            except Exception as exc:result.update(status='tool_error',reason=f'artifact capture failed: {exc}')
        result['calls']=len(calls.events) if calls else 0
        result['elapsed_seconds']=round(time.monotonic()-start,3)
        if result.get('differences'):result['severity']='critical' if any(x.get('scope') in ('delivery_version','layout','style') for x in result['differences']) else 'normal'
        dump(folder/'result.json',result)
        if client and hasattr(client,'close'):client.close()
        if service:service._background_executor.shutdown(wait=False,cancel_futures=True)


STATUS_ZH={'passed':'通过','failed':'失败','blocked':'环境或次数限制导致没测成',
           'needs_review':'需要人检查','tool_error':'测试工具出错'}
KIND_ZH={'normal':'正常任务','clarification':'追问或修改','refusal':'应拒绝的要求','fault':'故意制造错误'}
CAPABILITY_ZH={'lookup':'查找匹配','normalize':'文本整理','aggregate':'分组汇总','sort_rank':'排序排名',
               'clean':'筛选去重','formula':'公式','conversion':'文件转换','transform':'数值与字段处理',
               'boundary':'拒绝不支持的要求','protection':'故障保护'}
SOURCE_ZH={'synthetic':'按规则生成','research_reconstructed':'根据调研问题重建表格',
           'spreadsheetbench_adapted':'使用公开原文件'}
STAGE_ZH={'setup':'准备文件','planning':'理解要求、生成计划','approval':'确认计划前',
          'execution':'处理文件','checking':'检查结果','configuration':'读取模型配置',
          'environment':'连接模型服务','batch_limit':'等待运行','runner_timeout':'等待超时','worker_exit':'测试进程退出'}


def readable_reason(result):
    """Explain outcomes to a reader; keep original machine records unchanged."""
    reason=result.get('reason','')
    if reason=='independent output checks' and result.get('status')=='failed':
        differences=result.get('differences',[])
        if differences:
            first=differences[0]
            if all(k in first for k in ('sheet','row','column')):
                return f"结果文件里“{first['sheet']}”第{first['row']}行、第{first['column']}列应为{first.get('expected')!r}，实际为{first.get('actual')!r}。完整差异见该例result.json。"
            if first.get('scope')=='layout':
                return f"产品处理文件后，原工作表“{first['sheet']}”的{'、'.join(first.get('changed',['布局设置']))}改变了；案例只允许修改指定结果区域。具体前后值见该例result.json。"
            if first.get('scope')=='style':
                return f"产品处理文件后，原工作表“{first['sheet']}”的{first['cell']}格式改变了；具体差异见该例result.json。"
            if first.get('scope')=='delivery_version':return '结果文件与产品记录的交付版本不一致，具体路径和文件校验值见该例result.json。'
    if reason=='plan does not match authorized task; not approved':
        notes=[]
        def explain(a,b):
            if not isinstance(a,dict) or not isinstance(b,dict):return
            if a.get('output_sheet')!=b.get('output_sheet'):
                notes.append(f"结果表应叫“{a.get('output_sheet')}”，计划里却叫“{b.get('output_sheet')}”")
            for left,right in zip(a.get('metrics',[]),b.get('metrics',[])):
                if left.get('output_column')!=right.get('output_column'):
                    notes.append(f"结果列应叫“{left.get('output_column')}”，计划里却叫“{right.get('output_column')}”")
            for left,right in zip(a.get('map_rules',[]),b.get('map_rules',[])):
                extra={k:v for k,v in right.get('mapping',{}).items() if k not in left.get('mapping',{})}
                if extra:notes.append('程序多加了未要求的替换：'+ '、'.join(f'“{k}”改成{v}' for k,v in extra.items()))
            if a.get('constraints')!=b.get('constraints'):
                notes.append('原检查器认为两份操作说明不同；需查看复查说明，可能只是中英文说法不同')
            if a.get('task_type')=='composite' and b.get('task_type')=='composite':
                for left,right in zip(a['operations'],b['operations']):explain(left,right)
        for diff in result.get('differences',[]):explain(diff.get('expected'),diff.get('actual'))
        notes=list(dict.fromkeys(notes))
        return ('；'.join(notes) if notes else '计划中的表、列、规则或步骤与案例要求不同，具体差异见该例result.json')+'。测试在确认前停止，没有执行本次文件修改。'
    messages={
        'independent output checks':'结果文件通过了另外计算的答案检查。' if result['status']=='passed' else '结果文件与正确答案不同，具体单元格或格式差异见该例result.json。',
        'explicit refusal':'产品明确拒绝了当前不支持的要求，没有生成结果文件。',
        'unsupported request not refused':f'产品{result.get("observed_action","没有明确拒绝")}，没有按案例要求拒绝。测试没有确认执行，也没有交付新文件。',
        'no declared answer for this clarification':'产品问了脚本没准备答案的问题，测试停止，尚未执行文件修改，需要人看对话。',
        'required clarification was skipped':'案例要求产品先追问，但产品直接给了计划；需确认是真漏问，还是案例本来不需要追问。',
        'revision case did not reach initial plan':'原本要在计划生成后修改要求，但第一份计划就没有生成，后面的修改没测到。',
        'revision did not invalidate old plan':'用户改了要求，计划编号却没有更新，需要检查旧计划是否仍然能执行。',
        'repeated clarification or round limit':'产品重复问同一个问题，或超过允许的问答轮数，测试停止。',
        'product returned planning error':'产品生成计划时出错，没有执行文件修改；查看对话和错误日志定位原因。',
        'model configuration missing':'没有完整模型配置，未调用模型。',
        'model transport unavailable':'没有连上模型服务，无法判断模型回答是否正确。',
        'injected transport failure safely handled':'测试故意让模型超时或返回无法读取的内容；产品提示失败，没有生成新的正式结果文件。',
        'failure not safely handled':'测试故意让模型请求失败，产品随后表现不符合预期，需要查看状态记录。',
        'injected failure blocked publication':'测试故意让操作出错或让计划对应的版本失效；产品停止，没有生成新的正式结果文件。',
        'injected protection did not stop execution':'故意制造了应当停止的情况，但产品仍继续执行。',
        'executable plan differs from declared rules':'准备执行的步骤与计划中的规则对不上，测试没有确认执行。',
        'harness killed timed-out worker; not product cancellation':'等待超过测试上限，工具结束了这一例；这不表示产品的取消功能正常。',
        'original input modified':'原始上传文件被改动，需要优先排查。',
        'case input checksum mismatch':'输入文件与案例记录的文件不一致，测试工具停止运行。',
        'call_limit':'本批请求次数已用完，这一例没有运行。',
        'batch_timeout':'本批等待时间已用完，这一例没有运行。',
    }
    return messages.get(reason,'测试停止，具体错误为：'+str(reason))


def summarize(run_dir: Path):
    meta=json.loads((run_dir/'run.json').read_text(encoding='utf-8'))
    results=[json.loads(p.read_text(encoding='utf-8')) for p in sorted(run_dir.glob('*/result.json'))]
    cases={}
    for path in run_dir.glob('*/case/case.json'):
        case=Case.model_validate_json(path.read_text(encoding='utf-8'))
        cases[case.id]=case
    for result in results:
        case=cases.get(result['case_id'])
        if case:
            result['request']=case.request
            result['source_type']=case.source['type']
            result['input_file_modified']=result.get('reason')=='original input modified'
            result['new_version_delivered']=bool(result.get('output'))
            if case.kind=='refusal' and result.get('status')=='failed':
                conversation_path=run_dir/case.id/'conversation.json'
                if conversation_path.is_file():
                    conversation=json.loads(conversation_path.read_text(encoding='utf-8'))
                    actions=[message.get('kind') for message in conversation if message.get('role')=='assistant']
                    result['observed_action']='生成了处理计划' if 'plan' in actions else '提出了追问' if 'clarification' in actions else '给了其他回复'
    review_path=run_dir/'review.json'
    reviews=json.loads(review_path.read_text(encoding='utf-8')) if review_path.exists() else {}
    for r in results:
        if r['case_id'] in reviews:
            r['reviewed_status']=reviews[r['case_id']]['status'];r['review_reason']=reviews[r['case_id']]['reason']
    with (run_dir/'results.jsonl').open('w',encoding='utf-8') as f:
        for r in results:f.write(json.dumps(r,ensure_ascii=False,default=str)+'\n')
    fields=['case_id','kind','capability','source_type','mode','status','reviewed_status','review_reason','stage','severity','calls','elapsed_seconds','request','new_version_delivered','reason','reproduce']
    with (run_dir/'summary.csv').open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.DictWriter(f,fieldnames=fields,extrasaction='ignore');w.writeheader();w.writerows(results)
    def counts(values):return '、'.join(f'{STATUS_ZH.get(k,k)}{v}例' for k,v in Counter(values).items())
    mode='调用真实模型' if meta['mode']=='live' else '不调用模型，直接使用预先写好的操作规则'
    stop={'completed':'已运行完所选案例','call_limit':'请求次数已用完','batch_timeout':'等待时间已用完',
          'stopped_after_transport_failures':'多次没有取得网络响应，已停止这一批'}
    lines=['# 第一轮逐例测试结果','',f'测试方式：{mode}。','',
           f"选了{meta['selected_count']}例，保存了{len(results)}例结果。原始记录：{counts(r['status'] for r in results)}。",
           '本批怎样结束：'+stop.get(meta.get('stop_reason','completed'),meta.get('stop_reason',''))+'。']
    if meta['mode']=='live':lines+=['',f"共记录{sum(r.get('calls',0) for r in results)}次请求尝试，包括未取得响应的请求，不等于这些请求都已到达模型服务或已计费。"]
    else:lines+=['','这里不检查真实模型能否读懂用户的话。故意模拟模型超时的日志也不是云端请求。']
    lines+=['','本批公式只检查公式文字，没有检查实际计算值。','',
            '|案例|用户要求|原始结果|停在哪一步|程序实际做了什么|有没有交付新文件版本|','|---|---|---|---|---|---|']
    for r in results:
        reason=readable_reason(r).replace('|','/').replace('\n',' ')
        request=r.get('request','见该例case.json').replace('|','/').replace('\n',' ')
        delivered='有' if r.get('new_version_delivered') else '没有'
        lines.append(f"|{r['case_id']}|{request}|{STATUS_ZH.get(r['status'],r['status'])}|{STAGE_ZH.get(r.get('stage'),r.get('stage','未记录'))}|{reason}|{delivered}|")
    if reviews:
        lines+=['','## 复查后怎样看这些结果','',f"复查结果：{counts(r.get('reviewed_status',r['status']) for r in results)}。上表和原始result.json仍保留第一次判断，下面说明哪里需要纠正。",'']
        for ident,review in reviews.items():lines.append(f"- {ident}：{STATUS_ZH.get(review['status'],review['status'])}。{review['reason']}")
    (run_dir/'report.md').write_text('\n'.join(lines)+'\n',encoding='utf-8')
    return results


def run_suite(out: Path,mode,ids=None,kinds=None,max_tasks=20,max_calls=60,batch_seconds=1800,case_seconds=300,rerun=False):
    from datetime import datetime
    max_tasks=min(max_tasks,20);max_calls=min(max_calls,60)
    batch_seconds=min(batch_seconds,1800);case_seconds=min(case_seconds,300)
    out=protected_output(out);cases=load_cases(out/'cases')
    cases=[c for c in cases if (not ids or c.id in ids) and (not kinds or c.kind in kinds) and mode in c.modes]
    if ids and set(ids)-{c.id for c in cases}:raise ValueError('unknown or excluded case ID')
    if not cases:raise ValueError('no selected cases')
    if mode=='live':
        previous={} if rerun else completed_cases(out,'live')
        cases=[c for c in cases if previous.get(c.id)!=case_signature(c)]
        # Interleave capabilities so early batches do not test just one family.
        buckets={}
        for c in cases:buckets.setdefault((c.kind,c.capability),[]).append(c)
        for group in buckets.values():group.sort(key=lambda c:c.id)
        ordered=[]
        while any(buckets.values()):
            for key in sorted(buckets):
                if buckets[key]:ordered.append(buckets[key].pop(0))
        cases=ordered[:min(max_tasks,20)]
        if not cases:
            summarize_suite(out)
            return None
    run_dir=out/'runs'/(datetime.now().strftime('%Y%m%d-%H%M%S-%f')+'-'+mode);run_dir.mkdir(parents=True)
    try:revision=subprocess.run(['git','rev-parse','HEAD'],cwd=APP,capture_output=True,text=True).stdout.strip() or 'unavailable (not a git checkout)'
    except OSError:revision='unavailable'
    import hashlib
    code_hash=hashlib.sha256()
    for p in sorted([*(APP/'src/sheetpilot').rglob('*.py'),*(APP/'evals/phase1').glob('*.py')]):
        code_hash.update(str(p.relative_to(APP)).encode());code_hash.update(p.read_bytes())
    meta={'mode':mode,'selected_count':len(cases),'selected_ids':[c.id for c in cases],'case_version':'1.0',
          'code_sha256':code_hash.hexdigest(),
          'code_revision':revision,'max_calls':min(max_calls,60) if mode=='live' else 0,
          'max_tasks':min(max_tasks,20) if mode=='live' else len(cases),'batch_seconds':min(batch_seconds,1800)}
    dump(run_dir/'run.json',meta);used=0;started=time.monotonic()
    for case in cases:
        if mode=='live' and used>=min(max_calls,60):meta['stop_reason']='call_limit';break
        remaining=min(batch_seconds,1800)-(time.monotonic()-started)
        if remaining<=0:meta['stop_reason']='batch_timeout';break
        folder=run_dir/case.id;folder.mkdir()
        frozen=folder/'case';frozen.mkdir()
        import shutil
        source=out/'cases'/case.id/case.input
        if sha(source)!=case.input_sha256:raise ValueError(f'{case.id}: source changed during run')
        shutil.copyfile(source,frozen/case.input);dump(frozen/'case.json',case.model_dump(mode='json'))
        cmd=[sys.executable,'-m','evals.phase1','worker','--case',str(frozen/'case.json'),
             '--folder',str(folder),'--mode',mode,'--call-limit',str(min(case.max_calls,max_calls-used) if mode=='live' else 6)]
        env=dict(os.environ,PYTHONIOENCODING='utf-8')
        with (folder/'process.log').open('w',encoding='utf-8') as log:
            proc=subprocess.Popen(cmd,cwd=APP,env=env,stdout=log,stderr=subprocess.STDOUT)
            try:proc.wait(timeout=min(case_seconds,case.max_seconds,remaining))
            except subprocess.TimeoutExpired:
                proc.kill();proc.wait()
                dump(folder/'result.json',{'case_id':case.id,'kind':case.kind,'capability':case.capability,'mode':mode,
                     'status':'failed','stage':'runner_timeout','severity':'normal','reason':'harness killed timed-out worker; not product cancellation',
                     'reproduce':f'{sys.executable} -m evals.phase1 run --mode {mode} --ids {case.id}'})
        if not (folder/'result.json').exists():dump(folder/'result.json',{'case_id':case.id,'kind':case.kind,'mode':mode,'status':'tool_error','stage':'worker_exit','reason':f'worker exited {proc.returncode}'})
        events=json.loads((folder/'calls.json').read_text(encoding='utf-8')) if (folder/'calls.json').exists() else []
        actual_calls=len(events);used+=actual_calls if mode=='live' else 0
        result=json.loads((folder/'result.json').read_text(encoding='utf-8'));result['calls']=actual_calls;dump(folder/'result.json',result)
        print(f"{case.id}：{STATUS_ZH.get(result['status'],result['status'])}（请求尝试{actual_calls}次）",flush=True)
    meta.setdefault('stop_reason','completed');meta['actual_calls']=used;dump(run_dir/'run.json',meta)
    for c in cases:
        if not (run_dir/c.id/'result.json').exists():
            dump(run_dir/c.id/'result.json',{'case_id':c.id,'kind':c.kind,'capability':c.capability,'mode':mode,
                 'status':'blocked','stage':'batch_limit','reason':meta['stop_reason'],'calls':0})
    summarize(run_dir);summarize_suite(out);return run_dir


def case_signature(case):
    import hashlib
    return hashlib.sha256(json.dumps(case.model_dump(mode='json'),sort_keys=True,ensure_ascii=False).encode('utf-8')).hexdigest()


def completed_cases(out,mode):
    completed={}
    for run in sorted((out/'runs').glob(f'*-{mode}')):
        for result_path in run.glob('*/result.json'):
            try:
                result=json.loads(result_path.read_text(encoding='utf-8'))
                frozen=result_path.parent/'case/case.json'
                if result.get('stage')=='batch_limit' or (result.get('status')=='blocked' and result.get('stage') in ('configuration','environment')) or not frozen.is_file():continue
                case=Case.model_validate_json(frozen.read_text(encoding='utf-8'))
                completed[case.id]=case_signature(case)
            except (ValueError,KeyError):continue
    return completed


def summarize_suite(out):
    """Latest attempted result for each case; never count unrun batch placeholders."""
    out=protected_output(out);cases=load_cases(out/'cases')
    latest={};current={c.id:case_signature(c) for c in cases}
    for run in sorted((out/'runs').iterdir()) if (out/'runs').exists() else []:
        if not run.is_dir() or not (run/'run.json').is_file():continue
        meta=json.loads((run/'run.json').read_text(encoding='utf-8'))
        for path in sorted(run.glob('*/result.json')):
            result=json.loads(path.read_text(encoding='utf-8'))
            if result.get('stage')=='batch_limit':continue
            frozen=path.parent/'case/case.json'
            if not frozen.is_file():continue
            case=Case.model_validate_json(frozen.read_text(encoding='utf-8'))
            if current.get(case.id)!=case_signature(case):continue
            result['run_dir']=str(run);result['source_type']=case.source['type']
            latest[(meta['mode'],case.id)]=result
    summary=out/'suite-summary.json'
    counts={}
    for (mode,_),r in latest.items():counts.setdefault(mode,Counter())[r['status']]+=1
    finished_live=completed_cases(out,'live')
    payload={'total_cases':len(cases),'by_mode':{mode:dict(v) for mode,v in counts.items()},
             'live_remaining':sum('live' in c.modes and finished_live.get(c.id)!=case_signature(c) for c in cases),
             'results':list(latest.values())}
    dump(summary,payload)
    with (out/'suite-summary.csv').open('w',encoding='utf-8-sig',newline='') as f:
        fields=['mode','case_id','kind','capability','source_type','status','stage','calls','reason','run_dir']
        writer=csv.DictWriter(f,fieldnames=fields,extrasaction='ignore');writer.writeheader();writer.writerows(payload['results'])
    lines=['# 第一阶段累计测试结果','',f'案例总数：{len(cases)}。','']
    for mode in ('offline','live'):
        subset=[r for r in latest.values() if r['mode']==mode]
        lines.append(f'## {"表格工具测试" if mode=="offline" else "完整流程测试"}')
        lines.append('')
        lines.append(f'已实际运行{len(subset)}例。'+('剩余'+str(payload['live_remaining'])+'例待运行。' if mode=='live' else ''))
        for label,field in [('结果','status'),('能力','capability'),('来源','source_type'),('停止阶段','stage')]:
            values=Counter(str(r.get(field,'未知')) for r in subset)
            names={'status':STATUS_ZH,'capability':CAPABILITY_ZH,'source_type':SOURCE_ZH,'stage':STAGE_ZH}[field]
            lines.append(f'{label}：'+('；'.join(f'{names.get(k,k)} {v}例' for k,v in sorted(values.items())) if values else '尚无结果')+'。')
        lines.append('')
    (out/'suite-report.md').write_text('\n'.join(lines),encoding='utf-8')
    return payload
