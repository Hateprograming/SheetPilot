from __future__ import annotations

import argparse
import json
import os
import socket
import subprocess
import sys
import time
from dataclasses import replace
from pathlib import Path

APP=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(APP/'src'))
from .common import Case,dump,protected_output


def serve(case_path,folder,port):
    # Set isolation before importing web_api, which constructs its service at import.
    os.environ['SHEETPILOT_WEB_DATA_DIR']=str(folder/'workspace')
    from sheetpilot import web_api
    from sheetpilot.service import SheetPilotService
    from .runner import ScriptedClient
    case=Case.model_validate_json(case_path.read_text(encoding='utf-8'))
    web_api.service._background_executor.shutdown(wait=False)
    web_api.settings=replace(web_api.settings,api_key='',base_url='',model='phase1-browser-double')
    web_api.service=SheetPilotService(web_api.settings,ScriptedClient(case))
    import uvicorn
    uvicorn.run(web_api.app,host='127.0.0.1',port=port)


def page_test(case_path,folder,port):
    from playwright.sync_api import sync_playwright
    from sheetpilot.schemas import TaskEnvelope,TaskPlan,WorkbookProfile,AgentTurnResult,ConversationMessage,ConversationRole,MessageKind,ConversationState
    from .runner import guard_plan
    from .oracle import snapshot,compare,unchanged_metadata
    case=Case.model_validate_json(case_path.read_text(encoding='utf-8'))
    result={'case_id':case.id,'mode':'browser-offline','status':'tool_error','separate_from_service_counts':True}
    errors=[];started=time.monotonic()
    with sync_playwright() as p:
        browser=p.chromium.launch(headless=True)
        context=browser.new_context(accept_downloads=True,viewport={'width':1440,'height':1000})
        page=context.new_page();page.set_default_timeout(45000)
        page.on('pageerror',lambda e:errors.append(str(e)))
        try:
            page.goto(f'http://127.0.0.1:{port}',wait_until='networkidle')
            page.locator('input[type=file]').set_input_files(str(case_path.parent/case.input))
            check=page.locator('input[type=checkbox]');assert not check.is_checked();check.check()
            page.get_by_role('button',name='分析这个工作簿').click()
            page.get_by_label('任务描述').fill(case.request)
            page.get_by_role('button',name='发送',exact=True).click()
            page.locator('.plan-block').wait_for()
            page.screenshot(path=str(folder/'plan.png'),full_page=True)
            task_id=page.url.rsplit('/',1)[-1]
            workspace=context.request.get(f'http://127.0.0.1:{port}/api/workbooks/{task_id}/workspace').json()
            dump(folder/'planned-workspace.json',workspace)
            metadata=next(m['metadata'] for m in reversed(workspace['messages']) if m['kind']=='plan')
            env=TaskEnvelope.model_validate(metadata['envelope'])
            plan=dict(metadata['plan'],input_sha256=case.input_sha256,base_sha256=case.input_sha256)
            profile=dict(workspace['profile'],sha256=case.input_sha256)
            turn=AgentTurnResult(state=ConversationState.PLAN_READY,envelope=env,plan=TaskPlan.model_validate(plan),
                                 reply=ConversationMessage(task_id=task_id,role=ConversationRole.ASSISTANT,kind=MessageKind.PLAN,content='browser'))
            diffs=guard_plan(case,turn,WorkbookProfile.model_validate(profile))
            assert not diffs,f'plan blocked: {diffs}'
            page.get_by_role('button',name='确认并执行',exact=True).click()
            block=page.get_by_label('本次处理结果')
            block.get_by_role('heading',name='本次结果 · v1').wait_for()
            if case.id=='L1':
                block.get_by_text('按类型查看异常并定位原始行',exact=True).click()
                block.get_by_role('button',name='定位该行').first.click()
                page.wait_for_function("document.querySelector('.preview-table-block tbody tr td')?.textContent === '4'")
            with page.expect_download() as download:block.get_by_role('link',name='下载本次结果 v1',exact=True).click()
            target=folder/'download.xlsx';download.value.save_as(target)
            diffs=compare(case.checks['workbook'],snapshot(target))
            diffs+=unchanged_metadata(case_path.parent/case.input,target,case.checks['workbook'])
            assert not diffs,f'output mismatch: {diffs[:5]}'
            assert not errors,errors
            result.update(status='passed',reason='real backend and browser, explicit model substitute',differences=diffs)
        except Exception as exc:
            result.update(status='failed',reason=str(exc),error_type=type(exc).__name__)
        finally:
            page.screenshot(path=str(folder/'result.png'),full_page=True)
            dump(folder/'console.json',errors);result['elapsed_seconds']=time.monotonic()-started;dump(folder/'result.json',result)
            context.close();browser.close()
    return 0 if result['status']=='passed' else 1


def run_browser(out):
    from datetime import datetime
    root=out/'browser'/datetime.now().strftime('%Y%m%d-%H%M%S');root.mkdir(parents=True)
    # Direct Popen makes Windows teardown reliable and preserves server stderr.
    # The skill helper uses shell=True, whose cmd child is not the server process.
    results=[]
    for ident in ['L1','N1','A1','C2']:
        folder=root/ident;folder.mkdir();case=out/'cases'/ident/'case.json'
        with socket.socket() as s:s.bind(('127.0.0.1',0));port=s.getsockname()[1]
        server_cmd=[sys.executable,'-m','evals.phase1.browser','serve','--case',str(case),'--folder',str(folder),'--port',str(port)]
        cmd=[sys.executable,'-m','evals.phase1.browser','page','--case',str(case),'--folder',str(folder),'--port',str(port)]
        with (folder/'process.log').open('w',encoding='utf-8') as log:
            server=subprocess.Popen(server_cmd,cwd=APP,stdout=log,stderr=subprocess.STDOUT,
                                    env=dict(os.environ,PYTHONIOENCODING='utf-8'),creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
            try:
                deadline=time.monotonic()+40;ready=False
                while time.monotonic()<deadline and server.poll() is None:
                    try:
                        with socket.create_connection(('127.0.0.1',port),timeout=.5):ready=True;break
                    except OSError:time.sleep(.2)
                if ready:subprocess.run(cmd,cwd=APP,stdout=log,stderr=subprocess.STDOUT,timeout=120,env=dict(os.environ,PYTHONIOENCODING='utf-8'))
            except subprocess.TimeoutExpired:pass
            finally:
                server.terminate()
                try:server.wait(timeout=5)
                except subprocess.TimeoutExpired:server.kill();server.wait()
        if not (folder/'result.json').exists():dump(folder/'result.json',{'case_id':ident,'status':'blocked','reason':'browser/server unavailable; see process.log','mode':'browser-offline'})
        r=json.loads((folder/'result.json').read_text(encoding='utf-8'));results.append(r);print(f"browser {ident}: {r['status']}",flush=True)
    dump(root/'summary.json',results)
    print(root)
    return 0 if all(r['status']=='passed' for r in results) else 1


if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('action',choices=['serve','page']);p.add_argument('--case',type=Path,required=True)
    p.add_argument('--folder',type=Path,required=True);p.add_argument('--port',type=int,required=True);a=p.parse_args()
    protected_output(a.folder)
    if a.action=='serve':serve(a.case,a.folder,a.port)
    else:raise SystemExit(page_test(a.case,a.folder,a.port))
