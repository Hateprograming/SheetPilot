from pathlib import Path
from io import BytesIO
from copy import deepcopy
from openpyxl import Workbook
from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT/'ui-test-artifacts'/'p0'
OUT.mkdir(parents=True, exist_ok=True)


def main():
    wb=Workbook()
    sheet=wb.active; sheet.title='订单'
    sheet.append(['客户编号','订单额'])
    for i in range(1,56): sheet.append([str(i),i*10])
    source=wb.create_sheet('客户'); source.append(['客户编号','客户等级'])
    source.append(['1','A'])
    stream=BytesIO(); wb.save(stream); wb.close()
    with sync_playwright() as p:
        browser=p.chromium.launch(headless=True)
        context=browser.new_context(accept_downloads=True,viewport={'width':1440,'height':1000})
        page=context.new_page()
        errors=[]
        page.on('pageerror', lambda error: errors.append(str(error)))
        page.goto('http://127.0.0.1:8769',wait_until='networkidle')
        page.locator('input[type=file]').set_input_files({'name':'p0-sample.xlsx','mimeType':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','buffer':stream.getvalue()})
        assert not page.locator('input[type=checkbox]').is_checked()
        page.locator('input[type=checkbox]').check()
        page.get_by_role('button',name='分析这个工作簿').click()
        page.get_by_label('任务描述').fill('按客户编号，将客户等级匹配到订单')
        page.get_by_role('button',name='发送',exact=True).click()
        page.locator('.plan-block').wait_for(timeout=30000)
        page.screenshot(path=str(OUT/'plan.png'),full_page=True)
        # Inspect rendered buttons instead of relying on internal component state.
        buttons=page.locator('.plan-block button').all_text_contents()
        target=next(label for label in buttons if '执行' in label or '确认' in label)
        page.locator('.plan-block').get_by_role('button',name=target.strip(),exact=True).click()
        result=page.get_by_label('本次处理结果')
        result.get_by_role('heading',name='本次结果 · v1').wait_for(timeout=30000)
        workspace_url=page.url
        task_id=workspace_url.rsplit('/',1)[-1]
        snapshot=context.request.get(f'http://127.0.0.1:8769/api/workbooks/{task_id}/workspace').json()
        assert result.get_by_text('已完成以下检查',exact=True).is_visible()
        result.get_by_text('按类型查看异常并定位原始行',exact=True).click()
        result.get_by_text('共 54 条相关异常',exact=False).wait_for()
        result.get_by_role('button',name='下一页',exact=True).click()
        result.get_by_text('Excel 第 53 行',exact=False).wait_for()
        result.get_by_role('button',name='定位该行').first.click()
        page.wait_for_function("document.querySelector('.preview-table-block tbody tr td')?.textContent === '53'")
        result.get_by_text('这次是否完成目标？反馈与净省时',exact=True).click()
        result.get_by_label('目标完成情况',exact=True).select_option('completed')
        result.get_by_label('原方法完整耗时（分钟）',exact=True).fill('20')
        result.get_by_label('本次完整耗时（分钟，含描述、等待、复核与交付）',exact=True).fill('8')
        result.get_by_role('button',name='保存反馈',exact=True).click()
        result.get_by_text('净省时：12 分钟',exact=False).wait_for()
        with page.expect_download() as download:
            result.get_by_role('link',name='下载本次结果 v1',exact=True).click()
        assert download.value.suggested_filename.startswith('SheetPilot_v1_')
        download.value.save_as(OUT/'verified-result.xlsx')
        page.screenshot(path=str(OUT/'result-feedback.png'),full_page=True)
        # Every new request gets its own delivery state; old v1 is history.
        page.get_by_label('任务描述').fill('继续处理客户等级')
        page.get_by_role('button',name='发送',exact=True).click()
        page.wait_for_function("!document.querySelector('[aria-label=\"本次处理结果\"]')?.textContent.includes('下载本次结果 v1')")
        verify_additional_states(context,page,snapshot,workspace_url,task_id)
        assert not errors, errors
        print('Browser passed: explicit consent, planning, execution, paged exceptions, exact Excel row, version download, feedback, new-operation isolation.')
        context.close(); browser.close()


def verify_additional_states(context,page,snapshot,url,task_id):
    """Deterministic UI failure states; backend authorization is tested separately."""
    state={'workspace':deepcopy(snapshot),'expired':False,'decisions':[],'resumed':False}
    def workspace_route(route):
        if state['expired']:
            route.fulfill(status=404,json={'message':'工作簿不存在或已不可用。'})
        else:
            route.fulfill(json=state['workspace'])
    context.route(f'**/api/workbooks/{task_id}/workspace',workspace_route)
    def authorization_route(route):
        approved=route.request.post_data_json['approved']
        state['decisions'].append(approved)
        state['workspace']['data_authorization']=None
        state['workspace']['task']['status']='awaiting_approval' if approved else 'needs_clarification'
        state['workspace']['current_operation']['status']=state['workspace']['task']['status']
        route.fulfill(status=202,json={'status':'queued' if approved else 'denied'})
    context.route(f'**/api/workbooks/{task_id}/authorizations/*',authorization_route)
    def prepare_grant():
        state['workspace']['task']['status']='awaiting_data_authorization'
        state['workspace']['current_operation'].update(status='awaiting_data_authorization',summary=None,result_version=None,delivery_status='not_ready')
        state['workspace']['data_authorization']={'grant_id':'demo-grant','purpose':'合成异常值授权检查','count':1,'version':1,'rows':[{'sheet':'订单','excel_row':3,'target_values':{'尺寸':'12×?'}}]}
        page.goto(url,wait_until='networkidle')
    prepare_grant()
    page.get_by_label('异常数据授权').get_by_text('查看将发送的目标值与位置',exact=True).click()
    page.get_by_label('异常数据授权').get_by_text('尺寸：12×?',exact=True).wait_for()
    page.screenshot(path=str(OUT/'authorization.png'),full_page=True)
    page.get_by_role('button',name='不发送，手动指定规则',exact=True).click()
    page.get_by_label('异常数据授权').wait_for(state='detached')
    assert state['decisions']==[False]
    prepare_grant()
    page.get_by_role('button',name='允许本次发送并继续',exact=True).click()
    page.get_by_label('异常数据授权').wait_for(state='detached')
    assert state['decisions']==[False,True]
    state['workspace']['task']['status']='paused'
    state['workspace']['current_operation'].update(status='paused',can_resume=True)
    state['workspace']['active_run']={'status':'paused','error':{'safe_message':'达到执行时间预算，已在安全检查点暂停；可主动恢复。'}}
    def resume(route):
        state['resumed']=True
        state['workspace']['task']['status']='executing'
        route.fulfill(status=202,json={'status':'queued'})
    context.route(f'**/api/workbooks/{task_id}/run/resume',resume)
    page.goto(url,wait_until='networkidle')
    page.get_by_text('达到执行时间预算',exact=False).wait_for()
    page.screenshot(path=str(OUT/'timeout-recovery.png'),full_page=True)
    page.get_by_role('button',name='从检查点重试',exact=True).click()
    page.locator('.recovery-block').wait_for(state='detached')
    assert state['resumed']
    state['expired']=True
    page.goto(url,wait_until='networkidle')
    page.get_by_text('工作簿不存在或已不可用。',exact=True).wait_for()
    assert page.get_by_role('link',name='下载本次结果 v1',exact=True).count()==0
    print('Browser UI states passed: exact grant preview, deny, approve, safe timeout resume, expired-workspace access.')


if __name__=='__main__':
    main()
