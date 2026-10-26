from __future__ import annotations

import json
import time
from dataclasses import replace
from datetime import UTC, datetime, timedelta

import pytest

from sheetpilot.errors import IntentValidationError
from sheetpilot.p0 import FeedbackInput
from test_v3_service import make_service, create_task, composite_envelope


def execute(service, task_id):
    envelope = composite_envelope()
    plan, valid = service.make_task_plan(task_id, envelope)
    assert valid.valid
    service.approve(task_id, 'approve')
    return service.execute_task_plan(task_id, envelope, plan)


def test_delivery_is_bound_to_operation_and_rechecks_file(tmp_path, workbook_factory):
    service = make_service(tmp_path)
    task_id = create_task(service, workbook_factory())
    output, _, _ = execute(service, task_id)
    current = service.get_workspace(task_id)['current_operation']
    assert current['delivery_status'] == 'ready'
    assert current['result_version'] == 1
    assert len(current['summary']['steps']) == 3
    output.write_bytes(b'corrupt')
    assert service.get_workspace(task_id)['current_operation']['delivery_status'] == 'unavailable'
    service.p0.begin(task_id)
    assert service.get_workspace(task_id)['current_operation']['result_version'] is None


def test_budget_pauses_at_verified_boundary_and_can_resume(tmp_path, workbook_factory, monkeypatch):
    service = make_service(tmp_path)
    task_id = create_task(service, workbook_factory())
    envelope = composite_envelope()
    plan, _ = service.make_task_plan(task_id, envelope)
    service.approve(task_id, 'approve')
    clock = [0.0]
    monkeypatch.setattr('sheetpilot.service.time.monotonic', lambda: clock[0])
    tool = service.tools.get(plan.steps[0].tool_name)
    original = tool.execute
    def slow(*args):
        result = original(*args)
        clock[0] = 601
        return result
    monkeypatch.setattr(tool, 'execute', slow)
    with pytest.raises(IntentValidationError, match='预算'):
        service.execute_task_plan(task_id, envelope, plan)
    run = service.repository.latest_plan_run(task_id)
    assert run['status'] == 'paused' and run['current_step'] == 1
    assert service.repository.get_task(task_id)['current_version'] == 0
    op_id = service.p0.current(task_id)['operation_id']
    monkeypatch.setattr(tool, 'execute', original)
    service.resume_current_run(task_id)
    assert service.p0.current(task_id)['operation_id'] == op_id
    assert service.get_workspace(task_id)['current_operation']['delivery_status'] == 'ready'


def test_feedback_upsert_and_no_synthetic_time_saving(tmp_path, workbook_factory):
    service = make_service(tmp_path)
    task_id = create_task(service, workbook_factory())
    execute(service, task_id)
    op = service.p0.current(task_id)
    service.p0.save_feedback(task_id, op['operation_id'], FeedbackInput(outcome='partial'))
    assert service.get_workspace(task_id)['current_operation']['metrics']['net_saved_minutes'] is None
    service.p0.save_feedback(task_id, op['operation_id'], FeedbackInput(outcome='completed', baseline_minutes=20, actual_total_minutes=8))
    assert service.get_workspace(task_id)['current_operation']['metrics']['net_saved_minutes'] == 12
    with service.repository.connect() as db:
        assert db.execute('SELECT COUNT(*) FROM p0_feedback').fetchone()[0] == 1


def test_cleanup_tombstones_failures_and_retries(tmp_path, workbook_factory, monkeypatch):
    service = make_service(tmp_path)
    task_id = create_task(service, workbook_factory())
    op_id = service.p0.begin(task_id)
    service.p0.state(op_id, 'failed')
    original = __import__('shutil').rmtree
    def fail(*args, **kwargs):
        raise PermissionError('locked')
    monkeypatch.setattr('sheetpilot.p0.shutil.rmtree', fail)
    with pytest.raises(IntentValidationError, match='清理尚未完成'):
        service.delete_workbook(task_id)
    assert service.repository.get_task(task_id)['cleanup_status'] == 'pending'
    with pytest.raises(KeyError):
        service.get_download(task_id)
    monkeypatch.setattr('sheetpilot.p0.shutil.rmtree', original)
    service.p0.cleanup()
    assert service.repository.get_task(task_id) is None
    with service.repository.connect() as db:
        assert db.execute('SELECT COUNT(*) FROM p0_operations').fetchone()[0] == 0


def test_polling_does_not_renew_retention_but_download_does(tmp_path, workbook_factory):
    service = make_service(tmp_path)
    task_id = create_task(service, workbook_factory())
    with service.repository.connect() as db:
        db.execute("UPDATE tasks SET expires_at='2000-01-01T00:00:00+00:00' WHERE task_id=?", (task_id,))
    service.get_workspace(task_id)
    assert service.repository.get_task(task_id)['expires_at'].startswith('2000')
    service.get_download(task_id)
    assert datetime.fromisoformat(service.repository.get_task(task_id)['expires_at']) > datetime.now(UTC)+timedelta(days=6)


def test_planning_timeout_discards_late_result(tmp_path, workbook_factory):
    from sheetpilot.llm import FakeLLMClient
    service = make_service(tmp_path)
    task_id = create_task(service, workbook_factory())
    class Slow(FakeLLMClient):
        def parse_task(self, *args, **kwargs):
            time.sleep(.12)
            return composite_envelope()
    service.llm = Slow(composite_envelope())
    service.settings = replace(service.settings, planning_budget_seconds=.03)
    job = service.enqueue_message(task_id, '统一编号，然后汇总，最后排序')
    service._background_futures[job['job_id']].result(timeout=10)
    assert service.repository.get_task(task_id)['status'] == 'failed'
    assert service.p0.current(task_id)['status'] == 'failed'
    assert service.repository.get_task(task_id)['plan_json'] is None


def test_model_egress_cannot_use_profile_or_assistant_history_as_bypass(tmp_path, workbook_factory):
    from sheetpilot.llm import QwenCompatibleClient
    from sheetpilot.schemas import LookupTaskSpec
    from sheetpilot.p0 import authorized_model_rows
    service = make_service(tmp_path)
    task_id = create_task(service, workbook_factory())
    profile = service.load_profile(task_id)
    client = object.__new__(QwenCompatibleClient)
    captured = []
    client._complete_json = lambda system, user, schema: captured.append(json.loads(user)) or composite_envelope()
    client.parse_task('按客户编号汇总', profile, previous_envelope=composite_envelope(),
                      conversation=[{'role':'assistant','content':'SECRET_PATIENT_82'}, {'role':'user','content':'按编号汇总'}],
                      validation_errors=['SECRET_PATIENT_82'], workbook_context={'raw_values':['SECRET_PATIENT_82']})
    assert 'SECRET_PATIENT_82' not in json.dumps(captured[-1])
    assert captured[-1]['authorized_exception_rows'] == []
    token = authorized_model_rows.set([{'sheet':'订单','excel_row':2,'target_values':{'尺寸':'授权示例'}}])
    try:
        client.parse_task('确认', profile)
        assert captured[-1]['authorized_exception_rows'][0]['excel_row'] == 2
    finally:
        authorized_model_rows.reset(token)
    client.parse_lookup('匹配', profile)
    assert 'sample_values' not in json.dumps(captured[-1])


def test_authorization_is_bound_to_operation_and_exact_data(tmp_path, workbook_factory):
    from sheetpilot.schemas import ConversationRole, MessageKind
    service = make_service(tmp_path)
    task_id = create_task(service, workbook_factory())
    op_id = service.p0.begin(task_id)
    message = service._add_message(task_id, ConversationRole.USER, MessageKind.TEXT, '处理异常')
    service.p0.link(task_id, op_id, 'planning', 'test-job', message.id)
    service._add_message(task_id, ConversationRole.ASSISTANT, MessageKind.CLARIFICATION, '确认规则',
                         {'review':{'rows':[{'sheet':'订单','row':2,'raw_values':{'尺寸':'12*?'}}]}})
    service.p0.state(op_id, 'awaiting_data_authorization')
    preview = service.p0.authorization_preview(task_id)
    assert preview['rows'][0]['target_values'] == {'尺寸':'12*?'}
    with pytest.raises(IntentValidationError):
        service.p0.decide_authorization(task_id, 'other-grant', True)
    service.p0.decide_authorization(task_id, preview['grant_id'], True)
    rows = service.p0.review_rows(task_id, op_id)
    assert service.p0.authorization(task_id, op_id, rows)['status'] == 'approved'
    assert service.p0.authorization(task_id, op_id, [{**rows[0], 'excel_row':3}])['status'] == 'pending'
    op2 = service.p0.begin(task_id)
    assert service.p0.authorization(task_id, op2, rows)['status'] == 'pending'


def test_cleanup_skips_active_and_removes_graph_checkpoints(tmp_path, workbook_factory):
    import sqlite3
    service = make_service(tmp_path)
    task_id = create_task(service, workbook_factory())
    with service.repository.connect() as db:
        db.execute("UPDATE tasks SET expires_at='2000-01-01',status='executing' WHERE task_id=?",(task_id,))
    service.p0.cleanup()
    assert service.repository.get_task(task_id) is not None
    checkpoint = service.settings.data_dir/'langgraph-v3-checkpoints.sqlite3'
    with sqlite3.connect(checkpoint) as db:
        db.execute('CREATE TABLE checkpoints(thread_id TEXT,content TEXT)')
        db.execute('INSERT INTO checkpoints VALUES(?,?)', (task_id+':plan:r1','private'))
        db.execute('INSERT INTO checkpoints VALUES(?,?)', ('unrelated:plan:r1','keep'))
    service.repository.update_task(task_id,status='profiled')
    service.p0.cleanup()
    assert service.repository.get_task(task_id) is None
    with sqlite3.connect(checkpoint) as db:
        assert db.execute('SELECT thread_id FROM checkpoints').fetchall() == [('unrelated:plan:r1',)]


def test_exception_paging_and_source_excel_row(tmp_path, workbook_factory):
    import csv
    service = make_service(tmp_path)
    task_id = create_task(service, workbook_factory())
    execute(service, task_id)
    op = service.p0.current(task_id)
    artifact = service.files.artifact_path(task_id,'p0_exceptions.csv')
    with artifact.open('w',encoding='utf-8-sig',newline='') as stream:
        writer=csv.DictWriter(stream,fieldnames=['工作表','行号','源列','原始值','问题代码','说明'])
        writer.writeheader()
        writer.writerows([{'工作表':'订单','行号':i+2,'源列':'订单额','原始值':'private','问题代码':'invalid','说明':'检查格式'} for i in range(53)])
    summary = json.loads(op['summary_json']); summary['exception_artifacts']=[artifact.name]
    with service.repository.connect() as db:
        db.execute('UPDATE p0_operations SET summary_json=? WHERE operation_id=?',(json.dumps(summary),op['operation_id']))
    first=service.p0.exceptions(task_id,op['operation_id'],limit=50)
    last=service.p0.exceptions(task_id,op['operation_id'],offset=50,limit=50)
    assert first['total']==53 and len(first['rows'])==50
    assert last['rows'][0]['excel_row']==52 and len(last['rows'])==3
    assert 'private' not in json.dumps(first)
    preview=service.get_workbook_preview(task_id,0,'订单',2,start_row=3)
    assert preview['start_row']==3 and preview['rows'][0]['订单额']==20


def test_restore_does_not_offer_previous_operation_result_as_current(tmp_path, workbook_factory):
    service = make_service(tmp_path)
    task_id = create_task(service, workbook_factory())
    execute(service, task_id)
    previous = service.p0.current(task_id)['operation_id']
    service.restore_version(task_id,0)
    current = service.get_workspace(task_id)['current_operation']
    assert current['operation_id'] != previous
    assert current['result_version'] is None
    assert current['delivery_status'] == 'not_ready'


def test_duplicate_execution_cannot_downgrade_committed_operation(tmp_path, workbook_factory):
    from sheetpilot.schemas import TaskPlan
    service=make_service(tmp_path)
    task_id=create_task(service,workbook_factory())
    execute(service,task_id)
    plan=TaskPlan.model_validate_json(service.repository.get_task(task_id)['plan_json'])
    with pytest.raises(IntentValidationError):
        service.execute_task_plan(task_id,composite_envelope(),plan)
    assert service.get_workspace(task_id)['current_operation']['delivery_status']=='ready'


def test_interrupted_operation_is_recoverable_after_restart(tmp_path, workbook_factory):
    service = make_service(tmp_path)
    task_id = create_task(service, workbook_factory())
    op_id = service.p0.begin(task_id)
    service.repository.update_task(task_id,status='analyzing_request')
    restarted = make_service(tmp_path)
    assert restarted.p0.current(task_id)['status'] == 'failed'
    assert restarted.repository.get_task(task_id)['status'] == 'failed'
    assert restarted.p0.current(task_id)['operation_id'] == op_id


def test_authorization_gate_pauses_before_model_then_allows_only_bound_rows(tmp_path, workbook_factory):
    from sheetpilot.llm import FakeLLMClient
    from sheetpilot.p0 import authorized_model_rows
    from sheetpilot.schemas import ConversationRole,MessageKind
    service = make_service(tmp_path)
    task_id = create_task(service,workbook_factory())
    seen=[]
    class Capture(FakeLLMClient):
        def parse_task(self,*args,**kwargs):
            seen.append(list(authorized_model_rows.get()))
            return composite_envelope()
    service.llm=Capture(composite_envelope())
    op_id=service.p0.begin(task_id)
    user=service._add_message(task_id,ConversationRole.USER,MessageKind.TEXT,'补充规则')
    service.p0.link(task_id,op_id,'planning','seed',user.id)
    service._add_message(task_id,ConversationRole.ASSISTANT,MessageKind.CLARIFICATION,'需要确认',{'review':{'rows':[{'sheet':'订单','row':2,'raw_values':{'订单额':'受限值'}}]}})
    service.repository.update_task(task_id,status='needs_clarification')
    job=service.enqueue_message(task_id,'请处理这些异常')
    service._background_futures[job['job_id']].result(timeout=10)
    assert not seen
    assert service.repository.get_task(task_id)['status']=='awaiting_data_authorization'
    grant=service.p0.authorization_preview(task_id)
    service.p0.decide_authorization(task_id,grant['grant_id'],True)
    job=service.enqueue_message(task_id,'请处理这些异常')
    service._background_futures[job['job_id']].result(timeout=10)
    assert seen[0][0]['target_values']=={'订单额':'受限值'}
    assert service.p0.current(task_id)['operation_id']==op_id
