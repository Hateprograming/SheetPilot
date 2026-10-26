from __future__ import annotations

import pytest
from fastapi import HTTPException

from sheetpilot import web_api
from sheetpilot.p0 import FeedbackInput
from test_web_api import api_service, upload, run


def test_p0_endpoints_enforce_ownership(api_service):
    task_id = upload(api_service, 'a'*32)
    op_id = api_service.p0.begin(task_id)
    api_service.p0.state(op_id, 'failed')
    for coroutine in (
        web_api.save_feedback(task_id, op_id, FeedbackInput(outcome='partial'), 'b'*32),
        web_api.decide_authorization(task_id, 'grant', web_api.AuthorizationDecision(approved=True), 'b'*32),
        web_api.get_exceptions(task_id, op_id, None, 0, 50, 'b'*32),
    ):
        with pytest.raises(HTTPException) as error:
            run(coroutine)
        assert error.value.status_code == 404


def test_feedback_rejects_unknown_operation_and_accepts_failed_operation(api_service):
    from sheetpilot.errors import IntentValidationError
    owner = 'a'*32
    task_id = upload(api_service, owner)
    op_id = api_service.p0.begin(task_id)
    api_service.p0.state(op_id, 'failed')
    assert run(web_api.save_feedback(task_id, op_id, FeedbackInput(outcome='not_completed'), owner))['status'] == 'saved'
    with pytest.raises(IntentValidationError):
        run(web_api.save_feedback(task_id, 'unknown', FeedbackInput(outcome='completed'), owner))


def test_policy_never_exposes_url_credentials_or_key(api_service, monkeypatch):
    from dataclasses import replace
    settings = replace(api_service.settings, base_url='https://username:password@example.org/path?key=SECRET', api_key='SECRET')
    monkeypatch.setattr(web_api,'settings',settings)
    payload = run(web_api.data_policy())
    assert payload['recipient'] == 'example.org'
    assert all(x not in str(payload) for x in ('SECRET','password','username'))
