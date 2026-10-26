from __future__ import annotations

from io import BytesIO

import pytest
from fastapi import HTTPException, Response, UploadFile
from openpyxl import Workbook

from sheetpilot import web_api
from sheetpilot.config import Settings
from sheetpilot.errors import FileRejectedError, IntentValidationError
from sheetpilot.service import SheetPilotService


def workbook_bytes() -> bytes:
    stream = BytesIO()
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "订单明细"
    sheet.append(["订单编号", "金额"])
    sheet.append(["SO-001", 120])
    sheet.append(["SO-002", 240])
    workbook.save(stream)
    workbook.close()
    return stream.getvalue()


@pytest.fixture
def api_service(tmp_path, monkeypatch):
    settings = Settings(
        project_root=tmp_path,
        data_dir=tmp_path / "web-data",
        database_path=tmp_path / "web-data" / "sheetpilot.sqlite3",
        llm_provider="fake",
        api_key="",
        base_url="",
        model="fake",
        max_file_bytes=20 * 1024 * 1024,
        max_sheet_rows=100_000,
        session_secret="test-session-secret",
    )
    service = SheetPilotService(settings)
    monkeypatch.setattr(web_api, "settings", settings)
    monkeypatch.setattr(web_api, "service", service)
    return service


def upload(service: SheetPilotService, owner_id: str) -> str:
    profile = service.create_task("orders.xlsx", workbook_bytes(), owner_id=owner_id)
    return service.task_id_from_profile(profile)


def run(coroutine):
    try:
        coroutine.send(None)
    except StopIteration as completed:
        return completed.value
    raise AssertionError("The endpoint unexpectedly suspended during this unit test")


def test_anonymous_sessions_are_isolated(api_service):
    owner = "a" * 32
    stranger = "b" * 32
    task_id = upload(api_service, owner)

    assert run(web_api.get_workspace(task_id, owner))["task"]["task_id"] == task_id
    with pytest.raises(HTTPException, match="工作簿不存在"):
        run(web_api.get_workspace(task_id, stranger))
    with pytest.raises(HTTPException, match="工作簿不存在"):
        run(web_api.download_current(task_id, stranger))
    assert run(web_api.list_workbooks(owner))[0]["task_id"] == task_id
    assert run(web_api.list_workbooks(stranger)) == []


def test_workbook_list_includes_older_owned_workbooks(api_service):
    owner, stranger = "a" * 32, "b" * 32
    owned = {upload(api_service, owner) for _ in range(14)}
    foreign = upload(api_service, stranger)
    listed = run(web_api.list_workbooks(owner))
    assert {item["task_id"] for item in listed} == owned
    assert foreign not in {item["task_id"] for item in listed}
    assert [item["updated_at"] for item in listed] == sorted((item["updated_at"] for item in listed), reverse=True)


def test_workspace_removes_server_paths_and_supports_preview(api_service):
    owner = "a" * 32
    task_id = upload(api_service, owner)
    workspace = run(web_api.get_workspace(task_id, owner))
    serialized = str(workspace)

    assert "original_path" not in serialized
    assert "current_path" not in serialized
    assert "owner_id" not in serialized
    preview = run(web_api.get_preview(task_id, "订单明细", None, 20, owner))
    assert preview["rows"][0]["订单编号"] == "SO-001"


def test_upload_requires_consent_and_supported_file(api_service):
    owner = "a" * 32
    file = UploadFile(filename="orders.xlsx", file=BytesIO(workbook_bytes()))
    with pytest.raises(HTTPException, match="确认数据处理说明"):
        run(web_api.create_workbook(Response(), file, False, owner))

    with pytest.raises(FileRejectedError):
        api_service.create_task("notes.txt", b"not a workbook", owner_id=owner)


def test_restore_download_and_session_signature(api_service):
    owner = "a" * 32
    task_id = upload(api_service, owner)
    signed = web_api._sign_session(owner)

    assert web_api._read_session(signed) == owner
    assert web_api._read_session(signed + "x") is None
    restored = run(web_api.restore_version(task_id, 0, owner))
    assert restored["status"] == "restored"
    assert run(web_api.download_current(task_id, owner)).status_code == 200
    assert run(web_api.download_original(task_id, owner)).status_code == 200
    assert run(web_api.download_version(task_id, 0, owner)).status_code == 200


def test_rename_and_delete_workbook_are_owner_scoped(api_service):
    owner = "a" * 32
    stranger = "b" * 32
    task_id = upload(api_service, owner)
    task_dir = api_service.files.task_dir(task_id)

    renamed = run(web_api.rename_workbook(
        task_id, web_api.RenameWorkbookRequest(name="华东区销售明细"), owner
    ))
    assert renamed["name"] == "华东区销售明细.xlsx"
    assert api_service.repository.get_task(task_id)["original_name"] == "华东区销售明细.xlsx"

    with pytest.raises(HTTPException, match="工作簿不存在"):
        run(web_api.delete_workbook(task_id, stranger))

    assert run(web_api.delete_workbook(task_id, owner)) == {"status": "deleted"}
    assert api_service.repository.get_task(task_id) is None
    assert not task_dir.exists()


def test_executing_workbook_cannot_be_deleted(api_service):
    owner = "a" * 32
    task_id = upload(api_service, owner)
    api_service.repository.update_task(task_id, status="executing")

    with pytest.raises(IntentValidationError, match="正在处理"):
        run(web_api.delete_workbook(task_id, owner))
    assert api_service.repository.get_task(task_id) is not None
