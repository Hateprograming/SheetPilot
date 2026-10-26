from __future__ import annotations

import hashlib
import hmac
import json
import logging
import mimetypes
import os
import secrets
import asyncio
from contextlib import asynccontextmanager, suppress
from dataclasses import replace
from pathlib import Path
from typing import Any
from urllib.parse import quote

from fastapi import BackgroundTasks, Depends, FastAPI, File, Form, HTTPException, Request, Response, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from .config import Settings
from .errors import SheetPilotError
from .service import SheetPilotService
from .p0 import FeedbackInput


LOGGER = logging.getLogger(__name__)
COOKIE_NAME = "sheetpilot_session"
ROOT = Path(__file__).resolve().parents[2]
settings = Settings.load(ROOT)
web_data_raw = Path(os.getenv("SHEETPILOT_WEB_DATA_DIR", "./data_web"))
web_data_dir = web_data_raw if web_data_raw.is_absolute() else ROOT / web_data_raw
settings = replace(
    settings,
    data_dir=web_data_dir.resolve(),
    database_path=(web_data_dir / "sheetpilot.sqlite3").resolve(),
)
service = SheetPilotService(settings)
@asynccontextmanager
async def lifespan(_: FastAPI):
    await asyncio.to_thread(service.p0.cleanup)
    async def sweep():
        while True:
            await asyncio.sleep(3600)
            await asyncio.to_thread(service.p0.cleanup)
    cleanup_task = asyncio.create_task(sweep())
    try:
        yield
    finally:
        cleanup_task.cancel()
        with suppress(asyncio.CancelledError):
            await cleanup_task


app = FastAPI(title="SheetPilot Web API", version="2.1.0", lifespan=lifespan)


class MessageRequest(BaseModel):
    text: str = Field(min_length=1, max_length=10_000)


class ClarificationRequest(BaseModel):
    label: str = Field(min_length=1, max_length=300)


class ExecutePlanRequest(BaseModel):
    acknowledgements: list[str] = Field(default_factory=list)


class RenameWorkbookRequest(BaseModel):
    name: str = Field(min_length=1, max_length=180)


class AuthorizationDecision(BaseModel):
    approved: bool


def _sign_session(owner_id: str) -> str:
    signature = hmac.new(
        settings.session_secret.encode("utf-8"), owner_id.encode("ascii"), hashlib.sha256
    ).hexdigest()
    return f"{owner_id}.{signature}"


def _read_session(raw: str | None) -> str | None:
    if not raw or "." not in raw:
        return None
    owner_id, signature = raw.rsplit(".", 1)
    if len(owner_id) != 32 or any(char not in "0123456789abcdef" for char in owner_id):
        return None
    return owner_id if hmac.compare_digest(_sign_session(owner_id), raw) else None


async def current_owner(request: Request, response: Response) -> str:
    owner_id = _read_session(request.cookies.get(COOKIE_NAME))
    if owner_id is None:
        owner_id = secrets.token_hex(16)
        response.set_cookie(
            COOKIE_NAME,
            _sign_session(owner_id),
            httponly=True,
            secure=request.url.scheme == "https",
            samesite="lax",
            max_age=60 * 60 * 24 * 30,
            path="/",
        )
    return owner_id


def require_owned_task(task_id: str, owner_id: str) -> dict[str, Any]:
    task = service.repository.get_task(task_id)
    if not task or task.get("owner_id") != owner_id or task.get("cleanup_status") == "pending":
        raise HTTPException(status_code=404, detail="工作簿不存在或已不可用。")
    return task


def _safe_workspace(workspace: dict[str, Any]) -> dict[str, Any]:
    blocked = {
        "owner_id", "original_path", "current_path", "path", "checkpoint_path",
        "last_checkpoint_path", "input_sha256", "output_sha256", "original_sha256",
        "base_sha256", "fencing_token", "spec_json", "plan_json", "changeset_json",
        "verification_json", "profile_json", "delivery_json",
    }

    def clean(value: Any) -> Any:
        if isinstance(value, list):
            return [clean(item) for item in value]
        if not isinstance(value, dict):
            return value
        result: dict[str, Any] = {}
        for key, item in value.items():
            if key in blocked or key.endswith("_path") or key.endswith("sha256") or key == "commit_id":
                continue
            if key == "error_json":
                if item:
                    try:
                        error = json.loads(item) if isinstance(item, str) else item
                    except (TypeError, json.JSONDecodeError):
                        error = {}
                    result["error"] = {
                        "safe_message": error.get("safe_message", "处理没有完成，正式版本保持不变。"),
                        "category": error.get("category"),
                    }
                continue
            result[key] = clean(item)
        return result

    return clean(workspace)


def _download_response(name: str, content: bytes, media_type: str | None = None) -> Response:
    encoded = quote(name)
    if media_type is None and name.casefold().endswith(".csv"):
        media_type = "text/csv"
    return Response(
        content=content,
        media_type=media_type or mimetypes.guess_type(name)[0] or "application/octet-stream",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded}"},
    )


@app.exception_handler(SheetPilotError)
async def sheetpilot_error_handler(_: Request, exc: SheetPilotError) -> JSONResponse:
    return JSONResponse(
        status_code=400,
        content={"code": type(exc).__name__, "message": exc.message, "detail": exc.detail},
    )


@app.exception_handler(Exception)
async def unexpected_error_handler(_: Request, exc: Exception) -> JSONResponse:
    LOGGER.exception("Unhandled web API error", exc_info=exc)
    return JSONResponse(
        status_code=500,
        content={
            "code": "InternalError",
            "message": "这次操作没有完成，文件和上一正式版本保持不变。",
            "detail": None,
        },
    )


@app.get("/api/health")
async def health() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/api/session")
async def session(owner_id: str = Depends(current_owner)) -> dict[str, str]:
    return {"status": "ready", "session": owner_id[:8]}


@app.get("/api/data-policy")
async def data_policy() -> dict[str, Any]:
    from urllib.parse import urlsplit
    host = urlsplit(settings.base_url).hostname or "尚未配置"
    return {"recipient": host, "provider": settings.llm_provider, "model": settings.model,
            "retention_days": settings.retention_days,
            "default_sent": ["你输入的任务文字与近期用户对话", "文件名、相关工作表及表头", "行列数、字段类型及采样统计"],
            "extra_consent": "异常目标值需要另行授权；不发送完整文件或整行。",
            "deletion": "可主动删除；闲置到期自动清理。保留期只约束本应用，不代替模型服务的数据政策。"}


@app.put("/api/workbooks/{task_id}/operations/{operation_id}/feedback")
async def save_feedback(task_id: str, operation_id: str, payload: FeedbackInput, owner_id: str = Depends(current_owner)):
    require_owned_task(task_id, owner_id)
    service.p0.save_feedback(task_id, operation_id, payload)
    return {"status": "saved"}


@app.post("/api/workbooks/{task_id}/authorizations/{grant_id}", status_code=202)
async def decide_authorization(task_id: str, grant_id: str, payload: AuthorizationDecision, owner_id: str = Depends(current_owner)):
    require_owned_task(task_id, owner_id)
    service.p0.decide_authorization(task_id, grant_id, payload.approved)
    if payload.approved:
        run = service.repository.latest_planning_run(task_id)
        message = service.repository.get_message(run["request_message_id"])
        return service.enqueue_message(task_id, message['content'])
    return {"status": "denied", "message": "未发送异常值。请手动指定处理规则，或停止本次处理。"}


@app.get("/api/workbooks/{task_id}/operations/{operation_id}/exceptions")
async def get_exceptions(task_id: str, operation_id: str, issue: str | None = None, offset: int = 0, limit: int = 50, owner_id: str = Depends(current_owner)):
    require_owned_task(task_id, owner_id)
    return service.p0.exceptions(task_id, operation_id, issue, max(0, offset), max(1, min(50, limit)))


@app.post("/api/workbooks", status_code=201)
async def create_workbook(
    response: Response,
    file: UploadFile = File(...),
    consent: bool = Form(...),
    owner_id: str = Depends(current_owner),
) -> dict[str, Any]:
    if not consent:
        raise HTTPException(status_code=400, detail="请先确认数据处理说明。")
    content = await file.read(settings.max_file_bytes + 1)
    profile = service.create_task(file.filename or "workbook.xlsx", content, owner_id=owner_id)
    task_id = service.task_id_from_profile(profile)
    service.repository.log_event(task_id, "data_policy_accepted", {"policy_version": "p0-v1", "categories": ["user_text", "workbook_structure", "statistics"], "retention_days": settings.retention_days})
    return {"task_id": task_id, "workspace": _safe_workspace(service.get_workspace(task_id))}


@app.get("/api/workbooks")
async def list_workbooks(owner_id: str = Depends(current_owner)) -> list[dict[str, Any]]:
    return service.recent_workspaces(limit=None, owner_id=owner_id)


@app.get("/api/workbooks/{task_id}/workspace")
async def get_workspace(task_id: str, owner_id: str = Depends(current_owner)) -> dict[str, Any]:
    require_owned_task(task_id, owner_id)
    return _safe_workspace(service.get_workspace(task_id))


@app.patch("/api/workbooks/{task_id}")
async def rename_workbook(
    task_id: str, payload: RenameWorkbookRequest, owner_id: str = Depends(current_owner)
) -> dict[str, str]:
    require_owned_task(task_id, owner_id)
    return {"name": service.rename_workbook(task_id, payload.name)}


@app.delete("/api/workbooks/{task_id}")
async def delete_workbook(
    task_id: str, owner_id: str = Depends(current_owner)
) -> dict[str, str]:
    require_owned_task(task_id, owner_id)
    service.delete_workbook(task_id)
    return {"status": "deleted"}


@app.get("/api/workbooks/{task_id}/preview")
async def get_preview(
    task_id: str,
    sheet: str,
    version: int | None = None,
    limit: int = 50,
    owner_id: str = Depends(current_owner),
    start_row: int | None = None,
) -> dict[str, Any]:
    require_owned_task(task_id, owner_id)
    return service.get_workbook_preview(task_id, version, sheet, limit, start_row=start_row)


@app.post("/api/workbooks/{task_id}/messages", status_code=202)
async def send_message(
    task_id: str, payload: MessageRequest, owner_id: str = Depends(current_owner)
) -> dict[str, Any]:
    require_owned_task(task_id, owner_id)
    return service.enqueue_message(task_id, payload.text.strip())


@app.post("/api/workbooks/{task_id}/clarifications/custom")
async def custom_clarification(
    task_id: str, payload: ClarificationRequest, owner_id: str = Depends(current_owner)
) -> dict[str, str]:
    require_owned_task(task_id, owner_id)
    service.select_custom_clarification_input(task_id, payload.label)
    return {"status": "ready"}


@app.post("/api/workbooks/{task_id}/plan/execute", status_code=202)
async def execute_plan(
    task_id: str, payload: ExecutePlanRequest, owner_id: str = Depends(current_owner)
) -> dict[str, Any]:
    require_owned_task(task_id, owner_id)
    service.p0.touch(task_id)
    return service.enqueue_current_plan_execution(task_id, payload.acknowledgements)


@app.post("/api/workbooks/{task_id}/plan/cancel")
async def cancel_plan(task_id: str, owner_id: str = Depends(current_owner)) -> dict[str, str]:
    require_owned_task(task_id, owner_id)
    service.cancel_current_plan(task_id)
    return {"status": "cancelled"}


@app.post("/api/workbooks/{task_id}/run/{action}", status_code=202)
async def control_run(
    task_id: str,
    action: str,
    background_tasks: BackgroundTasks,
    owner_id: str = Depends(current_owner),
) -> dict[str, str]:
    require_owned_task(task_id, owner_id)
    service.p0.touch(task_id)
    if action in {"pause", "cancel"}:
        service.control_current_run(task_id, action)
    elif action == "resume":
        background_tasks.add_task(service.resume_current_run, task_id)
    elif action == "skip":
        background_tasks.add_task(service.skip_optional_step, task_id)
    else:
        raise HTTPException(status_code=404, detail="不支持的运行操作。")
    return {"status": "queued", "action": action}


@app.post("/api/workbooks/{task_id}/versions/{version}/restore")
async def restore_version(
    task_id: str, version: int, owner_id: str = Depends(current_owner)
) -> dict[str, Any]:
    require_owned_task(task_id, owner_id)
    service.p0.touch(task_id)
    profile = service.restore_version(task_id, version)
    return {"status": "restored", "profile": profile.model_dump(mode="json")}


@app.get("/api/workbooks/{task_id}/download")
async def download_current(task_id: str, owner_id: str = Depends(current_owner)) -> Response:
    require_owned_task(task_id, owner_id)
    return _download_response(*service.get_download(task_id))


@app.get("/api/workbooks/{task_id}/original")
async def download_original(task_id: str, owner_id: str = Depends(current_owner)) -> Response:
    require_owned_task(task_id, owner_id)
    return _download_response(*service.get_download(task_id, original=True))


@app.get("/api/workbooks/{task_id}/versions/{version}/download")
async def download_version(
    task_id: str, version: int, owner_id: str = Depends(current_owner)
) -> Response:
    require_owned_task(task_id, owner_id)
    return _download_response(*service.get_version_download(task_id, version))


@app.get("/api/workbooks/{task_id}/artifacts/{name}")
async def download_artifact(
    task_id: str, name: str, owner_id: str = Depends(current_owner)
) -> Response:
    require_owned_task(task_id, owner_id)
    artifact = service.files.artifact_path(task_id, name)
    return _download_response(*service.get_artifact_download(task_id, str(artifact)))


frontend_dist = ROOT / "web" / "dist"
if frontend_dist.is_dir():
    assets = frontend_dist / "assets"
    if assets.is_dir():
        app.mount("/assets", StaticFiles(directory=assets), name="assets")

    @app.get("/{path:path}", include_in_schema=False)
    async def spa(path: str) -> FileResponse:
        candidate = (frontend_dist / path).resolve()
        if path and candidate.is_file() and frontend_dist.resolve() in candidate.parents:
            return FileResponse(candidate)
        return FileResponse(frontend_dist / "index.html")


def main() -> None:
    import uvicorn

    uvicorn.run("sheetpilot.web_api:app", host="0.0.0.0", port=8000, workers=1)
