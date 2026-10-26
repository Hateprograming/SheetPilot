from __future__ import annotations

from pathlib import Path

from playwright.sync_api import Page, sync_playwright


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "ui-test-artifacts" / "web-redesign"
SAMPLE = ROOT.parent / "Test" / "v1.0_cases" / "01_跨表精确匹配.xlsx"


def assert_type_scale(page: Page) -> None:
    violations = page.evaluate(
        """
        () => [...document.querySelectorAll('body *')].flatMap((element) => {
          const directText = [...element.childNodes]
            .filter((node) => node.nodeType === Node.TEXT_NODE)
            .map((node) => node.textContent.trim()).join(' ')
          if (!directText || element.getClientRects().length === 0) return []
          const size = Number.parseFloat(getComputedStyle(element).fontSize)
          return size < 13.9 || size > 20.1
            ? [{ tag: element.tagName, text: directText.slice(0, 60), size }]
            : []
        })
        """
    )
    assert not violations, f"Typography outside 14-20px: {violations[:10]}"


def capture(page: Page, name: str) -> None:
    page.screenshot(path=str(OUTPUT / name), full_page=True)
    assert_type_scale(page)
    assert page.evaluate("document.documentElement.scrollWidth <= window.innerWidth")


def mock_workspace(status: str) -> dict:
    messages = [
        {"id": 1, "role": "user", "kind": "text", "content": "按订单编号去重，每组保留日期最新的一条。", "metadata": {}, "created_at": "2026-10-16T10:00:00"},
    ]
    plan = {
        "revision": 1,
        "base_version": 0,
        "risk_level": "low",
        "confirmation_requirements": [],
        "steps": [
            {"id": "s1", "description": "按订单编号识别重复记录", "expected_changes": {"estimated_rows": 248, "modified_sheets": ["订单明细"]}},
            {"id": "s2", "description": "每组保留订单日期最新的一条", "expected_changes": {"estimated_rows": 248, "modified_sheets": ["订单明细"]}},
            {"id": "s3", "description": "检查编号唯一性和总行数", "expected_changes": {"modified_sheets": ["订单明细"]}},
        ],
    }
    if status in {"awaiting_approval", "executing", "failed"}:
        messages.append({"id": 2, "role": "assistant", "kind": "plan", "content": "我已经整理好处理方式。", "metadata": {"plan": plan, "envelope": {"goal": "删除重复订单并保留最新记录"}}})
    if status == "needs_clarification":
        messages.append({"id": 2, "role": "assistant", "kind": "clarification", "content": "有两条记录的订单日期相同，需要你确认保留规则。", "metadata": {"choices": ["保留在表格中靠后的一条", "保留靠前的一条"], "review": {"total": 2, "rows": [{"Sheet": "订单明细", "行号": 18, "订单编号": "A10008"}]}}})
    verification = {"status": "passed", "checks": [{"name": "订单编号唯一性", "detail": "未发现重复值", "passed": True}]}
    result = {"rows_affected": 248, "warnings": [], "details": {"steps": plan["steps"]}}
    if status == "completed":
        messages.append({"id": 2, "role": "assistant", "kind": "result", "content": "已移除 248 条重复记录，每个订单编号都保留了日期最新的一条。", "metadata": {"result": result, "verification": verification}})
    planning = {"job_id": "job-1", "status": "running", "stage": "preview"} if status == "analyzing_request" else None
    return {
        "task": {"task_id": "mock", "original_name": "2024年Q3华东区销售明细.xlsx", "current_version": 1 if status == "completed" else 0, "status": status},
        "profile": {"sheets": [{"name": "订单明细", "max_row": 12401, "max_column": 5, "header_row": 1}], "risks": []},
        "messages": messages,
        "versions": ([{"version": 0, "created_at": "2026-10-16T09:30:00", "parent_version": None}, {"version": 1, "created_at": "2026-10-16T10:10:00", "parent_version": 0}] if status == "completed" else [{"version": 0, "created_at": "2026-10-16T09:30:00", "parent_version": None}]),
        "result": result if status == "completed" else None,
        "verification": verification if status == "completed" else None,
        "active_run": ({"status": "running"} if status == "executing" else {"status": "failed", "error": {"safe_message": "处理在核验前暂停，正式版本没有改变。"}} if status == "failed" else None),
        "tool_runs": ([{"plan_step_id": "s1", "status": "completed"}] if status == "executing" else []),
        "planning_run": planning,
        "activity_events": ([{"id": 1, "run_kind": "planning", "run_id": "job-1", "status": "completed", "label": "已读取工作簿结构"}, {"id": 2, "run_kind": "planning", "run_id": "job-1", "status": "running", "label": "正在预检受影响的列和异常记录"}] if status == "analyzing_request" else []),
    }


def verify_mock_states(browser) -> None:
    current = {"status": "analyzing_request"}
    context = browser.new_context(viewport={"width": 1440, "height": 900})

    def handle(route) -> None:
        url = route.request.url
        if url.endswith("/api/session"):
            route.fulfill(json={"status": "ok"})
        elif "/api/workbooks/mock/workspace" in url:
            route.fulfill(json=mock_workspace(current["status"]))
        elif "/api/workbooks/mock/preview" in url:
            route.fulfill(json={"sheet": "订单明细", "version": 1 if current["status"] == "completed" else 0, "headers": ["订单编号", "客户名称", "区域", "订单日期", "金额"], "rows": [{"订单编号": "A10001", "客户名称": "上海益达商贸", "区域": "华东", "订单日期": "2026-07-01", "金额": 12800}], "total_rows": 12152})
        elif url.endswith("/api/workbooks"):
            route.fulfill(json=[{"task_id": "mock", "name": "2024年Q3华东区销售明细.xlsx", "version": 0, "status": current["status"], "updated_at": "2026-10-16T10:00:00"}])
        else:
            route.continue_()

    context.route("**/api/**", handle)
    page = context.new_page()
    for status in ["analyzing_request", "needs_clarification", "awaiting_approval", "executing", "failed", "completed"]:
        current["status"] = status
        page.goto("http://127.0.0.1:8000/workbooks/mock", wait_until="networkidle")
        page.locator(".workspace-page").wait_for()
        capture(page, f"state-{status}.png")
    context.close()


def main() -> None:
    OUTPUT.mkdir(parents=True, exist_ok=True)
    console_errors: list[str] = []
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        context = browser.new_context(viewport={"width": 1920, "height": 1080})
        page = context.new_page()
        page.on(
            "console",
            lambda message: console_errors.append(message.text)
            if message.type == "error" else None,
        )
        page.goto("http://127.0.0.1:8000", wait_until="networkidle")
        page.get_by_text("How can I help with you today?", exact=True).wait_for()
        capture(page, "home-1920.png")

        page.locator('input[type="file"]').set_input_files(str(SAMPLE))
        page.get_by_role("button", name="分析这个工作簿").click()
        page.wait_for_url("**/workbooks/**", timeout=30_000)
        page.get_by_role("heading", name="工作簿预览").wait_for(timeout=30_000)
        page.locator(".table-wrap table").wait_for(timeout=30_000)
        capture(page, "workspace-1920.png")
        workspace_url = page.url

        page.get_by_role("link", name="版本记录").click()
        page.get_by_role("heading", name="版本轨迹").wait_for()
        page.locator(".version-row").first.wait_for()
        capture(page, "versions-1920.png")

        compact = browser.new_context(viewport={"width": 1440, "height": 900})
        compact_page = compact.new_page()
        compact_page.goto("http://127.0.0.1:8000", wait_until="networkidle")
        capture(compact_page, "home-1440.png")
        compact.close()

        isolated = browser.new_context(viewport={"width": 1440, "height": 900})
        isolated_page = isolated.new_page()
        isolated_page.goto(workspace_url, wait_until="networkidle")
        isolated_page.get_by_role("heading", name="无法打开工作簿").wait_for()
        isolated.close()

        verify_mock_states(browser)

        assert not console_errors, f"Browser console errors: {console_errors}"
        context.close()
        browser.close()

    print(f"screenshots={OUTPUT}")
    print("font_range=14-20px")
    print("console_errors=0")


if __name__ == "__main__":
    main()
