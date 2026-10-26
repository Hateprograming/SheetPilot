from __future__ import annotations

from pathlib import Path

from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parents[1]
ARTIFACTS = ROOT / "ui-test-artifacts"
WORKBOOK = ROOT.parent / "Test" / "v0.2_cases" / "TC01_跨表匹配.xlsx"
ARTIFACTS.mkdir(exist_ok=True)

with sync_playwright() as playwright:
    browser = playwright.chromium.launch(headless=True)
    page = browser.new_page(viewport={"width": 1440, "height": 1000})
    console_errors: list[str] = []
    page_errors: list[str] = []
    page.on("console", lambda message: console_errors.append(message.text) if message.type == "error" else None)
    page.on("pageerror", lambda error: page_errors.append(str(error)))
    page.goto("http://127.0.0.1:8512")
    page.wait_for_load_state("networkidle")
    page.get_by_text("上传工作簿，然后告诉我你想得到什么", exact=True).wait_for()
    page.locator('input[type="file"]').set_input_files(str(WORKBOOK))
    page.get_by_text("我同意将任务描述", exact=False).click()
    page.get_by_role("button", name="分析这个工作簿").click()
    page.get_by_text("和你的工作簿对话", exact=True).wait_for(timeout=15_000)

    body = page.locator("body").inner_text()
    page.screenshot(path=str(ARTIFACTS / "chat_ui_after_upload.png"), full_page=True)
    assert "需要补充的列（1–3列）" not in body
    assert "任务类型" not in body
    assert "我已经完成本地结构分析" in body
    assert page.locator('[data-testid="stChatInput"]').is_visible()

    sidebar_button = page.get_by_role("button", name="keyboard_double_arrow_right")
    if sidebar_button.count():
        sidebar_button.click()
        page.get_by_text("查看工作簿内容", exact=True).click()
        page.wait_for_timeout(500)
    after = page.locator("body").inner_text()
    assert "和你的工作簿对话" in after
    assert len(after.strip()) > 200
    print(
        {
            "chat_visible": True,
            "legacy_forms_absent": True,
            "body_chars": len(after),
            "console_errors": console_errors,
            "page_errors": page_errors,
        }
    )
    assert not console_errors
    assert not page_errors
    browser.close()
