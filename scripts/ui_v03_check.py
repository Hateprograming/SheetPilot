from __future__ import annotations

from pathlib import Path

from playwright.sync_api import sync_playwright


OUTPUT = Path(__file__).resolve().parents[1] / "ui-test-artifacts" / "v03_upload.png"


with sync_playwright() as playwright:
    browser = playwright.chromium.launch(headless=True)
    page = browser.new_page(viewport={"width": 1440, "height": 1000})
    console_errors: list[str] = []
    page.on("console", lambda message: console_errors.append(message.text) if message.type == "error" else None)
    page.goto("http://127.0.0.1:8511")
    page.wait_for_load_state("networkidle")
    page.get_by_text("上传工作簿，然后告诉我你想得到什么").wait_for()
    page.get_by_text("v1.0 · Skill 驱动的表格工作流 Agent").wait_for()
    page.screenshot(path=str(OUTPUT), full_page=True)
    application_errors = [
        item for item in console_errors
        if "metrics config" not in item and "ERR_NETWORK_ACCESS_DENIED" not in item
    ]
    assert not application_errors, application_errors
    browser.close()

print(f"UI check passed: {OUTPUT}")
