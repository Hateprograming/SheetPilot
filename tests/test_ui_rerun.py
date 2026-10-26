from __future__ import annotations

import ast
from pathlib import Path

from streamlit.testing.v1 import AppTest


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def test_widget_rerun_keeps_upload_page_visible():
    app = AppTest.from_file(PROJECT_ROOT / "app.py", default_timeout=10).run()

    assert not app.exception
    assert any("上传工作簿，然后告诉我你想得到什么" in block.value for block in app.markdown)
    assert app.checkbox

    app.checkbox[0].check().run()

    assert not app.exception
    assert any("上传工作簿，然后告诉我你想得到什么" in block.value for block in app.markdown)
    assert app.checkbox[0].value is True


def test_result_preview_does_not_shadow_message_argument():
    tree = ast.parse((PROJECT_ROOT / "app_chat.py").read_text(encoding="utf-8"))
    function = next(
        node for node in tree.body
        if isinstance(node, ast.FunctionDef) and node.name == "render_result_preview"
    )
    reassigned = {
        node.id for node in ast.walk(function)
        if isinstance(node, ast.Name) and isinstance(node.ctx, ast.Store)
    }
    assert "message" not in reassigned
