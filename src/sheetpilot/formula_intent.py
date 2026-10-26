"""Translate explicitly stated column arithmetic without evaluating user code."""
from __future__ import annotations

import ast
import math
import re
import unicodedata


def column_expression(target: str, source: str) -> str | None:
    if not target:
        return None
    text = unicodedata.normalize("NFKC", source).translate(str.maketrans({"×": "*", "÷": "/", "−": "-"}))
    name = re.escape(unicodedata.normalize("NFKC", target))
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if len(lines) > 1:
        latest = lines[-1].removeprefix("=").strip()
        previous = "\n".join(lines[:-1])
        if (re.search(name, previous) and any(cue in previous for cue in ("公式", "填入", "=", "计算"))
                and "=" not in latest and re.search(r"[+*/-]", latest)
                and not re.search(r"[，,。；;]", latest)):
            # An answer to the prior target's formula question may contain only
            # the expression. The same extraction is used by policy validation.
            return latest if len(latest) <= 2048 else None
    # Consume only a complete user clause, never a model-authored expression.
    patterns = [
        rf"[“‘\"']?{name}[”’\"']?\s*(?:这一列|这列|列)?\s*(?:中|里)?\s*(?:填入|填写|写入|设置为|设为|计算为|等于|=)\s*(.+)",
        rf"(?:在|给)\s*[“‘\"']?{name}[”’\"']?\s*(?:这一列|这列|列)\s*(?:中|里)?\s*(?:填入|填写|写入)\s*(.+)",
    ]
    expression = None
    for clause in reversed(re.split(r"[，,。；;\n]", text)):
        for pattern in patterns:
            match = re.search(pattern, clause)
            if match:
                expression = match.group(1).strip().removeprefix("公式").strip()
                break
        if expression:
            break
    if not expression or len(expression) > 2048:
        return None
    return expression.lstrip("=").strip()


def column_formula(profile, sheet_name: str, target: str, source: str, row: int | None = None) -> str | None:
    sheet = next((s for s in profile.sheets if s.name == sheet_name), None)
    selected_target = None
    target_header = target
    if sheet and not any((col.display_header or col.name) == target for col in sheet.columns):
        selected_target = next((col for col in sheet.columns if col.letter == target.upper()), None)
        if selected_target:
            letter = re.escape(selected_target.letter)
            if not re.search(rf"(?:\b{letter}\s*列|(?m:^\s*{letter}\s*$))", source, re.IGNORECASE):
                return None
            target_header = selected_target.display_header or selected_target.name
    expression = column_expression(target_header, source)
    if sheet is None or not expression:
        return None
    headers: dict[str, list] = {}
    for col in sheet.columns:
        header = unicodedata.normalize("NFKC", col.display_header or col.name)
        headers.setdefault(header, []).append(col)
    if selected_target is None and len(headers.get(unicodedata.normalize("NFKC", target_header), [])) != 1:
        return None
    refs = {}
    # Tokenize names before parsing, so headers containing spaces or punctuation work.
    tokens = sorted((h for h in headers if h), key=len, reverse=True)
    token_pattern = "|".join(re.escape(h) for h in tokens)
    if not token_pattern:
        return None
    ambiguous = False
    def replace(match):
        nonlocal ambiguous
        header = match.group(0)
        cols = headers[header]
        if len(cols) != 1:
            ambiguous = True
        key = f"_c{len(refs)}"
        refs[key] = f"{cols[0].letter}{row or (sheet.header_row or 1) + 1}"
        return key
    expression = re.sub(token_pattern, replace, expression)
    def replace_letter(match):
        nonlocal ambiguous
        letter = match.group(1).upper()
        if not any(col.letter == letter for col in sheet.columns):
            ambiguous = True
        key = f"_c{len(refs)}"
        refs[key] = f"{letter}{row or (sheet.header_row or 1) + 1}"
        return key
    expression = re.sub(r"\b([A-Za-z]{1,3})\s*列", replace_letter, expression)
    expression = re.sub(r"[“”‘’\"']", "", expression)
    if ambiguous or not refs:
        return None
    try:
        tree = ast.parse(expression, mode="eval")
    except (SyntaxError, ValueError, RecursionError):
        return None
    def render(node):
        if isinstance(node, ast.Name) and node.id in refs:
            return refs[node.id]
        if isinstance(node, ast.Constant) and type(node.value) in (int, float) and math.isfinite(node.value):
            return str(node.value)
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.UAdd, ast.USub)):
            return ("+" if isinstance(node.op, ast.UAdd) else "-") + render(node.operand)
        if isinstance(node, ast.BinOp) and type(node.op) in (ast.Add, ast.Sub, ast.Mult, ast.Div):
            op = {ast.Add: "+", ast.Sub: "-", ast.Mult: "*", ast.Div: "/"}[type(node.op)]
            return f"({render(node.left)}{op}{render(node.right)})"
        raise ValueError("unsupported arithmetic")
    try:
        return "=" + render(tree.body)
    except (ValueError, RecursionError, OverflowError):
        return None
