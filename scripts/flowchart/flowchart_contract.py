from __future__ import annotations

import argparse
import json
import re
from typing import Any

FLOWCHART_TEXT_MAX_LEN = 4
ALLOWED_TEXT = re.compile(r"^[一-鿿A-Za-z0-9]+$")
INPUT_VERBS: tuple[str, ...] = ("编辑", "写入", "输入", "填写", "录入", "添加", "新建", "创建")

TEMPLATE_ROLES: dict[int, tuple[str, ...]] = {
    1: ("start", "process", "process", "decision", "process", "process", "process", "end"),
    2: ("start", "input", "process", "decision", "process", "process", "end"),
    3: ("start", "process", "process", "decision", "process", "process", "end"),
    4: ("start", "process", "process", "decision", "process", "process", "end"),
}


def validate_flowchart_payload(payload: dict[str, Any]) -> list[str]:
    errs: list[str] = []
    template = payload.get("template")
    nodes = payload.get("nodes")
    title = payload.get("title")

    if template not in TEMPLATE_ROLES:
        errs.append("template must be 1, 2, 3, or 4")
        return errs
    if not isinstance(title, str) or not title.strip():
        errs.append("title must be a non-empty string")
    if not isinstance(nodes, list):
        errs.append("nodes must be an array")
        return errs

    expected = TEMPLATE_ROLES[int(template)]
    if len(nodes) != len(expected):
        errs.append(f"template {template} requires exactly {len(expected)} nodes")
        return errs

    for idx, node in enumerate(nodes):
        if not isinstance(node, dict):
            errs.append(f"nodes[{idx}] must be an object")
            continue
        role = node.get("role")
        text = node.get("text")
        exp = expected[idx]

        if role != exp:
            errs.append(f"nodes[{idx}].role must be {exp}")
        if not isinstance(text, str) or not text.strip():
            errs.append(f"nodes[{idx}].text must be non-empty string")
            continue
        ts = text.strip()
        if len(ts) > FLOWCHART_TEXT_MAX_LEN:
            errs.append(f"nodes[{idx}].text max {FLOWCHART_TEXT_MAX_LEN} chars")
        if not ALLOWED_TEXT.fullmatch(ts):
            errs.append(f"nodes[{idx}].text contains invalid symbols or whitespace")
        if exp == "decision" and not ts.startswith("是否"):
            errs.append(f'nodes[{idx}].text decision must start with "是否"')
        if exp == "input" and not any(k in ts for k in INPUT_VERBS):
            errs.append(f"nodes[{idx}].text input must contain an action verb")
        if exp == "process":
            matched = [k for k in INPUT_VERBS if k in ts]
            if matched:
                errs.append(
                    f"nodes[{idx}].text is input-like ({'/'.join(matched)}), "
                    f"must use the input slot, not a process rectangle"
                )
    return errs


def main() -> None:
    parser = argparse.ArgumentParser(description="Validate flowchart payload contract")
    parser.add_argument("--json", required=True, help="JSON text for flowchart data object")
    args = parser.parse_args()

    payload = json.loads(args.json)
    errors = validate_flowchart_payload(payload)
    if errors:
        print(json.dumps({"ok": False, "errors": errors}, ensure_ascii=False))
        raise SystemExit(1)
    print(json.dumps({"ok": True, "template": payload["template"]}, ensure_ascii=False))


if __name__ == "__main__":
    main()
