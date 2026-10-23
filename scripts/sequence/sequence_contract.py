from __future__ import annotations

import argparse
import json
from typing import Any

LANE_USER = "用户"
LANE_FRONTEND = "前端"
LANE_DB = "数据库"
LANE_ADMIN = "管理员"
LANE_CONTROLLER = "控制器"
LANE_SERVICE = "服务层"
LANE_MAPPER = "Mapper"
LANE_ROUTER = "路由层"
LANE_REPOSITORY = "仓储层"
LANE_STORAGE = "存储层"

FLOW_SPRING_22: tuple[tuple[str, str], ...] = (
    (LANE_ADMIN, LANE_FRONTEND),
    (LANE_FRONTEND, LANE_CONTROLLER),
    (LANE_CONTROLLER, LANE_SERVICE),
    (LANE_SERVICE, LANE_MAPPER),
    (LANE_MAPPER, LANE_DB),
    (LANE_DB, LANE_DB),
    (LANE_DB, LANE_MAPPER),
    (LANE_MAPPER, LANE_SERVICE),
    (LANE_SERVICE, LANE_CONTROLLER),
    (LANE_CONTROLLER, LANE_FRONTEND),
    (LANE_FRONTEND, LANE_ADMIN),
) * 2

FLOW_PYTHON_22: tuple[tuple[str, str], ...] = (
    (LANE_USER, LANE_FRONTEND),
    (LANE_FRONTEND, LANE_ROUTER),
    (LANE_ROUTER, LANE_SERVICE),
    (LANE_SERVICE, LANE_REPOSITORY),
    (LANE_REPOSITORY, LANE_STORAGE),
    (LANE_STORAGE, LANE_STORAGE),
    (LANE_STORAGE, LANE_REPOSITORY),
    (LANE_REPOSITORY, LANE_SERVICE),
    (LANE_SERVICE, LANE_ROUTER),
    (LANE_ROUTER, LANE_FRONTEND),
    (LANE_FRONTEND, LANE_USER),
) * 2

FLOW_5_18: tuple[tuple[str, str], ...] = (
    (LANE_ADMIN, LANE_FRONTEND),
    (LANE_FRONTEND, LANE_CONTROLLER),
    (LANE_CONTROLLER, LANE_SERVICE),
    (LANE_SERVICE, LANE_DB),
    (LANE_DB, LANE_DB),
    (LANE_DB, LANE_SERVICE),
    (LANE_SERVICE, LANE_CONTROLLER),
    (LANE_CONTROLLER, LANE_FRONTEND),
    (LANE_FRONTEND, LANE_ADMIN),
) * 2

FLOW_5_18_PY: tuple[tuple[str, str], ...] = (
    (LANE_ADMIN, LANE_FRONTEND),
    (LANE_FRONTEND, LANE_ROUTER),
    (LANE_ROUTER, LANE_SERVICE),
    (LANE_SERVICE, LANE_DB),
    (LANE_DB, LANE_DB),
    (LANE_DB, LANE_SERVICE),
    (LANE_SERVICE, LANE_ROUTER),
    (LANE_ROUTER, LANE_FRONTEND),
    (LANE_FRONTEND, LANE_ADMIN),
) * 2

FLOW_4_14_SERVICE: tuple[tuple[str, str], ...] = (
    (LANE_ADMIN, LANE_FRONTEND),
    (LANE_FRONTEND, LANE_SERVICE),
    (LANE_SERVICE, LANE_DB),
    (LANE_DB, LANE_DB),
    (LANE_DB, LANE_SERVICE),
    (LANE_SERVICE, LANE_FRONTEND),
    (LANE_FRONTEND, LANE_ADMIN),
) * 2

FLOW_4_14_CONTROLLER: tuple[tuple[str, str], ...] = (
    (LANE_ADMIN, LANE_FRONTEND),
    (LANE_FRONTEND, LANE_CONTROLLER),
    (LANE_CONTROLLER, LANE_DB),
    (LANE_DB, LANE_DB),
    (LANE_DB, LANE_CONTROLLER),
    (LANE_CONTROLLER, LANE_FRONTEND),
    (LANE_FRONTEND, LANE_ADMIN),
) * 2

TEMPLATE_FLOW: dict[str, tuple[tuple[str, str], ...]] = {
    "template_1": FLOW_SPRING_22,
    "template_2": FLOW_PYTHON_22,
    "template_4": FLOW_5_18,
    "template_5": FLOW_5_18_PY,
    "template_6": FLOW_4_14_SERVICE,
    "template_7": FLOW_4_14_CONTROLLER,
}

SEQUENCE_TEXT_MAX_LEN = 6
FRONTEND_LANE_MAX_LEN = 8


def validate_sequence_payload(payload: dict[str, Any]) -> list[str]:
    errs: list[str] = []
    template = payload.get("template")
    messages = payload.get("messages")
    title = payload.get("title")
    frontend_lane = payload.get("frontend_lane")

    if template not in TEMPLATE_FLOW:
        errs.append(
            "template must be one of "
            + ", ".join(sorted(TEMPLATE_FLOW))
        )
        return errs
    if not isinstance(title, str) or not title.strip():
        errs.append("title must be a non-empty string")
    if not isinstance(messages, list):
        errs.append("messages must be an array")
        return errs

    if not isinstance(frontend_lane, str) or not frontend_lane.strip():
        errs.append("frontend_lane must be a non-empty string (e.g. 课程管理页面)")
    elif frontend_lane.strip() == LANE_FRONTEND:
        errs.append('frontend_lane must be a concrete page name, not "前端"')
    elif len(frontend_lane.strip()) > FRONTEND_LANE_MAX_LEN:
        errs.append(
            f"frontend_lane max {FRONTEND_LANE_MAX_LEN} chars, got {len(frontend_lane.strip())}"
        )

    role = payload.get("role")
    if not isinstance(role, str) or not role.strip():
        errs.append("role must be an explicit non-empty string (first lane role)")
    elif "\n" in role or "\r" in role:
        errs.append("role must not contain line breaks")

    expected = TEMPLATE_FLOW[template]
    if len(messages) != len(expected):
        errs.append(f"{template} requires exactly {len(expected)} messages")
        return errs

    for idx, msg in enumerate(messages):
        if not isinstance(msg, dict):
            errs.append(f"messages[{idx}] must be an object")
            continue
        frm = msg.get("from")
        to = msg.get("to")
        text = msg.get("text")
        exp_from, exp_to = expected[idx]

        if not isinstance(frm, str) or not isinstance(to, str):
            errs.append(f"messages[{idx}] from/to must be strings")
        else:
            if frm.strip() != exp_from or to.strip() != exp_to:
                errs.append(
                    f"messages[{idx}] must be {exp_from}->{exp_to}, got {frm!r}->{to!r}"
                )
        if not isinstance(text, str) or not text.strip():
            errs.append(f"messages[{idx}].text must be non-empty string")
        else:
            ln = len(text.strip())
            if ln > SEQUENCE_TEXT_MAX_LEN:
                errs.append(
                    f"messages[{idx}].text max {SEQUENCE_TEXT_MAX_LEN} chars, got {ln}"
                )
    return errs


def main() -> None:
    parser = argparse.ArgumentParser(description="Validate sequence payload contract")
    parser.add_argument("--json", required=True, help="JSON text for sequence data object")
    args = parser.parse_args()

    payload = json.loads(args.json)
    errors = validate_sequence_payload(payload)
    if errors:
        print(json.dumps({"ok": False, "errors": errors}, ensure_ascii=False))
        raise SystemExit(1)
    print(json.dumps({"ok": True, "template": payload["template"]}, ensure_ascii=False))


if __name__ == "__main__":
    main()
