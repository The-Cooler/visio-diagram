"""
时序图固定泳道与每条消息的 from→to 约定（按 template/sequence/*.vsdx 实际连接确认）。

  template_1: 22 条（6泳道：管理员/前端/控制器/服务层/Mapper/数据库）  Spring Boot 六层
  template_2: 22 条（6泳道：用户/前端/路由层/服务层/仓储层/存储层）    Python 六层
  template_4: 18 条（5泳道：管理员/前端/控制器/服务层/数据库）          Spring Boot 五层
  template_5: 18 条（5泳道：管理员/前端/路由层/服务层/数据库）          Python 五层
  template_6: 14 条（4泳道：管理员/前端/服务层/数据库）                通用四泳道
  template_7: 14 条（4泳道：管理员/前端/控制器/数据库）                通用四泳道

参与者名称须与 Visio 模板中泳道文案一致。
"""

from __future__ import annotations

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

# template_1: Spring Boot 六层，11 条 × 2 轮
TEMPLATE_1_ROUND1: tuple[tuple[str, str], ...] = (
    (LANE_ADMIN, LANE_FRONTEND),       # 1
    (LANE_FRONTEND, LANE_CONTROLLER),  # 2
    (LANE_CONTROLLER, LANE_SERVICE),   # 3
    (LANE_SERVICE, LANE_MAPPER),       # 4
    (LANE_MAPPER, LANE_DB),            # 5
    (LANE_DB, LANE_DB),                # 6  数据库自循环
    (LANE_DB, LANE_MAPPER),            # 7  返回
    (LANE_MAPPER, LANE_SERVICE),       # 8  返回
    (LANE_SERVICE, LANE_CONTROLLER),   # 9  返回
    (LANE_CONTROLLER, LANE_FRONTEND),  # 10 返回
    (LANE_FRONTEND, LANE_ADMIN),       # 11 返回
)
TEMPLATE_1_SLOTS: tuple[tuple[str, str], ...] = TEMPLATE_1_ROUND1 + TEMPLATE_1_ROUND1

# template_2: Python 六层，11 条 × 2 轮
TEMPLATE_2_ROUND1: tuple[tuple[str, str], ...] = (
    (LANE_USER, LANE_FRONTEND),        # 1
    (LANE_FRONTEND, LANE_ROUTER),      # 2
    (LANE_ROUTER, LANE_SERVICE),       # 3
    (LANE_SERVICE, LANE_REPOSITORY),   # 4
    (LANE_REPOSITORY, LANE_STORAGE),   # 5
    (LANE_STORAGE, LANE_STORAGE),      # 6  存储自循环
    (LANE_STORAGE, LANE_REPOSITORY),   # 7  返回
    (LANE_REPOSITORY, LANE_SERVICE),   # 8  返回
    (LANE_SERVICE, LANE_ROUTER),       # 9  返回
    (LANE_ROUTER, LANE_FRONTEND),      # 10 返回
    (LANE_FRONTEND, LANE_USER),        # 11 返回
)
TEMPLATE_2_SLOTS: tuple[tuple[str, str], ...] = TEMPLATE_2_ROUND1 + TEMPLATE_2_ROUND1

# template_4: 五层 9 条 × 2 轮（无 Mapper）
TEMPLATE_4_ROUND1: tuple[tuple[str, str], ...] = (
    (LANE_ADMIN, LANE_FRONTEND),        # 1
    (LANE_FRONTEND, LANE_CONTROLLER),   # 2
    (LANE_CONTROLLER, LANE_SERVICE),    # 3
    (LANE_SERVICE, LANE_DB),            # 4
    (LANE_DB, LANE_DB),                 # 5  数据库自循环
    (LANE_DB, LANE_SERVICE),            # 6  返回
    (LANE_SERVICE, LANE_CONTROLLER),    # 7  返回
    (LANE_CONTROLLER, LANE_FRONTEND),   # 8  返回
    (LANE_FRONTEND, LANE_ADMIN),        # 9  返回
)
TEMPLATE_4_SLOTS: tuple[tuple[str, str], ...] = TEMPLATE_4_ROUND1 + TEMPLATE_4_ROUND1

# template_5: 五层 9 条 × 2 轮（无仓储层，Python 分层）
TEMPLATE_5_ROUND1: tuple[tuple[str, str], ...] = (
    (LANE_ADMIN, LANE_FRONTEND),        # 1
    (LANE_FRONTEND, LANE_ROUTER),       # 2
    (LANE_ROUTER, LANE_SERVICE),        # 3
    (LANE_SERVICE, LANE_DB),            # 4
    (LANE_DB, LANE_DB),                 # 5  数据库自循环
    (LANE_DB, LANE_SERVICE),            # 6  返回
    (LANE_SERVICE, LANE_ROUTER),        # 7  返回
    (LANE_ROUTER, LANE_FRONTEND),       # 8  返回
    (LANE_FRONTEND, LANE_ADMIN),        # 9  返回
)
TEMPLATE_5_SLOTS: tuple[tuple[str, str], ...] = TEMPLATE_5_ROUND1 + TEMPLATE_5_ROUND1

# template_6: 四泳道 7 条 × 2 轮（管理员/前端/服务层/数据库）
TEMPLATE_6_ROUND1: tuple[tuple[str, str], ...] = (
    (LANE_ADMIN, LANE_FRONTEND),        # 1
    (LANE_FRONTEND, LANE_SERVICE),      # 2
    (LANE_SERVICE, LANE_DB),            # 3
    (LANE_DB, LANE_DB),                 # 4  数据库自循环
    (LANE_DB, LANE_SERVICE),            # 5  返回
    (LANE_SERVICE, LANE_FRONTEND),      # 6  返回
    (LANE_FRONTEND, LANE_ADMIN),        # 7  返回
)
TEMPLATE_6_SLOTS: tuple[tuple[str, str], ...] = TEMPLATE_6_ROUND1 + TEMPLATE_6_ROUND1

# template_7: 四泳道 7 条 × 2 轮（管理员/前端/控制器/数据库）
TEMPLATE_7_ROUND1: tuple[tuple[str, str], ...] = (
    (LANE_ADMIN, LANE_FRONTEND),        # 1
    (LANE_FRONTEND, LANE_CONTROLLER),   # 2
    (LANE_CONTROLLER, LANE_DB),         # 3
    (LANE_DB, LANE_DB),                 # 4  数据库自循环
    (LANE_DB, LANE_CONTROLLER),         # 5  返回
    (LANE_CONTROLLER, LANE_FRONTEND),   # 6  返回
    (LANE_FRONTEND, LANE_ADMIN),        # 7  返回
)
TEMPLATE_7_SLOTS: tuple[tuple[str, str], ...] = TEMPLATE_7_ROUND1 + TEMPLATE_7_ROUND1

# 模板说明：用于让用户选择模板（ask 时展示）。lanes 为该模板泳道顺序。
TEMPLATE_INFO: dict[str, dict[str, Any]] = {
    "template_1": {
        "name": "Spring Boot 六层",
        "lanes": ["管理员", "前端", "控制器", "服务层", "Mapper", "数据库"],
        "messages": 22,
        "desc": "含控制器、服务层、Mapper 的标准 Spring Boot 分层，两轮请求-响应",
    },
    "template_2": {
        "name": "Python 六层",
        "lanes": ["用户", "前端", "路由层", "服务层", "仓储层", "存储层"],
        "messages": 22,
        "desc": "含路由层、服务层、仓储层、存储层的 Python 分层，两轮请求-响应",
    },
    "template_4": {
        "name": "Spring Boot 五层",
        "lanes": ["管理员", "前端", "控制器", "服务层", "数据库"],
        "messages": 18,
        "desc": "省去 Mapper，控制器直接访问服务层，两轮请求-响应",
    },
    "template_5": {
        "name": "Python 五层",
        "lanes": ["管理员", "前端", "路由层", "服务层", "数据库"],
        "messages": 18,
        "desc": "省去仓储层/存储层，路由层直接访问服务层，两轮请求-响应",
    },
    "template_6": {
        "name": "通用四泳道（服务层）",
        "lanes": ["管理员", "前端", "服务层", "数据库"],
        "messages": 14,
        "desc": "最简四泳道，前端直连服务层，Spring Boot 与 Python 都能用",
    },
    "template_7": {
        "name": "通用四泳道（控制器）",
        "lanes": ["管理员", "前端", "控制器", "数据库"],
        "messages": 14,
        "desc": "推荐。最简四泳道，前端经控制器访问数据库，Spring Boot 与 Python 都能用",
    },
}

FRONTEND_LANE = "前端"


def first_lane_name(template: str) -> str:
    """模板第一泳道的默认名称（管理员 / 用户）。"""
    return TEMPLATE_INFO[template]["lanes"][0]


SEQUENCE_TEXT_MAX_LEN = 6
_TEMPLATE_VALUES = frozenset(TEMPLATE_INFO)


def normalize_sequence_template(value: Any) -> str | None:
    if not isinstance(value, str):
        return None
    v = value.strip()
    if v in _TEMPLATE_VALUES:
        return v
    return None


def slots_for_template(template: str) -> tuple[tuple[str, str], ...]:
    if template == "template_1":
        return TEMPLATE_1_SLOTS
    if template == "template_2":
        return TEMPLATE_2_SLOTS
    if template == "template_4":
        return TEMPLATE_4_SLOTS
    if template == "template_5":
        return TEMPLATE_5_SLOTS
    if template == "template_6":
        return TEMPLATE_6_SLOTS
    if template == "template_7":
        return TEMPLATE_7_SLOTS
    raise ValueError(
        f'template 必须为 {sorted(_TEMPLATE_VALUES)} 之一，当前: {template!r}'
    )


def template_options() -> list[dict[str, Any]]:
    """给 ask 用：每个模板的名称、泳道、消息数和说明。"""
    return [
        {
            "template": tid,
            "name": info["name"],
            "lanes": " / ".join(info["lanes"]),
            "messages": info["messages"],
            "desc": info["desc"],
        }
        for tid, info in TEMPLATE_INFO.items()
    ]


FRONTEND_LANE_MAX_LEN = 8


def validate_frontend_lane(lane: Any) -> list[str]:
    """前端泳道替换文字：非空、不能含换行、不能就是“前端”本身、不超过 8 个字。"""
    if not isinstance(lane, str) or not lane.strip():
        return ["sequence: data.frontend_lane 须为非空字符串"]
    v = lane.strip()
    if v == FRONTEND_LANE:
        return [
            f"sequence: data.frontend_lane 须替换为具体页面名（如“课程管理页面”），不能仍为 {FRONTEND_LANE!r}"
        ]
    if any(c in "\r\n" for c in v):
        return ["sequence: data.frontend_lane 不能含换行"]
    if len(v) > FRONTEND_LANE_MAX_LEN:
        return [
            f"sequence: data.frontend_lane 至多 {FRONTEND_LANE_MAX_LEN} 个字"
            f"（当前 {len(v)}），请删减"
        ]
    return []


def validate_role(role: Any, template: str) -> list[str]:
    """
    第一泳道角色：必须显式给出且非空、不含换行。
    允许与模板默认名相同（例如图里的角色本来就叫“管理员”），
    这里只保证字段被明确填写，不限制具体取值。
    """
    if not isinstance(role, str) or not role.strip():
        return [
            "sequence: data.role 须显式填写第一泳道的角色名"
            f"（模板默认 {first_lane_name(template)!r}；如与默认相同也请照填）"
        ]
    if any(c in "\r\n" for c in role.strip()):
        return ["sequence: data.role 不能含换行"]
    return []


def validate_sequence_text(index: int, text: Any) -> list[str]:
    """校验单条消息 text；index 为 0 起的消息下标，用于错误定位。"""
    if not isinstance(text, str) or not text.strip():
        return [f"sequence: messages[{index}].text 须为非空字符串"]
    ts = text.strip()
    if len(ts) > SEQUENCE_TEXT_MAX_LEN:
        return [
            f"sequence: messages[{index}].text 至多 {SEQUENCE_TEXT_MAX_LEN} 个字"
            f"（当前 {len(ts)}），请删减"
        ]
    return []


def validate_sequence_data(data: dict[str, Any]) -> list[str]:
    """校验 type=sequence 的 data；返回错误列表（空表示通过）。"""
    errs: list[str] = []

    title = data.get("title")
    if not title or not isinstance(title, str) or not title.strip():
        errs.append("sequence: data.title 必须为非空字符串（用作保存文件名）")
    else:
        forbidden = '<>:"/\\|?*'
        t = title.strip()
        if any(c in forbidden or ord(c) < 32 for c in t):
            errs.append(
                "sequence: data.title 含非法文件名字符（不含 \\ / : * ? \" < > | 等）"
            )

    if "trip" in data:
        errs.append(
            "sequence: 不支持 data.trip，请使用 data.template（见 TEMPLATE_INFO）"
        )

    template = normalize_sequence_template(data.get("template"))
    if template is None:
        errs.append(
            f"sequence: data.template 必须为 {sorted(_TEMPLATE_VALUES)} 之一"
        )
        return errs

    errs.extend(validate_frontend_lane(data.get("frontend_lane")))
    errs.extend(validate_role(data.get("role"), template))

    try:
        expected = slots_for_template(template)
    except ValueError as e:
        errs.append(f"sequence: {e}")
        return errs

    messages = data.get("messages")
    if not isinstance(messages, list):
        errs.append("sequence: data.messages 必须为数组")
        return errs

    n_exp = len(expected)
    if len(messages) != n_exp:
        errs.append(
            f"sequence: data.messages 须恰好 {n_exp} 条（当前 template={template}），当前 {len(messages)} 条"
        )
        return errs

    for i, m in enumerate(messages):
        if not isinstance(m, dict):
            errs.append(f"sequence: messages[{i}] 必须是对象")
            continue
        for k in ("from", "to", "text"):
            v = m.get(k)
            if not isinstance(v, str) or not v.strip():
                errs.append(f"sequence: messages[{i}].{k} 必须为非空字符串")
        errs.extend(validate_sequence_text(i, m.get("text")))

        exp_from, exp_to = expected[i]
        mf, mt = m.get("from"), m.get("to")
        if isinstance(mf, str) and isinstance(mt, str):
            if mf.strip() != exp_from or mt.strip() != exp_to:
                errs.append(
                    f"sequence: messages[{i}] 须为 {exp_from} -> {exp_to}（序号 {i + 1}），"
                    f"当前为 {mf.strip()!r} -> {mt.strip()!r}"
                )

    return errs
