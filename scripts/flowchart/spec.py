"""
流程图模板 1 / 2 / 3 / 4 的节点顺序与 role 约定（按 template/flowchart/*.vsdx 实际形状确认）。

模板1（8节点）：开始 → 处理 → 处理 → 判断 → 处理(是) → 处理(否) → 处理(汇合) → 结束
模板2（7节点）：开始 → 输入 → 处理 → 判断 → 处理(是) → 处理 → 结束
模板3（7节点）：开始 → 处理 → 处理 → 判断 → 处理(是) → 处理 → 结束
模板4（7节点）：开始 → 处理 → 处理 → 判断 → 处理(是) → 处理 → 结束
"""

from __future__ import annotations

from typing import Any
import re

FLOWCHART_TEXT_MAX_LEN = 4
_FLOWCHART_ALLOWED_TEXT = re.compile(r"^[一-鿿A-Za-z0-9]+$")

# 输入类动作词：出现在 input 槽位是必须的，出现在 process（矩形）槽位则说明放错了形状
INPUT_VERBS: tuple[str, ...] = ("编辑", "写入", "输入", "填写", "录入", "添加", "新建", "创建")

# 与 Visio 模板占位顺序一致：第 i 个可替换形状对应 nodes[i]
TEMPLATE_1_ROLES: tuple[str, ...] = (
    "start",  # 圆角·开始
    "process",  # 矩形·处理
    "process",
    "decision",  # 菱形·判断
    "process",  # 是分支
    "process",  # 否分支
    "process",  # 汇合
    "end",  # 圆角·结束
)

TEMPLATE_2_ROLES: tuple[str, ...] = (
    "start",  # 圆角·开始
    "input",  # 平行四边形·输入
    "process",  # 矩形·处理
    "decision",  # 菱形·判断
    "process",  # 是分支
    "process",
    "end",  # 圆角·结束
)

TEMPLATE_3_ROLES: tuple[str, ...] = (
    "start",  # 圆角·开始
    "process",  # 矩形·处理
    "process",  # 矩形·处理
    "decision",  # 菱形·判断
    "process",  # 是分支
    "process",
    "end",  # 圆角·结束
)

TEMPLATE_4_ROLES: tuple[str, ...] = (
    "start",  # 圆角·开始
    "process",  # 矩形·处理
    "process",  # 矩形·处理
    "decision",  # 菱形·判断
    "process",  # 是分支
    "process",
    "end",  # 圆角·结束
)

ROLE_SHAPE_HINT: dict[str, str] = {
    "start": "圆角矩形·开始",
    "end": "圆角矩形·结束",
    "input": "平行四边形·输入",
    "process": "矩形·处理",
    "decision": "菱形·判断",
}


def roles_for_template(template: int) -> tuple[str, ...]:
    if template == 1:
        return TEMPLATE_1_ROLES
    if template == 2:
        return TEMPLATE_2_ROLES
    if template == 3:
        return TEMPLATE_3_ROLES
    if template == 4:
        return TEMPLATE_4_ROLES
    raise ValueError("template 须为 1、2、3 或 4")


def normalize_flowchart_template(template: Any) -> int | None:
    if template == 1 or template == "1":
        return 1
    if template == 2 or template == "2":
        return 2
    if template == 3 or template == "3":
        return 3
    if template == 4 or template == "4":
        return 4
    return None


def validate_flowchart_text(index: int, role: str, text: Any) -> list[str]:
    """校验单个节点的 text；index 为 0 起的节点下标，用于错误定位。"""
    errs: list[str] = []
    if not isinstance(text, str) or not text.strip():
        errs.append(f"flowchart: nodes[{index}].text 须为非空字符串")
        return errs

    ts = text.strip()
    if len(ts) > FLOWCHART_TEXT_MAX_LEN:
        errs.append(
            f"flowchart: nodes[{index}].text 至多 {FLOWCHART_TEXT_MAX_LEN} 个字（当前 {len(ts)}），请删减"
        )
    if not _FLOWCHART_ALLOWED_TEXT.fullmatch(ts):
        errs.append(
            f"flowchart: nodes[{index}].text 不允许符号或空白，仅允许中文/英文/数字，当前为 {ts!r}"
        )
    if role == "decision" and not ts.startswith("是否"):
        errs.append(
            f'flowchart: nodes[{index}].text 判断框必须为判断句式，且以"是否"开头（如"是否登录"），当前为 {ts!r}'
        )
    if role == "input":
        has_verb = any(kw in ts for kw in INPUT_VERBS)
        if not has_verb:
            errs.append(
                f'flowchart: nodes[{index}].text 输入框必须包含动作动词（如"编辑"、"写入"、"输入"等），当前为 {ts!r}'
            )
    if role == "process":
        matched = [kw for kw in INPUT_VERBS if kw in ts]
        if matched:
            errs.append(
                f'flowchart: nodes[{index}].text 是输入类内容（含"{"、".join(matched)}"），不能放在矩形处理框；'
                f"模板有输入框（平行四边形）时移到该槽位，没有输入框时改成非录入类描述，当前为 {ts!r}"
            )
    return errs


def validate_flowchart_content(data: dict[str, Any]) -> list[str]:
    """
    校验 data 中的 template、nodes（不含 title）。
    nodes[i].role 须与模板第 i 个槽位一致；text 字数由 len() 判定。
    """
    errs: list[str] = []
    t_raw = data.get("template")
    t = normalize_flowchart_template(t_raw)
    if t is None:
        errs.append("flowchart: data.template 须为 1、2、3 或 4")
        return errs

    try:
        expected = roles_for_template(t)
    except ValueError as e:
        errs.append(f"flowchart: {e}")
        return errs

    nodes = data.get("nodes")
    if not isinstance(nodes, list):
        errs.append("flowchart: data.nodes 必须为数组")
        return errs

    n_exp = len(expected)
    if len(nodes) != n_exp:
        errs.append(f"flowchart: 模板{t} 须恰好 {n_exp} 个节点，当前 {len(nodes)} 个")
        return errs

    for i, node in enumerate(nodes):
        if not isinstance(node, dict):
            errs.append(f"flowchart: nodes[{i}] 必须是对象")
            continue
        role = node.get("role")
        exp = expected[i]
        if not isinstance(role, str) or role.strip() != exp:
            hint = ROLE_SHAPE_HINT.get(exp, exp)
            errs.append(
                f'flowchart: nodes[{i}].role 须为 "{exp}"（{hint}），'
                f"不可与菱形判断/输入/矩形处理槽位混用；当前 {role!r}"
            )
        errs.extend(validate_flowchart_text(i, exp, node.get("text")))

    return errs
