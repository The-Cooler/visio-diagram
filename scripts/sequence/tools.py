"""
时序图工具：按 template/sequence/template_1/2/4/5/6/7 填充 Visio。

  模板（须自备）：
    template_1.vsdx — 22 条（Spring Boot 六层）
    template_2.vsdx — 22 条（Python 六层）
    template_4.vsdx — 18 条（Spring Boot 五层）
    template_5.vsdx — 18 条（Python 五层）
    template_6.vsdx — 14 条（通用四泳道：管理员/前端/服务层/数据库）
    template_7.vsdx — 14 条（通用四泳道：管理员/前端/控制器/数据库）

占位文案形如 "1：消息1" … "N：消息N"；写入时替换为 "N：<text>"。

前端泳道名称会按 data.frontend_lane 替换为具体页面名（如“课程管理页面”），
只改泳道文案，不改样式、尺寸和其余泳道。

字数限制由代码用 len() 判定（SEQUENCE_TEXT_MAX_LEN=6），不依赖模型判断。
生成后若某条消息字数不对，用 sequence_set_message 只改该条，不要重画整张图。

输出路径（代码内固定）：{project_root}/draw/sequence/{title}.vsdx
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

import pythoncom
import win32com.client

from .slots import (
    FRONTEND_LANE,
    SEQUENCE_TEXT_MAX_LEN,
    first_lane_name,
    normalize_sequence_template,
    slots_for_template,
    template_options,
    validate_frontend_lane,
    validate_role,
    validate_sequence_data,
    validate_sequence_text,
)

_PAGE_XML = re.compile(r"^visio/pages/page\d+\.xml$")
_PLACEHOLDER = re.compile(r"^(\d+)[.：]\s*消息\d+$")
_TITLE_FORBIDDEN = frozenset('<>:"/\\|?*')


def _lane_text_re(lane_name: str) -> re.Pattern[str]:
    """泳道形状的 Text 形如 ":前端"（可带换行），匹配冒号前缀 + 该泳道名。"""
    return re.compile(r"^\s*:?\s*" + re.escape(lane_name) + r"\s*$")


def _skill_root() -> Path:
    return Path(__file__).resolve().parents[2]


def _lname(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _xml_text(shape: ET.Element) -> str:
    for child in shape:
        if _lname(child.tag) == "Text":
            return "".join(child.itertext()).strip()
    return ""


def _read_page_root(path: Path) -> ET.Element:
    with zipfile.ZipFile(path) as zf:
        names = sorted(n for n in zf.namelist() if _PAGE_XML.fullmatch(n))
        if not names:
            raise ValueError(f"{path} 内未找到页面 XML")
        return ET.fromstring(zf.read(names[0]))


def _slot_shape_ids(template: str) -> dict[int, str]:
    """模板占位编号 → 形状 ID（只读模板 XML，不修改模板）。"""
    max_slot = len(slots_for_template(template))
    root = _read_page_root(_skill_root() / "template" / "sequence" / f"{template}.vsdx")
    ids: dict[int, str] = {}
    for el in root.iter():
        if _lname(el.tag) != "Shape":
            continue
        sid = el.attrib.get("ID")
        m = _PLACEHOLDER.match(_xml_text(el))
        if not sid or not m:
            continue
        v = int(m.group(1))
        if 1 <= v <= max_slot and v not in ids:
            ids[v] = sid
    return ids


def _lane_shape_ids(template: str) -> dict[str, str]:
    """泳道名 → 形状 ID（只读模板 XML，不修改模板）。"""
    root = _read_page_root(_skill_root() / "template" / "sequence" / f"{template}.vsdx")
    ids: dict[str, str] = {}
    for el in root.iter():
        if _lname(el.tag) != "Shape":
            continue
        if el.attrib.get("Type") != "Group":
            continue
        sid = el.attrib.get("ID")
        text = _xml_text(el).strip().lstrip(":").strip()
        if sid and text and text not in ids:
            ids[text] = sid
    return ids


def _validate_title(title: Any) -> list[str]:
    if not title or not isinstance(title, str) or not title.strip():
        return ["sequence: data.title 必须为非空字符串（用作保存文件名）"]
    t = title.strip()
    if any(c in _TITLE_FORBIDDEN or ord(c) < 32 for c in t):
        return ['sequence: data.title 含非法文件名字符（不含 \\ / : * ? " < > | 等）']
    return []


def _resolve_root(project_root: str | None) -> Path:
    return Path(project_root).resolve() if project_root else Path.cwd().resolve()


def _output_path(project_root: str | None, title: str) -> Path:
    return _resolve_root(project_root) / "draw" / "sequence" / f"{title.strip()}.vsdx"


def _iter_visio_shapes(container: Any) -> Any:
    try:
        for i in range(1, int(container.Shapes.Count) + 1):
            sh = container.Shapes.ItemU(i)
            yield sh
            try:
                if int(sh.Shapes.Count) > 0:
                    yield from _iter_visio_shapes(sh)
            except Exception:
                continue
    except Exception:
        return


def sequence_check(
    template: str,
    messages: list,
    frontend_lane: str | None = None,
    role: str | None = None,
) -> dict:
    """用代码校验 messages：流向、条数、每条文案 ≤ SEQUENCE_TEXT_MAX_LEN 字、前端泳道名与第一泳道角色。"""
    errs = validate_sequence_data(
        {
            "title": "check",
            "template": template,
            "messages": messages,
            "frontend_lane": frontend_lane,
            "role": role,
        }
    )
    lengths: list[int | None] = []
    if isinstance(messages, list):
        for m in messages:
            text = m.get("text") if isinstance(m, dict) else None
            lengths.append(len(text.strip()) if isinstance(text, str) else None)
    return {
        "ok": not errs,
        "errors": errs,
        "max_length": SEQUENCE_TEXT_MAX_LEN,
        "lengths": lengths,
        "frontend_lane": (frontend_lane or "").strip() or None,
        "role": (role or "").strip() or None,
    }


def sequence_verify(
    project_root: str | None,
    title: str,
    template: str,
    frontend_lane: str | None = None,
    role: str | None = None,
) -> dict:
    """
    从已生成的 vsdx 读回每条消息文字，用代码复核字数；给了 frontend_lane 时一并复核泳道名。
    返回 messages: [{index, text, length}]；errors 非空表示需要单条修改。
    """
    resolved = normalize_sequence_template(template)
    if resolved is None:
        return {"ok": False, "errors": ["sequence: template 非法"]}
    terr = _validate_title(title)
    if terr:
        return {"ok": False, "errors": terr}
    out_path = _output_path(project_root, title)
    if not out_path.exists():
        return {"ok": False, "errors": [f"输出文件不存在: {out_path}"]}

    try:
        id_map = _slot_shape_ids(resolved)
        root = _read_page_root(out_path)
        by_id = {
            el.attrib.get("ID"): _xml_text(el)
            for el in root.iter()
            if _lname(el.tag) == "Shape"
        }
    except Exception as e:
        return {"ok": False, "errors": [f"读取输出文件失败: {e}"]}

    rows: list[dict] = []
    errors: list[str] = []
    for slot in range(1, len(slots_for_template(resolved)) + 1):
        text = by_id.get(id_map.get(slot, ""), "")
        body = re.sub(r"^\d+[.：]\s*", "", text).strip()
        length = len(body)
        rows.append({"index": slot, "text": body, "length": length})
        if length == 0:
            errors.append(f"sequence: 消息 {slot}（messages[{slot - 1}]）文字为空")
        elif length > SEQUENCE_TEXT_MAX_LEN:
            errors.append(
                f"sequence: 消息 {slot}（messages[{slot - 1}]）{length} 字，超过 {SEQUENCE_TEXT_MAX_LEN} 字"
            )

    lane_check: dict | None = None
    if frontend_lane:
        want = frontend_lane.strip()
        lane_id = _lane_shape_ids(resolved).get(FRONTEND_LANE)
        actual = by_id.get(lane_id or "", "").strip().lstrip(":").strip()
        lane_check = {"expected": want, "actual": actual}
        if actual != want:
            errors.append(f"sequence: 前端泳道应为 {want!r}，实际为 {actual!r}")

    role_check: dict | None = None
    if role:
        want = role.strip()
        lane_id = _lane_shape_ids(resolved).get(first_lane_name(resolved))
        actual = by_id.get(lane_id or "", "").strip().lstrip(":").strip()
        role_check = {"expected": want, "actual": actual}
        if actual != want:
            errors.append(f"sequence: 第一泳道角色应为 {want!r}，实际为 {actual!r}")

    return {
        "ok": not errors,
        "errors": errors,
        "max_length": SEQUENCE_TEXT_MAX_LEN,
        "messages": rows,
        "frontend_lane": lane_check,
        "role": role_check,
    }


def sequence_draw(
    project_root: str | None,
    user: str,
    title: str,
    messages: list,
    template: str,
    frontend_lane: str | None = None,
    role: str | None = None,
    overwrite: bool = False,
) -> dict:
    """
    绘制时序图；保存路径固定为 project_root/draw/sequence/title.vsdx。

    Args:
        project_root: 项目根目录绝对路径（与用户消息一致）
        user: 顶层请求中的 user（保留参数，不参与路径）
        title: 图标题（文件名主体，与校验通过的 data.title 一致）
        messages: 与校验通过的 data.messages 一致（每条含 from、to、text）
        template: "template_1" / "template_2" / "template_4" / "template_5" / "template_6" / "template_7"
        frontend_lane: 前端泳道显示名（如“课程管理页面”），只替换泳道文案，不动样式
        role: 第一泳道显示的角色名（如“教师”），只替换泳道文案，不动样式
        overwrite: 目标文件已存在时必须显式为 True（先向用户确认覆盖范围）
    """
    verr = validate_sequence_data(
        {
            "title": title,
            "template": template,
            "messages": messages,
            "frontend_lane": frontend_lane,
            "role": role,
        }
    )
    if verr:
        return {"success": False, "message": "; ".join(verr)}
    resolved_template = normalize_sequence_template(template)
    if resolved_template is None:
        return {
            "success": False,
            "message": "sequence: template 非法",
        }
    template_path = str(
        _skill_root() / "template" / "sequence" / f"{resolved_template}.vsdx"
    )

    if not os.path.exists(template_path):
        return {
            "success": False,
            "message": f"模板不存在: {template_path}",
        }

    out_path = _output_path(project_root, title)
    abs_output = str(out_path)
    if out_path.exists() and not overwrite:
        return {
            "success": False,
            "message": f"目标文件已存在: {abs_output}；确认覆盖后再加 --overwrite",
        }

    # 统一使用中文冒号格式：N：消息文本
    sep = "："
    lines: list[str] = []
    for i, m in enumerate(messages):
        raw = m.get("text", "")
        text = raw.strip() if isinstance(raw, str) else ""
        lines.append(f"{i + 1}{sep}{text}")

    pythoncom.CoInitialize()
    visio = None
    try:
        out_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(template_path, abs_output)

        visio = win32com.client.Dispatch("Visio.Application")
        visio.Visible = False
        doc = visio.Documents.Open(abs_output)
        page = doc.Pages.ItemU(1)

        replaced = 0
        lane_replaced = {"role": 0, "frontend": 0}
        role_default = first_lane_name(resolved_template)
        role_re = _lane_text_re(role_default)
        frontend_re = _lane_text_re(FRONTEND_LANE)
        for shape in _iter_visio_shapes(page):
            try:
                original = shape.Text or ""
            except Exception:
                continue
            m = _PLACEHOLDER.match(original.strip())
            if m:
                idx = int(m.group(1)) - 1
                if 0 <= idx < len(lines):
                    shape.Text = lines[idx]
                    replaced += 1
                continue
            # 只替换泳道文案，保留冒号前缀与原有格式
            if role and role_re.match(original):
                shape.Text = original.replace(role_default, role.strip(), 1)
                lane_replaced["role"] += 1
            elif frontend_lane and frontend_re.match(original):
                shape.Text = original.replace(FRONTEND_LANE, frontend_lane.strip(), 1)
                lane_replaced["frontend"] += 1

        if replaced == 0:
            doc.Close()
            visio.Quit()
            try:
                os.remove(abs_output)
            except OSError:
                pass
            return {
                "success": False,
                "message": "未替换任何文本：未找到“N：消息N”占位。已取消生成输出文件。",
            }

        missing = [
            name
            for name, count in (
                (f"第一泳道 {role_default!r}", lane_replaced["role"]),
                (f"前端泳道 {FRONTEND_LANE!r}", lane_replaced["frontend"]),
            )
            if count == 0
        ]
        if missing:
            doc.Close()
            visio.Quit()
            try:
                os.remove(abs_output)
            except OSError:
                pass
            return {
                "success": False,
                "message": f"未找到{'、'.join(missing)}。已取消生成输出文件。",
            }

        doc.Save()
        doc.Close()
        visio.Quit()

        return {
            "success": True,
            "message": f"已保存（template={resolved_template}）",
        }

    except Exception as e:
        if visio is not None and hasattr(visio, "Quit"):
            try:
                visio.Quit()
            except Exception:
                pass
        if os.path.exists(abs_output):
            try:
                os.remove(abs_output)
            except OSError:
                pass
        err_msg = str(e)
        if "DiskFull" in err_msg or "Permission" in err_msg:
            return {
                "success": False,
                "message": f"保存失败: {err_msg}，可能磁盘已满或文件被占用",
            }
        if "无法运行" in err_msg or "unexpected" in err_msg.lower():
            return {
                "success": False,
                "message": "Visio 操作失败，可能模板文件已损坏或正在被其他程序占用",
            }
        return {"success": False, "message": f"失败: {e}"}

    finally:
        pythoncom.CoUninitialize()


def sequence_set_message(
    project_root: str | None,
    title: str,
    template: str,
    index: int,
    text: str,
) -> dict:
    """
    只修改已生成图中的某条消息文字（index 从 1 开始，对应 messages 下标 index-1）。
    校验字数后落盘，不重新生成整张图。
    """
    resolved = normalize_sequence_template(template)
    if resolved is None:
        return {"success": False, "message": "sequence: template 非法"}
    terr = _validate_title(title)
    if terr:
        return {"success": False, "message": "; ".join(terr)}

    n_exp = len(slots_for_template(resolved))
    if not isinstance(index, int) or not 1 <= index <= n_exp:
        return {
            "success": False,
            "message": f"sequence: index 须为 1～{n_exp}，当前 {index!r}",
        }

    errs = validate_sequence_text(index - 1, text)
    if errs:
        return {"success": False, "message": "; ".join(errs)}

    out_path = _output_path(project_root, title)
    if not out_path.exists():
        return {"success": False, "message": f"输出文件不存在: {out_path}"}

    target_id = _slot_shape_ids(resolved).get(index)
    if target_id is None:
        return {"success": False, "message": f"模板中未找到消息 {index} 的占位"}

    abs_output = str(out_path)
    pythoncom.CoInitialize()
    visio = None
    try:
        visio = win32com.client.Dispatch("Visio.Application")
        visio.Visible = False
        doc = visio.Documents.Open(abs_output)
        page = doc.Pages.ItemU(1)

        found = None
        for shape in _iter_visio_shapes(page):
            try:
                if str(shape.ID) == str(target_id):
                    found = shape
                    break
            except Exception:
                continue
        if found is None:
            doc.Close()
            visio.Quit()
            return {"success": False, "message": f"输出图中未找到形状 ID {target_id}"}

        found.Text = f"{index}：{text.strip()}"
        doc.Save()
        doc.Close()
        visio.Quit()
        return {"success": True, "message": f"消息 {index} 已改为 {text.strip()!r}"}
    except Exception as e:
        if visio is not None and hasattr(visio, "Quit"):
            try:
                visio.Quit()
            except Exception:
                pass
        return {"success": False, "message": f"失败: {e}"}
    finally:
        pythoncom.CoUninitialize()


def _load_json(path: str) -> dict:
    return json.loads(Path(path).read_text(encoding="utf-8"))


def main() -> None:
    parser = argparse.ArgumentParser(description="Sequence templates/check/draw/verify/set-text")
    sub = parser.add_subparsers(dest="cmd", required=True)

    sub.add_parser("templates", help="列出所有模板的泳道、消息数和说明")

    p_check = sub.add_parser("check", help="校验 JSON 中每条消息字数、流向和前端泳道名")
    p_check.add_argument("--json", required=True)

    p_draw = sub.add_parser("draw", help="按 JSON 生成 vsdx")
    p_draw.add_argument("--json", required=True)
    p_draw.add_argument("--project-root", help="项目根；省略时使用当前工作目录")
    p_draw.add_argument("--overwrite", action="store_true", help="目标文件已存在时允许覆盖")

    p_verify = sub.add_parser("verify", help="从已生成的 vsdx 复核字数")
    p_verify.add_argument("--json", required=True)
    p_verify.add_argument("--project-root", help="项目根；省略时使用当前工作目录")

    p_set = sub.add_parser("set-text", help="只修改图中某条消息的文字")
    p_set.add_argument("--project-root", help="项目根；省略时使用当前工作目录")
    p_set.add_argument("--title", required=True)
    p_set.add_argument("--template", required=True)
    p_set.add_argument("--index", required=True, type=int)
    p_set.add_argument("--text", required=True)

    args = parser.parse_args()
    if args.cmd == "templates":
        result = {"ok": True, "templates": template_options()}
    elif args.cmd == "check":
        data = _load_json(args.json)
        result = sequence_check(
            data.get("template"), data.get("messages"), data.get("frontend_lane"), data.get("role")
        )
    elif args.cmd == "draw":
        data = _load_json(args.json)
        result = sequence_draw(
            args.project_root,
            data.get("user", ""),
            data.get("title", ""),
            data.get("messages"),
            data.get("template"),
            data.get("frontend_lane"),
            data.get("role"),
            overwrite=args.overwrite,
        )
    elif args.cmd == "verify":
        data = _load_json(args.json)
        result = sequence_verify(
            args.project_root,
            data.get("title", ""),
            data.get("template"),
            data.get("frontend_lane"),
            data.get("role"),
        )
    else:
        result = sequence_set_message(
            args.project_root, args.title, args.template, args.index, args.text
        )
    print(json.dumps(result, ensure_ascii=False))
    if not (result.get("ok") or result.get("success")):
        raise SystemExit(1)


if __name__ == "__main__":
    main()
