"""
流程图：按模板 1/2/3/4 填充 Visio，并提供字数校验与单形状改字。

占位：节点形状 Text 仅为数字 `1`…`N`（N=模板槽位数；模板1为8，模板2/3/4为7）。
先做空白/全角数字规范化后再解析。

字数限制由代码用 len() 判定（FLOWCHART_TEXT_MAX_LEN=4），不依赖模型判断。
生成后若某个形状字数不对，用 flowchart_set_node 只改该形状，不要重画整张图。

输出：{project_root}/draw/flowchart/{title}.vsdx
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

from .spec import (
    FLOWCHART_TEXT_MAX_LEN,
    normalize_flowchart_template,
    roles_for_template,
    validate_flowchart_content,
    validate_flowchart_text,
)

_TITLE_FORBIDDEN = frozenset('<>:"/\\|?*')

# 全角数字 → 半角，便于识别 Visio 里常见的「１」「２」
_FW_DIGITS = str.maketrans("０１２３４５６７８９", "0123456789")
_PAGE_XML = re.compile(r"^visio/pages/page\d+\.xml$")


def _skill_root() -> Path:
    return Path(__file__).resolve().parents[2]


def _template_path(template: int) -> Path:
    return _skill_root() / "template" / "flowchart" / f"template_{template}.vsdx"


def _normalize_placeholder_raw(text: str) -> str:
    """去掉空白/零宽字符，全角数字转半角。"""
    s = (text or "").strip().translate(_FW_DIGITS)
    return re.sub(r"[\s​ \r\n\t]+", "", s)


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


def _slot_shape_ids(template: int) -> dict[int, str]:
    """模板占位编号 → 形状 ID（只读模板 XML，不修改模板）。"""
    max_slot = len(roles_for_template(template))
    root = _read_page_root(_template_path(template))
    ids: dict[int, str] = {}
    for el in root.iter():
        if _lname(el.tag) != "Shape":
            continue
        sid = el.attrib.get("ID")
        compact = _normalize_placeholder_raw(_xml_text(el))
        if not sid or not re.fullmatch(r"\d+", compact):
            continue
        v = int(compact)
        if 1 <= v <= max_slot and v not in ids:
            ids[v] = sid
    return ids


def _iter_visio_shapes(container: Any) -> Any:
    """深度优先遍历 Page 或组合内的所有 Shape。"""
    try:
        count = int(container.Shapes.Count)
    except Exception:
        return
    for i in range(1, count + 1):
        try:
            sh = container.Shapes.ItemU(i)
        except Exception:
            continue
        yield sh
        try:
            sub = int(sh.Shapes.Count)
        except Exception:
            sub = 0
        if sub > 0:
            yield from _iter_visio_shapes(sh)


def _parse_flowchart_placeholder_index(text: str, max_slot: int) -> int | None:
    """占位序号 1..max_slot：规范化后整段须为十进制数字。"""
    compact = _normalize_placeholder_raw(text)
    if not compact or not re.fullmatch(r"\d+", compact):
        return None
    v = int(compact)
    if 1 <= v <= max_slot:
        return v
    return None


def _validate_flowchart_title(title: str) -> list[str]:
    errs: list[str] = []
    if not title or not isinstance(title, str) or not title.strip():
        errs.append("flowchart: title 必须为非空字符串（用作文件名）")
        return errs
    t = title.strip()
    if any(c in _TITLE_FORBIDDEN or ord(c) < 32 for c in t):
        errs.append('flowchart: title 含非法文件名字符（不含 \\ / : * ? " < > | 等）')
    return errs


def flowchart_check(template: int, nodes: list) -> dict:
    """
    用代码校验 nodes：槽位数、role 顺序、每条文案 ≤ FLOWCHART_TEXT_MAX_LEN 字。
    返回每个节点的实际字数，便于定位需要单改的形状。
    """
    ti = normalize_flowchart_template(template)
    errs = validate_flowchart_content({"template": template, "nodes": nodes})
    lengths: list[int | None] = []
    if isinstance(nodes, list):
        for node in nodes:
            text = node.get("text") if isinstance(node, dict) else None
            lengths.append(len(text.strip()) if isinstance(text, str) else None)
    return {
        "ok": not errs,
        "errors": errs,
        "max_length": FLOWCHART_TEXT_MAX_LEN,
        "lengths": lengths,
        "template": ti,
    }


def _resolve_root(project_root: str | None) -> Path:
    return Path(project_root).resolve() if project_root else Path.cwd().resolve()


def _output_path(project_root: str | None, title: str) -> Path:
    return _resolve_root(project_root) / "draw" / "flowchart" / f"{title.strip()}.vsdx"


def flowchart_verify(project_root: str | None, title: str, template: int) -> dict:
    """
    从已生成的 vsdx 读回每个槽位文字，用代码复核字数。
    返回 shapes: [{index, text, length}]；errors 非空表示需要单形状修改。
    """
    ti = normalize_flowchart_template(template)
    if ti is None:
        return {"ok": False, "errors": ["flowchart: template 须为 1、2、3 或 4"]}
    terr = _validate_flowchart_title(title)
    if terr:
        return {"ok": False, "errors": terr}
    out_path = _output_path(project_root, title)
    if not out_path.exists():
        return {"ok": False, "errors": [f"输出文件不存在: {out_path}"]}

    try:
        id_map = _slot_shape_ids(ti)
        root = _read_page_root(out_path)
        by_id = {
            el.attrib.get("ID"): _xml_text(el)
            for el in root.iter()
            if _lname(el.tag) == "Shape"
        }
    except Exception as e:
        return {"ok": False, "errors": [f"读取输出文件失败: {e}"]}

    shapes: list[dict] = []
    errors: list[str] = []
    for slot in range(1, len(roles_for_template(ti)) + 1):
        text = by_id.get(id_map.get(slot, ""), "")
        length = len(text.strip())
        shapes.append({"index": slot, "text": text, "length": length})
        if length == 0:
            errors.append(f"flowchart: 形状 {slot}（nodes[{slot - 1}]）文字为空")
        elif length > FLOWCHART_TEXT_MAX_LEN:
            errors.append(
                f"flowchart: 形状 {slot}（nodes[{slot - 1}]）{length} 字，超过 {FLOWCHART_TEXT_MAX_LEN} 字"
            )
    return {
        "ok": not errors,
        "errors": errors,
        "max_length": FLOWCHART_TEXT_MAX_LEN,
        "shapes": shapes,
    }


def flowchart_draw(
    project_root: str | None,
    user: str,
    title: str,
    template: int,
    nodes: list,
    overwrite: bool = False,
) -> dict:
    """
    绘制流程图；保存路径固定为 project_root/draw/flowchart/title.vsdx。

    Args:
        project_root: 项目根目录绝对路径
        user: 顶层 user（保留参数，不参与路径）
        title: data.title
        template: 1、2、3 或 4
        nodes: data.nodes（已通过校验）
        overwrite: 目标文件已存在时必须显式为 True（先向用户确认覆盖范围）
    """
    terr = _validate_flowchart_title(title)
    cerr = validate_flowchart_content({"template": template, "nodes": nodes})
    if terr or cerr:
        return {"success": False, "message": "; ".join(terr + cerr)}

    skills_dir = _skill_root()
    template_dir = skills_dir / "template" / "flowchart"
    ti = normalize_flowchart_template(template)
    if ti is None:
        return {"success": False, "message": "flowchart: template 须为 1、2、3 或 4"}
    template_path = str(_template_path(ti))
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

    lines: list[str] = []
    for node in nodes:
        raw = node.get("text", "")
        lines.append(raw.strip() if isinstance(raw, str) else "")

    max_slot = len(lines)
    pythoncom.CoInitialize()
    visio = None
    doc = None
    try:
        out_path.parent.mkdir(parents=True, exist_ok=True)
        # Always edit a copied working file; never modify template files in-place.
        shutil.copy2(template_path, abs_output)

        visio = win32com.client.Dispatch("Visio.Application")
        visio.Visible = False
        doc = visio.Documents.Open(abs_output)
        page = doc.Pages.ItemU(1)

        indexed: list[tuple[int, Any]] = []
        for shape in _iter_visio_shapes(page):
            try:
                raw = shape.Text or ""
            except Exception:
                continue
            idx = _parse_flowchart_placeholder_index(raw, max_slot)
            if idx is None:
                continue
            indexed.append((idx, shape))

        replaced = 0
        for idx, shape in indexed:
            pos = idx - 1
            if not (0 <= pos < len(lines)):
                continue
            try:
                shape.Text = lines[pos]
                replaced += 1
            except Exception:
                pass

        if replaced == 0:
            try:
                doc.Close()
            except Exception:
                try:
                    doc.Close()
                except Exception:
                    pass
            try:
                visio.Quit()
            except Exception:
                pass
            try:
                os.remove(abs_output)
            except OSError:
                pass
            return {
                "success": False,
                "message": (
                    f"未替换任何文本：未找到占位（形状 Text 须为 1～{max_slot} 的纯数字，"
                    f"可在组合内子形状上）。已取消生成输出文件。"
                ),
            }

        doc.Save()
        doc.Close()
        visio.Quit()
        return {"success": True, "message": "已保存"}
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


def flowchart_set_node(
    project_root: str | None,
    title: str,
    template: int,
    index: int,
    text: str,
) -> dict:
    """
    只修改已生成图中的某个形状文字（index 从 1 开始，对应槽位号）。
    校验该文字字数与 role 规则后落盘，不重新生成整张图。
    """
    ti = normalize_flowchart_template(template)
    if ti is None:
        return {"success": False, "message": "flowchart: template 须为 1、2、3 或 4"}
    terr = _validate_flowchart_title(title)
    if terr:
        return {"success": False, "message": "; ".join(terr)}

    roles = roles_for_template(ti)
    if not isinstance(index, int) or not 1 <= index <= len(roles):
        return {
            "success": False,
            "message": f"flowchart: index 须为 1～{len(roles)}，当前 {index!r}",
        }

    errs = validate_flowchart_text(index - 1, roles[index - 1], text)
    if errs:
        return {"success": False, "message": "; ".join(errs)}

    out_path = _output_path(project_root, title)
    if not out_path.exists():
        return {"success": False, "message": f"输出文件不存在: {out_path}"}

    target_id = _slot_shape_ids(ti).get(index)
    if target_id is None:
        return {"success": False, "message": f"模板中未找到形状 {index} 的占位"}

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

        found.Text = text.strip()
        doc.Save()
        doc.Close()
        visio.Quit()
        return {"success": True, "message": f"形状 {index} 已改为 {text.strip()!r}"}
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
    parser = argparse.ArgumentParser(description="Flowchart check/draw/verify/set-text")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_check = sub.add_parser("check", help="校验 JSON 中每个形状字数")
    p_check.add_argument("--json", required=True)

    p_draw = sub.add_parser("draw", help="按 JSON 生成 vsdx")
    p_draw.add_argument("--json", required=True)
    p_draw.add_argument("--project-root", help="项目根；省略时使用当前工作目录")
    p_draw.add_argument("--overwrite", action="store_true", help="目标文件已存在时允许覆盖")

    p_verify = sub.add_parser("verify", help="从已生成的 vsdx 复核字数")
    p_verify.add_argument("--json", required=True)
    p_verify.add_argument("--project-root", help="项目根；省略时使用当前工作目录")

    p_set = sub.add_parser("set-text", help="只修改图中某个形状的文字")
    p_set.add_argument("--project-root", help="项目根；省略时使用当前工作目录")
    p_set.add_argument("--title", required=True)
    p_set.add_argument("--template", required=True, type=int)
    p_set.add_argument("--index", required=True, type=int)
    p_set.add_argument("--text", required=True)

    args = parser.parse_args()
    if args.cmd == "check":
        data = _load_json(args.json)
        result = flowchart_check(data.get("template"), data.get("nodes"))
    elif args.cmd == "draw":
        data = _load_json(args.json)
        result = flowchart_draw(
            args.project_root,
            data.get("user", ""),
            data.get("title", ""),
            data.get("template"),
            data.get("nodes"),
            overwrite=args.overwrite,
        )
    elif args.cmd == "verify":
        data = _load_json(args.json)
        result = flowchart_verify(args.project_root, data.get("title", ""), data.get("template"))
    else:
        result = flowchart_set_node(
            args.project_root, args.title, args.template, args.index, args.text
        )
    print(json.dumps(result, ensure_ascii=False))
    if not (result.get("ok") or result.get("success")):
        raise SystemExit(1)


if __name__ == "__main__":
    main()
