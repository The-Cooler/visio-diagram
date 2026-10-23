"""
ER 图工具 - 从 JSON 文件读取实体数据，绘制 ER 实体到 Visio 文件。

JSON 路径: {project_root}/json/er/{entity}.json
输出路径: {project_root}/draw/er/{entity}.vsdx
project_root 未指定时使用当前工作目录。
"""

import json as _json
import os
import shutil
from pathlib import Path

import pythoncom
import win32com.client

from .er_contract import normalize_attrs, validate_er_payload

ER_TEMPLATE_ATTR_SLOTS = 10


def _resolve_root(project_root: str | None) -> Path:
    return Path(project_root).resolve() if project_root else Path.cwd().resolve()


def er_draw(project_root: str | None, json_path: str, user: str = None) -> dict:
    """从 JSON 文件读取实体数据，绘制 ER 实体图。

    JSON 格式: {"entity": "实体名", "attributes": [{"name": "属性", "type": "pk"}, ...]}
    JSON 位置: {project_root}/json/er/{entity}.json
    输出位置: {project_root}/draw/er/{entity}.vsdx
    project_root 为 None 时使用当前工作目录；user 参数保留兼容，不参与路径。
    """
    rp = _resolve_root(project_root)
    jp = Path(json_path)

    if not jp.exists():
        return {"success": False, "message": f"JSON 文件不存在: {json_path}"}
    try:
        payload = _json.loads(jp.read_text(encoding="utf-8-sig"))
    except _json.JSONDecodeError as e:
        return {"success": False, "message": f"JSON 解析失败: {e}"}

    entity = payload.get("entity")
    attributes = payload.get("attributes")
    if not entity or not attributes:
        return {"success": False, "message": "JSON 须包含 entity 和 attributes"}

    errors = validate_er_payload(payload)
    if errors:
        return {"success": False, "message": "; ".join(errors)}

    ent = entity.strip()
    template_dir = Path(__file__).resolve().parents[2] / "template" / "er"
    out_dir = rp / "draw" / "er"
    abs_output = str(out_dir / f"{ent}.vsdx")

    draw_attrs = normalize_attrs(attributes)
    template_name = "ER_odd.vsdx" if len(attributes) % 2 == 1 else "ER_even.vsdx"
    template_path = str(template_dir / template_name)
    if not os.path.exists(template_path):
        return {"success": False, "message": f"模板不存在: {template_path}"}

    pythoncom.CoInitialize()
    visio = None
    doc = None
    try:
        out_dir.mkdir(parents=True, exist_ok=True)
        # 复制后用副本落图，模板原件始终不动
        shutil.copy2(template_path, abs_output)

        visio = win32com.client.Dispatch("Visio.Application")
        visio.Visible = False
        doc = visio.Documents.Open(abs_output)
        page = doc.Pages.ItemU(1)

        entities = []
        attrs = []
        for i in range(1, page.Shapes.Count + 1):
            shape = page.Shapes.ItemU(i)
            if "流程图过程" in shape.Name:
                entities.append(shape)
            elif "ER图-属性" in shape.Name:
                attrs.append(shape)

        def get_pos(s):
            try:
                return (s.get_Cells("PinY").ResultIU, s.get_Cells("PinX").ResultIU)
            except Exception:
                return (0, 0)

        def get_num(s):
            try:
                return int(s.Text)
            except Exception:
                return 0

        entities.sort(key=get_pos)
        attrs.sort(key=get_num)

        if entities:
            entities[0].Text = ent

        for i, attr_info in enumerate(draw_attrs):
            if i < len(attrs):
                attrs[i].Text = attr_info.get("name", "")

        for i in range(len(draw_attrs), len(attrs)):
            attrs[i].Delete()

        doc.Save()
        doc.Close()
        doc = None
        visio.Quit()
        visio = None

        return {"success": True, "message": "已保存"}

    except Exception as e:
        if doc is not None:
            try:
                doc.Close()
            except Exception:
                pass
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
