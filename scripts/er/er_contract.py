from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
from pathlib import Path
from typing import Any

ER_TEMPLATE_ATTR_SLOTS = 10
_CJK = re.compile(r"[\u4e00-\u9fff]")


def has_cjk(text: str) -> bool:
    return bool(_CJK.search(text))


def is_id_exempt(name: str) -> bool:
    return name.strip().casefold() == "id"


def choose_template_name(attr_count: int) -> str:
    return "ER_odd.vsdx" if attr_count % 2 == 1 else "ER_even.vsdx"


def validate_er_payload(payload: dict[str, Any]) -> list[str]:
    errs: list[str] = []
    entity = payload.get("entity")
    attrs = payload.get("attributes")

    if not isinstance(entity, str) or not entity.strip():
        errs.append("entity must be a non-empty string")
    elif not has_cjk(entity.strip()):
        errs.append("entity must contain CJK")

    if not isinstance(attrs, list):
        errs.append("attributes must be an array")
        return errs
    if len(attrs) == 0:
        errs.append("attributes must not be empty")
        return errs
    if len(attrs) > ER_TEMPLATE_ATTR_SLOTS:
        errs.append(f"attributes max {ER_TEMPLATE_ATTR_SLOTS}")

    pk_count = 0
    for idx, item in enumerate(attrs):
        if not isinstance(item, dict):
            errs.append(f"attributes[{idx}] must be an object")
            continue
        name = item.get("name")
        if not isinstance(name, str) or not name.strip():
            errs.append(f"attributes[{idx}].name must be non-empty string")
        elif not is_id_exempt(name) and not has_cjk(name.strip()):
            errs.append(f"attributes[{idx}].name must contain CJK or be ID/id")
        if item.get("type") == "pk":
            pk_count += 1
    if pk_count > 1:
        errs.append("only one attribute can be pk")
    return errs


def normalize_attrs(attrs: list[dict[str, Any]]) -> list[dict[str, Any]]:
    pk = None
    rest: list[dict[str, Any]] = []
    for item in attrs:
        if item.get("type") == "pk" and pk is None:
            pk = item
        else:
            rest.append(item)
    ordered = ([pk] if pk else []) + rest
    return ordered[:ER_TEMPLATE_ATTR_SLOTS]


def validate_diagram_request(request: str | dict[str, Any]) -> dict[str, Any]:
    if isinstance(request, str):
        try:
            obj = json.loads(request.lstrip("\ufeff"))
        except json.JSONDecodeError as exc:
            return {
                "valid": False,
                "errors": [f"JSON parse failed: {exc}"],
                "parsed": None,
                "hint": "Pass a valid top-level ER JSON object.",
            }
    else:
        obj = request

    if not isinstance(obj, dict):
        return {
            "valid": False,
            "errors": ["root must be a JSON object"],
            "parsed": None,
            "hint": "Root must contain user, type, and data.",
        }

    errors: list[str] = []
    user = obj.get("user")
    if not isinstance(user, str) or not user.strip():
        errors.append("user must be a non-empty string")

    dtype = obj.get("type")
    if dtype != "er":
        errors.append('type must be "er"')

    data = obj.get("data")
    if not isinstance(data, dict):
        errors.append("data must be an object")
        data = {}

    if dtype == "er":
        errors.extend(validate_er_payload(data))

    valid = len(errors) == 0
    return {
        "valid": valid,
        "errors": errors,
        "parsed": obj,
        "hint": (
            "ER request is valid; call er_draw(project_root, user, entity, attributes)."
            if valid
            else "Fix errors and call validate_diagram_request again."
        ),
    }


def _skill_root() -> Path:
    return Path(__file__).resolve().parents[2]


def _shape_position(shape: Any) -> tuple[float, float]:
    try:
        return (
            float(shape.get_Cells("PinY").ResultIU),
            float(shape.get_Cells("PinX").ResultIU),
        )
    except Exception:
        return (0.0, 0.0)


def _shape_text_int(shape: Any) -> int:
    try:
        return int(str(shape.Text).strip())
    except Exception:
        return 0


def er_draw(
    project_root: str,
    user: str,
    entity: str,
    attributes: list[dict[str, Any]],
) -> dict[str, Any]:
    if not isinstance(project_root, str) or not project_root.strip():
        return {"success": False, "message": "project_root must be a non-empty string"}
    if not isinstance(user, str) or not user.strip():
        return {"success": False, "message": "user must be a non-empty string"}

    errors = validate_er_payload({"entity": entity, "attributes": attributes})
    if errors:
        return {"success": False, "message": "; ".join(errors)}

    root = Path(project_root).resolve()
    ent = entity.strip()
    draw_attrs = normalize_attrs(attributes)
    template_path = _skill_root() / "template" / "er" / choose_template_name(len(attributes))
    if not template_path.exists():
        return {"success": False, "message": f"template not found: {template_path}"}

    out_dir = root / "data" / user.strip() / "er"
    abs_output = out_dir / f"{ent}.vsdx"

    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        return {"success": False, "message": f"Visio COM dependencies unavailable: {exc}"}

    pythoncom.CoInitialize()
    visio = None
    doc = None
    try:
        out_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(template_path, abs_output)

        visio = win32com.client.Dispatch("Visio.Application")
        visio.Visible = False
        doc = visio.Documents.Open(str(abs_output))
        page = doc.Pages.ItemU(1)

        entities: list[Any] = []
        attrs: list[Any] = []
        for i in range(1, page.Shapes.Count + 1):
            shape = page.Shapes.ItemU(i)
            if "流程图过程" in shape.Name:
                entities.append(shape)
            elif "ER图-属性" in shape.Name:
                attrs.append(shape)

        entities.sort(key=_shape_position)
        attrs.sort(key=_shape_text_int)

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
        return {"success": True, "message": f"已保存: {abs_output}"}

    except Exception as exc:
        if doc is not None:
            try:
                doc.Close()
            except Exception:
                pass
        if visio is not None:
            try:
                visio.Quit()
            except Exception:
                pass
        if abs_output.exists():
            try:
                abs_output.unlink()
            except OSError:
                pass
        err_msg = str(exc)
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
        return {"success": False, "message": f"失败: {exc}"}
    finally:
        pythoncom.CoUninitialize()


def process_er_request(project_root: str, request: str | dict[str, Any]) -> dict[str, Any]:
    validation = validate_diagram_request(request)
    if not validation["valid"]:
        return {"success": False, "validation": validation, "message": "validation failed"}

    parsed = validation["parsed"]
    data = parsed["data"]
    draw = er_draw(project_root, parsed["user"], data["entity"], data["attributes"])
    return {"success": bool(draw.get("success")), "validation": validation, "draw": draw}


def _validate_data_json(json_text: str) -> dict[str, Any]:
    try:
        payload = json.loads(json_text.lstrip("\ufeff"))
    except json.JSONDecodeError as exc:
        return {"ok": False, "errors": [f"JSON parse failed: {exc}"]}
    if not isinstance(payload, dict):
        return {"ok": False, "errors": ["data JSON must be an object"]}
    errors = validate_er_payload(payload)
    if errors:
        return {"ok": False, "errors": errors}
    return {
        "ok": True,
        "template": choose_template_name(len(payload["attributes"])),
        "normalized_attributes": normalize_attrs(payload["attributes"]),
    }


def _read_json_arg(value: str | None, file_path: str | None) -> str:
    if value and file_path:
        raise ValueError("provide JSON text or JSON file, not both")
    if file_path:
        return Path(file_path).read_text(encoding="utf-8-sig")
    if value == "-":
        return sys.stdin.read().lstrip("\ufeff")
    if value:
        return value
    raise ValueError("missing JSON input")


def main() -> None:
    parser = argparse.ArgumentParser(description="Validate or draw a Visio ER request")
    parser.add_argument("--json", help="JSON text for ER data object only")
    parser.add_argument("--json-file", help="UTF-8 file containing ER data object JSON")
    parser.add_argument("--request-json", help="JSON text for top-level ER request")
    parser.add_argument("--request-json-file", help="UTF-8 file containing top-level ER request JSON")
    parser.add_argument("--project-root", help="Project root for drawing output")
    parser.add_argument("--draw", action="store_true", help="Draw after top-level validation")
    args = parser.parse_args()

    data_input = bool(args.json) or bool(args.json_file)
    request_input = bool(args.request_json) or bool(args.request_json_file)
    if data_input == request_input:
        parser.error("provide exactly one of --json/--json-file or --request-json/--request-json-file")
    if data_input and args.draw:
        parser.error("--draw requires --request-json")
    if args.draw and not args.project_root:
        parser.error("--project-root is required with --draw")

    try:
        if data_input:
            json_text = _read_json_arg(args.json, args.json_file)
            result = _validate_data_json(json_text)
        elif args.draw:
            json_text = _read_json_arg(args.request_json, args.request_json_file)
            result = process_er_request(args.project_root, json_text)
        else:
            json_text = _read_json_arg(args.request_json, args.request_json_file)
            result = validate_diagram_request(json_text)
            if result["valid"]:
                data = result["parsed"]["data"]
                result["template"] = choose_template_name(len(data["attributes"]))
                result["normalized_attributes"] = normalize_attrs(data["attributes"])
    except ValueError as exc:
        parser.error(str(exc))

    print(json.dumps(result, ensure_ascii=False))
    ok = result.get("ok", result.get("valid", result.get("success", False)))
    if not ok:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
