from __future__ import annotations

import argparse
import json
import math
import re
import zipfile
from collections import defaultdict, deque
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

VISIO_NS = "{http://schemas.microsoft.com/office/visio/2012/main}"
TEXT_SPACE = re.compile(r"\s+")


def lname(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def clean_text(text: str) -> str:
    return TEXT_SPACE.sub("", text or "").strip()


def cell_value(shape: ET.Element, name: str) -> str | None:
    for cell in shape:
        if lname(cell.tag) == "Cell" and cell.attrib.get("N") == name:
            return cell.attrib.get("V")
    return None


def as_float(value: str | None) -> float | None:
    if value is None:
        return None
    try:
        return float(value)
    except ValueError:
        return None


def shape_text(shape: ET.Element) -> str:
    for child in shape:
        if lname(child.tag) == "Text":
            return clean_text("".join(child.itertext()))
    return ""


def is_connector(shape: dict[str, Any], connector_ids: set[str]) -> bool:
    name = (shape.get("name") or "").lower()
    return (
        shape["id"] in connector_ids
        or shape.get("oned") == "1"
        or "connector" in name
        or "dynamic" in name
    )


def read_page_names(zf: zipfile.ZipFile) -> list[str]:
    return sorted(
        name
        for name in zf.namelist()
        if name.startswith("visio/pages/page")
        and name.endswith(".xml")
        and "/_rels/" not in name
    )


def parse_page(xml_bytes: bytes, page_name: str) -> dict[str, Any]:
    root = ET.fromstring(xml_bytes)
    raw_shapes: dict[str, dict[str, Any]] = {}

    for shape in root.iter():
        if lname(shape.tag) != "Shape":
            continue
        sid = shape.attrib.get("ID")
        if not sid:
            continue
        raw_shapes[sid] = {
            "id": sid,
            "name": shape.attrib.get("NameU") or shape.attrib.get("Name") or "",
            "type": shape.attrib.get("Type") or "",
            "master": shape.attrib.get("Master") or "",
            "oned": shape.attrib.get("OneD") or "",
            "text": shape_text(shape),
            "pin_x": as_float(cell_value(shape, "PinX")),
            "pin_y": as_float(cell_value(shape, "PinY")),
            "width": as_float(cell_value(shape, "Width")),
            "height": as_float(cell_value(shape, "Height")),
        }

    connects = []
    connector_ids: set[str] = set()
    endpoints: dict[str, dict[str, str]] = defaultdict(dict)
    for conn in root.iter():
        if lname(conn.tag) != "Connect":
            continue
        item = dict(conn.attrib)
        connects.append(item)
        from_sheet = item.get("FromSheet")
        to_sheet = item.get("ToSheet")
        from_cell = item.get("FromCell", "")
        if not from_sheet or not to_sheet:
            continue
        connector_ids.add(from_sheet)
        if from_cell.startswith("Begin"):
            endpoints[from_sheet]["from"] = to_sheet
        elif from_cell.startswith("End"):
            endpoints[from_sheet]["to"] = to_sheet

    shapes = list(raw_shapes.values())
    connector_shapes = [s for s in shapes if is_connector(s, connector_ids)]
    node_shapes = [s for s in shapes if not is_connector(s, connector_ids)]

    edges = []
    for connector_id, pair in endpoints.items():
        src = pair.get("from")
        dst = pair.get("to")
        if not src or not dst:
            continue
        src_shape = raw_shapes.get(src, {})
        dst_shape = raw_shapes.get(dst, {})
        conn_shape = raw_shapes.get(connector_id, {})
        edges.append(
            {
                "connector_id": connector_id,
                "from": src,
                "to": dst,
                "from_text": src_shape.get("text", ""),
                "to_text": dst_shape.get("text", ""),
                "text": conn_shape.get("text", ""),
            }
        )

    role_sequence = infer_role_sequence(node_shapes, edges)
    return {
        "page": page_name,
        "shape_count": len(shapes),
        "node_count": len(node_shapes),
        "connector_count": len(connector_shapes),
        "connect_count": len(connects),
        "text_slots": [s for s in node_shapes if s.get("text")],
        "edges": edges,
        "role_sequence": role_sequence,
    }


def role_for_shape(shape: dict[str, Any]) -> str:
    text = shape.get("text") or ""
    name = (shape.get("name") or "").lower()
    if text in {"开始", "start", "Start"}:
        return "start"
    if text in {"结束", "end", "End"}:
        return "end"
    if text.startswith("是否") or "decision" in name or "diamond" in name:
        return "decision"
    if any(word in text for word in ("输入", "录入", "填写", "编辑", "写入", "添加", "新建", "创建")):
        return "input"
    return "process"


def position_key(shape: dict[str, Any]) -> tuple[float, float, str]:
    y = shape.get("pin_y")
    x = shape.get("pin_x")
    return (-(y if y is not None else -math.inf), x if x is not None else math.inf, shape["id"])


def infer_role_sequence(nodes: list[dict[str, Any]], edges: list[dict[str, Any]]) -> list[dict[str, str]]:
    by_id = {s["id"]: s for s in nodes}
    if not nodes:
        return []

    outgoing: dict[str, list[str]] = defaultdict(list)
    incoming: dict[str, int] = defaultdict(int)
    for edge in edges:
        src = edge.get("from")
        dst = edge.get("to")
        if src in by_id and dst in by_id:
            outgoing[src].append(dst)
            incoming[dst] += 1

    starts = [sid for sid in by_id if incoming[sid] == 0]
    starts.sort(key=lambda sid: position_key(by_id[sid]))
    ordered: list[str] = []
    seen: set[str] = set()
    queue: deque[str] = deque(starts)
    while queue:
        sid = queue.popleft()
        if sid in seen:
            continue
        seen.add(sid)
        ordered.append(sid)
        for dst in sorted(outgoing.get(sid, []), key=lambda item: position_key(by_id[item])):
            if dst not in seen:
                queue.append(dst)

    for shape in sorted(nodes, key=position_key):
        if shape["id"] not in seen:
            ordered.append(shape["id"])

    return [
        {"id": sid, "role": role_for_shape(by_id[sid]), "text": by_id[sid].get("text", "")}
        for sid in ordered
    ]


def infer_type(pages: list[dict[str, Any]], forced: str | None) -> dict[str, Any]:
    if forced:
        return {"type": forced, "confidence": 1.0, "reason": "provided by user"}

    all_text = "".join(
        slot.get("text", "")
        for page in pages
        for slot in page.get("text_slots", [])
    )
    node_count = sum(page["node_count"] for page in pages)
    edge_count = sum(len(page["edges"]) for page in pages)
    roles = [item["role"] for page in pages for item in page.get("role_sequence", [])]

    lane_hits = sum(1 for word in ("用户", "前端", "后端", "数据库") if word in all_text)
    if lane_hits >= 3 and edge_count >= 4:
        return {"type": "sequence", "confidence": 0.82, "reason": "common sequence lanes and directed messages detected"}
    if "decision" in roles or edge_count >= 2:
        return {"type": "flowchart", "confidence": 0.72, "reason": "connectors or decision slots detected"}
    if node_count >= 2 and ("id" in all_text.lower() or edge_count == 0):
        return {"type": "er", "confidence": 0.58, "reason": "attribute-like text slots with low connector use"}
    return {"type": "unknown", "confidence": 0.0, "reason": "insufficient structural signal"}


def complexity(pages: list[dict[str, Any]]) -> dict[str, Any]:
    nodes = sum(page["node_count"] for page in pages)
    edges = sum(len(page["edges"]) for page in pages)
    score = nodes + edges * 1.5 + len(pages) * 2
    if score <= 12:
        level = "simple"
    elif score <= 28:
        level = "medium"
    else:
        level = "complex"
    return {"level": level, "score": round(score, 2), "nodes": nodes, "edges": edges}


def analyze_vsdx(path: Path, diagram_type: str | None = None) -> dict[str, Any]:
    path = path.expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(path)
    if path.suffix.lower() != ".vsdx":
        raise ValueError("source must be a .vsdx file")

    with zipfile.ZipFile(path) as zf:
        page_names = read_page_names(zf)
        master_names = sorted(name for name in zf.namelist() if name.startswith("visio/masters/master") and name.endswith(".xml"))
        pages = [parse_page(zf.read(name), name) for name in page_names]

    inferred = infer_type(pages, diagram_type)
    return {
        "ok": True,
        "file": str(path),
        "xml": {"pages": page_names, "masters": master_names},
        "summary": {
            "page_count": len(page_names),
            "master_count": len(master_names),
            "shape_count": sum(page["shape_count"] for page in pages),
            "node_count": sum(page["node_count"] for page in pages),
            "connector_count": sum(page["connector_count"] for page in pages),
            "connect_count": sum(page["connect_count"] for page in pages),
        },
        "inferred_diagram": inferred,
        "complexity": complexity(pages),
        "pages": pages,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Analyze Visio .vsdx XML structure")
    parser.add_argument("source", help="Path to a .vsdx template")
    parser.add_argument("--diagram-type", choices=("er", "flowchart", "sequence"))
    args = parser.parse_args()

    result = analyze_vsdx(Path(args.source), args.diagram_type)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
