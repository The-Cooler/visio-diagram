from __future__ import annotations

import argparse
import json
import re
import shutil
from pathlib import Path

from analyze_visio_template import analyze_vsdx

VALID_ID = re.compile(r"^[a-z0-9_]+$")


def skill_root() -> Path:
    return Path(__file__).resolve().parents[1]


def default_template_id(source: Path) -> str:
    stem = re.sub(r"[^a-zA-Z0-9]+", "_", source.stem).strip("_").lower()
    return f"custom_{stem or 'template'}"


def load_registry(path: Path) -> dict:
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    return {"version": 1, "templates": []}


def save_registry(path: Path, registry: dict) -> None:
    path.write_text(json.dumps(registry, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def register_template(
    source: Path,
    diagram_type: str,
    template_id: str | None,
    replace: bool,
    confirmed: bool,
) -> dict:
    if not confirmed:
        raise PermissionError("registration requires explicit user consent and --confirmed")

    root = skill_root()
    source = source.expanduser().resolve()
    if not source.exists():
        raise FileNotFoundError(source)
    if source.suffix.lower() != ".vsdx":
        raise ValueError("source must be a .vsdx file")

    tid = template_id or default_template_id(source)
    if not VALID_ID.fullmatch(tid):
        raise ValueError("template-id must use lowercase letters, digits, and underscores")

    target_dir = (root / "template" / diagram_type).resolve()
    target_dir.mkdir(parents=True, exist_ok=True)
    dest = (target_dir / f"{tid}.vsdx").resolve()
    if target_dir not in dest.parents:
        raise ValueError("resolved destination escaped template directory")
    if dest.exists() and not replace:
        raise FileExistsError(f"template already exists: {dest}")

    analysis = analyze_vsdx(source, diagram_type)
    shutil.copy2(source, dest)

    metadata = {
        "id": tid,
        "diagram_type": diagram_type,
        "file": str(dest.relative_to(root)).replace("\\", "/"),
        "source_file": str(source),
        "analysis": analysis,
    }
    meta_path = target_dir / f"{tid}.json"
    meta_path.write_text(json.dumps(metadata, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    registry_path = root / "template" / "registry.json"
    registry = load_registry(registry_path)
    registry["templates"] = [
        item
        for item in registry.get("templates", [])
        if not (item.get("diagram_type") == diagram_type and item.get("id") == tid)
    ]
    registry["templates"].append(
        {
            "id": tid,
            "diagram_type": diagram_type,
            "file": metadata["file"],
            "source": "user",
            "metadata": str(meta_path.relative_to(root)).replace("\\", "/"),
        }
    )
    save_registry(registry_path, registry)
    return metadata


def main() -> None:
    parser = argparse.ArgumentParser(description="Register a confirmed Visio template")
    parser.add_argument("source", help="Path to a .vsdx template")
    parser.add_argument("--diagram-type", required=True, choices=("er", "flowchart", "sequence"))
    parser.add_argument("--template-id")
    parser.add_argument("--replace", action="store_true")
    parser.add_argument("--confirmed", action="store_true")
    args = parser.parse_args()

    result = register_template(
        Path(args.source),
        args.diagram_type,
        args.template_id,
        args.replace,
        args.confirmed,
    )
    print(json.dumps({"ok": True, "template": result}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
