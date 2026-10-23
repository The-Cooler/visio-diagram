"""ER 图命令行入口：与流程图/时序图的使用方式一致。

用法（在 skill 目录执行）：
  python -m scripts.er <项目根>/json/er/<实体>.json [--project-root <项目根>]

JSON 放在 {project_root}/json/er/，输出到 {project_root}/draw/er/。
省略 --project-root 时使用当前工作目录。
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from .er_contract import validate_er_payload
from .tools import er_draw


def main() -> None:
    parser = argparse.ArgumentParser(description="Validate and draw an ER entity diagram")
    parser.add_argument("json_path", help="JSON path, usually <project_root>/json/er/<entity>.json")
    parser.add_argument("--project-root", help="项目根；省略时使用当前工作目录")
    parser.add_argument("--user", help="兼容旧调用，不参与路径")
    args = parser.parse_args()

    path = Path(args.json_path)
    if not path.exists():
        print(json.dumps({"success": False, "message": f"JSON 文件不存在: {path}"}, ensure_ascii=False))
        raise SystemExit(1)
    try:
        payload = json.loads(path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as e:
        print(json.dumps({"success": False, "message": f"JSON 解析失败: {e}"}, ensure_ascii=False))
        raise SystemExit(1)

    errors = validate_er_payload(payload)
    if errors:
        print(json.dumps({"ok": False, "errors": errors}, ensure_ascii=False))
        raise SystemExit(1)

    result = er_draw(args.project_root, str(path), user=args.user)
    print(json.dumps(result, ensure_ascii=False))
    if not result.get("success"):
        raise SystemExit(1)


if __name__ == "__main__":
    main()
