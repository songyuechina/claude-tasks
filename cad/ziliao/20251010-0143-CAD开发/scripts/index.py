# -*- coding: utf-8 -*-
"""
扫描当前任务文件夹，收集以 "# @script" 标注的脚本元数据，生成：
- artifacts/scripts_manifest.json
- docs/SCRIPTS.md（若 docs/ 不存在会自动创建）

标注示例（置于脚本文件顶部）：
    # @script id: cad_cli
    # @script name: CAD CLI Entrypoint
    # @script description: 启动/关闭CAD、DWG管理、基本操作
    # @script commands: start, stop, open, ...
    # @script usage: python cad_cli.py <command>
"""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Dict, List


ROOT = Path(__file__).resolve().parents[1]


def parse_script_header(path: Path) -> dict:
    meta: Dict[str, str] = {
        "path": str(path.relative_to(ROOT)).replace("\\", "/"),
        "id": "",
        "name": "",
        "description": "",
        "commands": "",
        "usage": "",
    }
    try:
        with path.open("r", encoding="utf-8", errors="ignore") as f:
            for i, line in enumerate(f):
                if i > 50:  # 只扫描前 50 行
                    break
                if line.startswith("# @script"):
                    # 形如: # @script key: value
                    parts = line.strip().split(" ", 2)
                    if len(parts) >= 3:
                        rest = parts[2]
                        if ":" in rest:
                            k, v = rest.split(":", 1)
                            meta[k.strip()] = v.strip()
    except Exception:
        pass
    # 仅保留标注过 id 或 name 的脚本
    if meta.get("id") or meta.get("name"):
        return meta
    return {}


def scan_scripts() -> List[dict]:
    metas: List[dict] = []
    ignore_dirs = {"artifacts", "__pycache__", ".git", ".venv", "venv"}
    for p in ROOT.rglob("*.py"):
        # 跳过已知忽略目录
        if any(part in ignore_dirs for part in p.parts):
            continue
        m = parse_script_header(p)
        if m:
            metas.append(m)
    return metas


def write_manifest(metas: List[dict]) -> None:
    out_dir = ROOT / "artifacts"
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "scripts_manifest.json").write_text(
        json.dumps({"scripts": metas}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def write_markdown(metas: List[dict]) -> None:
    docs_dir = ROOT / "docs"
    docs_dir.mkdir(parents=True, exist_ok=True)
    lines = [
        "# 已登记脚本一览",
        "",
        "- 本表由 `scripts/index.py` 扫描生成；建议每个脚本在顶部使用 # @script 元数据标注。",
        "",
    ]
    for m in metas:
        lines.append(f"- `{m.get('path')}` — {m.get('name') or m.get('id')}")
        desc = m.get("description")
        if desc:
            lines.append(f"  - 描述: {desc}")
        cmds = m.get("commands")
        if cmds:
            lines.append(f"  - 命令: {cmds}")
        usage = m.get("usage")
        if usage:
            lines.append(f"  - 用法: {usage}")
    (docs_dir / "SCRIPTS.md").write_text("\n".join(lines), encoding="utf-8")


def main() -> int:
    metas = scan_scripts()
    write_manifest(metas)
    write_markdown(metas)
    print(f"Indexed {len(metas)} scripts.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

