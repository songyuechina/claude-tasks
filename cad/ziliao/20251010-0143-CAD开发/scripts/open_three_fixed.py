# -*- coding: utf-8 -*-
# @script id: open_three_fixed
# @script name: Open Three DWGs Sequentially (TArch)
# @script description: 归为单文件不确定状态后，按顺序仅打开一次指定的三个 DWG（通过 CAD 基本操作.py 的 start_applicationV9 启动），并打印当前打开列表。
# @script usage: python scripts/open_three_fixed.py

from __future__ import annotations

from pathlib import Path
import sys as _sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in _sys.path:
    _sys.path.insert(0, str(ROOT))

from cad_automation import CadController


def main() -> int:
    paths = [
        r"D:/codex-tasks/20251010-0143-CAD开发/artifacts/dwgs/标准图签模版-SS.dwg",
        r"D:/codex-tasks/20251010-0143-CAD开发/artifacts/dwgs/4#栋【电气】_t3.dwg",
        r"D:/codex-tasks/20251010-0143-CAD开发/artifacts/dwgs/3#栋电气]_t3.dwg",
    ]
    print("TARGETS:")
    for p in paths:
        print(" -", p)

    ctl = CadController()
    # Converge to single-unsaved (will close processes -> start_applicationV9 -> settle)
    name = ctl.single_unsaved_state()
    print("STATE:", name or "<none>")

    ok_count = ctl.open_dwgs(paths)
    print("OPENED COUNT:", ok_count)
    names = ctl.list_open_documents()
    print("OPEN LIST:", ", ".join(names) if names else "<none>")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
