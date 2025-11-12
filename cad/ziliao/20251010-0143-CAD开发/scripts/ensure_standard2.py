# -*- coding: utf-8 -*-
# @script id: ensure_standard2
# @script name: Ensure Two-Doc Standard State
# @script description: 检测当前 CAD 进程与 DWG 打开数，必要时收敛到 00.dwg + 01.dwg 双文件确定状态
# @script usage: python scripts/ensure_standard2.py [a|b]

from __future__ import annotations

import sys
from pathlib import Path

import sys as _sys
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in _sys.path:
    _sys.path.insert(0, str(ROOT))
from cad_automation import CadController


def main() -> int:
    active = None
    if len(sys.argv) >= 2 and sys.argv[1] in ("a", "b"):
        active = sys.argv[1]

    ctl = CadController()
    # 检测当前打开 DWG 数
    names_before = ctl.list_open_documents()
    print("OPEN BEFORE:", ", ".join(names_before) if names_before else "<none>")

    # 若不是两个文件，则收敛到固定 00/01
    if len(names_before) != 2:
        print("[info] Not 2 docs; converging to 00/01")
        names = ctl.standardize_two_default(active=active)
        print("RESULT:", names)
        return 0

    # 如果已有两个文件，为了“确定状态”同样收敛为固定 00/01
    names = ctl.standardize_two_default(active=active)
    print("RESULT:", names)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
