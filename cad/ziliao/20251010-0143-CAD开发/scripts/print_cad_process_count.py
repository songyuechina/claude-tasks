# -*- coding: utf-8 -*-
# @script id: print_cad_process_count
# @script name: Print CAD Process Count
# @script description: 打印并返回本机当前 CAD 进程数量，便于快速检测。
# @script usage: python scripts/print_cad_process_count.py

from __future__ import annotations

from pathlib import Path
import sys as _sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in _sys.path:
    _sys.path.insert(0, str(ROOT))

from cad_automation import CadController


def main() -> int:
    ctl = CadController()
    cnt = ctl.cad_process_count(print_it=True)
    # Also print a plain line for convenience
    print(f"CAD PROCESS COUNT: {cnt}")
    return 0


if __name__ == '__main__':
    raise SystemExit(main())

