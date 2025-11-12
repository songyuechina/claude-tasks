# -*- coding: utf-8 -*-
# @script id: do_single_unsaved
# @script name: Force Single Unsaved State
# @script description: 关闭所有 CAD 进程并启动到“单文件未保存”状态，打印结果文件名与当前列表。
# @script usage: python scripts/do_single_unsaved.py

from __future__ import annotations

from pathlib import Path
import sys as _sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in _sys.path:
    _sys.path.insert(0, str(ROOT))

from cad_automation import CadController


def main() -> int:
    ctl = CadController()
    # 强制到“单文件未保存”状态
    name = ctl.single_unsaved_state()
    names = ctl.list_open_documents()
    print('ACTIVE:', name or '<none>')
    print('OPEN  :', ', '.join(names) if names else '<none>')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())

