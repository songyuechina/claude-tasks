# -*- coding: utf-8 -*-
# @script id: insert_block_once
# @script name: Insert AB into 123 (block, no explode)
# @script description: 以 123.dwg 为激活文件，将 ab.dwg 作为块插入到原点（不炸开）。
# @script usage: python scripts/insert_block_once.py

from __future__ import annotations

from pathlib import Path
import sys as _sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in _sys.path:
    _sys.path.insert(0, str(ROOT))

from cad_automation import CadController


def main() -> int:
    root = ROOT
    a = str((root / 'artifacts' / 'dwgs' / '123.dwg').resolve())
    b = str((root / 'artifacts' / 'dwgs' / 'ab.dwg').resolve())

    ctl = CadController()

    ctl.open_dwg(a)
    ctl.wait_quiescent(30)
    ops = ctl.ops
    pos = (0.0, 0.0, 0.0)

    # 使用 -INSERT 命令以避免外部脚本在控制台打印 Unicode 字符导致的编码异常
    print('[info] using -INSERT command (no explode)')
    ops.get_acad_doc()
    cmd = f"-INSERT\n\"{b}\"\n{pos[0]},{pos[1]},{pos[2]}\n1\n1\n0\n"
    ops.doc.SendCommand(cmd)

    ctl.wait_quiescent(60)
    ctl.save()
    print('DONE insert block into 123.dwg')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
