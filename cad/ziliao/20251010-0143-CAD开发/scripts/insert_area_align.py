# -*- coding: utf-8 -*-
# @script id: insert_area_align
# @script name: Insert AB with Area Alignment (no explode)
# @script description: 以矩形区域左下角对齐到目标点（不做裁剪，仅对齐插入）。
# @script usage: python scripts/insert_area_align.py

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

    # 矩形区域 (x1,y1)-(x2,y2)，这里只用于对齐计算，不做裁剪
    x1, y1, x2, y2 = 0.0, 0.0, 500.0, 500.0
    # 目标位置（原点）
    tx, ty, tz = 0.0, 0.0, 0.0
    blx, bly = (min(x1, x2), min(y1, y2))
    ins = (tx - blx, ty - bly, tz)

    ctl = CadController()
    ctl.open_dwg(a)
    ctl.wait_quiescent(30)
    ops = ctl.ops
    ops.get_acad_doc()
    cmd = f"-INSERT\n\"{b}\"\n{ins[0]},{ins[1]},{ins[2]}\n1\n1\n0\n"
    ops.doc.SendCommand(cmd)
    ctl.wait_quiescent(60)
    ctl.save()
    print('DONE insert area-aligned block into 123.dwg (no explode)')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())

