# -*- coding: utf-8 -*-
# @script id: test_flow
# @script name: CAD Smoke Test
# @script description: 以标准化状态启动CAD，打开/保存/画线并列出DWG
# @script usage: python scripts/test_flow.py
"""
本地验证流程：
- 确保仅一个 CAD 进程
- 创建/另存为 artifacts/dwgs/A.dwg
- 打开 A.dwg，再画一条线并保存
- 列出当前打开的 DWG 文件名
"""
from __future__ import annotations

import os
import time
from pathlib import Path

from cad_automation import CadController


def main() -> int:
    out_dir = Path("artifacts/dwgs")
    out_dir.mkdir(parents=True, exist_ok=True)
    a_path = str((out_dir / "A.dwg").resolve())

    ctl = CadController()
    assert ctl.start_tarch_v9(), "启动/复用 CAD 失败"

    # 回退到标准CAD状态（单进程单文档）。这里不指定文件，默认保留当前激活文档
    assert ctl.standardize_state(), "标准状态设置失败"

    # 收敛至单进程
    ctl.ensure_single_process()

    # 若不存在，使用 SaveAs 创建一个空白 DWG
    if not os.path.exists(a_path):
        ctl.save_as(a_path)
        time.sleep(1)

    # 打开 A.dwg
    assert ctl.open_dwg(a_path), "打开 A.dwg 失败"

    # 在当前图中画线并保存
    ctl.draw_line((0, 0, 0), (1000, 0, 0))
    ctl.save()

    # 列出当前打开 DWG
    names = ctl.list_open_documents()
    print("OPEN:", ", ".join(names))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
