#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""最终窗测试 - 简化版"""
import sys
import subprocess
import time
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_basic import close_all_cad_processes, start_applicationV9
from CAD_file_operations import new_file, draw_tarch_wall, insert_tarch_window, save_file
from CAD_coordination import ensure_single_process, wait_quiescent

print("="*60)
print("最终窗测试")
print("="*60)

# 启动dialog killer
dialog_killer = subprocess.Popen([sys.executable, "D:/claude-tasks/cad/scripts/cad_dialog_killer.py"])
time.sleep(2)

try:
    close_all_cad_processes()
    proc = start_applicationV9(PTH=r"C:\Tangent\TArchT20V9")
    ensure_single_process()
    wait_quiescent(min_quiet=2.0, timeout=30.0)

    new_file("D:/claude-tasks/tests/test_files/天正测试文件2.dwg")
    time.sleep(3)

    # 绘制墙体
    walls = [
        ((100000, 75000, 0), (105000, 75000, 0)),
        ((105000, 75000, 0), (102500, 79330, 0)),
        ((102500, 79330, 0), (100000, 75000, 0)),
    ]

    for p1, p2 in walls:
        draw_tarch_wall(p1, p2, thickness=200)
        time.sleep(2)

    # 插入窗
    windows = [
        ((102500, 75000, 0), 900, 3),
        ((103750, 77165, 0), 1200, 7),
        ((101250, 77165, 0), 2000, 9),
    ]

    for p, width, wtype in windows:
        print(f"\n插入窗 (宽度{width}, 类型{wtype})...")
        result = insert_tarch_window(p, width, wtype)
        print(f"结果: {result['success']}")
        time.sleep(3)

    save_file()
    print("\n测试完成！")

except Exception as e:
    print(f"\n错误: {e}")
    import traceback
    traceback.print_exc()

finally:
    close_all_cad_processes()
    dialog_killer.terminate()
