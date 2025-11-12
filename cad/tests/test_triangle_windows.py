#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""测试三角形墙体+窗"""
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_basic import close_all_cad_processes, start_applicationV9
from CAD_file_operations import open_file, draw_tarch_wall, insert_tarch_window, save_file
from CAD_coordination import ensure_single_process, wait_quiescent

print("="*60)
print("三角形墙体+窗测试")
print("="*60)

try:
    close_all_cad_processes()
    proc = start_applicationV9(PTH=r"C:\Tangent\TArchT20V9")
    ensure_single_process()
    wait_quiescent(min_quiet=2.0, timeout=30.0)

    open_file("D:/claude-tasks/tests/test_files/天正测试文件2.dwg")
    wait_quiescent(min_quiet=1.0, timeout=15.0)

    # 绘制三角形墙体（厚度200）
    print("\n绘制三角形墙体（厚度200）...")
    p1, p2, p3 = (100000, 75000, 0), (105000, 75000, 0), (102500, 79330, 0)

    draw_tarch_wall(p1, p2, thickness=200)
    draw_tarch_wall(p2, p3, thickness=200)
    draw_tarch_wall(p3, p1, thickness=200)

    # 插入窗
    print("\n插入窗...")
    windows = [
        ((102500, 75000, 0), 900, 3, "底边"),
        ((103750, 77165, 0), 1200, 7, "右边"),
        ((101250, 77165, 0), 2000, 9, "左边"),
    ]

    for i, (pos, width, wtype, name) in enumerate(windows, 1):
        print(f"\n  [{i}/3] {name}: 宽{width}, 类型{wtype}")
        result = insert_tarch_window(pos, width, wtype)
        print(f"    {'成功' if result['success'] else '失败'}")

    save_file()
    print("\n完成！")

except Exception as e:
    print(f"\n错误: {e}")
    import traceback
    traceback.print_exc()

finally:
    close_all_cad_processes()
