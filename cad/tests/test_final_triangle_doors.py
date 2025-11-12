#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""最终测试：三角形墙体+门"""
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_basic import close_all_cad_processes, start_applicationV9
from CAD_file_operations import open_file, draw_tarch_wall, insert_tarch_door, save_file
from CAD_coordination import ensure_single_process, wait_quiescent

print("="*60)
print("三角形墙体+门测试")
print("="*60)

try:
    close_all_cad_processes()
    proc = start_applicationV9(PTH=r"C:\Tangent\TArchT20V9")
    ensure_single_process()
    wait_quiescent(min_quiet=2.0, timeout=30.0)

    open_file("D:/claude-tasks/tests/test_files/天正测试文件1.dwg")
    wait_quiescent(min_quiet=1.0, timeout=15.0)

    # 绘制三角形墙体（厚度230）
    print("\n绘制三角形墙体...")
    p1, p2, p3 = (100000, 75000, 0), (105000, 75000, 0), (102500, 79330, 0)

    draw_tarch_wall(p1, p2, thickness=230)
    draw_tarch_wall(p2, p3, thickness=230)
    draw_tarch_wall(p3, p1, thickness=230)

    # 插入门（宽度950）
    print("\n插入门...")
    for i, pos in enumerate([(102500, 75000, 0), (103750, 77165, 0), (101250, 77165, 0)], 1):
        result = insert_tarch_door(pos, width=950)
        print(f"  门{i}: {'成功' if result['success'] else '失败'}")

    save_file()
    print("\n完成！")

except Exception as e:
    print(f"\n错误: {e}")

finally:
    close_all_cad_processes()
