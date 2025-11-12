#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""绘制三角形墙体并在中间插入门"""
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_basic import close_all_cad_processes, start_applicationV9
from CAD_file_operations import new_file, draw_tarch_wall, insert_tarch_door, save_file
from CAD_coordination import ensure_single_process, wait_quiescent

print("="*60)
print("绘制三角形墙体并插入门")
print("="*60)

try:
    # 准备环境
    print("\n[1] 准备环境...")
    close_all_cad_processes()
    proc = start_applicationV9(PTH=r"C:\Tangent\TArchT20V9")
    ensure_single_process()
    wait_quiescent(min_quiet=2.0, timeout=30.0)

    # 新建文件
    print("\n[2] 新建文件...")
    new_file("D:/claude-tasks/tests/test_files/三角形墙体测试.dwg")
    wait_quiescent(min_quiet=1.0, timeout=15.0)

    # 绘制三角形墙体（厚度230）
    print("\n[3] 绘制三角形墙体（厚度230）...")

    p1 = (100000, 75000, 0)
    p2 = (105000, 75000, 0)
    p3 = (102500, 79330, 0)

    print(f"  墙1: {p1[:2]} -> {p2[:2]}")
    draw_tarch_wall(p1, p2, thickness=230)

    print(f"  墙2: {p2[:2]} -> {p3[:2]}")
    draw_tarch_wall(p2, p3, thickness=230)

    print(f"  墙3: {p3[:2]} -> {p1[:2]}")
    draw_tarch_wall(p3, p1, thickness=230)

    print("  三角形墙体绘制完成")

    # 在每堵墙中间插入门（宽度950）
    print("\n[4] 在墙体中间插入门（宽度950）...")

    door_positions = [
        (102500, 75000, 0),  # 底边中点
        (103750, 77165, 0),  # 右边中点
        (101250, 77165, 0),  # 左边中点
    ]

    for i, pos in enumerate(door_positions, 1):
        print(f"\n  [{i}/3] 在位置 {pos[:2]} 插入门...")
        result = insert_tarch_door(pos, width=950)

        if result['success']:
            print(f"    成功 - 门宽度: {result['width']}")
        else:
            print(f"    失败")

    # 保存文件
    print("\n[5] 保存文件...")
    save_file()

    print("\n" + "="*60)
    print("测试完成！")
    print("文件保存在: D:/claude-tasks/tests/test_files/三角形墙体测试.dwg")
    print("="*60)

except Exception as e:
    print(f"\n[错误] {e}")
    import traceback
    traceback.print_exc()

finally:
    print("\n[清理] 关闭CAD...")
    close_all_cad_processes()
