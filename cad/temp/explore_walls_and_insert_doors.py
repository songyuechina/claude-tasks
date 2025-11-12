#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""探索墙体并插入门"""
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_basic import get_acad_doc, close_all_cad_processes, start_applicationV9, get_object_property
from CAD_file_operations import open_file, insert_tarch_door, draw_tarch_wall, save_file
from CAD_coordination import ensure_single_process, wait_quiescent

print("="*60)
print("探索墙体并插入门")
print("="*60)

try:
    # 准备环境
    print("\n[1] 准备环境...")
    close_all_cad_processes()
    proc = start_applicationV9(PTH=r"C:\Tangent\TArchT20V9")
    ensure_single_process()
    wait_quiescent(min_quiet=2.0, timeout=30.0)

    # 打开测试文件
    print("\n[2] 打开测试文件...")
    open_file("D:/claude-tasks/tests/test_files/天正测试文件1.dwg")
    wait_quiescent(min_quiet=1.0, timeout=15.0)

    # 查找墙体
    print("\n[3] 探索文件中的墙体...")
    acad, doc = get_acad_doc()
    ms = doc.ModelSpace

    walls = []
    for i in range(ms.Count):
        obj = ms.Item(i)
        if obj.ObjectName == "TDbWall":
            thickness = get_object_property(obj, 'Thickness')
            walls.append((obj, thickness))

    print(f"  找到 {len(walls)} 堵墙")

    if len(walls) == 0:
        print("\n  文件中没有墙体，创建三角形墙体...")

        # 创建三角形墙体（宽度230）
        p1 = (100000, 75000, 0)
        p2 = (105000, 75000, 0)
        p3 = (102500, 79330, 0)  # 等边三角形高度

        print(f"  绘制墙1: {p1[:2]} -> {p2[:2]}")
        draw_tarch_wall(p1, p2, thickness=230)

        print(f"  绘制墙2: {p2[:2]} -> {p3[:2]}")
        draw_tarch_wall(p2, p3, thickness=230)

        print(f"  绘制墙3: {p3[:2]} -> {p1[:2]}")
        draw_tarch_wall(p3, p1, thickness=230)

        print("  三角形墙体创建完成")

        # 重新获取墙体
        wait_quiescent(min_quiet=1.0, timeout=10.0)
        walls = []
        ms = doc.ModelSpace
        for i in range(ms.Count):
            obj = ms.Item(i)
            if obj.ObjectName == "TDbWall":
                thickness = get_object_property(obj, 'Thickness')
                walls.append((obj, thickness))

        print(f"  现在有 {len(walls)} 堵墙")

    else:
        print(f"  墙体厚度分布:")
        thickness_count = {}
        for wall, thickness in walls:
            thickness_count[thickness] = thickness_count.get(thickness, 0) + 1
        for t, count in sorted(thickness_count.items()):
            print(f"    厚度 {t}: {count} 堵")

    # 在墙体中间插入门
    print("\n[4] 在墙体中间位置插入门...")

    # 取前3堵墙（如果有的话）
    walls_to_use = walls[:3] if len(walls) >= 3 else walls

    # 计算门的插入位置（在墙的中间）
    # 由于无法直接获取墙的精确坐标，使用三角形的边中点
    door_positions = [
        (102500, 75000, 0),  # 底边中点
        (103750, 77165, 0),  # 右边中点
        (101250, 77165, 0),  # 左边中点
    ]

    for i, (wall, thickness) in enumerate(walls_to_use, 1):
        print(f"\n  [{i}/{len(walls_to_use)}] 在第{i}堵墙插入门...")
        print(f"    墙厚度: {thickness}")

        pos = door_positions[i-1] if i <= len(door_positions) else door_positions[0]
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
    print("="*60)

except Exception as e:
    print(f"\n[错误] {e}")
    import traceback
    traceback.print_exc()

finally:
    print("\n[清理] 关闭CAD...")
    close_all_cad_processes()
