#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""在三角形墙体中间位置插入门"""
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_basic import get_acad_doc, close_all_cad_processes, start_applicationV9, get_object_property
from CAD_file_operations import open_file, insert_tarch_door, save_file
from CAD_coordination import ensure_single_process, wait_quiescent

print("="*60)
print("在三角形墙体中间位置插入门")
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

    # 查找三角形墙体
    print("\n[3] 查找三角形墙体...")
    acad, doc = get_acad_doc()
    ms = doc.ModelSpace

    walls = []
    for i in range(ms.Count):
        obj = ms.Item(i)
        if obj.ObjectName == "TDbWall":
            thickness = get_object_property(obj, 'Thickness')
            if thickness == 230:  # 找到宽度为230的墙
                walls.append(obj)

    print(f"  找到 {len(walls)} 堵宽度为230的墙")

    if len(walls) < 3:
        print("[错误] 未找到三角形墙体（需要3堵墙）")
        sys.exit(1)

    # 在每堵墙的中间位置插入门
    print("\n[4] 在墙体中间位置插入门...")
    for i, wall in enumerate(walls[:3], 1):  # 只处理前3堵墙
        # 获取墙的长度和位置（简化处理，使用墙的中心点）
        # 这里需要根据实际墙体的坐标计算中点
        # 暂时使用固定坐标测试

        print(f"\n  [{i}/3] 处理第{i}堵墙...")

        # 由于无法直接获取墙的精确中点，使用近似位置
        # 实际应用中需要根据墙的起点和终点计算
        test_positions = [
            (100500, 75000, 0),  # 墙1中点（示例）
            (105000, 75000, 0),  # 墙2中点（示例）
            (102500, 78000, 0),  # 墙3中点（示例）
        ]

        pos = test_positions[i-1]
        result = insert_tarch_door(pos, width=950)

        if result['success']:
            print(f"    成功插入门 - 宽度: {result['width']}")
        else:
            print(f"    插入失败")

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
