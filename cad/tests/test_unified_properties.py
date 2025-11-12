#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""测试统一的对象属性访问模块"""
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_basic import get_acad_doc, close_all_cad_processes, start_applicationV9
from CAD_file_operations import open_file
from CAD_coordination import ensure_single_process, wait_quiescent, send_cmd_with_sync
from CAD_object_properties import cast_object, get_object_property, set_object_property

print("="*60)
print("测试统一对象属性访问模块")
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

    acad, doc = get_acad_doc()
    ms = doc.ModelSpace

    # 绘制一条测试直线
    print("\n[3] 绘制测试直线...")
    send_cmd_with_sync("_LINE\n0,0\n1000,1000\n\n", wait_after=1.0)

    # 遍历对象找到不同类型
    print("\n[4] 测试不同类型对象...")

    line_obj = None
    door_obj = None
    wall_obj = None

    for i in range(ms.Count):
        obj = ms.Item(i)
        obj_name = obj.ObjectName

        if obj_name == "AcDbLine" and line_obj is None:
            line_obj = obj
        elif obj_name == "TDbOpening" and door_obj is None:
            door_obj = obj
        elif obj_name == "TDbWall" and wall_obj is None:
            wall_obj = obj

        if line_obj and door_obj and wall_obj:
            break

    # 测试CAD直线对象
    if line_obj:
        print(f"\n{'='*60}")
        print("测试CAD直线 (AcDbLine)")
        print(f"{'='*60}")

        # 方法1: 使用cast_object转换
        line = cast_object(line_obj)
        start = line.StartPoint
        end = line.EndPoint

        print(f"使用cast_object:")
        print(f"  起点: ({start[0]:.2f}, {start[1]:.2f}, {start[2]:.2f})")
        print(f"  终点: ({end[0]:.2f}, {end[1]:.2f}, {end[2]:.2f})")

        # 方法2: 使用get_object_property
        start2 = get_object_property(line_obj, 'StartPoint')
        print(f"\n使用get_object_property:")
        print(f"  起点: ({start2[0]:.2f}, {start2[1]:.2f}, {start2[2]:.2f})")

    # 测试天正门对象
    if door_obj:
        print(f"\n{'='*60}")
        print("测试天正门 (TDbOpening)")
        print(f"{'='*60}")

        width = get_object_property(door_obj, 'Width')
        height = get_object_property(door_obj, 'Height')
        door_type = get_object_property(door_obj, 'Type')

        print(f"使用get_object_property:")
        print(f"  宽度: {width}")
        print(f"  高度: {height}")
        print(f"  类型: {door_type}")

    # 测试天正墙对象
    if wall_obj:
        print(f"\n{'='*60}")
        print("测试天正墙 (TDbWall)")
        print(f"{'='*60}")

        thickness = get_object_property(wall_obj, 'Thickness')
        wall_type = get_object_property(wall_obj, 'WallType')

        print(f"使用get_object_property:")
        print(f"  厚度: {thickness}")
        print(f"  墙类型: {wall_type}")

    print(f"\n{'='*60}")
    print("测试完成！统一模块同时支持CAD对象和天正对象")
    print(f"{'='*60}")

except Exception as e:
    print(f"\n错误: {e}")
    import traceback
    traceback.print_exc()

finally:
    print("\n[清理] 关闭CAD...")
    close_all_cad_processes()
