#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""测试CAD_tarch_properties模块"""
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_basic import get_acad_doc, close_all_cad_processes, start_applicationV9
from CAD_file_operations import open_file
from CAD_coordination import ensure_single_process, wait_quiescent
from CAD_tarch_properties import (
    get_tarch_property,
    set_tarch_property,
    get_all_tarch_properties,
    TARCH_PROPERTY_MAP
)

print("="*60)
print("测试CAD_tarch_properties模块")
print("="*60)

try:
    # 关闭并启动CAD
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

    # 查找天正门和天正墙
    print("\n[3] 查找天正对象...")
    door = None
    wall = None

    for i in range(ms.Count):
        obj = ms.Item(i)
        if obj.ObjectName == "TDbOpening" and door is None:
            door = obj
        elif obj.ObjectName == "TDbWall" and wall is None:
            wall = obj

        if door and wall:
            break

    # 测试天正门
    if door:
        print(f"\n{'='*60}")
        print("测试天正门 (TDbOpening)")
        print(f"{'='*60}")

        # 获取属性
        width = get_tarch_property(door, 'Width')
        height = get_tarch_property(door, 'Height')
        door_type = get_tarch_property(door, 'Type')

        print(f"原始属性:")
        print(f"  宽度 (Width): {width}")
        print(f"  高度 (Height): {height}")
        print(f"  类型 (Type): {door_type}")

        # 修改属性
        print(f"\n修改宽度为1500...")
        set_tarch_property(door, 'Width', 1500)
        new_width = get_tarch_property(door, 'Width')
        print(f"  新宽度: {new_width}")

        # 恢复
        set_tarch_property(door, 'Width', width)
        print(f"  已恢复为: {width}")

    # 测试天正墙
    if wall:
        print(f"\n{'='*60}")
        print("测试天正墙 (TDbWall)")
        print(f"{'='*60}")

        # 获取属性
        thickness = get_tarch_property(wall, 'Thickness')
        wall_type = get_tarch_property(wall, 'WallType')
        material = get_tarch_property(wall, 'Material')

        print(f"属性:")
        print(f"  厚度 (Thickness): {thickness}")
        print(f"  墙类型 (WallType): {wall_type}")
        print(f"  材质 (Material): {material}")

    # 显示支持的属性映射
    print(f"\n{'='*60}")
    print("支持的属性映射")
    print(f"{'='*60}")
    for obj_type, props in TARCH_PROPERTY_MAP.items():
        print(f"\n{obj_type}:")
        for prop_name, dispid in sorted(props.items(), key=lambda x: x[1]):
            print(f"  {prop_name:15s} -> DISPID {dispid}")

    print(f"\n{'='*60}")
    print("测试完成！")
    print(f"{'='*60}")

except Exception as e:
    print(f"\n错误: {e}")
    import traceback
    traceback.print_exc()

finally:
    print("\n[清理] 关闭CAD...")
    close_all_cad_processes()
