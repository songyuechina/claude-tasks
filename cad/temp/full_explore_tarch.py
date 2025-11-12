#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""完整探索天正对象属性"""
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_basic import get_acad_doc, close_all_cad_processes, start_applicationV9
from CAD_file_operations import open_file
from CAD_coordination import ensure_single_process, wait_quiescent
import pythoncom

def explore_tarch_object(obj, obj_name):
    """探索天正对象的所有DISPID属性"""
    print(f"\n{'='*60}")
    print(f"探索 {obj_name}")
    print(f"{'='*60}")
    print(f"ObjectName: {obj.ObjectName}")

    oleobj = obj._oleobj_
    properties = {}

    # 尝试DISPID 1-100
    for dispid in range(1, 101):
        try:
            result = oleobj.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYGET, True)
            if result is not None:
                properties[dispid] = result
        except:
            pass

    print(f"\n找到 {len(properties)} 个属性:")
    for dispid, value in sorted(properties.items()):
        # 截断过长的值
        val_str = str(value)
        if len(val_str) > 50:
            val_str = val_str[:50] + "..."
        print(f"  DISPID {dispid:3d}: {val_str}")

    return properties

print("="*60)
print("完整探索天正对象属性")
print("="*60)

try:
    # 关闭所有CAD
    print("\n[1] 关闭所有CAD进程...")
    close_all_cad_processes()

    # 启动CAD
    print("\n[2] 启动CAD...")
    proc = start_applicationV9(PTH=r"C:\Tangent\TArchT20V9")
    ensure_single_process()
    wait_quiescent(min_quiet=2.0, timeout=30.0)

    # 打开测试文件
    print("\n[3] 打开测试文件...")
    open_file("D:/claude-tasks/tests/test_files/天正测试文件1.dwg")
    wait_quiescent(min_quiet=1.0, timeout=15.0)

    acad, doc = get_acad_doc()
    ms = doc.ModelSpace

    print(f"\n模型空间对象数量: {ms.Count}")

    # 遍历对象
    door_props = None
    wall_props = None

    for i in range(ms.Count):
        obj = ms.Item(i)
        obj_name = obj.ObjectName

        if obj_name == "TDbOpening" and door_props is None:
            door_props = explore_tarch_object(obj, "天正门窗 (TDbOpening)")

        elif obj_name == "TDbWall" and wall_props is None:
            wall_props = explore_tarch_object(obj, "天正墙 (TDbWall)")

        if door_props and wall_props:
            break

    # 总结
    print(f"\n{'='*60}")
    print("属性映射总结")
    print(f"{'='*60}")

    if door_props:
        print(f"\nTDbOpening (天正门窗):")
        print(f"  DISPID 2  -> Width (宽度)")
        print(f"  DISPID 10 -> Height (高度)")

    if wall_props:
        print(f"\nTDbWall (天正墙):")
        # 根据常见墙体属性推测
        if 2 in wall_props:
            print(f"  DISPID 2  -> 可能是Thickness (厚度): {wall_props[2]}")

except Exception as e:
    print(f"\n错误: {e}")
    import traceback
    traceback.print_exc()

finally:
    print("\n[清理] 关闭CAD...")
    close_all_cad_processes()
