#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""探索天正对象的所有属性"""
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_basic import get_acad_doc
from CAD_coordination import wait_quiescent
import pythoncom

def explore_tarch_object(obj, obj_name):
    """探索天正对象的所有DISPID属性"""
    print(f"\n{'='*60}")
    print(f"探索 {obj_name} 对象")
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
        print(f"  DISPID {dispid:3d}: {value}")

    return properties

wait_quiescent(min_quiet=2.0, timeout=15.0)

try:
    acad, doc = get_acad_doc()
    ms = doc.ModelSpace

    print(f"模型空间对象数量: {ms.Count}")

    # 遍历所有对象，找到天正门和天正墙
    door_props = None
    wall_props = None

    for i in range(ms.Count):
        obj = ms.Item(i)
        obj_name = obj.ObjectName

        if obj_name == "TDbOpening" and door_props is None:
            door_props = explore_tarch_object(obj, "天正门窗")

        elif obj_name == "TDbWall" and wall_props is None:
            wall_props = explore_tarch_object(obj, "天正墙")

        if door_props and wall_props:
            break

    # 总结
    print(f"\n{'='*60}")
    print("总结")
    print(f"{'='*60}")

    if door_props:
        print(f"\nTDbOpening (天正门窗) - {len(door_props)} 个属性")
        print("  关键属性:")
        print(f"    DISPID 2  = Width (宽度)")
        print(f"    DISPID 10 = Height (高度)")

    if wall_props:
        print(f"\nTDbWall (天正墙) - {len(wall_props)} 个属性")
        print("  需要进一步分析确定各属性含义")

except Exception as e:
    print(f"\n错误: {e}")
    import traceback
    traceback.print_exc()
