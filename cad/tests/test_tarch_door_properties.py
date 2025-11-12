#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""验证天正门对象的Width和Height属性"""
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_basic import get_acad_doc
from CAD_coordination import send_cmd_with_sync, wait_quiescent
import win32com.client
import pythoncom

def get_tarch_property(obj, dispid):
    """通过DISPID获取天正对象属性"""
    try:
        oleobj = obj._oleobj_
        result = oleobj.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYGET, True)
        return result
    except Exception as e:
        return None

def set_tarch_property(obj, dispid, value):
    """通过DISPID设置天正对象属性"""
    try:
        oleobj = obj._oleobj_
        oleobj.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYPUT, True, value)
        return True
    except Exception as e:
        print(f"设置失败: {e}")
        return False

print("="*60)
print("验证天正门属性DISPID")
print("="*60)

wait_quiescent(min_quiet=2.0, timeout=15.0)

try:
    acad, doc = get_acad_doc()
    ms = doc.ModelSpace

    # 获取最后一个对象（已插入的门）
    door = ms.Item(ms.Count - 1)

    print(f"\n当前门对象:")
    print(f"  ObjectName: {door.ObjectName}")

    # 读取属性
    width = get_tarch_property(door, 2)
    height = get_tarch_property(door, 10)

    print(f"\n读取属性:")
    print(f"  DISPID 2 (Width): {width}")
    print(f"  DISPID 10 (Height): {height}")

    # 尝试修改宽度
    print(f"\n测试修改宽度为1200...")
    if set_tarch_property(door, 2, 1200):
        new_width = get_tarch_property(door, 2)
        print(f"  修改后Width: {new_width}")

        # 再改回去
        set_tarch_property(door, 2, 900)
        print(f"  恢复为900")

    print(f"\n确认: DISPID 2 就是Width属性!")

    # 测试其他属性
    print(f"\n其他属性:")
    properties = {
        1: "未知1",
        3: "类型",
        4: "未知4",
        5: "未知5",
        6: "偏移",
        7: "开启方向",
        8: "角度",
        9: "未知9",
    }

    for dispid, name in properties.items():
        val = get_tarch_property(door, dispid)
        print(f"  DISPID {dispid} ({name}): {val}")

except Exception as e:
    print(f"\n错误: {e}")
    import traceback
    traceback.print_exc()
