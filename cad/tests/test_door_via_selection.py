#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""通过选择集获取刚插入的门"""
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_basic import get_acad_doc
from CAD_coordination import send_cmd_with_sync, wait_quiescent
import pythoncom
import time

def get_tarch_property(obj, dispid):
    """通过DISPID获取天正对象属性"""
    try:
        oleobj = obj._oleobj_
        result = oleobj.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYGET, True)
        return result
    except:
        return None

print("="*60)
print("通过选择集获取门对象")
print("="*60)

wait_quiescent(min_quiet=2.0, timeout=15.0)

try:
    acad, doc = get_acad_doc()

    # 先选择所有对象，记录数量
    print("\n创建选择集...")
    try:
        doc.SelectionSets.Item("TempSS").Delete()
    except:
        pass

    ss = doc.SelectionSets.Add("TempSS")
    ss.SelectOnScreen()  # 让用户选择刚插入的门

    print(f"选中对象数量: {ss.Count}")

    if ss.Count > 0:
        door = ss.Item(0)
        print(f"对象类型: {door.ObjectName}")

        if door.ObjectName == "TDbOpening":
            width = get_tarch_property(door, 2)
            height = get_tarch_property(door, 10)

            print(f"\n门的属性:")
            print(f"  Width (DISPID 2): {width}")
            print(f"  Height (DISPID 10): {height}")

    ss.Delete()

except Exception as e:
    print(f"\n错误: {e}")
    import traceback
    traceback.print_exc()
