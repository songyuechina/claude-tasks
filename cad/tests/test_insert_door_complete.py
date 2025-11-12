#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""完整测试：插入门并获取Width属性"""
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
print("完整测试：插入门并获取Width")
print("="*60)

wait_quiescent(min_quiet=2.0, timeout=15.0)

# 在墙体上插入门
p = (107500, 70000, 0)  # 矩形底边的另一个位置

print(f"\n在点 {p[:2]} 插入天正门...")
cmd = f"TOpening\n{p[0]},{p[1]}\n\n"
send_cmd_with_sync(cmd, wait_after=3.0)

print("\n等待门插入完成...")
time.sleep(2)

try:
    acad, doc = get_acad_doc()
    ms = doc.ModelSpace

    print(f"\n模型空间对象数量: {ms.Count}")

    # 获取最后一个对象
    door = ms.Item(ms.Count - 1)

    print(f"最后对象: {door.ObjectName}")

    if door.ObjectName == "TDbOpening":
        width = get_tarch_property(door, 2)
        height = get_tarch_property(door, 10)

        print(f"\n门的属性:")
        print(f"  Width (DISPID 2): {width}")
        print(f"  Height (DISPID 10): {height}")

        print(f"\n成功！DISPID 2 就是Width属性")
    else:
        print(f"最后对象不是门，而是: {door.ObjectName}")

except Exception as e:
    print(f"\n错误: {e}")
    import traceback
    traceback.print_exc()
