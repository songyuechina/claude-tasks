#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""使用InvokeTypes研究天正门对象"""
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_coordination import wait_quiescent
import win32com.client
import pythoncom

print("="*60)
print("使用InvokeTypes研究天正门对象")
print("="*60)

wait_quiescent(min_quiet=2.0, timeout=15.0)

try:
    acad = win32com.client.Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    ms = doc.ModelSpace
    last_obj = ms.Item(ms.Count - 1)

    print(f"\n对象: {last_obj.ObjectName}")

    # 获取COM对象的IDispatch接口
    oleobj = last_obj._oleobj_

    print(f"\n尝试通过IDispatch获取Width...")

    # 尝试不同的DISPID
    for dispid in range(1, 100):
        try:
            result = oleobj.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYGET, True)
            if result is not None:
                print(f"  DISPID {dispid}: {result}")
        except:
            pass

except Exception as e:
    print(f"\n错误: {e}")
    import traceback
    traceback.print_exc()
