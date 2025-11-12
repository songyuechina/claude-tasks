#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""使用晚期绑定研究天正门对象"""
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_coordination import wait_quiescent
import win32com.client

print("="*60)
print("使用晚期绑定研究天正门对象")
print("="*60)

wait_quiescent(min_quiet=2.0, timeout=15.0)

try:
    # 使用晚期绑定获取CAD对象
    acad = win32com.client.Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    ms = doc.ModelSpace

    # 获取最后一个对象（刚插入的门）
    last_obj = ms.Item(ms.Count - 1)

    print(f"\n对象信息:")
    print(f"  ObjectName: {last_obj.ObjectName}")
    print(f"  Handle: {last_obj.Handle}")

    # 尝试直接访问Width（晚期绑定）
    print(f"\n尝试访问Width属性（晚期绑定）...")
    try:
        width = last_obj.Width
        print(f"  Width: {width}")
    except Exception as e:
        print(f"  失败: {e}")

    # 尝试访问其他可能的属性
    print(f"\n尝试访问其他属性...")
    attrs_to_try = ['Width', 'Height', 'Depth', 'OpeningWidth', 'DoorWidth',
                    'Size', 'Dimension', 'Parameters']

    for attr in attrs_to_try:
        try:
            val = getattr(last_obj, attr)
            print(f"  {attr}: {val}")
        except:
            pass

    # 尝试使用GetXData
    print(f"\n尝试获取扩展数据...")
    try:
        xdata_type = []
        xdata_val = []
        last_obj.GetXData("", xdata_type, xdata_val)
        print(f"  XData Type: {xdata_type}")
        print(f"  XData Val: {xdata_val}")
    except Exception as e:
        print(f"  GetXData失败: {e}")

except Exception as e:
    print(f"\n错误: {e}")
    import traceback
    traceback.print_exc()
