#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""研究天正门对象属性"""
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_basic import get_acad_doc
from CAD_coordination import send_cmd_with_sync, wait_quiescent
import win32com.client

print("="*60)
print("研究天正门对象属性")
print("="*60)

wait_quiescent(min_quiet=2.0, timeout=15.0)

# 在墙体上的一个点插入门
p = (105000, 70000, 0)  # 矩形底边中点

print(f"\n在点 {p[:2]} 插入天正门...")

# 发送TOpening命令
cmd = f"TOpening\n{p[0]},{p[1]}\n\n"
send_cmd_with_sync(cmd, wait_after=2.0)

print("\n获取最后生成的对象...")

try:
    acad, doc = get_acad_doc()
    ms = doc.ModelSpace

    # 获取最后一个对象
    last_obj = ms.Item(ms.Count - 1)

    print(f"\n对象信息:")
    print(f"  ObjectName: {last_obj.ObjectName}")
    print(f"  EntityName: {last_obj.EntityName if hasattr(last_obj, 'EntityName') else 'N/A'}")
    print(f"  Handle: {last_obj.Handle}")

    # 尝试直接访问Width
    print(f"\n尝试直接访问Width属性...")
    try:
        width = last_obj.Width
        print(f"  Width (直接访问): {width}")
    except Exception as e:
        print(f"  直接访问失败: {e}")

    # 尝试使用QueryInterface
    print(f"\n尝试使用QueryInterface转换...")
    try:
        from win32com.client import CastTo
        cast_obj = CastTo(last_obj, "ITDbOpening")
        width = cast_obj.Width
        print(f"  Width (ITDbOpening): {width}")
    except Exception as e:
        print(f"  ITDbOpening转换失败: {e}")

    # 尝试GetProperty
    print(f"\n尝试使用GetProperty...")
    try:
        width = last_obj.GetProperty("Width")
        print(f"  Width (GetProperty): {width}")
    except Exception as e:
        print(f"  GetProperty失败: {e}")

    # 列出所有可用属性
    print(f"\n列出对象的所有属性和方法:")
    for attr in dir(last_obj):
        if not attr.startswith('_'):
            try:
                val = getattr(last_obj, attr)
                if not callable(val):
                    print(f"  {attr}: {val}")
            except:
                pass

except Exception as e:
    print(f"\n错误: {e}")
    import traceback
    traceback.print_exc()
