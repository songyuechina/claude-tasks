#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
检查测试文件中的对象
"""

import sys
from pathlib import Path

sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_basic import li, start_applicationV9, get_acad_doc
from CAD_file_operations import copy_file_content_pywin32
from CAD_coordination import ensure_single_process, wait_quiescent
import win32com.client

def main():
    print("="*70)
    print("检查测试文件中的对象")
    print("="*70)

    # 1. 关闭所有CAD
    print("\n[1/4] 关闭所有CAD进程...")
    ensure_single_process()

    # 2. 启动CAD
    print("\n[2/4] 启动CAD...")
    start_applicationV9(PTH=r"C:\Tangent\TArchT20V9")
    wait_quiescent(min_quiet=1.0, timeout=30.0)

    # 3. 打开文件
    print("\n[3/4] 打开测试文件...")
    acad = win32com.client.GetActiveObject("AutoCAD.Application")
    while acad.Documents.Count > 0:
        acad.Documents.Item(0).Close(False)

    test_file = "D:/claude-tasks/cad/Function_testing/insert_tarch_window-1.dwg"
    acad.Documents.Open(test_file)
    wait_quiescent(min_quiet=1.0, timeout=15.0)

    # 4. 连接并检查对象
    print("\n[4/4] 检查对象...")
    li()

    _, doc = get_acad_doc()
    ms = doc.ModelSpace

    print(f"\n模型空间中共有 {ms.Count} 个对象:\n")

    for i in range(ms.Count):
        obj = ms.Item(i)
        print(f"对象 {i}:")
        print(f"  ObjectName: {obj.ObjectName}")
        print(f"  Layer: {obj.Layer}")
        try:
            print(f"  Handle: {obj.Handle}")
        except:
            pass
        print()

    # 插入MC_yuan.dwg
    print("\n插入MC_yuan.dwg后的对象:")
    print("="*70)

    copy_file_content_pywin32("D:/claude-tasks/cad/xitongwenjian/MC_yuan.dwg", test_file)
    li()

    print(f"\n模型空间中共有 {ms.Count} 个对象:\n")

    for i in range(ms.Count):
        obj = ms.Item(i)
        print(f"对象 {i}:")
        print(f"  ObjectName: {obj.ObjectName}")
        print(f"  Layer: {obj.Layer}")
        try:
            print(f"  Handle: {obj.Handle}")
        except:
            pass

        # 如果是墙体，打印位置信息
        if "wall" in obj.ObjectName.lower() or "qiang" in obj.Layer.lower():
            try:
                print(f"  [墙体] 尝试获取位置信息...")
                # 尝试获取坐标
                if hasattr(obj, 'StartPoint'):
                    print(f"  StartPoint: {obj.StartPoint}")
                if hasattr(obj, 'EndPoint'):
                    print(f"  EndPoint: {obj.EndPoint}")
            except Exception as e:
                print(f"  无法获取位置: {e}")
        print()

    print("\n检查完成!")


if __name__ == "__main__":
    main()
