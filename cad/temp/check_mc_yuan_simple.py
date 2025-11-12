#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""快速检查MC_yuan.dwg的对象"""
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_basic import close_all_cad_processes, start_applicationV9, get_acad_doc
from CAD_file_operations import open_file
from CAD_coordination import ensure_single_process, wait_quiescent

try:
    close_all_cad_processes()
    proc = start_applicationV9(PTH=r"C:\Tangent\TArchT20V9")
    ensure_single_process()
    wait_quiescent(min_quiet=2.0, timeout=30.0)

    # 直接打开MC_yuan.dwg
    print("打开MC_yuan.dwg...")
    open_file("D:/claude-tasks/cad/xitongwenjian/MC_yuan.dwg")
    wait_quiescent(min_quiet=2.0, timeout=15.0)

    _, doc = get_acad_doc()
    ms = doc.ModelSpace

    print(f"\n总对象数: {ms.Count}")

    # 收集所有对象的坐标信息
    objects_info = []
    for i in range(min(ms.Count, 100)):  # 只检查前100个
        try:
            obj = ms.Item(i)
            bbox = obj.GetBoundingBox()
            p1, p2 = bbox[0], bbox[1]

            objects_info.append({
                'index': i,
                'type': obj.ObjectName,
                'x_min': min(p1[0], p2[0]),
                'x_max': max(p1[0], p2[0]),
                'y_min': min(p1[1], p2[1]),
                'y_max': max(p1[1], p2[1])
            })
        except:
            pass

    if objects_info:
        # 打印坐标范围
        all_x_min = min(o['x_min'] for o in objects_info)
        all_x_max = max(o['x_max'] for o in objects_info)
        all_y_min = min(o['y_min'] for o in objects_info)
        all_y_max = max(o['y_max'] for o in objects_info)

        print(f"\n整体坐标范围:")
        print(f"X: {all_x_min:.0f} ~ {all_x_max:.0f}")
        print(f"Y: {all_y_min:.0f} ~ {all_y_max:.0f}")

        # 打印前20个对象的详细信息
        print(f"\n前20个对象:")
        for obj in objects_info[:20]:
            print(f"  [{obj['index']}] {obj['type']:20s} X:{obj['x_min']:8.0f}~{obj['x_max']:8.0f} Y:{obj['y_min']:9.0f}~{obj['y_max']:9.0f}")

except Exception as e:
    print(f"错误: {e}")
    import traceback
    traceback.print_exc()

finally:
    close_all_cad_processes()
