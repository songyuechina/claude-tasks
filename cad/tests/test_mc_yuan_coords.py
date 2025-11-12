#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""测试MC_yuan.dwg的窗坐标"""
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_basic import close_all_cad_processes, start_applicationV9, get_acad_doc
from CAD_file_operations import new_file, insert_file_exploded, save_file
from CAD_coordination import ensure_single_process, wait_quiescent

print("="*60)
print("测试MC_yuan.dwg坐标")
print("="*60)

try:
    close_all_cad_processes()
    proc = start_applicationV9(PTH=r"C:\Tangent\TArchT20V9")
    ensure_single_process()
    wait_quiescent(min_quiet=2.0, timeout=30.0)

    # 新建文件
    new_file("D:/claude-tasks/tests/test_files/MC_yuan_test.dwg")
    wait_quiescent(min_quiet=1.0, timeout=15.0)

    # 插入MC_yuan.dwg
    print("\n插入MC_yuan.dwg...")
    insert_file_exploded("D:/claude-tasks/cad/xitongwenjian/MC_yuan.dwg", 0, 0, 0)
    wait_quiescent(min_quiet=2.0, timeout=15.0)

    # 获取所有对象并打印坐标范围
    _, doc = get_acad_doc()
    ms = doc.ModelSpace

    print(f"\n模型空间对象数量: {ms.Count}")

    if ms.Count > 0:
        # 获取边界
        min_x, max_x = float('inf'), float('-inf')
        min_y, max_y = float('inf'), float('-inf')

        for i in range(ms.Count):
            try:
                obj = ms.Item(i)
                bbox = obj.GetBoundingBox()
                p1, p2 = bbox[0], bbox[1]

                min_x = min(min_x, p1[0], p2[0])
                max_x = max(max_x, p1[0], p2[0])
                min_y = min(min_y, p1[1], p2[1])
                max_y = max(max_y, p1[1], p2[1])
            except:
                pass

        print(f"\n坐标范围:")
        print(f"  X: {min_x:.0f} 到 {max_x:.0f}")
        print(f"  Y: {min_y:.0f} 到 {max_y:.0f}")

        # 测试选择窗类型3的坐标
        print(f"\n测试选择窗类型3...")
        p1 = (39060, 527163, 0)
        p2 = (41693, 524872, 0)
        print(f"  选择坐标: {p1[:2]} 到 {p2[:2]}")

        from CAD_basic import select_entities_in_window
        selected = select_entities_in_window(p1[0], p1[1], p2[0], p2[1], ty=2.0, select_mode="_W")
        print(f"  选中对象数: {len(selected) if selected else 0}")

        if selected:
            for i, obj in enumerate(selected[:5]):  # 只显示前5个
                try:
                    print(f"    对象{i+1}: {obj.ObjectName}")
                except:
                    print(f"    对象{i+1}: Unknown")

    save_file()
    print("\n完成！文件已保存")

except Exception as e:
    print(f"\n错误: {e}")
    import traceback
    traceback.print_exc()

finally:
    close_all_cad_processes()
