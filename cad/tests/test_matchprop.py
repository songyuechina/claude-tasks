#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""测试transfer_props_by_matchprop函数"""
import sys
import subprocess
import time
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_basic import (
    close_all_cad_processes, start_applicationV9, get_acad_doc,
    last_obj, get_object_property, transfer_props_by_matchprop
)
from CAD_file_operations import new_file, draw_tarch_wall, save_file
from CAD_coordination import ensure_single_process, wait_quiescent, send_cmd_with_sync

print("="*60)
print("测试transfer_props_by_matchprop")
print("="*60)

dialog_killer = subprocess.Popen([sys.executable, "D:/claude-tasks/cad/scripts/cad_dialog_killer.py"])
time.sleep(2)

try:
    close_all_cad_processes()
    proc = start_applicationV9(PTH=r"C:\Tangent\TArchT20V9")
    ensure_single_process()
    wait_quiescent(min_quiet=2.0, timeout=30.0)

    new_file("D:/claude-tasks/tests/test_files/天正测试文件3.dwg")
    time.sleep(3)

    # 绘制一堵墙
    draw_tarch_wall((100000, 75000, 0), (105000, 75000, 0), thickness=200)
    time.sleep(2)

    # 插入源窗（类型3）
    print("\n步骤1: 插入源窗（类型3）...")
    cmd = f"TOpening\n102000,75000\n\n"
    send_cmd_with_sync(cmd, wait_after=2.0)
    time.sleep(3)
    wait_quiescent(min_quiet=1.0, timeout=10.0)
    src_window = last_obj()
    print(f"源窗对象: {src_window.ObjectName}")

    # 从MC_yuan获取类型3的窗并复制Type属性
    from CAD_file_operations import insert_file_as_block
    _, doc = get_acad_doc()
    ms = doc.ModelSpace
    count_before = ms.Count

    insert_file_as_block("D:/claude-tasks/cad/xitongwenjian/MC_yuan.dwg", 0, 0, 0)
    wait_quiescent(min_quiet=1.0, timeout=10.0)
    send_cmd_with_sync("_EXPLODE\n_L\n\n\n\n\n\n\n\n\n", wait_after=2.0)
    wait_quiescent(min_quiet=1.0, timeout=10.0)

    # 找类型3的窗
    mc_window = None
    for i in range(ms.Count):
        try:
            obj = ms.Item(i)
            if obj.ObjectName == "TDbOpening" and obj != src_window:
                bbox = obj.GetBoundingBox()
                cy = (bbox[0][1] + bbox[1][1]) / 2
                if 506900 <= cy <= 507200:
                    mc_window = obj
                    break
        except:
            pass

    if mc_window:
        print(f"找到MC窗: {mc_window.ObjectName}")
        src_type = get_object_property(mc_window, 'Type')
        print(f"MC窗Type属性: {src_type}")

        # 测试transfer_props_by_matchprop
        print("\n步骤2: 测试transfer_props_by_matchprop...")
        print(f"源窗Type前: {get_object_property(src_window, 'Type')}")
        result = transfer_props_by_matchprop(mc_window, src_window)
        print(f"transfer_props_by_matchprop结果: {result}")
        print(f"源窗Type后: {get_object_property(src_window, 'Type')}")

        # 清理MC对象
        for i in range(ms.Count - 1, count_before - 1, -1):
            try:
                obj = ms.Item(i)
                if obj != src_window:
                    obj.Delete()
            except:
                pass

    # 插入目标门
    print("\n步骤3: 插入目标门...")
    cmd = f"TOpening\n103000,75000\n\n"
    send_cmd_with_sync(cmd, wait_after=2.0)
    time.sleep(3)
    wait_quiescent(min_quiet=1.0, timeout=10.0)
    target_door = last_obj()
    print(f"目标门对象: {target_door.ObjectName}")
    print(f"目标门Type前: {get_object_property(target_door, 'Type')}")

    # 测试从源窗到目标门的属性传递
    print("\n步骤4: 从源窗传递属性到目标门...")
    result = transfer_props_by_matchprop(src_window, target_door)
    print(f"transfer_props_by_matchprop结果: {result}")
    print(f"目标门Type后: {get_object_property(target_door, 'Type')}")

    save_file()
    print("\n测试完成！")

except Exception as e:
    print(f"\n错误: {e}")
    import traceback
    traceback.print_exc()

finally:
    close_all_cad_processes()
    dialog_killer.terminate()
