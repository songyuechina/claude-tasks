#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""简化的天正窗绘制函数 - 假设MC_yuan.dwg已在文件中"""

import sys
import time
from pathlib import Path

sys.path.append(str(Path(__file__).parent))

def insert_tarch_window_simple(p, window_type, width, tolerance=10):
    """
    在指定位置插入天正窗（假设MC_yuan.dwg已在文件中）

    Args:
        p: 插入点 (x, y, z)
        window_type: 窗类型图层名
        width: 窗宽度
        tolerance: 位置容差（检查是否已有门）

    Returns:
        dict: {'success': bool, 'window': 窗对象, 'width': 宽度}
    """
    import win32com.client
    from CAD_basic import stc, transfer_props_by_matchprop, get_object_property, set_object_property, last_obj
    from CAD_coordination import send_cmd_with_sync, wait_quiescent

    valid_types = ["jz-menlianchuang", "jz-dong", "jz-gaochuang", "jz-baiyechuang",
                   "jz-tuchuang", "jz-pingchuang", "jz-zimumen", "jz-juanlianmen",
                   "jz-tuilamen", "jz-shuangmen"]
    if window_type not in valid_types:
        return {'success': False, 'window': None, 'width': None}

    acad = win32com.client.GetActiveObject("AutoCAD.Application")
    doc = acad.ActiveDocument
    ms = doc.ModelSpace

    # 检查该位置是否已有门
    for i in range(ms.Count):
        try:
            obj = ms.Item(i)
            if obj.ObjectName == "TDbOpening":
                ins_pt = obj.InsertionPoint
                if abs(ins_pt[0] - p[0]) < tolerance and abs(ins_pt[1] - p[1]) < tolerance:
                    print(f"[跳过] 位置 {p} 已有门")
                    return {'success': False, 'window': None, 'width': None, 'reason': 'door_exists'}
        except:
            pass

    # 插入门
    send_cmd_with_sync(f"TOpening\n{p[0]},{p[1]}\n\n", wait_after=2.0)
    wait_quiescent(min_quiet=1.0, timeout=10.0)
    time.sleep(1.0)

    door = last_obj()
    if door.ObjectName != "TDbOpening":
        return {'success': False, 'window': None, 'width': None}

    # 获取窗对象
    window = stc(window_type)
    if isinstance(window, list):
        if len(window) == 0:
            print(f"[错误] 未找到 {window_type} 图层的窗对象，请先插入MC_yuan.dwg")
            send_cmd_with_sync("u\n", wait_after=1.0)
            return {'success': False, 'window': None, 'width': None}
        window = window[0]

    # 传递属性
    success = transfer_props_by_matchprop(window, door)
    if not success:
        send_cmd_with_sync("u\n", wait_after=1.0)
        return {'success': False, 'window': None, 'width': None}

    # 设置宽度
    set_object_property(door, 'Width', width)
    actual_width = get_object_property(door, 'Width')

    return {'success': True, 'window': door, 'width': actual_width, 'layer': door.Layer}

if __name__ == "__main__":
    # 测试
    result = insert_tarch_window_simple((5000, 0, 0), "jz-pingchuang", 1200)
    print(f"结果: {result}")
