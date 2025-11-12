#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""简化的天正窗测试 - 使用单一窗类型"""
import sys, time
from pathlib import Path
sys.path.append(str(Path(__file__).parent))

from CAD_file_operations import (start_cad_session, new_file, save_file_as,
                                  draw_tarch_wall, insert_tarch_door,
                                  copy_file_content_pywin32, save_file,
                                  restore_to_uncertain_state)
from CAD_coordination import wait_quiescent
from CAD_basic import transfer_props_by_matchprop, get_object_property, set_object_property
import win32com.client

print("="*60)
print("简化的天正窗测试")
print("="*60)

# 1. 启动天正CAD
print("\n[1/7] 启动天正CAD...")
start_cad_session()

# 2. 新建文件
print("\n[2/7] 新建文件...")
new_file()
wait_quiescent(min_quiet=1.0, timeout=10.0)

# 3. 绘制三角形墙
print("\n[3/7] 绘制三角形墙...")
p1, p2, p3 = (0, 0, 0), (15000, 0, 0), (7500, 13000, 0)
draw_tarch_wall(p1, p2)
draw_tarch_wall(p2, p3)
draw_tarch_wall(p3, p1)
wait_quiescent(min_quiet=1.0, timeout=10.0)
print("[OK] 墙已绘制")

# 4. 保存文件
print("\n[4/7] 保存文件...")
save_path = "D:/claude-tasks/tests/test_files/天正窗测试文件.dwg"
save_file_as(save_path)

# 5. 插入MC_yuan.dwg
print("\n[5/7] 插入MC_yuan.dwg...")
copy_file_content_pywin32('D:/claude-tasks/cad/xitongwenjian/MC_yuan.dwg', save_path)
wait_quiescent(min_quiet=1.0, timeout=10.0)
print("[OK] MC_yuan.dwg已插入")

# 6. 插入3个门
print("\n[6/7] 插入3个门...")
r1 = insert_tarch_door((5000, 0, 0), width=1200)
r2 = insert_tarch_door((12500, 4333, 0), width=900)
r3 = insert_tarch_door((5000, 8667, 0), width=1400)
print(f"门1: {r1['success']}, 门2: {r2['success']}, 门3: {r3['success']}")

# 7. 将门转换为窗
print("\n[7/7] 将门转换为窗...")
acad = win32com.client.GetActiveObject("AutoCAD.Application")
ms = acad.ActiveDocument.ModelSpace

# 查找3种窗对象（从MC_yuan.dwg）
windows = {'jz-tuilamen': None, 'jz-gaochuang': None, 'jz-shuangmen': None}
doors = []

for i in range(ms.Count):
    try:
        obj = ms.Item(i)
        if obj.ObjectName == "TDbOpening":
            layer = obj.Layer
            if layer in windows and windows[layer] is None:
                windows[layer] = obj
                print(f"  找到窗: {layer}")
            elif layer == 'WINDOW':
                width = get_object_property(obj, 'Width')
                doors.append((obj, width))
    except:
        pass

print(f"  找到 {len(doors)} 个门")

# 传递属性并设置宽度
for door, original_width in doors:
    if abs(original_width - 1200) < 10:
        window_type, target_width = 'jz-tuilamen', 1200
    elif abs(original_width - 900) < 10:
        window_type, target_width = 'jz-gaochuang', 900
    elif abs(original_width - 1400) < 10:
        window_type, target_width = 'jz-shuangmen', 1400
    else:
        continue

    window_src = windows.get(window_type)
    if window_src:
        success = transfer_props_by_matchprop(window_src, door)
        if success:
            set_object_property(door, 'Width', target_width)
            actual_width = get_object_property(door, 'Width')
            print(f"  门({original_width}) -> {window_type}(宽度{actual_width}): 成功")
        else:
            print(f"  门({original_width}) -> {window_type}: 失败")

# 保存
save_file()

print("\n" + "="*60)
print("测试完成！")
print("="*60)

# 恢复状态
restore_to_uncertain_state()
