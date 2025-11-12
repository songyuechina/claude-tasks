#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""最终的天正窗测试 - 使用门+MC_yuan+属性传递"""
import sys, time
from pathlib import Path
sys.path.append(str(Path(__file__).parent))

from CAD_file_operations import (start_cad_session, new_file, save_file_as,
                                  draw_tarch_wall, insert_tarch_door,
                                  copy_file_content_pywin32, save_file,
                                  restore_to_uncertain_state)
from CAD_coordination import wait_quiescent
from CAD_basic import transfer_props_by_matchprop, get_object_property
import win32com.client

print("="*60)
print("最终的天正窗测试")
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

# 查找窗对象（从MC_yuan.dwg）
windows = {
    'jz-tuilamen': None,
    'jz-gaochuang': None,
    'jz-shuangmen': None
}

for i in range(ms.Count):
    try:
        obj = ms.Item(i)
        if obj.ObjectName == "TDbOpening":
            layer = obj.Layer
            if layer in windows and windows[layer] is None:
                windows[layer] = obj
                print(f"  找到窗: {layer}")
    except:
        pass

# 查找门对象
doors = []
for i in range(ms.Count):
    try:
        obj = ms.Item(i)
        if obj.ObjectName == "TDbOpening":
            layer = obj.Layer
            if layer not in windows:
                doors.append((obj, layer))
    except:
        pass

print(f"  找到 {len(doors)} 个门")

# 传递属性
results = []
for door, layer in doors:
    # 根据宽度匹配窗类型
    width = get_object_property(door, 'Width')
    if abs(width - 1200) < 10:
        window_type = 'jz-tuilamen'
    elif abs(width - 900) < 10:
        window_type = 'jz-gaochuang'
    elif abs(width - 1400) < 10:
        window_type = 'jz-shuangmen'
    else:
        continue

    window = windows.get(window_type)
    if window:
        success = transfer_props_by_matchprop(window, door)
        results.append({'width': width, 'type': window_type, 'success': success})
        print(f"  门(宽度{width}) -> {window_type}: {success}")

# 保存
save_file()

print("\n" + "="*60)
print("测试完成！")
for r in results:
    print(f"  宽度{r['width']} -> {r['type']}: {r['success']}")
print("="*60)

# 恢复状态
restore_to_uncertain_state()
