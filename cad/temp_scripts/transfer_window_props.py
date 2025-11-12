#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""使用 stc 获取对象并传递窗属性给门"""

import sys
import io
from pathlib import Path

# 重新配置 stdout 为 utf-8 编码，避免 emoji 字符问题
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

sys.path.append(str(Path(__file__).parent))

from CAD_basic import stc, transfer_props_by_matchprop
from CAD_file_operations import save_file

print("=" * 60)
print("使用 stc 传递窗属性给门")
print("=" * 60)

# 1. 使用 stc 获取 jz-pingchuang 图层的窗对象
print("\n[1/3] 使用 stc('jz-pingchuang') 获取窗对象...")
window = stc('jz-pingchuang')
if isinstance(window, list):
    window = window[0]
print(f"✓ 窗对象: {window.ObjectName}, 图层: {window.Layer}")

# 2. 使用 stc 获取 WINDOW 图层的3个门对象
print("\n[2/3] 使用 stc('WINDOW') 获取门对象...")
doors = stc('WINDOW')
if not isinstance(doors, list):
    doors = [doors]
print(f"✓ 找到 {len(doors)} 个门对象")

# 3. 保存门的原始宽度
print("\n[3/6] 保存门的原始宽度...")
from CAD_basic import get_object_property
door_widths = []
for i, door in enumerate(doors, 1):
    width = get_object_property(door, 'Width')
    door_widths.append(width)
    print(f"  门 {i} 原始宽度: {width}")

# 4. 使用 transfer_props_by_matchprop 传递属性
print("\n[4/6] 使用 transfer_props_by_matchprop 传递属性...")
for i, door in enumerate(doors, 1):
    print(f"  正在处理第 {i} 个门...")
    transfer_props_by_matchprop(window, door)
    print(f"  ✓ 第 {i} 个门属性传递完成")

# 5. 恢复窗的宽度为门的原始宽度
print("\n[5/6] 恢复窗的宽度为门的原始宽度...")
from CAD_basic import set_object_property
for i, (door, original_width) in enumerate(zip(doors, door_widths), 1):
    set_object_property(door, 'Width', original_width)
    print(f"  窗 {i} 宽度已恢复为: {original_width}")

# 保存文件
print("\n[6/6] 保存文件...")
save_file()

print("\n" + "=" * 60)
print("✓ 任务完成！3个门已变成窗，宽度已恢复")
print("=" * 60)
