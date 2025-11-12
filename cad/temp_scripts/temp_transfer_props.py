#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""临时脚本：传递窗属性给门"""

import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent))

from CAD_basic import stc, transfer_props_by_matchprop
from CAD_file_operations import save_file

print("=" * 60)
print("传递窗属性给门")
print("=" * 60)

# 1. 获取 jz-pingchuang 图层的窗对象
print("\n[1/3] 获取 jz-pingchuang 图层的窗对象...")
window = stc('jz-pingchuang')
print(f"窗对象: {window.ObjectName}, 图层: {window.Layer}")

# 2. 获取 WINDOW 图层的3个门对象
print("\n[2/3] 获取 WINDOW 图层的门对象...")
doors = stc('WINDOW')
if isinstance(doors, list):
    print(f"找到 {len(doors)} 个门")
else:
    doors = [doors]
    print(f"找到 1 个门")

# 3. 传递属性
print("\n[3/3] 传递属性...")
for i, door in enumerate(doors, 1):
    print(f"正在传递属性给第 {i} 个门...")
    transfer_props_by_matchprop(window, door)
    print(f"第 {i} 个门属性已传递")

# 保存文件
print("\n保存文件...")
save_file()
print("\n" + "=" * 60)
print("完成！")
print("=" * 60)
