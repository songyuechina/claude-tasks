#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""测试简化的天正窗函数"""

import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent))

from insert_tarch_window_simple import insert_tarch_window_simple
from CAD_file_operations import new_file, save_file_as, draw_tarch_wall, copy_file_content_pywin32
from CAD_coordination import wait_quiescent

print("=" * 60)
print("测试简化的天正窗函数")
print("=" * 60)

# 1. 创建新文件
print("\n[1/5] 创建新文件...")
new_file()
wait_quiescent(min_quiet=1.0, timeout=10.0)

# 2. 绘制三角形墙
print("\n[2/5] 绘制三角形墙...")
p1, p2, p3 = (0, 0, 0), (15000, 0, 0), (7500, 13000, 0)
draw_tarch_wall(p1, p2)
draw_tarch_wall(p2, p3)
draw_tarch_wall(p3, p1)
wait_quiescent(min_quiet=1.0, timeout=10.0)

# 3. 保存文件
print("\n[3/5] 保存文件...")
save_path = "D:/claude-tasks/tests/test_files/天正窗测试文件.dwg"
save_file_as(save_path)

# 4. 插入MC_yuan.dwg（一次性）
print("\n[4/5] 插入MC_yuan.dwg...")
copy_file_content_pywin32('D:/claude-tasks/cad/xitongwenjian/MC_yuan.dwg', save_path)
wait_quiescent(min_quiet=1.0, timeout=10.0)

# 5. 插入3个窗
print("\n[5/5] 插入窗...")
wall1_pos = (p1[0] + (p2[0] - p1[0]) / 3, p1[1] + (p2[1] - p1[1]) / 3, 0)
wall2_pos = (p2[0] + (p3[0] - p2[0]) / 3, p2[1] + (p3[1] - p2[1]) / 3, 0)
wall3_pos = (p3[0] + (p1[0] - p3[0]) / 3, p3[1] + (p1[1] - p3[1]) / 3, 0)

r1 = insert_tarch_window_simple(wall1_pos, "jz-tuilamen", 1200)
r2 = insert_tarch_window_simple(wall2_pos, "jz-gaochuang", 900)
r3 = insert_tarch_window_simple(wall3_pos, "jz-shuangmen", 1400)

from CAD_file_operations import save_file
save_file()

print("\n" + "=" * 60)
print("测试完成！")
print(f"窗1: {r1}")
print(f"窗2: {r2}")
print(f"窗3: {r3}")
print("=" * 60)
