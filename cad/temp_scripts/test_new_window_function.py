#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""测试新的天正窗绘制函数"""

import sys
from pathlib import Path

sys.path.append(str(Path(__file__).parent))

from insert_tarch_window_new import insert_tarch_window_new
from CAD_file_operations import new_file, save_file_as, draw_tarch_wall
from CAD_coordination import wait_quiescent

print("=" * 60)
print("测试新的天正窗绘制函数")
print("=" * 60)

# 步骤1: 创建新文件
print("\n[1/4] 创建新文件...")
new_file()
wait_quiescent(min_quiet=1.0, timeout=10.0)
print("✓ 新文件已创建")

# 步骤2: 绘制三角形墙
print("\n[2/4] 绘制三角形墙...")
# 定义三角形的三个顶点（边长 >= 12000）
p1 = (0, 0, 0)
p2 = (15000, 0, 0)
p3 = (7500, 13000, 0)

# 绘制三条墙
draw_tarch_wall(p1, p2)
draw_tarch_wall(p2, p3)
draw_tarch_wall(p3, p1)
wait_quiescent(min_quiet=1.0, timeout=10.0)
print("✓ 三角形墙已绘制")

# 保存文件（必须先保存，否则copy_file_content_pywin32无法工作）
print("\n[2.5/4] 保存文件...")
save_path = "D:/claude-tasks/tests/test_files/天正窗测试文件.dwg"
save_file_as(save_path)
print(f"✓ 文件已保存到: {save_path}")

# 步骤3: 在每堵墙的1/3位置插入窗
print("\n[3/4] 在墙的1/3位置插入窗...")

# 墙1 (p1-p2): 在1/3位置插入jz-tuilamen，宽度1200
wall1_pos = (p1[0] + (p2[0] - p1[0]) / 3, p1[1] + (p2[1] - p1[1]) / 3, 0)
print(f"\n在位置 {wall1_pos} 插入 jz-tuilamen...")
result1 = insert_tarch_window_new(wall1_pos, "jz-tuilamen", 1200)
if result1['success']:
    print(f"✓ 窗1 插入成功: {result1['layer']}, 宽度={result1['width']}")
else:
    print("✗ 窗1 插入失败")

# 墙2 (p2-p3): 在1/3位置插入jz-gaochuang，宽度900
wall2_pos = (p2[0] + (p3[0] - p2[0]) / 3, p2[1] + (p3[1] - p2[1]) / 3, 0)
print(f"\n在位置 {wall2_pos} 插入 jz-gaochuang...")
result2 = insert_tarch_window_new(wall2_pos, "jz-gaochuang", 900)
if result2['success']:
    print(f"✓ 窗2 插入成功: {result2['layer']}, 宽度={result2['width']}")
else:
    print("✗ 窗2 插入失败")

# 墙3 (p3-p1): 在1/3位置插入jz-shuangmen，宽度1400
wall3_pos = (p3[0] + (p1[0] - p3[0]) / 3, p3[1] + (p1[1] - p3[1]) / 3, 0)
print(f"\n在位置 {wall3_pos} 插入 jz-shuangmen...")
result3 = insert_tarch_window_new(wall3_pos, "jz-shuangmen", 1400)
if result3['success']:
    print(f"✓ 窗3 插入成功: {result3['layer']}, 宽度={result3['width']}")
else:
    print("✗ 窗3 插入失败")

# 步骤4: 保存文件（最终保存）
print("\n[4/4] 最终保存文件...")
from CAD_file_operations import save_file
save_file()
print(f"✓ 文件已保存")

print("\n" + "=" * 60)
print("✓ 测试完成！")
print(f"成功插入的窗:")
if result1['success']:
    print(f"  - 窗1: {result1['layer']}, 宽度={result1['width']}")
if result2['success']:
    print(f"  - 窗2: {result2['layer']}, 宽度={result2['width']}")
if result3['success']:
    print(f"  - 窗3: {result3['layer']}, 宽度={result3['width']}")
print("=" * 60)
