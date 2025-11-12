#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""简单的天正门测试"""
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent))

from CAD_file_operations import (start_cad_session, new_file, save_file_as,
                                  draw_tarch_wall, insert_tarch_door, save_file,
                                  restore_to_uncertain_state)
from CAD_coordination import wait_quiescent, send_cmd_with_sync

print("="*60)
print("简单的天正门测试")
print("="*60)

# 1. 启动天正CAD
print("\n[1/6] 启动天正CAD...")
start_cad_session()

# 2. 新建文件
print("\n[2/6] 新建文件...")
new_file()
wait_quiescent(min_quiet=1.0, timeout=10.0)

# 3. 绘制三角形墙
print("\n[3/6] 绘制三角形墙...")
p1, p2, p3 = (0, 0, 0), (15000, 0, 0), (7500, 13000, 0)
draw_tarch_wall(p1, p2)
draw_tarch_wall(p2, p3)
draw_tarch_wall(p3, p1)
wait_quiescent(min_quiet=1.0, timeout=10.0)
print("[OK] 墙已绘制")

# 4. 保存文件
print("\n[4/6] 保存文件...")
save_path = "D:/claude-tasks/tests/test_files/天正门测试文件.dwg"
save_file_as(save_path)

# 5. 插入3个门
print("\n[5/6] 插入3个门...")
r1 = insert_tarch_door((5000, 0, 0), width=1200)
r2 = insert_tarch_door((12500, 4333, 0), width=900)
r3 = insert_tarch_door((5000, 8667, 0), width=1400)

# 6. 保存
print("\n[6/6] 保存文件...")
save_file()

print("\n" + "="*60)
print("测试完成！")
print(f"门1 (1200): {r1['success']}, 宽度={r1['width']}")
print(f"门2 (900): {r2['success']}, 宽度={r2['width']}")
print(f"门3 (1400): {r3['success']}, 宽度={r3['width']}")
print("="*60)

# 恢复状态
restore_to_uncertain_state()
