#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""测试绘制矩形墙体"""
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_file_operations import draw_tarch_wall, save_file
from CAD_coordination import wait_quiescent
import time

print("="*60)
print("测试绘制矩形墙体")
print("="*60)

# 定义矩形的四个顶点（逆时针方向）
p1 = (100000, 70000, 0)  # 左下
p2 = (110000, 70000, 0)  # 右下
p3 = (110000, 80000, 0)  # 右上
p4 = (100000, 80000, 0)  # 左上

print(f"\n矩形顶点:")
print(f"  左下: {p1}")
print(f"  右下: {p2}")
print(f"  右上: {p3}")
print(f"  左上: {p4}")

print("\n等待CAD空闲...")
wait_quiescent(min_quiet=2.0, timeout=15.0)

# 绘制四条边
edges = [
    (p1, p2, "底边"),
    (p2, p3, "右边"),
    (p3, p4, "顶边"),
    (p4, p1, "左边")
]

for i, (start, end, name) in enumerate(edges, 1):
    print(f"\n[{i}/4] 绘制{name}: {start[:2]} -> {end[:2]}")

    result = draw_tarch_wall(start, end, thickness=240)

    if result:
        print(f"      成功")
    else:
        print(f"      失败")
        break

    # 每条边之间等待一下
    time.sleep(1)

print("\n保存文件...")
save_file()

print("\n" + "="*60)
print("矩形墙体绘制完成！")
print("="*60)
