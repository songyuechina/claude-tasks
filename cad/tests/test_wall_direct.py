#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""直接测试墙体绘制函数（不启动CAD，假设已打开）"""
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_file_operations import draw_tarch_wall, save_file
from CAD_coordination import wait_quiescent
import time

print("="*60)
print("测试墙体绘制功能")
print("="*60)

print("\n等待CAD进入空闲状态...")
wait_quiescent(min_quiet=2.0, timeout=15.0)

p1 = (88800.42585138, 77306.33788321, 0)
p2 = (94193.69907482, 82695.99449027, 0)

print(f"\n绘制墙体:")
print(f"  起点: {p1}")
print(f"  终点: {p2}")
print(f"  墙厚: 320")

result = draw_tarch_wall(p1, p2, thickness=320)

if result:
    print("\n保存文件...")
    time.sleep(2)  # 等待墙体绘制完成
    save_file()
    print("\n测试完成！文件已保存")
else:
    print("\n测试失败")
