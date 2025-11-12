#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""测试insert_tarch_door函数"""
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_file_operations import insert_tarch_door, save_file
from CAD_coordination import wait_quiescent

print("="*60)
print("测试insert_tarch_door函数")
print("="*60)

wait_quiescent(min_quiet=2.0, timeout=15.0)

# 在矩形墙体的底边插入几个门
positions = [
    ((101000, 70000, 0), 1000, 2100, "默认门"),
    ((103000, 70000, 0), 1200, 2400, "大门"),
    ((108000, 70000, 0), 800, 2000, "小门"),
]

for i, (pos, w, h, name) in enumerate(positions, 1):
    print(f"\n[{i}/{len(positions)}] 插入{name}...")
    print(f"  位置: {pos[:2]}")
    print(f"  尺寸: 宽{w} x 高{h}")

    result = insert_tarch_door(pos, width=w, height=h)

    if result['success']:
        print(f"  成功 - 实际宽度:{result['width']}, 高度:{result['height']}")
    else:
        print(f"  失败")

print("\n保存文件...")
save_file()

print("\n" + "="*60)
print("测试完成！")
print("="*60)
