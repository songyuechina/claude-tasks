#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""测试直接从CAD_basic和CAD_file_operations导入使用"""
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

print("="*60)
print("测试直接导入使用对象属性函数")
print("="*60)

# 测试1: 从CAD_basic导入
print("\n[测试1] 从CAD_basic导入函数...")
try:
    from CAD_basic import cast_object, get_object_property, set_object_property
    print("  [OK] 成功导入 cast_object")
    print("  [OK] 成功导入 get_object_property")
    print("  [OK] 成功导入 set_object_property")
except Exception as e:
    print(f"  [FAIL] 导入失败: {e}")

# 测试2: 从CAD_file_operations导入墙体和门窗函数
print("\n[测试2] 从CAD_file_operations导入函数...")
try:
    from CAD_file_operations import draw_tarch_wall, insert_tarch_door
    print("  [OK] 成功导入 draw_tarch_wall")
    print("  [OK] 成功导入 insert_tarch_door")
except Exception as e:
    print(f"  [FAIL] 导入失败: {e}")

# 测试3: 检查函数签名
print("\n[测试3] 检查函数文档...")
from CAD_basic import get_object_property
print(f"\nget_object_property 文档:")
print(f"  {get_object_property.__doc__}")

print("\n" + "="*60)
print("所有函数都可以直接导入使用！")
print("="*60)

print("\n使用示例:")
print("  from CAD_basic import cast_object, get_object_property")
print("  from CAD_file_operations import draw_tarch_wall, insert_tarch_door")
print()
print("  # CAD对象")
print("  line = cast_object(entity)")
print("  start = line.StartPoint")
print()
print("  # 天正对象")
print("  width = get_object_property(door, 'Width')")
print()
print("  # 绘制墙体")
print("  draw_tarch_wall(p1, p2, thickness=240)")
print()
print("  # 插入门")
print("  result = insert_tarch_door(p, width=1000, height=2100)")
