#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简单测试 cad_zt_two
"""

import sys
import time
sys.path.append("D:/claude-tasks/cad/scripts")

from CAD_basic import jingchengshu_wenjian, get_open_document_names, li
from CAD_file_operations import cad_zt_zero, cad_zt_two

print("=" * 60)
print("测试 cad_zt_two")
print("=" * 60)

# 清理
print("\n[1] 清理环境")
cad_zt_zero()
time.sleep(2)
print(f"  CAD进程数: {jingchengshu_wenjian()}")

# 测试：从0开始 -> 2个文件
print("\n[2] 调用 cad_zt_two")
cad_zt_two("D:/claude-tasks/cad/xitongwenjian/0.dwg", "D:/claude-tasks/cad/xitongwenjian/1.dwg")
time.sleep(15)

print("\n[3] 检查结果")
shu = jingchengshu_wenjian()
print(f"  CAD进程数: {shu}")

if shu == 1:
    try:
        li()
        time.sleep(2)
        docs = get_open_document_names()
        print(f"  打开文件数: {len(docs)}")
        print(f"  文件: {docs}")

        if len(docs) == 2:
            print("\n成功: 有2个文件")
        else:
            print(f"\n失败: 应该有2个文件，实际有{len(docs)}个")
    except Exception as e:
        print(f"  错误: {e}")
else:
    print(f"\n失败: CAD进程数应该是1，实际是{shu}")

# 清理
print("\n[4] 清理")
cad_zt_zero()
time.sleep(2)

print("\n完成")
