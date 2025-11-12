#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试修复后的 cad_zt_two 函数
"""

import sys
import time
sys.path.append("D:/claude-tasks/cad/scripts")

from CAD_basic import jingchengshu_wenjian, get_open_document_names, li
from CAD_file_operations import cad_zt_zero, cad_zt_oneb, cad_zt_two

print("=" * 60)
print("测试 cad_zt_two 函数")
print("=" * 60)

# 清理环境
print("\n[准备] 关闭所有CAD")
cad_zt_zero()
time.sleep(2)

# 测试1: 从0个进程开始
print("\n[测试1] 从0个进程开始 -> 2个文件")
cad_zt_two("D:/claude-tasks/cad/xitongwenjian/0.dwg", "D:/claude-tasks/cad/xitongwenjian/1.dwg")
time.sleep(10)
shu = jingchengshu_wenjian()
print(f"  CAD进程数: {shu}")
if shu == 1:
    li()
    time.sleep(1)
    docs = get_open_document_names()
    print(f"  打开文件数: {len(docs)}")
    print(f"  文件: {docs}")
    if len(docs) == 2:
        print("  ✅ 测试1通过")
    else:
        print("  ❌ 测试1失败")

# 测试2: 从1个文件开始 -> 2个文件
print("\n[测试2] 从1个文件开始 -> 2个文件")
cad_zt_two("D:/claude-tasks/tests/test_files/A.dwg", "D:/claude-tasks/tests/test_files/B.dwg")
time.sleep(10)
shu = jingchengshu_wenjian()
print(f"  CAD进程数: {shu}")
if shu == 1:
    li()
    time.sleep(1)
    docs = get_open_document_names()
    print(f"  打开文件数: {len(docs)}")
    print(f"  文件: {docs}")
    if len(docs) == 2:
        print("  ✅ 测试2通过")
    else:
        print("  ❌ 测试2失败")

# 测试3: 从2个文件开始 -> 保持2个文件
print("\n[测试3] 从2个文件开始 -> 保持2个文件")
cad_zt_two("D:/claude-tasks/cad/xitongwenjian/0.dwg", "D:/claude-tasks/cad/xitongwenjian/2.dwg")
time.sleep(10)
shu = jingchengshu_wenjian()
print(f"  CAD进程数: {shu}")
if shu == 1:
    li()
    time.sleep(1)
    docs = get_open_document_names()
    print(f"  打开文件数: {len(docs)}")
    print(f"  文件: {docs}")
    if len(docs) == 2:
        print("  ✅ 测试3通过")
    else:
        print("  ❌ 测试3失败")

# 清理
print("\n[清理] 关闭所有CAD")
cad_zt_zero()
time.sleep(2)

print("\n" + "=" * 60)
print("测试完成")
print("=" * 60)
