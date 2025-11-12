#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试6个CAD状态管理函数
"""

import sys
import time
sys.path.append("D:/claude-tasks/cad/scripts")

from CAD_basic import jingchengshu_wenjian, get_open_document_names
from CAD_file_operations import (
    cad_zt_zero, cad_zt_oneb, cad_zt_oned,
    cad_zt_two, cad_zt_much
)

def print_status():
    """打印当前CAD状态"""
    shu = jingchengshu_wenjian()
    print(f"  当前CAD进程数: {shu}")
    if shu == 1:
        try:
            docs = get_open_document_names()
            print(f"  当前打开的文件数: {len(docs)}")
            print(f"  文件列表: {docs}")
        except:
            print("  无法获取打开的文件列表")
    print()

def wait_stable(seconds=3):
    """等待系统稳定"""
    time.sleep(seconds)

print("=" * 60)
print("CAD状态管理函数测试")
print("=" * 60)
print()

# 测试1: cad_zt_zero()
print("[测试1] cad_zt_zero() - 关闭所有CAD")
print("-" * 60)
cad_zt_zero()
wait_stable(2)
print_status()

# 测试2: cad_zt_oneb()
print("[测试2] cad_zt_oneb() - 1进程+1空白文件")
print("-" * 60)
cad_zt_oneb()
wait_stable(5)
print_status()

# 测试3: cad_zt_oned()
print("[测试3] cad_zt_oned() - 1进程+1指定文件")
print("-" * 60)
print("  使用文件: D:/claude-tasks/cad/xitongwenjian/0.dwg")
cad_zt_oned("D:/claude-tasks/cad/xitongwenjian/0.dwg")
wait_stable(5)
print_status()

# 测试4: cad_zt_two()
print("[测试4] cad_zt_two() - 1进程+2文件")
print("-" * 60)
print("  文件1: D:/claude-tasks/tests/test_files/A.dwg")
print("  文件2: D:/claude-tasks/tests/test_files/B.dwg")
cad_zt_two("D:/claude-tasks/tests/test_files/A.dwg", "D:/claude-tasks/tests/test_files/B.dwg")
wait_stable(8)
print_status()

# 测试5: cad_zt_much()
print("[测试5] cad_zt_much() - 1进程+多文件(>2)")
print("-" * 60)
print("  文件1: D:/claude-tasks/cad/xitongwenjian/0.dwg")
print("  文件2: D:/claude-tasks/cad/xitongwenjian/1.dwg")
print("  文件3: D:/claude-tasks/cad/xitongwenjian/2.dwg")
cad_zt_much(
    "D:/claude-tasks/cad/xitongwenjian/0.dwg",
    "D:/claude-tasks/cad/xitongwenjian/1.dwg",
    "D:/claude-tasks/cad/xitongwenjian/2.dwg"
)
wait_stable(10)
print_status()

# 测试6: 再次测试 cad_zt_oned() - 从多文件状态切换到单文件
print("[测试6] cad_zt_oned() - 从多文件切换到1文件")
print("-" * 60)
print("  使用文件: D:/claude-tasks/tests/test_files/窗测试.dwg")
cad_zt_oned("D:/claude-tasks/tests/test_files/窗测试.dwg")
wait_stable(5)
print_status()

# 测试7: 再次测试 cad_zt_two() - 从单文件状态切换到双文件
print("[测试7] cad_zt_two() - 从单文件切换到2文件")
print("-" * 60)
print("  文件1: D:/claude-tasks/tests/test_files/C_exploded.dwg")
print("  文件2: D:/claude-tasks/cad/xitongwenjian/1.dwg")
cad_zt_two("D:/claude-tasks/tests/test_files/C_exploded.dwg", "D:/claude-tasks/cad/xitongwenjian/1.dwg")
wait_stable(8)
print_status()

# 清理：关闭所有CAD
print("[清理] 关闭所有CAD进程")
print("-" * 60)
cad_zt_zero()
wait_stable(2)
print_status()

print("=" * 60)
print("测试完成")
print("=" * 60)
