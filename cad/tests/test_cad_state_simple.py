#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简化的CAD状态管理函数测试
"""

import sys
import time
sys.path.append("D:/claude-tasks/cad/scripts")

from CAD_basic import jingchengshu_wenjian, get_open_document_names, li
from CAD_file_operations import (
    cad_zt_zero, cad_zt_oneb, cad_zt_oned,
    cad_zt_two, cad_zt_much
)

def print_status():
    """打印当前CAD状态"""
    shu = jingchengshu_wenjian()
    print(f"  CAD进程数: {shu}")
    if shu == 1:
        try:
            li()
            time.sleep(1)
            docs = get_open_document_names()
            print(f"  打开文件数: {len(docs)}")
            if docs:
                print(f"  文件: {docs}")
        except Exception as e:
            print(f"  无法获取文件列表: {e}")
    print()

print("=" * 60)
print("CAD状态管理函数测试")
print("=" * 60)
print()

# 测试1: cad_zt_zero()
print("[测试1] cad_zt_zero()")
cad_zt_zero()
time.sleep(2)
print_status()

# 测试2: cad_zt_oneb()
print("[测试2] cad_zt_oneb()")
cad_zt_oneb()
time.sleep(8)
print_status()

# 测试3: cad_zt_two()
print("[测试3] cad_zt_two()")
cad_zt_two("D:/claude-tasks/cad/xitongwenjian/0.dwg", "D:/claude-tasks/cad/xitongwenjian/1.dwg")
time.sleep(10)
print_status()

# 测试4: cad_zt_much()
print("[测试4] cad_zt_much()")
cad_zt_much(
    "D:/claude-tasks/tests/test_files/A.dwg",
    "D:/claude-tasks/tests/test_files/B.dwg",
    "D:/claude-tasks/tests/test_files/窗测试.dwg"
)
time.sleep(12)
print_status()

# 测试5: cad_zt_oned()
print("[测试5] cad_zt_oned()")
cad_zt_oned("D:/claude-tasks/tests/test_files/C_exploded.dwg")
time.sleep(8)
print_status()

# 清理
print("[清理] cad_zt_zero()")
cad_zt_zero()
time.sleep(2)
print_status()

print("=" * 60)
print("测试完成")
print("=" * 60)
