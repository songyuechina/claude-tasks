#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""测试 insert_tarch_window V2"""

import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent))

from CAD_basic import start_applicationV9, li
from CAD_basic_operations import open_dwg_paradigm, save_current_dwg_paradigm
from CAD_coordination import ensure_single_process, wait_quiescent
import subprocess
import time

# 1. 关闭所有CAD进程
print("="*60)
print("1. 关闭所有CAD进程...")
ensure_single_process()

# 2. 启动弹窗治理脚本
print("\n2. 启动弹窗治理脚本...")
dialog_killer = Path(__file__).parent / "cad_dialog_killer.py"
subprocess.Popen([sys.executable, str(dialog_killer)],
                 creationflags=subprocess.CREATE_NEW_CONSOLE)
time.sleep(2)

# 3. 启动天正CAD
print("\n3. 启动天正CAD...")
start_applicationV9(PTH=r"C:\Tangent\TArchT20V9", max_retries=3, retry_delay=2.0)
wait_quiescent(min_quiet=1.0, timeout=30.0)

# 4. 打开测试文件
print("\n4. 打开测试文件...")
open_dwg_paradigm("D:/claude-tasks/cad/Function_testing/insert_tarch_window-1.dwg")
wait_quiescent(min_quiet=1.0, timeout=15.0)

# 5. 连接当前文件
print("\n5. 连接当前文件...")
li()

print("\n" + "="*60)
print("开始测试 insert_tarch_window V2")
print("="*60)

# 导入V2函数
from insert_tarch_window_v2 import insert_tarch_window

# 测试1
print("\n测试1: 插入 jz-gaochuang")
result1 = insert_tarch_window(
    p=(38612.86565445, 48750.63891910, 0),
    width=1200,
    height=1000,
    window_type="jz-gaochuang"
)
print(f"结果: {result1}")

if result1['success']:
    # 测试2
    print("\n测试2: 插入 jz-pingchuang")
    result2 = insert_tarch_window(
        p=(44695.30568975, 46646.78059028, 0),
        width=2400,
        height=1800,
        window_type="jz-pingchuang"
    )
    print(f"结果: {result2}")

    # 保存文件
    if result2['success']:
        print("\n保存文件...")
        save_current_dwg_paradigm()
        print("\n" + "="*60)
        print("测试完成！")
        print("="*60)
