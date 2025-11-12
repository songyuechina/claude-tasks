#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
sys.path.append("D:/claude-tasks/cad/scripts")

from CAD_basic import start_applicationV9, li
from CAD_coordination import ensure_single_process, wait_quiescent
from CAD_file_operations import insert_tarch_window
import win32com.client

# 1. 关闭所有CAD
ensure_single_process()

# 2. 启动CAD
start_applicationV9(PTH=r"C:\Tangent\TArchT20V9", max_retries=3, retry_delay=2.0)
wait_quiescent(min_quiet=1.0, timeout=30.0)

# 3. 打开文件
acad = win32com.client.GetActiveObject("AutoCAD.Application")
while acad.Documents.Count > 0:
    acad.Documents.Item(0).Close(False)
acad.Documents.Open("D:/claude-tasks/cad/Function_testing/insert_tarch_window - 3.dwg")
wait_quiescent(min_quiet=1.0, timeout=15.0)

# 4. 连接文件
li()

# 5. 测试1
print("\n测试1: jz-gaochuang")
r1 = insert_tarch_window(
    p=(38612.86565445, 48750.63891910, 0),
    width=1200,
    height=1000,
    window_type="jz-gaochuang",
    delete_mc_yuan=False
)
print(f"结果: {r1['success']}")

# 6. 测试2
print("\n测试2: jz-pingchuang")
r2 = insert_tarch_window(
    p=(44695.30568975, 46646.78059028, 0),
    width=2400,
    height=1800,
    window_type="jz-pingchuang",
    delete_mc_yuan=True
)
print(f"结果: {r2['success']}")

# 7. 保存
acad.ActiveDocument.Save()
print("\n测试完成，文件已保存")
