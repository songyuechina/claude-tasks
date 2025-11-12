#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
import time
import win32com.client
from pathlib import Path

# 使用pywin32简单创建新文件
try:
    # 连接到CAD
    acad = win32com.client.GetActiveObject("AutoCAD.Application")

    # 创建新文档
    doc = acad.Documents.Add()
    time.sleep(1)

    # 保存文件
    output_path = "D:/claude-tasks/tests/test_files/天正测试文件3.dwg"
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    doc.SaveAs(output_path)
    print(f"[成功] 文件已创建: {output_path}")

except Exception as e:
    print(f"[错误] {e}")
