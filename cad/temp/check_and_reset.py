#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
import win32com.client
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_file_operations import restore_to_uncertain_state

try:
    # 连接到CAD
    acad = win32com.client.GetActiveObject("AutoCAD.Application")

    # 查看当前打开的文件数
    file_count = acad.Documents.Count
    print(f"[信息] 当前打开文件数: {file_count}")

    # 列出所有打开的文件
    for i in range(file_count):
        doc = acad.Documents.Item(i)
        print(f"  [{i}] {doc.Name}")

    # 恢复到单文件不确定状态
    print("\n[操作] 恢复到单文件不确定状态...")
    restore_to_uncertain_state()

    # 再次查看
    file_count = acad.Documents.Count
    print(f"\n[信息] 恢复后文件数: {file_count}")
    for i in range(file_count):
        doc = acad.Documents.Item(i)
        print(f"  [{i}] {doc.Name} (已保存: {not doc.Saved})")

except Exception as e:
    print(f"[错误] {e}")
