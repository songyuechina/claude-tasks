#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试 activate_document_by_name() 函数
测试文档激活功能
"""

import sys
import time
from pathlib import Path

# 添加脚本路径
SCRIPT_DIR = Path(__file__).parent.parent / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from CAD_file_operations import (
    activate_document_by_name,
    cad_zt_two,
    cad_zt_oneb
)

def test_activate_document():
    """测试文档激活功能"""
    print("=" * 60)
    print("测试 activate_document_by_name() 函数")
    print("=" * 60)

    try:
        # 1. 确保CAD状态为双文件确定状态（打开0.dwg和1.dwg）
        print("\n[步骤1] 确保CAD处于双文件确定状态...")
        file1 = "D:/claude-tasks/cad/xitongwenjian/0.dwg"
        file2 = "D:/claude-tasks/cad/xitongwenjian/1.dwg"

        cad_zt_two(file1, file2)
        time.sleep(2)
        print("[成功] CAD状态准备完成 (0.dwg 和 1.dwg 已打开)")

        # 2. 验证两个文件都已打开
        print("\n[步骤2] 验证文件已打开...")
        import win32com.client
        acad = win32com.client.GetActiveObject("AutoCAD.Application")

        open_files = []
        for doc in acad.Documents:
            filename = Path(doc.Name).name
            open_files.append(filename)
            print(f"  已打开: {filename}")

        if "0.dwg" not in open_files or "1.dwg" not in open_files:
            print("[错误] 未找到预期的文件")
            return False

        # 3. 激活0.dwg
        print("\n[步骤3] 激活 0.dwg...")
        result1 = activate_document_by_name("0.dwg")
        if not result1:
            print("[错误] 激活 0.dwg 失败")
            return False

        time.sleep(1)

        # 验证0.dwg是当前激活文档
        active_doc_name = Path(acad.ActiveDocument.Name).name
        if active_doc_name != "0.dwg":
            print(f"[错误] 当前激活文档不是 0.dwg，而是 {active_doc_name}")
            return False

        print(f"[成功] 已激活 0.dwg，当前激活文档: {active_doc_name}")

        # 4. 激活1.dwg
        print("\n[步骤4] 激活 1.dwg...")
        result2 = activate_document_by_name("1.dwg")
        if not result2:
            print("[错误] 激活 1.dwg 失败")
            return False

        time.sleep(1)

        # 验证1.dwg是当前激活文档
        active_doc_name = Path(acad.ActiveDocument.Name).name
        if active_doc_name != "1.dwg":
            print(f"[错误] 当前激活文档不是 1.dwg，而是 {active_doc_name}")
            return False

        print(f"[成功] 已激活 1.dwg，当前激活文档: {active_doc_name}")

        # 5. 再次激活0.dwg
        print("\n[步骤5] 再次激活 0.dwg...")
        result3 = activate_document_by_name("0.dwg")
        if not result3:
            print("[错误] 再次激活 0.dwg 失败")
            return False

        time.sleep(1)

        # 验证0.dwg是当前激活文档
        active_doc_name = Path(acad.ActiveDocument.Name).name
        if active_doc_name != "0.dwg":
            print(f"[错误] 当前激活文档不是 0.dwg，而是 {active_doc_name}")
            return False

        print(f"[成功] 已激活 0.dwg，当前激活文档: {active_doc_name}")

        # 6. 测试激活不存在的文件（应该返回False）
        print("\n[步骤6] 测试激活不存在的文件...")
        result4 = activate_document_by_name("nonexistent.dwg")
        if result4:
            print("[错误] 激活不存在的文件应该返回False，但返回了True")
            return False

        print("[成功] 激活不存在的文件正确返回False")

        print("\n" + "=" * 60)
        print("[完成] activate_document_by_name() 测试通过!")
        print("=" * 60)
        print("测试结果:")
        print("  - 激活 0.dwg: 成功")
        print("  - 激活 1.dwg: 成功")
        print("  - 再次激活 0.dwg: 成功")
        print("  - 激活不存在文件: 正确返回False")
        print("文档激活功能正常")

        return True

    except Exception as e:
        print(f"\n[错误] 测试异常: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        # 恢复到单文件不确定状态
        print("\n[清理] 恢复到单文件不确定状态...")
        try:
            cad_zt_oneb()
            print("[成功] CAD状态已恢复")
        except:
            pass

if __name__ == "__main__":
    success = test_activate_document()
    sys.exit(0 if success else 1)
