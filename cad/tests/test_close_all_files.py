#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试 close_all_files() 函数
测试关闭所有文件功能
"""

import sys
import time
from pathlib import Path

# 添加脚本路径
SCRIPT_DIR = Path(__file__).parent.parent / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from CAD_file_operations import (
    close_all_files,
    cad_zt_much,
    cad_zt_oneb
)

def test_close_all_files():
    """测试关闭所有文件功能"""
    print("=" * 60)
    print("测试 close_all_files() 函数")
    print("=" * 60)

    try:
        # 1. 创建多文件状态
        print("\n[步骤1] 创建多文件状态...")
        file1 = "D:/claude-tasks/cad/xitongwenjian/0.dwg"
        file2 = "D:/claude-tasks/cad/xitongwenjian/1.dwg"
        file3 = "D:/claude-tasks/cad/xitongwenjian/2.dwg"

        cad_zt_much(file1, file2, file3)
        time.sleep(3)

        # 验证多文件状态
        import win32com.client
        acad = win32com.client.GetActiveObject("AutoCAD.Application")

        open_files_before = []
        for doc in acad.Documents:
            filename = Path(doc.Name).name
            open_files_before.append(filename)
            print(f"  已打开: {filename}")

        if len(open_files_before) < 2:
            print(f"[警告] 打开的文件数少于预期: {len(open_files_before)}")
        else:
            print(f"[成功] 已打开 {len(open_files_before)} 个文件")

        # 2. 调用 close_all_files()
        print("\n[步骤2] 调用 close_all_files()...")
        result = close_all_files()

        if not result:
            print("[错误] close_all_files() 返回False")
            return False

        time.sleep(2)
        print("[成功] close_all_files() 执行完成")

        # 3. 验证所有文件已关闭
        print("\n[步骤3] 验证文件已关闭...")
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        doc_count_after = acad.Documents.Count

        print(f"  关闭前文档数: {len(open_files_before)}")
        print(f"  关闭后文档数: {doc_count_after}")

        # CAD可能会自动创建一个新的空白文档，所以允许0或1个文档
        if doc_count_after > 1:
            print(f"[错误] 关闭后仍有多个文档: {doc_count_after}")
            return False

        if doc_count_after == 0:
            print("[成功] 所有文档已关闭，无打开文档")
        elif doc_count_after == 1:
            # 检查是否是新的空白文档
            active_doc = acad.ActiveDocument
            doc_name = active_doc.Name
            doc_path = active_doc.Path
            print(f"[信息] 剩余1个文档: {doc_name}")
            print(f"[信息] 文档路径: {doc_path if doc_path else '(未保存)'}")

            # 验证不是之前打开的文件
            if doc_name in open_files_before:
                print(f"[错误] 文档 {doc_name} 未被关闭")
                return False

            print("[成功] 所有原有文档已关闭，CAD自动创建了新空白文档")

        # 4. 再次测试：对空白状态调用 close_all_files()
        print("\n[步骤4] 再次测试（对空白或单文件状态）...")
        result2 = close_all_files()

        if not result2:
            print("[错误] 第二次 close_all_files() 返回False")
            return False

        time.sleep(2)
        print("[成功] 第二次调用也正常")

        # 5. 再次创建多文件状态并关闭
        print("\n[步骤5] 第三次测试（再次创建多文件并关闭）...")
        cad_zt_much(file1, file2, file3)
        time.sleep(3)

        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        doc_count_before_third = acad.Documents.Count
        print(f"  第三次测试前文档数: {doc_count_before_third}")

        result3 = close_all_files()
        if not result3:
            print("[错误] 第三次 close_all_files() 返回False")
            return False

        time.sleep(2)

        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        doc_count_after_third = acad.Documents.Count
        print(f"  第三次测试后文档数: {doc_count_after_third}")

        if doc_count_after_third > 1:
            print(f"[错误] 第三次测试后仍有多个文档: {doc_count_after_third}")
            return False

        print("[成功] 第三次测试也正常")

        print("\n" + "=" * 60)
        print("[完成] close_all_files() 测试通过!")
        print("=" * 60)
        print("测试结果:")
        print(f"  - 第一次关闭: {len(open_files_before)} → {doc_count_after} 个文档")
        print(f"  - 第二次关闭: 对空白状态调用正常")
        print(f"  - 第三次关闭: {doc_count_before_third} → {doc_count_after_third} 个文档")
        print("关闭所有文件功能正常")

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
    success = test_close_all_files()
    sys.exit(0 if success else 1)
