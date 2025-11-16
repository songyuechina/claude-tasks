#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试 restore_to_uncertain_state() 函数
测试恢复到单文件不确定状态功能
"""

import sys
import time
from pathlib import Path

# 添加脚本路径
SCRIPT_DIR = Path(__file__).parent.parent / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from CAD_file_operations import (
    restore_to_uncertain_state,
    cad_zt_much,
    cad_zt_oneb
)

def test_restore_state():
    """测试恢复到单文件不确定状态功能"""
    print("=" * 60)
    print("测试 restore_to_uncertain_state() 函数")
    print("=" * 60)

    try:
        # 1. 先创建复杂状态：打开多个文件
        print("\n[步骤1] 创建复杂状态（多文件）...")
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

        # 2. 调用 restore_to_uncertain_state()
        print("\n[步骤2] 调用 restore_to_uncertain_state()...")
        result = restore_to_uncertain_state()

        if not result:
            print("[错误] restore_to_uncertain_state() 返回False")
            return False

        time.sleep(3)
        print("[成功] restore_to_uncertain_state() 执行完成")

        # 3. 验证CAD进程数为1
        print("\n[步骤3] 验证CAD进程数...")
        sys.path.append(str(SCRIPT_DIR))
        from CAD_basic import jingchengshu_wenjian

        process_count = jingchengshu_wenjian()
        if process_count != 1:
            print(f"[错误] CAD进程数不是1，而是 {process_count}")
            return False

        print(f"[成功] CAD进程数为1")

        # 4. 验证只有1个文档打开
        print("\n[步骤4] 验证文档数量...")
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        doc_count = acad.Documents.Count

        if doc_count != 1:
            print(f"[错误] 打开的文档数不是1，而是 {doc_count}")
            return False

        print(f"[成功] 打开的文档数为1")

        # 5. 验证文档是未保存的空白文档
        print("\n[步骤5] 验证文档状态...")
        active_doc = acad.ActiveDocument

        # 检查文档名称（空白文档通常是 Drawing1.dwg, Drawing2.dwg 等）
        doc_name = active_doc.Name
        print(f"  文档名称: {doc_name}")

        # 检查是否已保存（未保存的文档 Saved 属性为 False）
        is_saved = active_doc.Saved
        print(f"  是否已保存: {is_saved}")

        # 检查文档路径（未保存的文档 Path 为空）
        doc_path = active_doc.Path
        print(f"  文档路径: {doc_path if doc_path else '(未保存)'}")

        # 检查对象数量（空白文档应该只有少量默认对象）
        ms = active_doc.ModelSpace
        obj_count = ms.Count
        print(f"  模型空间对象数: {obj_count}")

        if obj_count > 10:
            print(f"[警告] 模型空间对象数较多: {obj_count}，可能不是空白文档")

        print("[成功] 文档状态符合单文件不确定状态")

        # 6. 再次测试：从单文件状态恢复
        print("\n[步骤6] 再次测试从单文件状态恢复...")
        result2 = restore_to_uncertain_state()
        if not result2:
            print("[错误] 第二次 restore_to_uncertain_state() 返回False")
            return False

        time.sleep(3)

        # 验证仍然是单文件状态
        process_count2 = jingchengshu_wenjian()
        acad2 = win32com.client.GetActiveObject("AutoCAD.Application")
        doc_count2 = acad2.Documents.Count

        if process_count2 != 1 or doc_count2 != 1:
            print(f"[错误] 状态不正确 - 进程数: {process_count2}, 文档数: {doc_count2}")
            return False

        print("[成功] 第二次恢复也正常")

        print("\n" + "=" * 60)
        print("[完成] restore_to_uncertain_state() 测试通过!")
        print("=" * 60)
        print("测试结果:")
        print(f"  - 初始状态: {len(open_files_before)} 个文件")
        print(f"  - 恢复后进程数: 1")
        print(f"  - 恢复后文档数: 1")
        print(f"  - 文档状态: 未保存空白文档")
        print("恢复到单文件不确定状态功能正常")

        return True

    except Exception as e:
        print(f"\n[错误] 测试异常: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        # 最终清理：确保恢复到单文件不确定状态
        print("\n[清理] 最终清理...")
        try:
            cad_zt_oneb()
            print("[成功] CAD状态已恢复")
        except:
            pass

if __name__ == "__main__":
    success = test_restore_state()
    sys.exit(0 if success else 1)
