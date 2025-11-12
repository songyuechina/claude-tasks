#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
import win32com.client
import pythoncom
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_file_operations import start_cad_session, restore_to_uncertain_state

def copy_file_content_pywin32(source_file, target_file):
    """使用pywin32将源文件的所有对象复制到目标文件"""
    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")

        # 打开目标文件
        target_doc = acad.Documents.Open(target_file)
        print(f"[打开] 目标文件: {target_file}")

        # 打开源文件
        source_doc = acad.Documents.Open(source_file)
        print(f"[打开] 源文件: {source_file}")

        # 获取源文件的所有对象
        source_ms = source_doc.ModelSpace
        objects_to_copy = []
        for i in range(source_ms.Count):
            objects_to_copy.append(source_ms.Item(i))

        print(f"[信息] 源文件共有 {len(objects_to_copy)} 个对象")

        # 准备对象数组
        obj_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, objects_to_copy)

        # 复制对象到目标文档
        target_ms = target_doc.ModelSpace
        id_pairs = source_doc.CopyObjects(obj_array, target_ms)

        print(f"[成功] 已复制对象到目标文件")

        # 关闭源文件（不保存）
        source_doc.Close(False)
        print(f"[关闭] 源文件")

        # 保存目标文件
        target_doc.Save()
        print(f"[保存] 目标文件")

        return True

    except Exception as e:
        print(f"[错误] {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    # 启动CAD会话
    start_cad_session()

    try:
        # 执行复制
        success = copy_file_content_pywin32(
            source_file="D:/claude-tasks/tests/test_files/B.dwg",
            target_file="D:/claude-tasks/tests/test_files/天正测试文件3.dwg"
        )

        if success:
            print("\n[完成] B.dwg内容已复制到天正测试文件3.dwg")
        else:
            print("\n[失败] 操作未成功")

    finally:
        # 恢复到单文件不确定状态
        restore_to_uncertain_state()
