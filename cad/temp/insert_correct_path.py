#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
import win32com.client
import pythoncom
import time
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_file_operations import start_cad_session, restore_to_uncertain_state

if __name__ == "__main__":
    start_cad_session()

    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")

        # 打开目标文件
        target_path = str(Path("D:/claude-tasks/tests/test_files/天正测试文件3.dwg").resolve())
        source_path = str(Path("D:/claude-tasks/tests/test_files/B.dwg").resolve())

        print(f"[路径] 目标: {target_path}")
        print(f"[路径] 源: {source_path}")

        target_doc = acad.Documents.Open(target_path)
        print(f"[打开] 目标文件")
        time.sleep(2)

        # 获取ModelSpace
        ms = target_doc.ModelSpace

        # 创建插入点
        insert_point = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0, 0.0])

        # 插入块（不使用 InsertBlock，而是使用 import 或其他方法）
        # 尝试使用 SendCommand 方法
        target_doc.SendCommand("-INSERT\n" + source_path + "\n0,0,0\n1\n1\n0\n")
        time.sleep(3)

        # 炸开最后插入的对象
        target_doc.SendCommand("_EXPLODE\n_L\n\n")
        time.sleep(2)

        # 保存
        target_doc.Save()
        print(f"[保存] 已保存文件")

        print("\n[完成] B.dwg内容已插入到天正测试文件3.dwg")

    except Exception as e:
        print(f"[错误] {e}")
        import traceback
        traceback.print_exc()

    finally:
        restore_to_uncertain_state()
