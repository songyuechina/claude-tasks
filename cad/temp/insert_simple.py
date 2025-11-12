#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
import win32com.client
import time
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_file_operations import start_cad_session, restore_to_uncertain_state

if __name__ == "__main__":
    start_cad_session()

    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")

        # 打开目标文件
        target_doc = acad.Documents.Open("D:/claude-tasks/tests/test_files/天正测试文件3.dwg")
        print(f"[打开] 目标文件")
        time.sleep(1)

        # 使用 InsertBlock 插入源文件（会炸开）
        ms = target_doc.ModelSpace
        insert_point = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [0, 0, 0])

        # 插入块
        block_ref = ms.InsertBlock(insert_point, "D:/claude-tasks/tests/test_files/B.dwg", 1.0, 1.0, 1.0, 0)
        print(f"[插入] 已插入块")

        # 炸开块
        block_ref.Explode()
        print(f"[炸开] 已炸开块")

        # 删除块引用
        block_ref.Delete()

        # 保存
        target_doc.Save()
        print(f"[保存] 已保存文件")

        print("\n[完成] B.dwg内容已炸开插入到天正测试文件3.dwg")

    except Exception as e:
        print(f"[错误] {e}")
        import traceback
        traceback.print_exc()

    finally:
        restore_to_uncertain_state()
