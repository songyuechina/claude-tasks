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

        # 确保路径正确
        target_path = str(Path("D:/claude-tasks/tests/test_files/天正测试文件3.dwg").resolve())
        source_path = str(Path("D:/claude-tasks/tests/test_files/B.dwg").resolve())

        print(f"[路径] 目标: {target_path}")
        print(f"[路径] 源: {source_path}")

        # 先打开源文件
        print(f"[打开] 源文件...")
        source_doc = acad.Documents.Open(source_path)
        time.sleep(3)  # 等待完全打开

        # 获取源文件的所有对象
        source_ms = source_doc.ModelSpace
        objects_count = source_ms.Count
        print(f"[信息] 源文件共有 {objects_count} 个对象")

        # 创建对象列表
        objects_to_copy = []
        for i in range(objects_count):
            try:
                obj = source_ms.Item(i)
                objects_to_copy.append(obj)
            except:
                pass

        print(f"[信息] 准备复制 {len(objects_to_copy)} 个对象")

        # 打开目标文件
        print(f"[打开] 目标文件...")
        time.sleep(2)  # 确保源文件完全加载
        target_doc = acad.Documents.Open(target_path)
        time.sleep(3)  # 等待完全打开

        # 获取目标ModelSpace
        target_ms = target_doc.ModelSpace

        # 准备对象数组
        obj_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, objects_to_copy)

        # 复制对象从源文档到目标文档
        print(f"[复制] 正在复制对象...")
        id_pairs = source_doc.CopyObjects(obj_array, target_ms)
        print(f"[成功] 对象已复制")

        # 关闭源文件（不保存）
        source_doc.Close(False)
        print(f"[关闭] 源文件")

        # 检查并炸开块引用
        print(f"[检查] 查找块引用...")
        blocks_to_explode = []
        for i in range(target_ms.Count):
            try:
                obj = target_ms.Item(i)
                if obj.ObjectName == "AcDbBlockReference":
                    blocks_to_explode.append(obj)
            except:
                pass

        if blocks_to_explode:
            print(f"[炸开] 发现 {len(blocks_to_explode)} 个块，正在炸开...")
            for block in blocks_to_explode:
                try:
                    block.Explode()
                    block.Delete()
                except:
                    pass
            print(f"[成功] 块已炸开")

        # 保存目标文件
        target_doc.Save()
        print(f"[保存] 目标文件")

        print("\n[完成] B.dwg内容已复制到天正测试文件3.dwg")

    except Exception as e:
        print(f"[错误] {e}")
        import traceback
        traceback.print_exc()

    finally:
        restore_to_uncertain_state()
