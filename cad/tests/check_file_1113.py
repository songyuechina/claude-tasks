#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""检查命令测试1113.dwg文件内容"""

import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent.parent / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from CAD_file_operations import open_file, close_file, cad_zt_zero
import win32com.client

def check_file_content():
    """检查文件中的对象"""
    print("="*60)
    print("检查命令测试1113.dwg文件内容")
    print("="*60)

    test_file = "D:/claude-tasks/cad/tests/test_files/命令测试1113.dwg"

    try:
        # 打开文件
        print(f"\n[1] 打开文件: {test_file}")
        if not open_file(test_file):
            print("[错误] 文件打开失败")
            return False

        # 获取文档对象
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        doc = acad.ActiveDocument
        model_space = doc.ModelSpace

        print(f"\n[2] 文件信息:")
        print(f"    文件名: {doc.Name}")
        print(f"    对象总数: {model_space.Count}")

        # 统计对象类型
        object_types = {}
        print(f"\n[3] 对象列表:")

        for i in range(model_space.Count):
            obj = model_space.Item(i)
            obj_name = obj.ObjectName

            if obj_name not in object_types:
                object_types[obj_name] = 0
            object_types[obj_name] += 1

            print(f"    [{i+1}] {obj_name}")

        print(f"\n[4] 对象类型统计:")
        for obj_type, count in object_types.items():
            print(f"    {obj_type}: {count}个")

        # 检查是否包含所有要求的对象
        print(f"\n[5] 任务完成情况检查:")
        required_objects = {
            'AcDbPolyline': '多段线',
            'AcDbCircle': '圆',
            'AcDbLine': '直线',
        }

        all_completed = True
        for obj_type, obj_desc in required_objects.items():
            if obj_type in object_types:
                print(f"    [√] {obj_desc} ({obj_type}) - 已完成")
            else:
                print(f"    [×] {obj_desc} ({obj_type}) - 未找到")
                all_completed = False

        # 检查天正墙
        has_tarch_wall = any('TDb' in obj_type or 'Tarch' in obj_type.lower()
                             for obj_type in object_types.keys())
        if has_tarch_wall:
            print(f"    [√] 天正墙 - 已完成")
        else:
            print(f"    [×] 天正墙 - 未找到")
            all_completed = False

        if all_completed:
            print("\n[成功] 所有绘图任务已完成！")
        else:
            print("\n[警告] 部分任务未完成")

        return all_completed

    except Exception as e:
        print(f"\n[错误] 检查异常: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        # 关闭文件
        print("\n[6] 关闭文件...")
        cad_zt_zero()

if __name__ == "__main__":
    success = check_file_content()
    sys.exit(0 if success else 1)
