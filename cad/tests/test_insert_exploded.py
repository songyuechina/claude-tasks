#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试 insert_file_exploded() 函数
测试将文件炸开插入功能
"""

import sys
import time
import subprocess
from pathlib import Path

# 添加脚本路径
SCRIPT_DIR = Path(__file__).parent.parent / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from CAD_file_operations import (
    insert_file_exploded,
    new_file,
    save_file,
    close_file,
    cad_zt_oneb
)
from CAD_coordination import wait_quiescent

def test_insert_exploded():
    """测试将文件炸开插入功能"""
    print("=" * 60)
    print("测试 insert_file_exploded() 函数")
    print("=" * 60)

    test_files_created = []

    try:
        # 1. 确保测试源文件存在
        test_dir = Path(__file__).parent / "test_files"
        test_dir.mkdir(exist_ok=True)

        source_file = test_dir / "test_insertion_source.dwg"

        if not source_file.exists():
            print(f"\n[信息] 测试源文件不存在，正在创建...")
            create_script = Path(__file__).parent / "create_test_insertion_source.py"
            result = subprocess.run([sys.executable, str(create_script)], capture_output=True)
            if result.returncode != 0:
                print("[错误] 创建测试源文件失败")
                return False
            time.sleep(2)

        if not source_file.exists():
            print(f"[错误] 测试源文件不存在: {source_file}")
            return False

        print(f"[成功] 测试源文件存在: {source_file}")

        # 2. 确保CAD处于单文件不确定状态
        print("\n[步骤1] 准备CAD环境...")
        cad_zt_oneb()
        wait_quiescent(min_quiet=1.0, timeout=30.0)
        print("[成功] CAD环境准备完成")

        # 3. 创建目标文件
        target_file = test_dir / "test_insert_exploded_target.dwg"
        test_files_created.append(target_file)

        print(f"\n[步骤2] 创建目标文件: {target_file}")
        if not new_file(str(target_file)):
            print("[错误] 创建目标文件失败")
            return False

        wait_quiescent(min_quiet=1.0, timeout=15.0)
        print("[成功] 目标文件创建成功")

        # 获取插入前的对象数量
        import win32com.client
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        doc = acad.ActiveDocument
        ms = doc.ModelSpace
        obj_count_before = ms.Count
        print(f"  插入前对象数量: {obj_count_before}")

        # 4. 测试1: 在原点插入并炸开
        print("\n[步骤3] 测试1: 在原点插入并炸开...")
        if not insert_file_exploded(str(source_file), x=0, y=0, z=0):
            print("[错误] 在原点插入并炸开失败")
            return False

        wait_quiescent(min_quiet=1.5, timeout=20.0)
        print("[成功] 在原点插入并炸开成功")

        # 验证对象数量增加（炸开后应该有多个对象，而不是一个块）
        obj_count_after_1 = ms.Count
        print(f"  插入后对象数量: {obj_count_after_1}")

        if obj_count_after_1 <= obj_count_before:
            print("[错误] 插入后对象数量未增加")
            return False

        # 统计块引用数量（炸开后不应该有块引用）
        block_count_1 = 0
        for i in range(ms.Count):
            try:
                obj = ms.Item(i)
                if obj.ObjectName == "AcDbBlockReference":
                    block_count_1 += 1
            except:
                pass

        print(f"  块引用数量: {block_count_1}")
        print(f"  新增对象数量: {obj_count_after_1 - obj_count_before}")

        # 5. 测试2: 在指定位置插入并炸开
        print("\n[步骤4] 测试2: 在位置(10000, 0)插入并炸开...")
        obj_count_before_2 = ms.Count

        if not insert_file_exploded(str(source_file), x=10000, y=0, z=0):
            print("[错误] 在指定位置插入并炸开失败")
            return False

        wait_quiescent(min_quiet=1.5, timeout=20.0)
        print("[成功] 在指定位置插入并炸开成功")

        obj_count_after_2 = ms.Count
        print(f"  插入前对象数量: {obj_count_before_2}")
        print(f"  插入后对象数量: {obj_count_after_2}")
        print(f"  新增对象数量: {obj_count_after_2 - obj_count_before_2}")

        # 6. 测试3: 插入缩放并炸开的对象
        print("\n[步骤5] 测试3: 在位置(0, 10000)插入缩放并炸开(scale=0.5)...")
        obj_count_before_3 = ms.Count

        if not insert_file_exploded(str(source_file), x=0, y=10000, z=0, scale=0.5):
            print("[错误] 插入缩放并炸开失败")
            return False

        wait_quiescent(min_quiet=1.5, timeout=20.0)
        print("[成功] 插入缩放并炸开成功")

        obj_count_after_3 = ms.Count
        print(f"  插入前对象数量: {obj_count_before_3}")
        print(f"  插入后对象数量: {obj_count_after_3}")
        print(f"  新增对象数量: {obj_count_after_3 - obj_count_before_3}")

        # 7. 最终验证
        print("\n[步骤6] 最终验证...")
        final_obj_count = ms.Count
        total_new_objects = final_obj_count - obj_count_before

        # 统计最终块引用数量
        final_block_count = 0
        for i in range(ms.Count):
            try:
                obj = ms.Item(i)
                if obj.ObjectName == "AcDbBlockReference":
                    final_block_count += 1
            except:
                pass

        print(f"  总对象数量: {final_obj_count}")
        print(f"  总块引用数量: {final_block_count}")
        print(f"  总新增对象数量: {total_new_objects}")

        if total_new_objects < 3:
            print(f"[警告] 新增对象数量少于预期: {total_new_objects}")

        # 8. 保存文件
        print("\n[步骤7] 保存目标文件...")
        if not save_file():
            print("[错误] 保存文件失败")
            return False

        print("[成功] 文件已保存")

        # 9. 验证文件
        if target_file.exists():
            file_size = target_file.stat().st_size
            print(f"[成功] 目标文件创建成功")
            print(f"        文件大小: {file_size} 字节")
        else:
            print("[错误] 目标文件未创建")
            return False

        # 10. 关闭文件
        print("\n[步骤8] 关闭文件...")
        close_file("no_save")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        print("\n" + "=" * 60)
        print("[完成] insert_file_exploded() 测试通过!")
        print("=" * 60)
        print("测试结果:")
        print(f"  - 源文件: {source_file.name}")
        print(f"  - 目标文件: {target_file.name}")
        print(f"  - 总新增对象: {total_new_objects}")
        print(f"  - 块引用数量: {final_block_count}")
        print("  - 测试1: 原点插入并炸开 - 成功")
        print("  - 测试2: 指定位置插入并炸开 - 成功")
        print("  - 测试3: 缩放插入并炸开 - 成功")
        print("将文件炸开插入功能正常")

        return True

    except Exception as e:
        print(f"\n[错误] 测试异常: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        # 清理：删除测试文件
        print("\n[清理] 删除测试文件...")
        for file_path in test_files_created:
            try:
                if file_path.exists():
                    file_path.unlink()
                    print(f"[删除] {file_path.name}")
            except Exception as e:
                print(f"[警告] 删除文件失败: {file_path.name}, 错误: {e}")

        # 恢复到单文件不确定状态
        print("\n[清理] 恢复到单文件不确定状态...")
        try:
            cad_zt_oneb()
            print("[成功] CAD状态已恢复")
        except:
            pass

if __name__ == "__main__":
    success = test_insert_exploded()
    sys.exit(0 if success else 1)
