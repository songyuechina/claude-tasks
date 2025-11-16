#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试 insert_file_as_block() 函数
测试将文件作为块插入功能
"""

import sys
import time
import subprocess
from pathlib import Path

# 添加脚本路径
SCRIPT_DIR = Path(__file__).parent.parent / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from CAD_file_operations import (
    insert_file_as_block,
    new_file,
    save_file,
    close_file,
    cad_zt_oneb
)
from CAD_coordination import wait_quiescent

def test_insert_as_block():
    """测试将文件作为块插入功能"""
    print("=" * 60)
    print("测试 insert_file_as_block() 函数")
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
        target_file = test_dir / "test_insert_as_block_target.dwg"
        test_files_created.append(target_file)

        print(f"\n[步骤2] 创建目标文件: {target_file}")
        if not new_file(str(target_file)):
            print("[错误] 创建目标文件失败")
            return False

        wait_quiescent(min_quiet=1.0, timeout=15.0)
        print("[成功] 目标文件创建成功")

        # 4. 测试1: 在原点插入块（默认参数）
        print("\n[步骤3] 测试1: 在原点插入块...")
        if not insert_file_as_block(str(source_file), x=0, y=0, z=0):
            print("[错误] 在原点插入块失败")
            return False

        wait_quiescent(min_quiet=1.0, timeout=15.0)
        print("[成功] 在原点插入块成功")

        # 5. 测试2: 在指定位置插入块
        print("\n[步骤4] 测试2: 在位置(10000, 0)插入块...")
        if not insert_file_as_block(str(source_file), x=10000, y=0, z=0):
            print("[错误] 在指定位置插入块失败")
            return False

        wait_quiescent(min_quiet=1.0, timeout=15.0)
        print("[成功] 在指定位置插入块成功")

        # 6. 测试3: 插入缩放的块
        print("\n[步骤5] 测试3: 在位置(0, 10000)插入缩放块(scale=0.5)...")
        if not insert_file_as_block(str(source_file), x=0, y=10000, z=0, scale=0.5):
            print("[错误] 插入缩放块失败")
            return False

        wait_quiescent(min_quiet=1.0, timeout=15.0)
        print("[成功] 插入缩放块成功")

        # 7. 测试4: 插入旋转的块
        print("\n[步骤6] 测试4: 在位置(10000, 10000)插入旋转块(rotation=45°)...")
        if not insert_file_as_block(str(source_file), x=10000, y=10000, z=0, rotation=45.0):
            print("[错误] 插入旋转块失败")
            return False

        wait_quiescent(min_quiet=1.0, timeout=15.0)
        print("[成功] 插入旋转块成功")

        # 8. 验证插入的块
        print("\n[步骤7] 验证插入的块...")
        import win32com.client
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        doc = acad.ActiveDocument
        ms = doc.ModelSpace

        # 统计块引用数量
        block_count = 0
        for i in range(ms.Count):
            try:
                obj = ms.Item(i)
                if obj.ObjectName == "AcDbBlockReference":
                    block_count += 1
            except:
                pass

        print(f"  模型空间中的块引用数量: {block_count}")

        if block_count < 4:
            print(f"[警告] 块引用数量少于预期: {block_count} < 4")
        else:
            print(f"[成功] 找到 {block_count} 个块引用")

        # 9. 保存文件
        print("\n[步骤8] 保存目标文件...")
        if not save_file():
            print("[错误] 保存文件失败")
            return False

        print("[成功] 文件已保存")

        # 10. 验证文件
        if target_file.exists():
            file_size = target_file.stat().st_size
            print(f"[成功] 目标文件创建成功")
            print(f"        文件大小: {file_size} 字节")
        else:
            print("[错误] 目标文件未创建")
            return False

        # 11. 关闭文件
        print("\n[步骤9] 关闭文件...")
        close_file("no_save")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        print("\n" + "=" * 60)
        print("[完成] insert_file_as_block() 测试通过!")
        print("=" * 60)
        print("测试结果:")
        print(f"  - 源文件: {source_file.name}")
        print(f"  - 目标文件: {target_file.name}")
        print(f"  - 插入块数量: {block_count}")
        print("  - 测试1: 原点插入 - 成功")
        print("  - 测试2: 指定位置插入 - 成功")
        print("  - 测试3: 缩放插入 - 成功")
        print("  - 测试4: 旋转插入 - 成功")
        print("将文件作为块插入功能正常")

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
    success = test_insert_as_block()
    sys.exit(0 if success else 1)
