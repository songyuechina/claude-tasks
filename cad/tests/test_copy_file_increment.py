#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试 copy_file_with_increment() 函数
测试文件复制并自动递增命名功能
"""

import sys
import time
from pathlib import Path

# 添加脚本路径
SCRIPT_DIR = Path(__file__).parent.parent / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from CAD_file_operations import (
    copy_file_with_increment,
    new_file,
    cad_zt_oneb
)

def test_copy_file_increment():
    """测试文件复制递增命名功能"""
    print("=" * 60)
    print("测试 copy_file_with_increment() 函数")
    print("=" * 60)

    created_files = []

    try:
        # 1. 确保CAD状态为单文件不确定状态
        print("\n[步骤1] 确保CAD处于单文件不确定状态...")
        cad_zt_oneb()
        print("[成功] CAD状态准备完成")

        # 2. 创建源文件用于测试
        test_dir = Path(__file__).parent / "test_files"
        test_dir.mkdir(exist_ok=True)

        source_file = test_dir / "test_increment_source.dwg"

        print(f"\n[步骤2] 创建源测试文件: {source_file}")
        if not new_file(str(source_file)):
            print("[错误] 创建源文件失败")
            return False

        created_files.append(source_file)
        time.sleep(2)
        print("[成功] 源文件创建成功")

        # 3. 第一次复制 - 应该创建 test_increment_source-1.dwg
        print("\n[步骤3] 第一次复制文件...")
        result1 = copy_file_with_increment(str(source_file))
        if result1 is None:
            print("[错误] 第一次复制失败")
            return False

        result1_path = Path(result1)
        if not result1_path.exists():
            print(f"[错误] 复制的文件不存在: {result1}")
            return False

        created_files.append(result1_path)
        print(f"[成功] 第一次复制成功: {result1_path.name}")

        # 验证文件名
        expected_name1 = "test_increment_source-1.dwg"
        if result1_path.name != expected_name1:
            print(f"[错误] 文件名不符合预期: {result1_path.name}, 期望: {expected_name1}")
            return False
        print(f"[验证] 文件名正确: {result1_path.name}")

        # 4. 第二次复制 - 应该创建 test_increment_source-2.dwg
        print("\n[步骤4] 第二次复制文件...")
        result2 = copy_file_with_increment(str(source_file))
        if result2 is None:
            print("[错误] 第二次复制失败")
            return False

        result2_path = Path(result2)
        if not result2_path.exists():
            print(f"[错误] 复制的文件不存在: {result2}")
            return False

        created_files.append(result2_path)
        print(f"[成功] 第二次复制成功: {result2_path.name}")

        # 验证文件名
        expected_name2 = "test_increment_source-2.dwg"
        if result2_path.name != expected_name2:
            print(f"[错误] 文件名不符合预期: {result2_path.name}, 期望: {expected_name2}")
            return False
        print(f"[验证] 文件名正确: {result2_path.name}")

        # 5. 第三次复制 - 应该创建 test_increment_source-3.dwg
        print("\n[步骤5] 第三次复制文件...")
        result3 = copy_file_with_increment(str(source_file))
        if result3 is None:
            print("[错误] 第三次复制失败")
            return False

        result3_path = Path(result3)
        if not result3_path.exists():
            print(f"[错误] 复制的文件不存在: {result3}")
            return False

        created_files.append(result3_path)
        print(f"[成功] 第三次复制成功: {result3_path.name}")

        # 验证文件名
        expected_name3 = "test_increment_source-3.dwg"
        if result3_path.name != expected_name3:
            print(f"[错误] 文件名不符合预期: {result3_path.name}, 期望: {expected_name3}")
            return False
        print(f"[验证] 文件名正确: {result3_path.name}")

        # 6. 验证所有文件都存在且大小合理
        print("\n[步骤6] 验证所有复制的文件...")
        for file_path in created_files:
            if not file_path.exists():
                print(f"[错误] 文件不存在: {file_path}")
                return False

            file_size = file_path.stat().st_size
            if file_size < 1000:  # DWG文件至少应该大于1KB
                print(f"[错误] 文件大小异常: {file_path}, 大小: {file_size} 字节")
                return False

            print(f"[验证] {file_path.name}: {file_size} 字节")

        print("\n" + "=" * 60)
        print("[完成] copy_file_with_increment() 测试通过!")
        print("=" * 60)
        print(f"源文件: {source_file.name}")
        print(f"复制文件1: {result1_path.name}")
        print(f"复制文件2: {result2_path.name}")
        print(f"复制文件3: {result3_path.name}")
        print("递增命名功能正常")

        return True

    except Exception as e:
        print(f"\n[错误] 测试异常: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        # 清理：删除测试文件
        print("\n[清理] 删除测试文件...")
        for file_path in created_files:
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
    success = test_copy_file_increment()
    sys.exit(0 if success else 1)
