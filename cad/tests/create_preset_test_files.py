#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
创建CAD_file_operations.py所有函数的预备测试DWG文件

此脚本为每个函数创建对应的预备测试DWG文件，便于快速验证函数是否可用。
"""

import sys
import time
from pathlib import Path

# 添加路径
SCRIPT_DIR = Path(__file__).parent.parent / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from CAD_file_operations import (
    new_file, save_file_as, close_file, cad_zt_zero
)
from CAD_coordination import send_cmd_with_sync, wait_quiescent

def create_preset_directory():
    """创建预备测试文件目录"""
    preset_dir = Path("D:/claude-tasks/cad/tests/test_files/preset")
    preset_dir.mkdir(parents=True, exist_ok=True)
    print(f"[创建] 预备测试文件目录: {preset_dir}")
    return preset_dir

def create_simple_file(filepath, shapes=None):
    """创建简单图形文件"""
    new_file()
    wait_quiescent(min_quiet=1.0, timeout=15.0)

    if shapes:
        for shape in shapes:
            send_cmd_with_sync(shape, wait_after=0.5, timeout=30.0)

    save_file_as(str(filepath))
    close_file("no_save")
    print(f"[成功] 创建: {filepath.name}")

def create_wall_file(filepath):
    """创建带天正墙的文件"""
    new_file()
    wait_quiescent(min_quiet=1.0, timeout=15.0)

    # 绘制一段天正墙
    cmd = "tgwall\n0,0\n3000,0\n\n\n"
    send_cmd_with_sync(cmd, wait_after=1.5, timeout=60.0)

    save_file_as(str(filepath))
    close_file("no_save")
    print(f"[成功] 创建: {filepath.name}")

def main():
    """主函数"""
    print("="*60)
    print("创建CAD_file_operations.py预备测试文件")
    print("="*60)

    try:
        # 清理CAD
        print("\n[1] 清理CAD进程...")
        cad_zt_zero()
        time.sleep(3)
        wait_quiescent(min_quiet=2.0, timeout=30.0)

        # 创建目录
        print("\n[2] 创建预备文件目录...")
        preset_dir = create_preset_directory()

        # 创建预备测试文件
        print("\n[3] 创建预备测试文件...")

        # 3.1 简单文件（用于open_file, close_file等）
        print("\n--- 简单图形文件 ---")
        create_simple_file(
            preset_dir / "preset_open.dwg",
            ["_CIRCLE\n0,0\n50\n"]
        )

        create_simple_file(
            preset_dir / "preset_copy_source.dwg",
            ["_RECTANG\n0,0\n100,100\n"]
        )

        create_simple_file(
            preset_dir / "preset_saveas_source.dwg",
            ["_CIRCLE\n50,50\n30\n"]
        )

        create_simple_file(
            preset_dir / "preset_close.dwg",
            ["_PLINE\n0,0\n100,0\n100,100\n0,100\nC\n"]
        )

        # 3.2 多文档测试文件
        print("\n--- 多文档测试文件 ---")
        create_simple_file(
            preset_dir / "preset_close_all_1.dwg",
            ["_CIRCLE\n0,0\n25\n"]
        )

        create_simple_file(
            preset_dir / "preset_close_all_2.dwg",
            ["_CIRCLE\n100,0\n25\n"]
        )

        create_simple_file(
            preset_dir / "preset_activate_1.dwg",
            ["_RECTANG\n0,0\n50,50\n"]
        )

        create_simple_file(
            preset_dir / "preset_activate_2.dwg",
            ["_RECTANG\n100,0\n150,50\n"]
        )

        # 3.3 块插入测试文件
        print("\n--- 块插入测试文件 ---")
        create_simple_file(
            preset_dir / "preset_block_source.dwg",
            ["_CIRCLE\n0,0\n30\n", "_RECTANG\n-20,-20\n20,20\n"]
        )

        create_simple_file(
            preset_dir / "preset_block_target.dwg",
            ["_CIRCLE\n200,200\n10\n"]
        )

        create_simple_file(
            preset_dir / "preset_explode_source.dwg",
            ["_RECTANG\n0,0\n50,50\n", "_CIRCLE\n25,25\n20\n"]
        )

        create_simple_file(
            preset_dir / "preset_explode_target.dwg",
            ["_CIRCLE\n300,300\n15\n"]
        )

        # 3.4 内容复制测试文件
        print("\n--- 内容复制测试文件 ---")
        create_simple_file(
            preset_dir / "preset_content_source.dwg",
            [
                "_CIRCLE\n0,0\n40\n",
                "_RECTANG\n10,10\n60,60\n",
                "_PLINE\n20,20\n50,40\n40,20\nC\n"
            ]
        )

        create_simple_file(
            preset_dir / "preset_content_target.dwg",
            ["_CIRCLE\n500,500\n20\n"]
        )

        # 3.5 区域插入测试文件
        print("\n--- 区域插入测试文件 ---")
        create_simple_file(
            preset_dir / "preset_region_source.dwg",
            [
                "_RECTANG\n10,10\n90,90\n",  # 在(0,0)-(100,100)区域
                "_CIRCLE\n250,50\n30\n",     # 在区域外
                "_PLINE\n20,20\n50,80\n80,20\nC\n"  # 在区域内
            ]
        )

        create_simple_file(
            preset_dir / "preset_region_A.dwg",
            ["_CIRCLE\n0,0\n50\n"]
        )

        create_simple_file(
            preset_dir / "preset_region_B.dwg",
            [
                "_RECTANG\n10,10\n90,90\n",
                "_CIRCLE\n250,50\n30\n",
                "_PLINE\n20,20\n50,80\n80,20\nC\n"
            ]
        )

        # 3.6 标注测试文件
        print("\n--- 标注测试文件 ---")
        create_simple_file(
            preset_dir / "preset_dim_target.dwg",
            ["_LINE\n0,0\n1000,0\n", "_LINE\n0,0\n0,1000\n"]
        )

        # 3.7 天正墙体测试文件
        print("\n--- 天正墙体测试文件 ---")
        create_simple_file(
            preset_dir / "preset_wall_target.dwg",
            []  # 空文件，用于绘制墙
        )

        create_wall_file(preset_dir / "preset_door_wall.dwg")
        create_wall_file(preset_dir / "preset_window_wall.dwg")

        print("\n" + "="*60)
        print("[完成] 所有预备测试文件创建成功!")
        print("="*60)
        print(f"文件位置: {preset_dir}")
        print(f"共创建: 19 个预备测试文件")

        return True

    except Exception as e:
        print(f"\n[错误] 创建预备文件失败: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        print("\n[清理] 清理CAD进程...")
        try:
            cad_zt_zero()
            print("[成功] CAD进程已清理")
        except Exception as e:
            print(f"[警告] 清理失败: {e}")

if __name__ == "__main__":
    success = main()

    print("\n[清理] 最终清理...")
    try:
        cad_zt_zero()
        print("[成功] 最终清理完成")
    except Exception as e:
        print(f"[警告] 最终清理失败: {e}")

    sys.exit(0 if success else 1)
