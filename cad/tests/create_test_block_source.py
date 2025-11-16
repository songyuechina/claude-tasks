#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
创建测试块源文件
"""

import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent.parent / "scripts"
SYSTEM_DIR = Path(__file__).parent.parent / "system"
sys.path.insert(0, str(SCRIPT_DIR))
sys.path.insert(0, str(SYSTEM_DIR))

from CAD_file_operations import cad_zt_zero, new_file, save_file, close_file
from CAD_basic import li_new, draw_line, draw_circle, draw_point
from CAD_coordination import wait_quiescent

def create_test_block_source():
    """创建测试块源文件"""
    print("创建测试块源文件...")

    try:
        # 准备环境
        cad_zt_zero()
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 创建文件
        block_file = "D:/claude-tasks/cad/tests/test_files/test_block_source.dwg"
        print(f"创建文件: {block_file}")
        new_file(block_file)
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 初始化
        li_new()

        # 绘制简单图形
        draw_point((0, 0, 0))
        draw_line((0, 0, 0), (1000, 0, 0))
        draw_line((1000, 0, 0), (1000, 1000, 0))
        draw_line((1000, 1000, 0), (0, 1000, 0))
        draw_line((0, 1000, 0), (0, 0, 0))
        draw_circle((500, 500, 0), 300)
        wait_quiescent(min_quiet=0.5, timeout=10.0)

        print("绘制完成，保存文件...")

        # 保存文件
        save_file()
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 关闭文件
        close_file("no_save")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        if Path(block_file).exists():
            print(f"成功创建测试块文件: {block_file}")
            print(f"文件大小: {Path(block_file).stat().st_size} 字节")
            return True
        else:
            print("文件创建失败")
            return False

    except Exception as e:
        print(f"错误: {e}")
        import traceback
        traceback.print_exc()
        return False
    finally:
        cad_zt_zero()

if __name__ == "__main__":
    success = create_test_block_source()
    sys.exit(0 if success else 1)
