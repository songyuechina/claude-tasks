#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试文件拷贝操作
"""

import sys
import time
from pathlib import Path

# 添加scripts目录到路径
sys.path.append(str(Path(__file__).parent.parent / "scripts"))

from CAD_basic_operations import (
    copy_dwg_to_dwg_paradigm,
    copy_region_to_dwg_paradigm,
    close_all_dwg_paradigm,
    new_dwg_enhanced,
    save_current_dwg_paradigm,
    close_current_dwg_paradigm
)
from CAD_coordination import (
    start_cad_with_dialog_killer,
    ensure_single_process,
    wait_quiescent,
    send_cmd_with_sync
)

def create_test_file_a():
    """创建测试文件A.dwg - 绘制矩形"""
    print("\n[创建] 创建测试文件A.dwg...")

    test_file = "D:/claude-tasks/tests/test_files/A.dwg"
    Path(test_file).parent.mkdir(parents=True, exist_ok=True)

    # 使用范式创建新文件
    if not new_dwg_enhanced(test_file):
        print("[错误] 创建A.dwg失败")
        return False

    # 绘制矩形 (0,0) 到 (100,50)
    print("[绘制] 绘制矩形...")
    send_cmd_with_sync("_RECTANG\n0,0\n100,50\n", wait_after=1.0)

    # 保存并关闭
    save_current_dwg_paradigm()
    close_current_dwg_paradigm("no_save")

    print("[成功] A.dwg创建完成")
    return True

def create_test_file_b():
    """创建测试文件B.dwg - 绘制三角形和圆形"""
    print("\n[创建] 创建测试文件B.dwg...")

    test_file = "D:/claude-tasks/tests/test_files/B.dwg"

    # 使用范式创建新文件
    if not new_dwg_enhanced(test_file):
        print("[错误] 创建B.dwg失败")
        return False

    # 绘制三角形 (使用PLINE)
    print("[绘制] 绘制三角形...")
    send_cmd_with_sync("_PLINE\n200,0\n250,50\n200,100\nC\n", wait_after=1.0)

    # 绘制圆形
    print("[绘制] 绘制圆形...")
    send_cmd_with_sync("_CIRCLE\n300,50\n30\n", wait_after=1.0)

    # 保存并关闭
    save_current_dwg_paradigm()
    close_current_dwg_paradigm("no_save")

    print("[成功] B.dwg创建完成")
    return True

def test_copy_full_file():
    """测试任务1: 完整文件拷贝"""
    print("\n" + "="*60)
    print("测试任务1: 完整文件拷贝")
    print("="*60)

    source = "D:/claude-tasks/tests/test_files/A.dwg"
    target = "D:/claude-tasks/tests/test_files/B_with_A_block.dwg"

    # 测试块模式
    print("\n[测试] 块模式拷贝...")
    if copy_dwg_to_dwg_paradigm(source, target, explode=False):
        print("[成功] 块模式拷贝成功")
    else:
        print("[失败] 块模式拷贝失败")
        return False

    # 测试炸开模式
    target_explode = "D:/claude-tasks/tests/test_files/B_with_A_exploded.dwg"
    print("\n[测试] 炸开模式拷贝...")
    if copy_dwg_to_dwg_paradigm(source, target_explode, explode=True):
        print("[成功] 炸开模式拷贝成功")
    else:
        print("[失败] 炸开模式拷贝失败")
        return False

    return True

def test_copy_region():
    """测试任务2: 区域拷贝"""
    print("\n" + "="*60)
    print("测试任务2: 区域拷贝")
    print("="*60)

    source = "D:/claude-tasks/tests/test_files/A.dwg"
    target = "D:/claude-tasks/tests/test_files/B_with_region.dwg"

    # 拷贝矩形的左半部分 (0,0)到(50,50) 到目标位置(150,150)
    print("\n[测试] 区域拷贝...")
    if copy_region_to_dwg_paradigm(source, target, 0, 0, 50, 50, 150, 150):
        print("[成功] 区域拷贝成功")
        return True
    else:
        print("[失败] 区域拷贝失败")
        return False

def main():
    """主测试函数"""
    print("="*60)
    print("CAD文件拷贝操作测试")
    print("="*60)

    # 启动CAD
    print("\n[启动] 启动CAD...")
    if not start_cad_with_dialog_killer():
        print("[错误] CAD启动失败")
        return False

    ensure_single_process()
    wait_quiescent(min_quiet=1.0, timeout=30.0)

    # 创建测试文件
    if not create_test_file_a():
        return False

    if not create_test_file_b():
        return False

    # 测试任务1
    if not test_copy_full_file():
        return False

    # 测试任务2
    if not test_copy_region():
        return False

    # 清理
    print("\n[清理] 关闭所有文件...")
    close_all_dwg_paradigm()

    print("\n" + "="*60)
    print("所有测试完成!")
    print("="*60)
    print("\n测试文件位置: D:/claude-tasks/tests/test_files/")
    print("- A.dwg: 源文件(矩形)")
    print("- B.dwg: 源文件(三角形和圆)")
    print("- B_with_A_block.dwg: 任务1结果(块模式)")
    print("- B_with_A_exploded.dwg: 任务1结果(炸开模式)")
    print("- B_with_region.dwg: 任务2结果(区域拷贝)")

    return True

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
