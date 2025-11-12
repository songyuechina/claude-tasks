#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试插入操作：块模式和炸开模式
"""

import sys
from pathlib import Path

sys.path.append(str(Path(__file__).parent.parent / "scripts"))

from CAD_basic import start_applicationV9, close_all_cad_processes
from CAD_basic_operations import (
    new_dwg_enhanced,
    save_as_dwg_paradigm,
    close_current_dwg_paradigm,
    insert_dwg_as_block_paradigm,
    insert_and_explode_paradigm
)
from CAD_coordination import (
    wait_quiescent,
    send_cmd_with_sync,
    ensure_single_process
)

def create_source_file():
    """创建源文件A.dwg - 包含一个矩形"""
    print("\n[创建] 创建源文件A.dwg...")

    test_file = "D:/claude-tasks/tests/test_files/A.dwg"
    Path(test_file).parent.mkdir(parents=True, exist_ok=True)

    # 新建空白文件
    if not new_dwg_enhanced():
        print("[错误] 新建文件失败")
        return False

    wait_quiescent(min_quiet=1.0, timeout=15.0)

    # 绘制矩形
    print("[绘制] 绘制矩形...")
    send_cmd_with_sync("_RECTANG\n0,0\n100,50\n", wait_after=1.0)

    # 另存为
    if not save_as_dwg_paradigm(test_file):
        print("[错误] 保存失败")
        return False

    # 关闭
    close_current_dwg_paradigm("no_save")
    wait_quiescent(min_quiet=1.0, timeout=15.0)

    print(f"[成功] 源文件创建: {test_file}")
    return True

def test_insert_as_block():
    """测试：作为块插入"""
    print("\n" + "="*60)
    print("测试1: 作为块插入")
    print("="*60)

    source = "D:/claude-tasks/tests/test_files/A.dwg"
    target = "D:/claude-tasks/tests/test_files/B_with_block.dwg"

    # 新建目标文件
    if not new_dwg_enhanced():
        return False
    wait_quiescent(min_quiet=1.0, timeout=15.0)

    # 插入为块
    print("\n[插入] 作为块插入...")
    if not insert_dwg_as_block_paradigm(source, (200, 100, 0)):
        print("[失败] 块插入失败")
        return False

    print("[成功] 块插入成功")

    # 保存
    if not save_as_dwg_paradigm(target):
        return False

    close_current_dwg_paradigm("no_save")
    wait_quiescent(min_quiet=1.0, timeout=15.0)

    print(f"[完成] 文件保存: {target}")
    return True

def test_insert_and_explode():
    """测试：插入并炸开"""
    print("\n" + "="*60)
    print("测试2: 插入并炸开")
    print("="*60)

    source = "D:/claude-tasks/tests/test_files/A.dwg"
    target = "D:/claude-tasks/tests/test_files/C_exploded.dwg"

    # 新建目标文件
    if not new_dwg_enhanced():
        return False
    wait_quiescent(min_quiet=1.0, timeout=15.0)

    # 插入并炸开
    print("\n[插入] 插入并炸开...")
    if not insert_and_explode_paradigm(source, (200, 100, 0)):
        print("[失败] 炸开插入失败")
        return False

    print("[成功] 炸开插入成功")

    # 保存
    if not save_as_dwg_paradigm(target):
        return False

    close_current_dwg_paradigm("no_save")
    wait_quiescent(min_quiet=1.0, timeout=15.0)

    print(f"[完成] 文件保存: {target}")
    return True

def main():
    """主测试函数"""
    print("="*60)
    print("测试插入操作：块模式 vs 炸开模式")
    print("="*60)

    try:
        # 启动CAD
        print("\n[启动] 启动CAD...")
        proc = start_applicationV9(PTH=r"C:\Tangent\TArchT20V9")
        if not proc:
            print("[错误] CAD启动失败")
            return False

        ensure_single_process()
        wait_quiescent(min_quiet=2.0, timeout=30.0)

        # 创建源文件
        if not create_source_file():
            return False

        # 测试块模式
        if not test_insert_as_block():
            return False

        # 测试炸开模式
        if not test_insert_and_explode():
            return False

        print("\n" + "="*60)
        print("所有测试完成!")
        print("="*60)
        print("\n测试结果:")
        print("- A.dwg: 源文件(矩形)")
        print("- B_with_block.dwg: 块模式插入")
        print("- C_exploded.dwg: 炸开模式插入")
        print("\n区别:")
        print("- 块模式: 插入的内容是一个整体块，可以选择、移动、删除")
        print("- 炸开模式: 插入的内容被分解为独立的图形对象")

        return True

    except Exception as e:
        print(f"\n[错误] 测试异常: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        print("\n[清理] 关闭所有CAD进程...")
        try:
            close_all_cad_processes()
            print("[成功] 环境已清理")
        except:
            pass

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
