#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试区域插入功能
"""

import sys
from pathlib import Path

sys.path.append(str(Path(__file__).parent.parent / "scripts"))

from CAD_file_operations import new_file, save_file_as, close_file, insert_region_from_file
from CAD_basic import start_applicationV9, close_all_cad_processes
from CAD_coordination import send_cmd_with_sync, wait_quiescent, ensure_single_process

def create_test_files():
    """创建测试文件A.dwg和B.dwg"""
    print("\n[创建] 创建测试文件...")

    # 创建A.dwg - 包含一个圆形
    print("\n[A.dwg] 创建文件A...")
    new_file()
    wait_quiescent(min_quiet=1.0, timeout=15.0)
    send_cmd_with_sync("_CIRCLE\n0,0\n50\n", wait_after=1.0)
    save_file_as("D:/claude-tasks/tests/test_files/A.dwg")
    close_file("no_save")

    # 创建B.dwg - 包含多个图形
    print("\n[B.dwg] 创建文件B...")
    new_file()
    wait_quiescent(min_quiet=1.0, timeout=15.0)

    # 在(0,0)到(100,100)区域绘制矩形
    send_cmd_with_sync("_RECTANG\n10,10\n90,90\n", wait_after=1.0)

    # 在(200,0)到(300,100)区域绘制圆形
    send_cmd_with_sync("_CIRCLE\n250,50\n30\n", wait_after=1.0)

    # 在(0,0)到(100,100)区域绘制三角形
    send_cmd_with_sync("_PLINE\n20,20\n50,80\n80,20\nC\n", wait_after=1.0)

    save_file_as("D:/claude-tasks/tests/test_files/B.dwg")
    close_file("no_save")

    print("[成功] 测试文件创建完成")

def test_region_insert():
    """测试区域插入功能"""
    print("\n" + "="*60)
    print("测试区域插入功能")
    print("="*60)

    # 打开A.dwg
    print("\n[步骤1] 打开A.dwg...")
    from CAD_file_operations import open_file
    open_file("D:/claude-tasks/tests/test_files/A.dwg")
    wait_quiescent(min_quiet=1.0, timeout=15.0)

    # 从B.dwg的(0,0)到(100,100)区域插入到A.dwg的(150,150)位置
    print("\n[步骤2] 从B.dwg插入区域...")
    print("  源区域: (0,0) 到 (100,100)")
    print("  目标位置: (150,150)")

    success = insert_region_from_file(
        source_file="D:/claude-tasks/tests/test_files/B.dwg",
        x1=0, y1=0,
        x2=100, y2=100,
        x3=150, y3=150,
        explode=True
    )

    if success:
        print("[成功] 区域插入成功")
    else:
        print("[失败] 区域插入失败")
        return False

    # 保存结果
    print("\n[步骤3] 保存结果...")
    from CAD_file_operations import save_file
    save_file()
    close_file("no_save")

    print("\n[完成] 测试完成")
    print("结果文件: D:/claude-tasks/tests/test_files/A.dwg")
    print("  - 原有圆形在(0,0)")
    print("  - 插入的矩形和三角形在(150,150)区域")

    return True

def main():
    """主测试函数"""
    print("="*60)
    print("区域插入功能测试")
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

        # 创建测试文件
        create_test_files()

        # 测试区域插入
        if not test_region_insert():
            return False

        print("\n" + "="*60)
        print("所有测试完成!")
        print("="*60)

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
