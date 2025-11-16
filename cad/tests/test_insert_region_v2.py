#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试重写后的 insert_region_from_file 函数 (v2)
使用临时副本文件方案
"""

import sys
import time
from pathlib import Path

# 添加路径
SCRIPT_DIR = Path(__file__).parent.parent / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from CAD_file_operations import (
    new_file, save_file_as, close_file, insert_region_from_file,
    open_file, save_file, cad_zt_zero
)
from CAD_coordination import send_cmd_with_sync, wait_quiescent

def create_test_files():
    """创建测试文件A.dwg和B.dwg"""
    print("\n" + "="*60)
    print("[创建测试文件]")
    print("="*60)

    test_dir = Path("D:/claude-tasks/cad/tests/test_files")
    test_dir.mkdir(parents=True, exist_ok=True)

    # 创建A.dwg - 包含一个圆形
    print("\n[A.dwg] 创建文件A...")
    new_file()
    wait_quiescent(min_quiet=1.0, timeout=15.0)
    send_cmd_with_sync("_CIRCLE\n0,0\n50\n", wait_after=1.0, timeout=30.0)
    save_file_as(str(test_dir / "A.dwg"))
    close_file("no_save")
    print("[成功] A.dwg 创建完成")

    # 创建B.dwg - 包含多个图形
    print("\n[B.dwg] 创建文件B...")
    new_file()
    wait_quiescent(min_quiet=1.0, timeout=15.0)

    # 在(0,0)到(100,100)区域绘制矩形
    send_cmd_with_sync("_RECTANG\n10,10\n90,90\n", wait_after=1.0, timeout=30.0)
    print("  - 矩形已绘制 (10,10)-(90,90)")

    # 在(200,0)到(300,100)区域绘制圆形（不在选择区域内）
    send_cmd_with_sync("_CIRCLE\n250,50\n30\n", wait_after=1.0, timeout=30.0)
    print("  - 圆形已绘制 中心(250,50) 半径30")

    # 在(0,0)到(100,100)区域绘制三角形
    send_cmd_with_sync("_PLINE\n20,20\n50,80\n80,20\nC\n", wait_after=1.0, timeout=30.0)
    print("  - 三角形已绘制")

    save_file_as(str(test_dir / "B.dwg"))
    close_file("no_save")
    print("[成功] B.dwg 创建完成")

    print("\n[完成] 测试文件创建完成")
    print(f"  - A.dwg: {test_dir / 'A.dwg'}")
    print(f"  - B.dwg: {test_dir / 'B.dwg'}")

def test_region_insert():
    """测试区域插入功能"""
    print("\n" + "="*60)
    print("[测试区域插入功能 - v2方案]")
    print("="*60)

    test_dir = Path("D:/claude-tasks/cad/tests/test_files")

    # 打开A.dwg
    print("\n[步骤1] 打开A.dwg...")
    if not open_file(str(test_dir / "A.dwg")):
        print("[错误] 打开A.dwg失败")
        return False
    wait_quiescent(min_quiet=1.0, timeout=15.0)
    print("[成功] A.dwg 已打开")

    # 从B.dwg的(0,0)到(100,100)区域插入到A.dwg的(150,150)位置
    print("\n[步骤2] 从B.dwg插入区域...")
    print("  源文件: B.dwg")
    print("  源区域: (0,0) -> (100,100)")
    print("  目标位置: (150,150)")
    print("  预期结果: 只插入矩形和三角形，不插入圆形(250,50)")

    start_time = time.time()
    success = insert_region_from_file(
        source_file=str(test_dir / "B.dwg"),
        x1=0, y1=0,
        x2=100, y2=100,
        x3=150, y3=150,
        explode=True
    )
    elapsed = time.time() - start_time

    if success:
        print(f"\n[成功] 区域插入成功")
        print(f"  执行时间: {elapsed:.2f} 秒")
    else:
        print(f"\n[失败] 区域插入失败")
        print(f"  执行时间: {elapsed:.2f} 秒")
        return False

    # 保存结果
    print("\n[步骤3] 保存结果...")
    save_file()
    result_file = test_dir / "A_with_region_v2.dwg"
    save_file_as(str(result_file))
    print(f"[成功] 结果已保存: {result_file}")

    close_file("no_save")
    print("[成功] 文件已关闭")

    print("\n" + "="*60)
    print("[测试完成]")
    print("="*60)
    print(f"结果文件: {result_file}")
    print("  - 原有圆形在 (0,0)")
    print("  - 插入的矩形和三角形在 (150,150) 区域")
    print("  - 不应该有圆形(250,50)")
    print(f"  - 总执行时间: {elapsed:.2f} 秒")

    return True

def main():
    """主测试函数"""
    print("="*60)
    print("insert_region_from_file v2 测试")
    print("="*60)
    print(f"新方案: 通过临时副本文件实现")

    try:
        # 测试开始前清理
        print("\n[清理] 测试开始前清理CAD进程...")
        cad_zt_zero()
        print("[成功] CAD进程已清理")

        # 等待CAD稳定
        time.sleep(3)
        wait_quiescent(min_quiet=2.0, timeout=30.0)

        # 创建测试文件
        create_test_files()

        # 测试区域插入
        if not test_region_insert():
            print("\n[失败] 测试未通过")
            return False

        print("\n" + "="*60)
        print("[通过] 所有测试通过!")
        print("="*60)

        return True

    except Exception as e:
        print(f"\n[错误] 测试异常: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        # 测试结束后清理
        print("\n[清理] 测试结束后清理CAD进程...")
        try:
            cad_zt_zero()
            print("[成功] CAD进程已清理")
        except Exception as e:
            print(f"[警告] 清理失败: {e}")

if __name__ == "__main__":
    success = main()

    # 任务完成前最终清理
    print("\n[清理] 任务完成前最终清理...")
    try:
        cad_zt_zero()
        print("[成功] 最终清理完成")
    except Exception as e:
        print(f"[警告] 最终清理失败: {e}")

    sys.exit(0 if success else 1)
