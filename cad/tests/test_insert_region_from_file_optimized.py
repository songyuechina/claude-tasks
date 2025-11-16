#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试优化后的 insert_region_from_file 函数
包含60秒心跳时间戳功能
"""

import sys
import time
import threading
from pathlib import Path
from datetime import datetime

# 添加路径
SCRIPT_DIR = Path(__file__).parent.parent / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from CAD_file_operations import (
    new_file, save_file_as, close_file, insert_region_from_file,
    open_file, save_file, cad_zt_zero
)
from CAD_coordination import send_cmd_with_sync, wait_quiescent

# 心跳标志
heartbeat_running = False
heartbeat_file = "D:/claude-tasks/cad/任务进度汇报.txt"

def write_heartbeat():
    """每60秒写入心跳时间戳"""
    global heartbeat_running
    while heartbeat_running:
        try:
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            with open(heartbeat_file, "a", encoding="utf-8") as f:
                f.write(f"{timestamp}\n")
            print(f"\n💓 [心跳] {timestamp}")
        except Exception as e:
            print(f"[警告] 心跳写入失败: {e}")
        time.sleep(60)  # 每60秒写入一次

def start_heartbeat():
    """启动心跳线程"""
    global heartbeat_running
    heartbeat_running = True

    # 清空并初始化心跳文件
    try:
        with open(heartbeat_file, "w", encoding="utf-8") as f:
            f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - 任务开始\n")
        print(f"✅ 心跳监控已启动")
    except Exception as e:
        print(f"[警告] 心跳文件初始化失败: {e}")

    # 启动心跳线程
    thread = threading.Thread(target=write_heartbeat, daemon=True)
    thread.start()
    return thread

def stop_heartbeat():
    """停止心跳线程"""
    global heartbeat_running
    heartbeat_running = False
    try:
        with open(heartbeat_file, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - 任务完成\n")
        print(f"✅ 心跳监控已停止")
    except Exception as e:
        print(f"[警告] 心跳文件写入失败: {e}")

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

    # 在(200,0)到(300,100)区域绘制圆形
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
    print("[测试区域插入功能]")
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
    result_file = test_dir / "A_with_region.dwg"
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
    print(f"  - 总执行时间: {elapsed:.2f} 秒")

    return True

def main():
    """主测试函数"""
    print("="*60)
    print("insert_region_from_file 优化测试")
    print("="*60)
    print(f"测试时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    heartbeat_thread = None

    try:
        # 启动心跳监控
        heartbeat_thread = start_heartbeat()

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
        print("✅ 所有测试通过!")
        print("="*60)

        return True

    except Exception as e:
        print(f"\n[错误] 测试异常: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        # 停止心跳监控
        stop_heartbeat()

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
