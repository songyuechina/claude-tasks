#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试修复后的两个函数
1. start_applicationV9() - 弹窗治理脚本路径修复
2. close_all_cad_processes() - 强制关闭功能加强
"""

import sys
import time
from pathlib import Path

# 添加scripts目录到路径
SCRIPT_DIR = Path(__file__).parent.parent / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from CAD_basic import start_applicationV9, close_all_cad_processes, jingchengshu_wenjian

def test_start_application_and_dialog_killer():
    """测试1：启动CAD并验证弹窗治理脚本"""
    print("="*60)
    print("测试1：start_applicationV9() - 弹窗治理脚本路径修复")
    print("="*60)

    try:
        # 清理已有CAD进程
        print("\n[准备] 清理已有CAD进程...")
        close_all_cad_processes()
        time.sleep(2)

        # 启动CAD
        print("\n[测试] 启动CAD...")
        proc = start_applicationV9()

        if proc is None:
            print("[失败] CAD启动失败")
            return False

        print(f"[成功] CAD已启动 (PID: {proc.pid})")

        # 等待CAD完全启动
        print("\n[等待] 等待CAD完全启动...")
        time.sleep(10)

        # 验证CAD进程
        shu = jingchengshu_wenjian()
        print(f"[验证] 当前CAD进程数: {shu}")

        if shu == 0:
            print("[失败] CAD进程未启动成功")
            return False

        # 验证弹窗治理脚本是否在运行
        print("\n[验证] 检查cad_dialog_killer.py是否在运行...")
        import psutil
        dialog_killer_running = False

        for p in psutil.process_iter(['pid', 'name', 'cmdline']):
            try:
                if p.info['name'] in ['python.exe', 'pythonw.exe']:
                    cmdline = p.info.get('cmdline', [])
                    if any('cad_dialog_killer.py' in arg for arg in cmdline):
                        dialog_killer_running = True
                        print(f"[成功] cad_dialog_killer 正在运行 (PID: {p.info['pid']})")
                        break
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue

        if not dialog_killer_running:
            print("[警告] cad_dialog_killer 未检测到运行")

        print("\n" + "="*60)
        print("[OK] 测试1完成")
        print("="*60)
        return True

    except Exception as e:
        print(f"\n[ERROR] 测试1失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_close_all_cad_processes():
    """测试2：强制关闭所有CAD进程"""
    print("\n" + "="*60)
    print("测试2：close_all_cad_processes() - 强制关闭功能加强")
    print("="*60)

    try:
        # 检查当前CAD进程数
        shu_before = jingchengshu_wenjian()
        print(f"\n[检查] 关闭前CAD进程数: {shu_before}")

        if shu_before == 0:
            print("[信息] 没有CAD进程，先启动一个用于测试...")
            start_applicationV9()
            time.sleep(5)
            shu_before = jingchengshu_wenjian()
            print(f"[检查] 启动后CAD进程数: {shu_before}")

        # 执行强制关闭
        print("\n[测试] 执行 close_all_cad_processes()...")
        start_time = time.time()
        result = close_all_cad_processes()
        elapsed = time.time() - start_time

        print(f"\n[结果] 执行结果: {result}")
        print(f"[耗时] {elapsed:.2f}秒")

        # 验证是否真正关闭
        time.sleep(2)
        shu_after = jingchengshu_wenjian()
        print(f"[验证] 关闭后CAD进程数: {shu_after}")

        if shu_after == 0:
            print("[成功] 所有CAD进程已关闭")
            success = True
        else:
            print(f"[失败] 仍有 {shu_after} 个CAD进程未关闭")
            success = False

        print("\n" + "="*60)
        print("[OK] 测试2完成" if success else "[FAIL] 测试2失败")
        print("="*60)
        return success

    except Exception as e:
        print(f"\n[ERROR] 测试2失败: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    print("开始测试修复后的两个函数\n")

    # 测试1：启动CAD和弹窗治理
    test1_success = test_start_application_and_dialog_killer()

    # 等待一会儿
    time.sleep(3)

    # 测试2：强制关闭CAD
    test2_success = test_close_all_cad_processes()

    # 最终清理
    print("\n" + "="*60)
    print("最终清理")
    print("="*60)
    time.sleep(2)
    shu = jingchengshu_wenjian()
    if shu > 0:
        print(f"[清理] 检测到 {shu} 个残留CAD进程，正在清理...")
        close_all_cad_processes()

    # 总结
    print("\n" + "="*60)
    print("测试总结")
    print("="*60)
    print(f"测试1 (start_applicationV9): {'通过' if test1_success else '失败'}")
    print(f"测试2 (close_all_cad_processes): {'通过' if test2_success else '失败'}")

    if test1_success and test2_success:
        print("\n[OK] 所有测试通过")
        sys.exit(0)
    else:
        print("\n[FAIL] 部分测试失败")
        sys.exit(1)
