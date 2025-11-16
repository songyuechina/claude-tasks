#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""测试命令1113 - 绘制多段线、矩形、圆、直线、天正墙"""

import sys
from pathlib import Path
import time

SCRIPT_DIR = Path(__file__).parent.parent / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from CAD_file_operations import (
    cad_zt_zero,
    new_file,
    save_file,
    draw_tarch_wall,
    start_cad_session
)
from CAD_coordination import send_cmd_with_sync, wait_quiescent

def test_command_1113():
    """测试绘制多种图形"""
    print("="*60)
    print("测试命令1113 - 绘制多段线、矩形、圆、直线、天正墙")
    print("="*60)

    # 测试开始前清理状态
    print("\n[1] 清理CAD状态...")
    cad_zt_zero()

    start_time = time.time()

    try:
        # 启动CAD会话
        print("\n[2] 启动CAD会话...")
        if not start_cad_session():
            print("[错误] CAD启动失败")
            return False
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 新建文件
        test_file = "D:/claude-tasks/cad/tests/test_files/命令测试1113.dwg"
        print(f"\n[3] 新建文件: {test_file}")
        if not new_file(test_file):
            print("[错误] 文件新建失败")
            return False
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 绘制多段线
        print("\n[4] 绘制多段线...")
        if not send_cmd_with_sync("_PLINE\n0,0\n1000,0\n1000,1000\n0,1000\n_C\n", wait_after=2.0, timeout=30.0):
            print("[错误] 多段线绘制失败")
        else:
            print("✅ 多段线绘制完成")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 绘制矩形
        print("\n[5] 绘制矩形...")
        if not send_cmd_with_sync("_RECTANG\n2000,0\n3000,1000\n", wait_after=2.0, timeout=30.0):
            print("[错误] 矩形绘制失败")
        else:
            print("✅ 矩形绘制完成")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 绘制圆
        print("\n[6] 绘制圆...")
        if not send_cmd_with_sync("_CIRCLE\n4000,500\n500\n", wait_after=2.0, timeout=30.0):
            print("[错误] 圆绘制失败")
        else:
            print("✅ 圆绘制完成")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 绘制直线
        print("\n[7] 绘制直线...")
        if not send_cmd_with_sync("_LINE\n5000,0\n6000,1000\n\n", wait_after=2.0, timeout=30.0):
            print("[错误] 直线绘制失败")
        else:
            print("✅ 直线绘制完成")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 绘制天正墙
        print("\n[8] 绘制天正墙...")
        wall_start = time.time()
        if not draw_tarch_wall((0, 2000, 0), (3000, 2000, 0), 240):
            print("[错误] 天正墙绘制失败")
        else:
            wall_elapsed = time.time() - wall_start
            print(f"✅ 天正墙绘制完成 (耗时: {wall_elapsed:.2f}秒)")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 保存文件
        print("\n[9] 保存文件...")
        if not save_file():
            print("[错误] 文件保存失败")
            return False
        print("✅ 文件保存成功")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        elapsed = time.time() - start_time
        print(f"\n[成功] 测试完成！执行时间: {elapsed:.2f}秒")
        print(f"[成功] 文件已保存: {test_file}")

        return True

    except Exception as e:
        print(f"\n[错误] 测试异常: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        # 测试结束后恢复状态
        print("\n[9] 恢复CAD状态...")
        cad_zt_zero()

if __name__ == "__main__":
    success = test_command_1113()
    sys.exit(0 if success else 1)
