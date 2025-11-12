#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试dim_by_points函数 - 倾斜对象标注
"""

import sys
from pathlib import Path

sys.path.append(str(Path(__file__).parent.parent / "scripts"))

from CAD_basic import start_applicationV9, close_all_cad_processes, dim_by_points
from CAD_file_operations import open_file
from CAD_coordination import ensure_single_process, wait_quiescent

def main():
    print("="*60)
    print("测试dim_by_points函数 - 倾斜对象标注")
    print("="*60)

    try:
        # 启动CAD
        print("\n[步骤1] 启动CAD...")
        proc = start_applicationV9(PTH=r"C:\Tangent\TArchT20V9")
        if not proc:
            print("[错误] CAD启动失败")
            return False

        ensure_single_process()
        wait_quiescent(min_quiet=2.0, timeout=30.0)

        # 打开测试文件
        test_file = "D:/claude-tasks/tests/test_files/天正测试文件1.dwg"
        print(f"\n[步骤2] 打开测试文件: {test_file}")
        if not open_file(test_file):
            print("[错误] 文件打开失败")
            return False

        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 测试点坐标
        p1 = (96347.4900202, 96347.4900202, 0)
        p2 = (109748.28777108, 69727.96879379, 0)
        p3 = (98706.45627348, 61997.28289969, 0)

        print(f"\n[步骤3] 执行倾斜对象标注")
        print(f"  起点: {p1}")
        print(f"  终点: {p2}")
        print(f"  标注位置: {p3}")

        # 执行标注
        result = dim_by_points(p1, p2, p3)

        if result:
            print("\n[成功] 标注完成")
        else:
            print("\n[失败] 标注失败")
            return False

        print("\n" + "="*60)
        print("测试完成!")
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
