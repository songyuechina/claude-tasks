#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
执行 insert_tarch_window 函数测试
按照即时对话.txt的要求执行测试
"""

import sys
import time
from pathlib import Path

# 添加scripts目录到路径
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_basic import start_applicationV9, li, wait_quiescent
from CAD_coordination import ensure_single_process
from test_insert_tarch_window import insert_tarch_window
from CAD_file_operations import open_file, save_file

def main():
    """主测试函数"""
    print("="*60)
    print("insert_tarch_window 函数测试 - 按照即时对话.txt要求")
    print("="*60)

    # 1. 确保单进程（关闭所有CAD）
    print("\n[步骤 1/7] 关闭所有CAD进程...")
    ensure_single_process()
    print("[完成] CAD进程已清理")

    # 2. 启动天正CAD（界面可见）
    print("\n[步骤 2/7] 启动天正CAD...")
    start_applicationV9(
        PTH=r"C:\Tangent\TArchT20V9",
        max_retries=3,
        retry_delay=2.0
    )
    wait_quiescent(min_quiet=1.0, timeout=30.0)
    print("[完成] 天正CAD已启动")

    # 3. 打开测试文件
    print("\n[步骤 3/7] 打开测试文件 insert_tarch_window-2.dwg...")
    test_file = "D:/claude-tasks/cad/Function_testing/insert_tarch_window-2.dwg"
    open_file(test_file)
    wait_quiescent(min_quiet=1.0, timeout=15.0)
    print("[完成] 测试文件已打开")

    # 4. 连接当前文件
    print("\n[步骤 4/7] 使用 li() 连接当前激活文件...")
    li()
    print("[完成] 已连接到当前激活文件")

    # 5. 测试1: 在指定位置插入 jz-gaochuang，宽1200，高1000，delete_mc_yuan=False
    print("\n[步骤 5/7] 测试1: 插入 jz-gaochuang 窗...")
    print("  位置: (38612.86565445, 48750.63891910, 0)")
    print("  宽度: 1200, 高度: 1000")
    print("  类型: jz-gaochuang")
    print("  delete_mc_yuan: False")

    result1 = insert_tarch_window(
        p=(38612.86565445, 48750.63891910, 0),
        width=1200,
        height=1000,
        window_type="jz-gaochuang",
        delete_mc_yuan=False
    )

    if result1['success']:
        print(f"[成功] 测试1完成 - 窗宽度:{result1['width']}, 高度:{result1['height']}")
    else:
        print(f"[失败] 测试1失败")
        return False

    time.sleep(2)

    # 6. 测试2: 在指定位置插入 jz-pingchuang，宽2400，高1800，delete_mc_yuan=True
    print("\n[步骤 6/7] 测试2: 插入 jz-pingchuang 窗...")
    print("  位置: (44695.30568975, 46646.78059028, 0)")
    print("  宽度: 2400, 高度: 1800")
    print("  类型: jz-pingchuang")
    print("  delete_mc_yuan: True")

    result2 = insert_tarch_window(
        p=(44695.30568975, 46646.78059028, 0),
        width=2400,
        height=1800,
        window_type="jz-pingchuang",
        delete_mc_yuan=True
    )

    if result2['success']:
        print(f"[成功] 测试2完成 - 窗宽度:{result2['width']}, 高度:{result2['height']}")
    else:
        print(f"[失败] 测试2失败")
        return False

    # 7. 保存文件
    print("\n[步骤 7/7] 保存测试文件...")
    save_file()
    print("[完成] 文件已保存")

    print("\n" + "="*60)
    print("✓ 测试完成！所有测试均成功！")
    print("="*60)
    print(f"\n测试文件已保存到: {test_file}")
    print("\n测试总结:")
    print("  - 测试1: jz-gaochuang (1200x1000) ✓")
    print("  - 测试2: jz-pingchuang (2400x1800) ✓")

    return True


if __name__ == "__main__":
    try:
        success = main()
        if not success:
            print("\n测试失败，请检查日志")
            sys.exit(1)
    except Exception as e:
        print(f"\n测试过程中发生错误: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
