#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""在单进程CAD中测试插入门"""
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_basic import get_acad_doc, close_all_cad_processes, start_applicationV9
from CAD_file_operations import open_file, insert_tarch_door, save_file
from CAD_coordination import ensure_single_process, wait_quiescent

print("="*60)
print("在单进程CAD中测试插入门")
print("="*60)

try:
    # 先关闭所有CAD进程
    print("\n[步骤1] 关闭所有CAD进程...")
    close_all_cad_processes()

    # 启动单个CAD进程
    print("\n[步骤2] 启动CAD...")
    proc = start_applicationV9(PTH=r"C:\Tangent\TArchT20V9")
    if not proc:
        print("[错误] CAD启动失败")
        sys.exit(1)

    ensure_single_process()
    wait_quiescent(min_quiet=2.0, timeout=30.0)

    # 打开测试文件
    test_file = "D:/claude-tasks/tests/test_files/天正测试文件1.dwg"
    print(f"\n[步骤3] 打开文件: {test_file}")
    open_file(test_file)
    wait_quiescent(min_quiet=1.0, timeout=15.0)

    # 验证当前文档
    acad, doc = get_acad_doc()
    print(f"  当前文档: {doc.Name}")

    # 插入门
    print(f"\n[步骤4] 插入天正门...")
    p = (101000, 70000, 0)
    result = insert_tarch_door(p, width=1000, height=2100)

    if result['success']:
        print(f"  [成功] 宽度:{result['width']}, 高度:{result['height']}")
    else:
        print(f"  [失败]")

    # 保存
    print(f"\n[步骤5] 保存文件...")
    save_file()

    print("\n" + "="*60)
    print("测试完成！")
    print("="*60)

except Exception as e:
    print(f"\n[错误] {e}")
    import traceback
    traceback.print_exc()

finally:
    print("\n[清理] 关闭所有CAD进程...")
    close_all_cad_processes()
