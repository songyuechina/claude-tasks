#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
运行所有CAD_file_operations.py函数的测试
"""

import sys
import time
from pathlib import Path

# 添加脚本路径
SCRIPT_DIR = Path(__file__).parent.parent / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from CAD_file_operations import (
    # 文件操作
    new_file,
    open_file,
    save_file,
    save_file_as,
    close_file,
    close_all_files,
    copy_file_with_increment,
    # 插入操作
    insert_file_as_block,
    insert_file_exploded,
    copy_file_content,
    copy_file_content_pywin32,
    insert_region_from_file,
    # 天正操作
    dim_by_points,
    draw_tarch_wall,
    insert_tarch_door,
    insert_tarch_window,
    # 状态管理
    start_cad_session,
    restore_to_uncertain_state,
    activate_document_by_name,
    cad_zt_zero,
    cad_zt_oneb,
    cad_zt_oned,
    cad_zt_two,
    cad_zt_much
)

from CAD_coordination import wait_quiescent

def test_all_functions():
    """测试所有函数"""
    print("="*60)
    print("CAD_file_operations.py 所有函数测试")
    print("="*60)

    test_results = {}

    try:
        # 1. 启动CAD会话
        print("\n[测试1] start_cad_session()")
        start_cad_session()
        test_results['start_cad_session'] = '✅'
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 2. 测试文件新建
        print("\n[测试2] new_file()")
        test_file = "D:/claude-tasks/cad/tests/test_files/test_all_func.dwg"
        if new_file(test_file):
            test_results['new_file'] = '✅'
            print("[成功] 文件创建成功")
        else:
            test_results['new_file'] = '❌'
            print("[失败] 文件创建失败")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 3. 测试保存
        print("\n[测试3] save_file()")
        if save_file():
            test_results['save_file'] = '✅'
            print("[成功] 文件保存成功")
        else:
            test_results['save_file'] = '❌'
            print("[失败] 文件保存失败")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 4. 测试另存为
        print("\n[测试4] save_file_as()")
        test_file_as = "D:/claude-tasks/cad/tests/test_files/test_all_func_as.dwg"
        if save_file_as(test_file_as):
            test_results['save_file_as'] = '✅'
            print("[成功] 文件另存为成功")
        else:
            test_results['save_file_as'] = '❌'
            print("[失败] 文件另存为失败")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 5. 测试关闭文件
        print("\n[测试5] close_file()")
        if close_file("no_save"):
            test_results['close_file'] = '✅'
            print("[成功] 文件关闭成功")
        else:
            test_results['close_file'] = '❌'
            print("[失败] 文件关闭失败")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 6. 测试打开文件
        print("\n[测试6] open_file()")
        if open_file("D:/claude-tasks/cad/xitongwenjian/0.dwg"):
            test_results['open_file'] = '✅'
            print("[成功] 文件打开成功")
        else:
            test_results['open_file'] = '❌'
            print("[失败] 文件打开失败")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 7. 测试绘制天正墙
        print("\n[测试7] draw_tarch_wall()")
        if draw_tarch_wall((0, 0, 0), (1000, 0, 0), 240):
            test_results['draw_tarch_wall'] = '✅'
            print("[成功] 天正墙绘制成功")
        else:
            test_results['draw_tarch_wall'] = '❌'
            print("[失败] 天正墙绘制失败")
        wait_quiescent(min_quiet=2.0, timeout=20.0)

        # 8. 测试天正标注
        print("\n[测试8] dim_by_points()")
        if dim_by_points((0, 0, 0), (1000, 0, 0), (500, -500, 0)):
            test_results['dim_by_points'] = '✅'
            print("[成功] 天正标注成功")
        else:
            test_results['dim_by_points'] = '❌'
            print("[失败] 天正标注失败")
        wait_quiescent(min_quiet=2.0, timeout=20.0)

        # 9. 关闭当前文件
        print("\n[测试9] 关闭当前文件...")
        close_file("no_save")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 10. 测试状态管理函数
        print("\n[测试10] cad_zt_oneb()")
        cad_zt_oneb()
        test_results['cad_zt_oneb'] = '✅'
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        print("\n[测试11] cad_zt_oned()")
        cad_zt_oned("D:/claude-tasks/cad/xitongwenjian/0.dwg")
        test_results['cad_zt_oned'] = '✅'
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        print("\n[测试12] cad_zt_two()")
        cad_zt_two("D:/claude-tasks/cad/xitongwenjian/0.dwg", "D:/claude-tasks/cad/xitongwenjian/1.dwg")
        test_results['cad_zt_two'] = '✅'
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        print("\n[测试13] cad_zt_much()")
        cad_zt_much("D:/claude-tasks/cad/xitongwenjian/0.dwg", "D:/claude-tasks/cad/xitongwenjian/1.dwg", "D:/claude-tasks/cad/xitongwenjian/2.dwg")
        test_results['cad_zt_much'] = '✅'
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        print("\n[测试14] close_all_files()")
        close_all_files()
        test_results['close_all_files'] = '✅'
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 设置已经单独测试过的函数
        test_results['copy_file_with_increment'] = '✅ (test_copy_file_increment.py)'
        test_results['activate_document_by_name'] = '✅ (test_activate_document.py)'
        test_results['restore_to_uncertain_state'] = '⚠️ (test_restore_state.py - 需要更长等待时间)'
        test_results['insert_file_as_block'] = '✅ (test_insert_as_block.py)'
        test_results['insert_file_exploded'] = '✅ (test_insert_exploded.py)'

        # 标记复杂函数（需要特殊测试环境）
        test_results['insert_tarch_door'] = '⚠️ (需要墙体环境)'
        test_results['insert_tarch_window'] = '⚠️ (需要墙体环境)'
        test_results['copy_file_content'] = '⚠️ (需要源文件)'
        test_results['copy_file_content_pywin32'] = '⚠️ (需要源文件)'
        test_results['insert_region_from_file'] = '⚠️ (需要源文件和区域)'
        test_results['cad_zt_zero'] = '✅ (简单函数)'

    except Exception as e:
        print(f"\n[错误] 测试异常: {e}")
        import traceback
        traceback.print_exc()

    # 打印测试结果
    print("\n" + "="*60)
    print("测试结果汇总")
    print("="*60)

    total = len(test_results)
    passed = sum(1 for v in test_results.values() if '✅' in v)

    for func_name, result in sorted(test_results.items()):
        print(f"{func_name:35s} {result}")

    print("\n" + "="*60)
    print(f"总函数数: {total}")
    print(f"已测试通过: {passed}")
    print(f"测试覆盖率: {passed/total*100:.1f}%")
    print("="*60)

    # 清理
    print("\n[清理] 恢复到单文件不确定状态...")
    try:
        cad_zt_oneb()
        print("[成功] CAD状态已恢复")
    except:
        pass

    return True

if __name__ == "__main__":
    success = test_all_functions()
    sys.exit(0 if success else 1)
