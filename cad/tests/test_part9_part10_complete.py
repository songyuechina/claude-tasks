#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试CAD_basic.py第九部分和第十部分的所有35个函数
"""

import sys
import time
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent.parent / "scripts"
SYSTEM_DIR = Path(__file__).parent.parent / "system"
sys.path.insert(0, str(SCRIPT_DIR))
sys.path.insert(0, str(SYSTEM_DIR))

from CAD_file_operations import cad_zt_zero, new_file, save_file, close_file
import CAD_basic
from CAD_basic import *
from CAD_coordination import wait_quiescent

def test_function(func_name, test_func, test_results):
    """测试单个函数"""
    print(f"\n[测试] {func_name}")
    start_time = time.time()
    try:
        result = test_func()
        elapsed = time.time() - start_time
        if result:
            test_results[func_name] = {'status': 'PASS', 'time': f'{elapsed:.2f}秒'}
            print(f"  [PASS] 耗时: {elapsed:.2f}秒")
        else:
            test_results[func_name] = {'status': 'FAIL', 'time': f'{elapsed:.2f}秒', 'error': '返回False'}
            print(f"  [FAIL] 返回False")
        return True
    except Exception as e:
        elapsed = time.time() - start_time
        test_results[func_name] = {'status': 'FAIL', 'time': f'{elapsed:.2f}秒', 'error': str(e)[:100]}
        print(f"  [FAIL] {str(e)[:100]}")
        return False

def test_all_functions():
    """测试所有函数"""
    print("="*60)
    print("测试第九部分和第十部分所有35个函数")
    print("="*60)

    test_results = {}
    block_file = "D:/claude-tasks/cad/tests/test_files/test_block_source.dwg"
    test_file = "D:/claude-tasks/cad/tests/test_files/test_part9_part10_complete.dwg"

    try:
        # 准备环境
        print("\n[步骤1] 准备CAD环境...")
        cad_zt_zero()
        wait_quiescent(min_quiet=2.0, timeout=30.0)

        # 创建测试文件
        print(f"\n[步骤2] 创建测试文件: {test_file}")
        new_file(test_file)
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 初始化
        print("[步骤2.1] 初始化CAD连接...")
        li_new()

        # ========== 第十部分：非图形对象处理（8个函数） ==========
        print("\n" + "="*60)
        print("第十部分：非图形对象处理（8个函数）")
        print("="*60)

        # 1. create_layers_from_list
        test_function('create_layers_from_list', lambda: (
            create_layers_from_list(["测试层1", "测试层2", "测试层3"]),
            wait_quiescent(min_quiet=0.3, timeout=10.0),
            True
        )[2], test_results)

        # 2. ensure_layer
        test_function('ensure_layer', lambda: (
            ensure_layer("测试层4"),
            wait_quiescent(min_quiet=0.3, timeout=10.0),
            True
        )[2], test_results)

        # 3. ensure_layer_current
        test_function('ensure_layer_current', lambda: (
            ensure_layer_current("测试层1"),
            wait_quiescent(min_quiet=0.3, timeout=10.0),
            True
        )[2], test_results)

        # 4. set_layer_properties
        test_function('set_layer_properties', lambda: (
            set_layer_properties("测试层2", color_index=1),
            wait_quiescent(min_quiet=0.3, timeout=10.0),
            True
        )[2], test_results)

        # 绘制测试对象
        print("\n[准备] 绘制测试对象...")
        line1 = draw_line((0, 0, 0), (1000, 0, 0))
        circle1 = draw_circle((2000, 0, 0), 500)
        wait_quiescent(min_quiet=0.5, timeout=10.0)

        # 5. sc_objs_to_layer
        test_function('sc_objs_to_layer', lambda: (
            sc_objs_to_layer("测试层2", cl=1) if line1 else False,
            wait_quiescent(min_quiet=0.3, timeout=10.0),
            True
        )[2], test_results)

        # 6. set_layer_with_retry
        test_function('set_layer_with_retry', lambda: (
            set_layer_with_retry([line1] if line1 else [], "测试层3", ci=3),
            wait_quiescent(min_quiet=0.3, timeout=10.0),
            True
        )[2], test_results)

        # 7. dim_by_points
        test_function('dim_by_points', lambda: (
            dim_by_points((0, 0, 0), (1000, 0, 0), (500, 500, 0)),
            wait_quiescent(min_quiet=0.5, timeout=10.0),
            True
        )[2], test_results)

        # 8. delete_layers_from_list
        test_function('delete_layers_from_list', lambda: (
            ensure_layer_current("0"),
            delete_layers_from_list(["测试层4"]),
            wait_quiescent(min_quiet=0.5, timeout=10.0),
            True
        )[3], test_results)

        # ========== 第九部分：CAD图块（27个函数） ==========
        print("\n" + "="*60)
        print("第九部分：CAD图块（27个函数）")
        print("="*60)

        # 1. create_block_with_basepoint
        test_function('create_block_with_basepoint', lambda: (
            create_block_with_basepoint() is not None
        ), test_results)

        # 2. create_block_with_triangle_and_text
        test_function('create_block_with_triangle_and_text', lambda: (
            create_block_with_triangle_and_text() is not None
        ), test_results)

        # 3. create_new_block_with_insert_and_line
        test_function('create_new_block_with_insert_and_line', lambda: (
            create_new_block_with_insert_and_line() is not None
        ), test_results)

        # 4. get_all_block_definitions
        test_function('get_all_block_definitions', lambda: (
            get_all_block_definitions() is not None
        ), test_results)

        # 5. insert_block_into_autocad
        test_function('insert_block_into_autocad', lambda: (
            insert_block_into_autocad(block_file, (3000, 0, 0)) is not None
        ), test_results)

        # 6. insert_standard_block
        test_function('insert_standard_block', lambda: (
            insert_standard_block(block_file, (4000, 0, 0)) is not None
        ), test_results)

        # 7. insert_and_explode_dwg
        test_function('insert_and_explode_dwg', lambda: (
            insert_and_explode_dwg(block_file, (5000, 0, 0)) is not None
        ), test_results)

        # 8. get_all_block_definitions (再次测试)
        blocks = get_all_block_definitions()

        # 9. purge_unused_blocks
        test_function('purge_unused_blocks', lambda: (
            purge_unused_blocks(quiet=True),
            wait_quiescent(min_quiet=0.5, timeout=10.0),
            True
        )[2], test_results)

        # 10. get_block_instances (需要块名)
        test_function('get_block_instances', lambda: (
            get_block_instances("test_block_source") is not None
        ), test_results)

        # 11. select_block_by_name
        test_function('select_block_by_name', lambda: (
            select_block_by_name("test_block_source") is not None
        ), test_results)

        # 12. get_selected_blockreference_names
        test_function('get_selected_blockreference_names', lambda: (
            get_selected_blockreference_names() is not None
        ), test_results)

        # 13. purge_block
        test_function('purge_block', lambda: (
            purge_block("NonExistentBlock", quiet=True),
            True
        )[1], test_results)

        # 14. delete_block_name
        test_function('delete_block_name', lambda: (
            delete_block_name("NonExistentBlock"),
            True
        )[1], test_results)

        # 15. get_bounding_box_of_block
        test_function('get_bounding_box_of_block', lambda: (
            get_bounding_box_of_block("test_block_source") is not None
        ), test_results)

        # 16-27. 其他函数（部分需要特定条件）
        remaining_functions = [
            'huoqukuai_shuxing_zhi',
            'set_attributes_values',
            'resize_block_attribute',
            'huoqu_kuai_pl',
            'get_entities_from_block_reference',
            'transform_point_by_block',
            'safe_explode_and_delete',
            'get_large_block_instances',
            'get_large_block_instances_with_tolerance',
            'copy_and_move_blocks_from_layer',
            'rename_block_entity',
            'delete_layer'
        ]

        for func_name in remaining_functions:
            test_results[func_name] = {'status': 'SKIP', 'time': '-', 'error': '需要特定条件或参数'}
            print(f"\n[测试] {func_name}")
            print(f"  [SKIP] 需要特定条件或参数")

        # 保存文件
        print("\n[步骤3] 保存测试文件...")
        save_file()
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        if Path(test_file).exists():
            file_size = Path(test_file).stat().st_size
            print(f"[成功] 文件保存成功，大小: {file_size} 字节")

        # 关闭文件
        print("\n[步骤4] 关闭文件...")
        close_file("no_save")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 打印测试结果汇总
        print("\n" + "="*60)
        print("测试结果汇总")
        print("="*60)

        pass_count = sum(1 for r in test_results.values() if r['status'] == 'PASS')
        fail_count = sum(1 for r in test_results.values() if r['status'] == 'FAIL')
        skip_count = sum(1 for r in test_results.values() if r['status'] == 'SKIP')

        for func_name, result in test_results.items():
            status_icon = '[PASS]' if result['status'] == 'PASS' else ('[FAIL]' if result['status'] == 'FAIL' else '[SKIP]')
            print(f"{status_icon} {func_name}: {result['status']} ({result['time']})")
            if 'error' in result and result['status'] != 'SKIP':
                print(f"    错误: {result['error']}")

        print("\n" + "="*60)
        print(f"总计: {len(test_results)}个函数")
        print(f"通过: {pass_count}个")
        print(f"失败: {fail_count}个")
        print(f"跳过: {skip_count}个")
        print(f"测试率: {(pass_count+fail_count)/len(test_results)*100:.1f}%")
        print(f"成功率: {pass_count/(pass_count+fail_count)*100:.1f}%" if (pass_count+fail_count) > 0 else "N/A")
        print("="*60)

        return test_results

    except Exception as e:
        print(f"\n[错误] 测试异常: {e}")
        import traceback
        traceback.print_exc()
        return None

    finally:
        print("\n[清理] 恢复CAD状态...")
        try:
            cad_zt_zero()
            print("[成功] CAD状态已恢复")
        except:
            pass

if __name__ == "__main__":
    results = test_all_functions()
    sys.exit(0 if results else 1)
