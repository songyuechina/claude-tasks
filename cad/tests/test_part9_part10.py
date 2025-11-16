#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试CAD_basic.py第九部分（CAD图块）和第十部分（非图形对象处理）的核心函数
"""

import sys
import time
from pathlib import Path

# 添加脚本路径
SCRIPT_DIR = Path(__file__).parent.parent / "scripts"
SYSTEM_DIR = Path(__file__).parent.parent / "system"
sys.path.insert(0, str(SCRIPT_DIR))
sys.path.insert(0, str(SYSTEM_DIR))

from CAD_file_operations import (
    cad_zt_zero,
    new_file,
    save_file,
    close_file
)

import CAD_basic
from CAD_basic import (
    li_new,
    create_block_with_basepoint,
    get_all_block_definitions,
    purge_block,
    purge_unused_blocks,
    create_layers_from_list,
    ensure_layer,
    ensure_layer_current,
    set_layer_properties,
    delete_layers_from_list,
    draw_line,
    draw_circle
)

from CAD_coordination import wait_quiescent

def test_part9_part10():
    """测试第九部分和第十部分核心函数"""
    print("="*60)
    print("测试CAD_basic.py第九部分（CAD图块）和第十部分（非图形对象处理）")
    print("="*60)

    test_results = {}

    try:
        # 测试开始前 - 调用cad_zt_zero()
        print("\n[步骤1] 准备CAD环境...")
        cad_zt_zero()
        wait_quiescent(min_quiet=2.0, timeout=30.0)

        # 创建测试文件
        test_file = "D:/claude-tasks/cad/tests/test_files/test_part9_part10.dwg"
        print(f"\n[步骤2] 创建测试文件: {test_file}")
        new_file(test_file)
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 初始化CAD连接
        print("[步骤2.1] 初始化CAD连接...")
        li_new()
        print("[成功] CAD连接已初始化")

        # ========== 第十部分测试：图层操作 ==========
        print("\n" + "="*60)
        print("第十部分测试：图层操作")
        print("="*60)

        # ===== 测试1: create_layers_from_list =====
        print("\n[测试1] create_layers_from_list - 批量创建图层")
        start_time = time.time()
        try:
            test_layers = ["测试图层1", "测试图层2", "测试图层3"]
            create_layers_from_list(test_layers)
            wait_quiescent(min_quiet=0.5, timeout=10.0)
            elapsed = time.time() - start_time
            test_results['create_layers_from_list'] = {'status': 'PASS', 'time': f'{elapsed:.2f}秒'}
            print(f"  [成功] 创建3个图层，耗时: {elapsed:.2f}秒")
        except Exception as e:
            test_results['create_layers_from_list'] = {'status': 'FAIL', 'time': '-', 'error': str(e)}
            print(f"  [失败] {e}")

        # ===== 测试2: ensure_layer =====
        print("\n[测试2] ensure_layer - 确保图层存在")
        start_time = time.time()
        try:
            ensure_layer("测试图层4")
            wait_quiescent(min_quiet=0.3, timeout=10.0)
            elapsed = time.time() - start_time
            test_results['ensure_layer'] = {'status': 'PASS', 'time': f'{elapsed:.2f}秒'}
            print(f"  [成功] 确保图层存在，耗时: {elapsed:.2f}秒")
        except Exception as e:
            test_results['ensure_layer'] = {'status': 'FAIL', 'time': '-', 'error': str(e)}
            print(f"  [失败] {e}")

        # ===== 测试3: ensure_layer_current =====
        print("\n[测试3] ensure_layer_current - 设置当前图层")
        start_time = time.time()
        try:
            ensure_layer_current("测试图层1")
            wait_quiescent(min_quiet=0.3, timeout=10.0)
            elapsed = time.time() - start_time
            test_results['ensure_layer_current'] = {'status': 'PASS', 'time': f'{elapsed:.2f}秒'}
            print(f"  [成功] 设置当前图层，耗时: {elapsed:.2f}秒")
        except Exception as e:
            test_results['ensure_layer_current'] = {'status': 'FAIL', 'time': '-', 'error': str(e)}
            print(f"  [失败] {e}")

        # ===== 测试4: set_layer_properties =====
        print("\n[测试4] set_layer_properties - 设置图层属性")
        start_time = time.time()
        try:
            set_layer_properties("测试图层2", color_index=1, linetype="Continuous")
            wait_quiescent(min_quiet=0.3, timeout=10.0)
            elapsed = time.time() - start_time
            test_results['set_layer_properties'] = {'status': 'PASS', 'time': f'{elapsed:.2f}秒'}
            print(f"  [成功] 设置图层属性（红色），耗时: {elapsed:.2f}秒")
        except Exception as e:
            test_results['set_layer_properties'] = {'status': 'FAIL', 'time': '-', 'error': str(e)}
            print(f"  [失败] {e}")

        # 在测试图层上绘制对象
        print("\n[准备] 在测试图层上绘制对象...")
        try:
            ensure_layer_current("测试图层1")
            line1 = draw_line((0, 0, 0), (1000, 0, 0))
            circle1 = draw_circle((2000, 0, 0), 500)
            wait_quiescent(min_quiet=0.5, timeout=10.0)
            print("  [成功] 在测试图层1上绘制了线和圆")
        except Exception as e:
            print(f"  [警告] 绘制对象失败: {e}")

        # ========== 第九部分测试：块操作 ==========
        print("\n" + "="*60)
        print("第九部分测试：块操作")
        print("="*60)

        # ===== 测试5: create_block_with_basepoint =====
        print("\n[测试5] create_block_with_basepoint - 创建带基点的块")
        start_time = time.time()
        try:
            block_obj = create_block_with_basepoint()
            wait_quiescent(min_quiet=0.5, timeout=10.0)

            if block_obj:
                elapsed = time.time() - start_time
                test_results['create_block_with_basepoint'] = {'status': 'PASS', 'time': f'{elapsed:.2f}秒'}
                print(f"  [成功] 创建块，耗时: {elapsed:.2f}秒")
            else:
                test_results['create_block_with_basepoint'] = {'status': 'FAIL', 'time': '-', 'error': '返回None'}
                print("  [失败] 未能创建块")
        except Exception as e:
            test_results['create_block_with_basepoint'] = {'status': 'FAIL', 'time': '-', 'error': str(e)}
            print(f"  [失败] {e}")

        # ===== 测试6: get_all_block_definitions =====
        print("\n[测试6] get_all_block_definitions - 获取所有块定义")
        start_time = time.time()
        try:
            blocks = get_all_block_definitions()
            elapsed = time.time() - start_time

            if blocks is not None:
                block_count = len(blocks) if hasattr(blocks, '__len__') else 'N/A'
                test_results['get_all_block_definitions'] = {'status': 'PASS', 'time': f'{elapsed:.2f}秒', 'result': f'{block_count}个块'}
                print(f"  [成功] 获取块定义，数量: {block_count}，耗时: {elapsed:.2f}秒")
            else:
                test_results['get_all_block_definitions'] = {'status': 'FAIL', 'time': '-', 'error': '返回None'}
                print("  [失败] 未能获取块定义")
        except Exception as e:
            test_results['get_all_block_definitions'] = {'status': 'FAIL', 'time': '-', 'error': str(e)}
            print(f"  [失败] {e}")

        # ===== 测试7: purge_unused_blocks =====
        print("\n[测试7] purge_unused_blocks - 清除未使用的块")
        start_time = time.time()
        try:
            purge_unused_blocks(quiet=True)
            wait_quiescent(min_quiet=0.5, timeout=10.0)
            elapsed = time.time() - start_time
            test_results['purge_unused_blocks'] = {'status': 'PASS', 'time': f'{elapsed:.2f}秒'}
            print(f"  [成功] 清除未使用的块，耗时: {elapsed:.2f}秒")
        except Exception as e:
            test_results['purge_unused_blocks'] = {'status': 'FAIL', 'time': '-', 'error': str(e)}
            print(f"  [失败] {e}")

        # ===== 测试8: delete_layers_from_list =====
        print("\n[测试8] delete_layers_from_list - 删除图层")
        start_time = time.time()
        try:
            # 先切换到0层，避免删除当前层
            ensure_layer_current("0")
            wait_quiescent(min_quiet=0.3, timeout=10.0)

            delete_layers_from_list(["测试图层3", "测试图层4"])
            wait_quiescent(min_quiet=0.5, timeout=10.0)
            elapsed = time.time() - start_time
            test_results['delete_layers_from_list'] = {'status': 'PASS', 'time': f'{elapsed:.2f}秒'}
            print(f"  [成功] 删除2个图层，耗时: {elapsed:.2f}秒")
        except Exception as e:
            test_results['delete_layers_from_list'] = {'status': 'FAIL', 'time': '-', 'error': str(e)}
            print(f"  [失败] {e}")

        # 保存文件
        print("\n[步骤3] 保存测试文件...")
        save_file()
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 验证文件
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

        for func_name, result in test_results.items():
            status_icon = '[PASS]' if result['status'] == 'PASS' else '[FAIL]'
            print(f"{status_icon} {func_name}: {result['status']} ({result['time']})")
            if 'error' in result:
                print(f"    错误: {result['error']}")
            if 'result' in result:
                print(f"    结果: {result['result']}")

        print("\n" + "="*60)
        print(f"总计: {len(test_results)}个函数")
        print(f"通过: {pass_count}个")
        print(f"失败: {fail_count}个")
        print(f"成功率: {pass_count/len(test_results)*100:.1f}%")
        print("="*60)

        return test_results

    except Exception as e:
        print(f"\n[错误] 测试异常: {e}")
        import traceback
        traceback.print_exc()
        return None

    finally:
        # 测试结束后 - 调用cad_zt_zero()
        print("\n[清理] 恢复CAD状态...")
        try:
            cad_zt_zero()
            print("[成功] CAD状态已恢复")
        except:
            pass

if __name__ == "__main__":
    results = test_part9_part10()

    # 判断是否成功
    if results:
        pass_count = sum(1 for r in results.values() if r['status'] == 'PASS')
        success = pass_count >= 6  # 至少6个函数通过
        sys.exit(0 if success else 1)
    else:
        sys.exit(1)
