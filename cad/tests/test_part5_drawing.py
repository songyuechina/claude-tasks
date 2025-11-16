#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试CAD_basic.py第五部分（线面分析）的核心绘图函数
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
    draw_point,
    draw_line,
    draw_circle,
    draw_regular_polygon,
    draw_lwpolyline,
    draw_polyline,
    compute_line_angle,
    distance,
    same_point,
    area_of,
    li_new
)

from CAD_coordination import wait_quiescent

def test_part5_drawing():
    """测试第五部分核心绘图函数"""
    print("="*60)
    print("测试CAD_basic.py第五部分（线面分析）核心绘图函数")
    print("="*60)

    test_results = {}

    try:
        # 测试开始前 - 调用cad_zt_zero()
        print("\n[步骤1] 准备CAD环境...")
        cad_zt_zero()
        wait_quiescent(min_quiet=2.0, timeout=30.0)

        # 创建测试文件
        test_file = "D:/claude-tasks/cad/tests/test_files/test_part5_drawing.dwg"
        print(f"\n[步骤2] 创建测试文件: {test_file}")
        new_file(test_file)
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 初始化CAD连接和全局变量
        print("[步骤2.1] 初始化CAD连接...")
        li_new()
        print("[成功] CAD连接已初始化")

        # ===== 测试1: draw_point =====
        print("\n[测试1] draw_point - 绘制点")
        start_time = time.time()
        try:
            pt1 = draw_point((0, 0, 0))
            pt2 = draw_point((1000, 0, 0))
            pt3 = draw_point((0, 1000, 0))
            wait_quiescent(min_quiet=0.5, timeout=10.0)

            if pt1 and pt2 and pt3:
                elapsed = time.time() - start_time
                test_results['draw_point'] = {'status': 'PASS', 'time': f'{elapsed:.2f}秒'}
                print(f"  [成功] 绘制3个点，耗时: {elapsed:.2f}秒")
            else:
                test_results['draw_point'] = {'status': 'FAIL', 'time': '-', 'error': '返回None'}
                print("  [失败] 未能创建点")
        except Exception as e:
            test_results['draw_point'] = {'status': 'FAIL', 'time': '-', 'error': str(e)}
            print(f"  [失败] {e}")

        # ===== 测试2: draw_line =====
        print("\n[测试2] draw_line - 绘制直线")
        start_time = time.time()
        try:
            line1 = draw_line((2000, 0, 0), (3000, 0, 0))
            line2 = draw_line((2000, 0, 0), (2000, 1000, 0))
            line3 = draw_line((2000, 1000, 0), (3000, 1000, 0))
            wait_quiescent(min_quiet=0.5, timeout=10.0)

            if line1 and line2 and line3:
                elapsed = time.time() - start_time
                test_results['draw_line'] = {'status': 'PASS', 'time': f'{elapsed:.2f}秒'}
                print(f"  [成功] 绘制3条直线，耗时: {elapsed:.2f}秒")
            else:
                test_results['draw_line'] = {'status': 'FAIL', 'time': '-', 'error': '返回None'}
                print("  [失败] 未能创建直线")
        except Exception as e:
            test_results['draw_line'] = {'status': 'FAIL', 'time': '-', 'error': str(e)}
            print(f"  [失败] {e}")

        # ===== 测试3: draw_circle =====
        print("\n[测试3] draw_circle - 绘制圆")
        start_time = time.time()
        try:
            circle1 = draw_circle((4000, 500, 0), 500)
            circle2 = draw_circle((5500, 500, 0), 300)
            wait_quiescent(min_quiet=0.5, timeout=10.0)

            if circle1 and circle2:
                elapsed = time.time() - start_time
                test_results['draw_circle'] = {'status': 'PASS', 'time': f'{elapsed:.2f}秒'}
                print(f"  [成功] 绘制2个圆，耗时: {elapsed:.2f}秒")
            else:
                test_results['draw_circle'] = {'status': 'FAIL', 'time': '-', 'error': '返回None'}
                print("  [失败] 未能创建圆")
        except Exception as e:
            test_results['draw_circle'] = {'status': 'FAIL', 'time': '-', 'error': str(e)}
            print(f"  [失败] {e}")

        # ===== 测试4: draw_regular_polygon =====
        print("\n[测试4] draw_regular_polygon - 绘制正多边形")
        start_time = time.time()
        try:
            poly1 = draw_regular_polygon((7000, 500, 0), 500, 6)  # 正六边形
            poly2 = draw_regular_polygon((8500, 500, 0), 400, 5)  # 正五边形
            wait_quiescent(min_quiet=0.5, timeout=10.0)

            if poly1 and poly2:
                elapsed = time.time() - start_time
                test_results['draw_regular_polygon'] = {'status': 'PASS', 'time': f'{elapsed:.2f}秒'}
                print(f"  [成功] 绘制正六边形和正五边形，耗时: {elapsed:.2f}秒")
            else:
                test_results['draw_regular_polygon'] = {'status': 'FAIL', 'time': '-', 'error': '返回None'}
                print("  [失败] 未能创建正多边形")
        except Exception as e:
            test_results['draw_regular_polygon'] = {'status': 'FAIL', 'time': '-', 'error': str(e)}
            print(f"  [失败] {e}")

        # ===== 测试5: draw_polyline =====
        print("\n[测试5] draw_polyline - 绘制多段线")
        start_time = time.time()
        try:
            # 闭合多段线（首尾点相同）
            vertices = [(0, 2000, 0), (1000, 2000, 0), (1000, 3000, 0), (0, 3000, 0), (0, 2000, 0)]
            poly_line = draw_polyline(vertices)
            wait_quiescent(min_quiet=0.5, timeout=10.0)

            if poly_line:
                elapsed = time.time() - start_time
                test_results['draw_polyline'] = {'status': 'PASS', 'time': f'{elapsed:.2f}秒'}
                print(f"  [成功] 绘制闭合多段线，耗时: {elapsed:.2f}秒")
            else:
                test_results['draw_polyline'] = {'status': 'FAIL', 'time': '-', 'error': '返回None'}
                print("  [失败] 未能创建多段线")
        except Exception as e:
            test_results['draw_polyline'] = {'status': 'FAIL', 'time': '-', 'error': str(e)}
            print(f"  [失败] {e}")

        # ===== 测试6: compute_line_angle =====
        print("\n[测试6] compute_line_angle - 计算线段角度")
        start_time = time.time()
        try:
            if line1:
                angle = compute_line_angle(line1)
                elapsed = time.time() - start_time

                if angle is not None:
                    test_results['compute_line_angle'] = {'status': 'PASS', 'time': f'{elapsed:.2f}秒', 'result': f'{angle:.2f}度'}
                    print(f"  [成功] 计算角度: {angle:.2f}度，耗时: {elapsed:.2f}秒")
                else:
                    test_results['compute_line_angle'] = {'status': 'FAIL', 'time': '-', 'error': '返回None'}
                    print("  [失败] 角度计算返回None")
            else:
                test_results['compute_line_angle'] = {'status': 'SKIP', 'time': '-', 'error': 'line1未创建'}
                print("  [跳过] line1未成功创建")
        except Exception as e:
            test_results['compute_line_angle'] = {'status': 'FAIL', 'time': '-', 'error': str(e)}
            print(f"  [失败] {e}")

        # ===== 测试7: distance =====
        print("\n[测试7] distance - 计算点距离")
        start_time = time.time()
        try:
            dist = distance((0, 0, 0), (1000, 0, 0))
            elapsed = time.time() - start_time

            expected = 1000.0
            if abs(dist - expected) < 0.01:
                test_results['distance'] = {'status': 'PASS', 'time': f'{elapsed:.4f}秒', 'result': f'{dist:.2f}'}
                print(f"  [成功] 距离: {dist:.2f}，耗时: {elapsed:.4f}秒")
            else:
                test_results['distance'] = {'status': 'FAIL', 'time': '-', 'error': f'期望{expected}，实际{dist}'}
                print(f"  [失败] 期望{expected}，实际{dist}")
        except Exception as e:
            test_results['distance'] = {'status': 'FAIL', 'time': '-', 'error': str(e)}
            print(f"  [失败] {e}")

        # ===== 测试8: same_point =====
        print("\n[测试8] same_point - 判断点是否相同")
        start_time = time.time()
        try:
            p1 = (1000, 2000, 0)
            p2 = (1000.3, 2000.2, 0)  # 容差内
            p3 = (2000, 2000, 0)  # 容差外

            result1 = same_point(p1, p2, tol=0.5)
            result2 = same_point(p1, p3, tol=0.5)
            elapsed = time.time() - start_time

            if result1 and not result2:
                test_results['same_point'] = {'status': 'PASS', 'time': f'{elapsed:.4f}秒'}
                print(f"  [成功] 容差判断正确，耗时: {elapsed:.4f}秒")
            else:
                test_results['same_point'] = {'status': 'FAIL', 'time': '-', 'error': f'result1={result1}, result2={result2}'}
                print(f"  [失败] 容差判断错误: result1={result1}, result2={result2}")
        except Exception as e:
            test_results['same_point'] = {'status': 'FAIL', 'time': '-', 'error': str(e)}
            print(f"  [失败] {e}")

        # ===== 测试9: area_of =====
        print("\n[测试9] area_of - 计算面积")
        start_time = time.time()
        try:
            # 1000x1000的正方形
            verts = [(0, 0), (1000, 0), (1000, 1000), (0, 1000)]
            area = area_of(verts)
            elapsed = time.time() - start_time

            expected = 1000000.0  # 1000*1000
            if abs(area - expected) < 1.0:
                test_results['area_of'] = {'status': 'PASS', 'time': f'{elapsed:.4f}秒', 'result': f'{area:.0f}'}
                print(f"  [成功] 面积: {area:.0f}，耗时: {elapsed:.4f}秒")
            else:
                test_results['area_of'] = {'status': 'FAIL', 'time': '-', 'error': f'期望{expected}，实际{area}'}
                print(f"  [失败] 期望{expected}，实际{area}")
        except Exception as e:
            test_results['area_of'] = {'status': 'FAIL', 'time': '-', 'error': str(e)}
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
        skip_count = sum(1 for r in test_results.values() if r['status'] == 'SKIP')

        for func_name, result in test_results.items():
            status_icon = '[PASS]' if result['status'] == 'PASS' else ('[FAIL]' if result['status'] == 'FAIL' else '[SKIP]')
            print(f"{status_icon} {func_name}: {result['status']} ({result['time']})")
            if 'error' in result:
                print(f"    错误: {result['error']}")
            if 'result' in result:
                print(f"    结果: {result['result']}")

        print("\n" + "="*60)
        print(f"总计: {len(test_results)}个函数")
        print(f"通过: {pass_count}个")
        print(f"失败: {fail_count}个")
        print(f"跳过: {skip_count}个")
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
    results = test_part5_drawing()

    # 判断是否成功
    if results:
        pass_count = sum(1 for r in results.values() if r['status'] == 'PASS')
        success = pass_count >= 7  # 至少7个函数通过
        sys.exit(0 if success else 1)
    else:
        sys.exit(1)
