#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试：绘制矩形墙体并在每堵墙中间插入平窗
创建日期: 2025-11-16
"""

import sys
import time
from pathlib import Path

# 添加scripts目录到路径
SCRIPT_DIR = Path(__file__).parent.parent / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from CAD_file_operations import (
    new_file, save_file_as, cad_zt_zero,
    draw_tarch_wall, insert_tarch_window
)

def test_rectangle_walls_with_windows():
    """测试：绘制矩形墙体并在每堵墙中间插入平窗"""

    output_file = "D:/claude-tasks/cad/tests/test_files/rectangle_walls_windows.dwg"

    print("="*60)
    print("测试：绘制矩形墙体并在每堵墙中间插入平窗")
    print("="*60)

    try:
        # 定义矩形的四个角点（5000x4000的矩形）
        p1 = (0, 0, 0)
        p2 = (5000, 0, 0)
        p3 = (5000, 4000, 0)
        p4 = (0, 4000, 0)

        # 墙体厚度
        wall_thickness = 240

        # 窗户参数
        window_type = "jz-pingchuang"  # 建筑-平窗
        window_width = 1200
        window_height = 1500

        print(f"\n[1] 新建文件")
        start_time = time.time()
        new_file()
        print(f"[OK] 新建文件成功 (耗时: {time.time() - start_time:.2f}秒)")

        # 绘制矩形的四堵墙
        walls_info = [
            {"name": "南墙", "p1": p1, "p2": p2},  # 底边
            {"name": "东墙", "p1": p2, "p2": p3},  # 右边
            {"name": "北墙", "p1": p3, "p2": p4},  # 上边
            {"name": "西墙", "p1": p4, "p2": p1},  # 左边
        ]

        print(f"\n[2] 绘制矩形墙体（4堵墙）")
        for i, wall in enumerate(walls_info, 1):
            print(f"\n  [{i}] 绘制{wall['name']}：{wall['p1']} -> {wall['p2']}")
            start_time = time.time()
            draw_tarch_wall(wall['p1'], wall['p2'], wall_thickness)
            elapsed = time.time() - start_time
            print(f"  [OK] {wall['name']}绘制完成 (耗时: {elapsed:.2f}秒)")
            time.sleep(1)  # 墙体间间隔

        print(f"\n[3] 在每堵墙中间插入平窗")

        # 计算每堵墙的中点作为窗户位置
        windows_info = [
            {
                "name": "南墙窗",
                "wall_p1": p1,
                "wall_p2": p2,
                "insert_point": ((p1[0] + p2[0])/2, p1[1], 0)
            },
            {
                "name": "东墙窗",
                "wall_p1": p2,
                "wall_p2": p3,
                "insert_point": (p2[0], (p2[1] + p3[1])/2, 0)
            },
            {
                "name": "北墙窗",
                "wall_p1": p3,
                "wall_p2": p4,
                "insert_point": ((p3[0] + p4[0])/2, p3[1], 0)
            },
            {
                "name": "西墙窗",
                "wall_p1": p4,
                "wall_p2": p1,
                "insert_point": (p4[0], (p4[1] + p1[1])/2, 0)
            },
        ]

        for i, window in enumerate(windows_info, 1):
            print(f"\n  [{i}] 在{window['name']}插入窗户")
            print(f"      墙体: {window['wall_p1']} -> {window['wall_p2']}")
            print(f"      插入点: {window['insert_point']}")
            print(f"      类型: {window_type}, 宽度: {window_width}, 高度: {window_height}")

            start_time = time.time()
            insert_tarch_window(
                window['insert_point'],  # 插入点
                window_width,  # 宽度
                window_height,  # 高度
                window_type  # 窗户类型
            )
            elapsed = time.time() - start_time
            print(f"  [OK] {window['name']}插入完成 (耗时: {elapsed:.2f}秒)")
            time.sleep(2)  # 窗户间间隔

        print(f"\n[4] 保存文件: {output_file}")
        start_time = time.time()
        save_file_as(output_file)
        print(f"[OK] 文件保存成功 (耗时: {time.time() - start_time:.2f}秒)")

        # 验证文件
        if Path(output_file).exists():
            file_size = Path(output_file).stat().st_size
            print(f"\n[OK] 文件已创建: {output_file}")
            print(f"   文件大小: {file_size:,} 字节")
        else:
            print(f"\n[ERROR] 文件未创建: {output_file}")
            return False

        print("\n" + "="*60)
        print("[OK] 测试成功完成")
        print("="*60)
        return True

    except Exception as e:
        print(f"\n[ERROR] 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        # 第2次清理：测试结束后
        print("\n[清理] 调用 cad_zt_zero() 清理CAD状态...")
        try:
            cad_zt_zero()
            print("[OK] CAD状态清理完成")
        except Exception as e:
            print(f"[WARN] CAD清理异常: {e}")

if __name__ == "__main__":
    # 第1次清理：测试开始前
    print("[清理] 测试开始前调用 cad_zt_zero() 清理CAD状态...")
    try:
        from CAD_file_operations import cad_zt_zero
        cad_zt_zero()
        print("[OK] CAD状态清理完成\n")
    except Exception as e:
        print(f"[WARN] CAD清理异常: {e}\n")

    # 执行测试
    success = test_rectangle_walls_with_windows()

    # 退出
    sys.exit(0 if success else 1)
