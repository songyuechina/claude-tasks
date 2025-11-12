#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
根据地理坐标创建第二张图的DWG文件
从2.jpg截图中提取的坐标，进行X/Y转换后绘制
"""

import sys
import time
from pathlib import Path

# 添加脚本路径
SCRIPT_DIR = Path(__file__).parent.parent / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from CAD_basic import start_applicationV9, close_all_cad_processes
from CAD_basic_operations import (
    new_dwg_enhanced,
    save_as_dwg_paradigm,
    close_current_dwg_paradigm
)
from CAD_coordination import (
    wait_quiescent,
    send_cmd_with_sync,
    ensure_single_process
)

# 从图片提取的地理坐标点 (X, Y)
# 注意: 根据要求,在DWG中要把X当Y,Y当X
# 坐标来自即时对话.txt中从2.jpg提取的信息
original_coords = [
    (3083641.673, 381378.081),  # P1
    (3083634.821, 381387.414),  # P2
    (3083644.548, 381393.696),  # P3
    (3083645.613, 381393.914),  # P4
    (3083648.003, 381392.157),  # P5
    (3083651.207, 381388.858),  # P6
]

# 坐标转换: X->Y, Y->X
def convert_coordinates(coords):
    """
    将地理坐标转换为CAD坐标
    X坐标值在dwg中当成y坐标，Y坐标值当成x坐标
    使用绝对坐标
    """
    if not coords:
        return []

    converted = []
    for x, y in coords:
        # X->Y, Y->X (绝对坐标)
        cad_x = y  # Y坐标值当成x坐标
        cad_y = x  # X坐标值当成y坐标
        converted.append((cad_x, cad_y))

    return converted

def draw_plot_map_2():
    """绘制第二张图的总图"""
    print("=" * 60)
    print("创建私宅总图 (第二张图)")
    print("=" * 60)

    try:
        # 1. 启动CAD
        print("\n[步骤1] 启动CAD...")
        proc = start_applicationV9(
            PTH=r"C:\Tangent\TArchT20V9",
            max_retries=3,
            retry_delay=2.0
        )
        if not proc:
            print("[错误] CAD启动失败")
            return False

        print("[成功] CAD启动成功")
        ensure_single_process()
        wait_quiescent(min_quiet=1.0, timeout=30.0)

        # 2. 新建文件
        print("\n[步骤2] 新建DWG文件...")
        if not new_dwg_enhanced():
            print("[错误] 新建文件失败")
            return False

        wait_quiescent(min_quiet=1.0, timeout=15.0)
        print("[成功] 文件创建成功")

        # 3. 转换坐标
        print("\n[步骤3] 转换地理坐标...")
        cad_coords = convert_coordinates(original_coords)

        print("[信息] 原始地理坐标:")
        for i, (x, y) in enumerate(original_coords, 1):
            print(f"  P{i}: X={x}, Y={y}")

        print("[信息] 转换后的CAD坐标(X<->Y, 绝对坐标):")
        for i, (x, y) in enumerate(cad_coords, 1):
            print(f"  P{i}: X={x:.3f}, Y={y:.3f}")

        # 4. 绘制点
        print("\n[步骤4] 绘制坐标点...")
        for i, (x, y) in enumerate(cad_coords, 1):
            # 使用POINT命令绘制点
            point_cmd = f"_POINT\n{x},{y}\n"
            if not send_cmd_with_sync(point_cmd, wait_after=0.5):
                print(f"[警告] P{i}绘制失败")
            else:
                print(f"[成功] P{i}已绘制")

        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 5. 用多段线连接所有点(按顺序P1->P2->P3->P4->P5->P6)
        print("\n[步骤5] 按顺序用直线连接所有点...")

        # 构建PLINE命令
        pline_cmd = "_PLINE\n"
        for x, y in cad_coords:
            pline_cmd += f"{x},{y}\n"
        pline_cmd += "\n"  # 不闭合多段线,只是连接

        if not send_cmd_with_sync(pline_cmd, wait_after=2.0, timeout=30.0):
            print("[错误] 绘制多段线失败")
            return False

        print("[成功] 所有点已按顺序用直线连接")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 6. 缩放到全图
        print("\n[步骤6] 调整视图...")
        send_cmd_with_sync("_ZOOM\n_E\n", wait_after=1.0)
        wait_quiescent(min_quiet=0.5, timeout=15.0)

        # 7. 保存文件
        output_file = str(Path(__file__).parent / "私宅总图_第二张图.dwg")
        print(f"\n[步骤7] 保存文件: {output_file}")

        if not save_as_dwg_paradigm(output_file):
            print("[错误] 保存文件失败")
            return False

        print(f"[成功] 文件已保存: 私宅总图_第二张图.dwg")

        # 8. 验证文件
        if Path(output_file).exists():
            file_size = Path(output_file).stat().st_size
            print(f"[成功] 文件创建成功")
            print(f"        文件大小: {file_size} 字节")
        else:
            print("[错误] 文件未创建")
            return False

        # 9. 清理：关闭文件
        print("\n[清理] 关闭文件...")
        close_current_dwg_paradigm("no_save")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        print("\n" + "=" * 60)
        print("[完成] 私宅总图(第二张图)创建成功!")
        print("=" * 60)
        print(f"文件位置: {output_file}")
        print(f"坐标点数: {len(original_coords)}")
        print("坐标转换: X坐标->Y轴, Y坐标->X轴")
        print("连接方式: 按顺序P1->P2->P3->P4->P5->P6连接")

        return True

    except Exception as e:
        print(f"\n[错误] 执行异常: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        # 最终清理
        print("\n[清理] 关闭所有CAD进程...")
        try:
            close_all_cad_processes()
            print("[成功] 环境已清理")
        except:
            pass

if __name__ == "__main__":
    success = draw_plot_map_2()
    sys.exit(0 if success else 1)
