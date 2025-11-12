#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
为第二张图的私宅总图添加地理坐标标注
在每个点旁边标注原始地理坐标(X,Y)
"""

import sys
import time
from pathlib import Path

# 添加脚本路径
SCRIPT_DIR = Path(__file__).parent / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from CAD_basic import start_applicationV9, close_all_cad_processes
from CAD_basic_operations import (
    open_dwg_paradigm,
    save_current_dwg_paradigm,
    close_current_dwg_paradigm
)
from CAD_coordination import (
    wait_quiescent,
    send_cmd_with_sync,
    ensure_single_process
)

# 第二张图的原始地理坐标
original_coords_2 = [
    (3083641.673, 381378.081),  # P1
    (3083634.821, 381387.414),  # P2
    (3083644.548, 381393.696),  # P3
    (3083645.613, 381393.914),  # P4
    (3083648.003, 381392.157),  # P5
    (3083651.207, 381388.858),  # P6
]

def add_labels_to_map2():
    """为第二张图添加坐标标注"""
    print("=" * 60)
    print("为私宅总图(第二张图)添加坐标标注")
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

        # 2. 打开文件
        dwg_file = str(Path(__file__).parent / "私宅总图_第二张图.dwg")
        print(f"\n[步骤2] 打开文件: {dwg_file}")
        if not open_dwg_paradigm(dwg_file):
            print("[错误] 打开文件失败")
            return False

        print("[成功] 文件已打开")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 3. 设置文字样式和高度
        print("\n[步骤3] 设置文字样式...")
        # 设置文字高度为14
        send_cmd_with_sync("_STYLE\n\n\n14\n1\n0\nN\nN\n", wait_after=1.0)
        wait_quiescent(min_quiet=0.5, timeout=15.0)
        print("[成功] 文字高度设置为14")

        # 4. 为每个点添加标注
        print("\n[步骤4] 添加坐标标注...")
        for i, (geo_x, geo_y) in enumerate(original_coords_2, 1):
            # 转换为CAD坐标
            cad_x = geo_y
            cad_y = geo_x

            # 创建标注文字 (格式: X=地理X,Y=地理Y)
            label_text = f"X={geo_x:.3f},Y={geo_y:.3f}"

            # 计算标注位置(在点的右上方偏移)
            label_x = cad_x + 2  # 向右偏移2个单位
            label_y = cad_y + 2  # 向上偏移2个单位

            # 使用TEXT命令添加标注
            text_cmd = f"_TEXT\n{label_x},{label_y}\n14\n0\n{label_text}\n"

            if not send_cmd_with_sync(text_cmd, wait_after=0.5):
                print(f"[警告] P{i}标注添加失败")
            else:
                print(f"[成功] P{i}标注已添加: {label_text}")

        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 5. 缩放到全图
        print("\n[步骤5] 调整视图...")
        send_cmd_with_sync("_ZOOM\n_E\n", wait_after=1.0)
        wait_quiescent(min_quiet=0.5, timeout=15.0)

        # 6. 保存文件
        print("\n[步骤6] 保存文件...")
        if not save_current_dwg_paradigm():
            print("[错误] 保存文件失败")
            return False

        print("[成功] 文件已保存")

        # 7. 关闭文件
        print("\n[清理] 关闭文件...")
        close_current_dwg_paradigm("no_save")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        print("\n" + "=" * 60)
        print("[完成] 坐标标注添加成功!")
        print("=" * 60)
        print(f"文件位置: {dwg_file}")
        print(f"标注点数: {len(original_coords_2)}")
        print("文字高度: 14")

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
    success = add_labels_to_map2()
    sys.exit(0 if success else 1)
