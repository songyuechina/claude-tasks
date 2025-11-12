#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
使用LEADER和MTEXT命令为私宅总图添加坐标标注
采用命令行方式，避免COM自动化问题
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

# 第一张图的地理坐标
original_coords_1 = [
    (3084591.190, 380592.128),  # 点1
    (3084586.139, 380581.243),  # 点2
    (3084577.068, 380585.452),  # 点3
    (3084582.119, 380596.337),  # 点4
]

def add_leader_labels():
    """使用LEADER命令为两张图添加坐标标注"""
    print("=" * 60)
    print("使用LEADER命令添加坐标标注")
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
        wait_quiescent(min_quiet=2.0, timeout=30.0)

        # 处理第一张图
        print("\n" + "=" * 60)
        print("[处理] 第一张图: 私宅总图.dwg")
        print("=" * 60)

        dwg_file_1 = str(Path(__file__).parent / "私宅总图.dwg")
        print(f"\n[步骤2] 打开文件: {dwg_file_1}")

        if not open_dwg_paradigm(dwg_file_1):
            print("[错误] 打开文件失败")
            return False

        print("[成功] 文件已打开")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        print("\n[步骤3] 为第一张图添加标注...")
        for i, (geo_x, geo_y) in enumerate(original_coords_1, 1):
            # 转换为CAD坐标 (X->Y, Y->X)
            cad_x = geo_y
            cad_y = geo_x

            # 创建标注文字
            label_text = f"X={geo_x:.3f},Y={geo_y:.3f}"

            # 计算引线位置(在点的右上方约14个单位)
            # 引线起点：被标注的点
            # 引线终点：右上45度方向约14个单位
            import math
            distance = 14.0
            angle = math.radians(45)  # 45度
            text_x = cad_x + distance * math.cos(angle)
            text_y = cad_y + distance * math.sin(angle)

            # 使用LEADER命令创建引线
            # _LEADER: 引线命令
            # 起点坐标 -> 终点坐标 -> 回车结束点选择 -> 输入标注文字 -> 回车结束
            leader_cmd = f"_LEADER\n{cad_x},{cad_y}\n{text_x},{text_y}\n\n{label_text}\n\n"

            if not send_cmd_with_sync(leader_cmd, wait_after=1.0):
                print(f"[警告] 点{i}标注添加失败")
            else:
                print(f"[成功] 点{i}标注已添加: {label_text}")

            time.sleep(0.5)

        wait_quiescent(min_quiet=1.0, timeout=15.0)

        print("\n[步骤4] 保存文件...")
        if not save_current_dwg_paradigm():
            print("[错误] 保存文件失败")
            return False
        print("[成功] 文件已保存")

        print("\n[清理] 关闭文件...")
        close_current_dwg_paradigm("no_save")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 处理第二张图
        print("\n" + "=" * 60)
        print("[处理] 第二张图: 私宅总图_第二张图.dwg")
        print("=" * 60)

        # 第二张图的地理坐标
        original_coords_2 = [
            (3083641.673, 381378.081),  # P1
            (3083634.821, 381387.414),  # P2
            (3083644.548, 381393.696),  # P3
            (3083645.613, 381393.914),  # P4
            (3083648.003, 381392.157),  # P5
            (3083651.207, 381388.858),  # P6
        ]

        dwg_file_2 = str(Path(__file__).parent / "私宅总图_第二张图.dwg")
        print(f"\n[步骤5] 打开文件: {dwg_file_2}")

        if not open_dwg_paradigm(dwg_file_2):
            print("[错误] 打开文件失败")
            return False

        print("[成功] 文件已打开")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        print("\n[步骤6] 为第二张图添加标注...")
        for i, (geo_x, geo_y) in enumerate(original_coords_2, 1):
            # 转换为CAD坐标 (X->Y, Y->X)
            cad_x = geo_y
            cad_y = geo_x

            # 创建标注文字
            label_text = f"X={geo_x:.3f},Y={geo_y:.3f}"

            # 计算引线位置(在点的右上方约14个单位)
            import math
            distance = 14.0
            angle = math.radians(45)  # 45度
            text_x = cad_x + distance * math.cos(angle)
            text_y = cad_y + distance * math.sin(angle)

            # 使用LEADER命令创建引线
            leader_cmd = f"_LEADER\n{cad_x},{cad_y}\n{text_x},{text_y}\n\n{label_text}\n\n"

            if not send_cmd_with_sync(leader_cmd, wait_after=1.0):
                print(f"[警告] P{i}标注添加失败")
            else:
                print(f"[成功] P{i}标注已添加: {label_text}")

            time.sleep(0.5)

        wait_quiescent(min_quiet=1.0, timeout=15.0)

        print("\n[步骤7] 保存文件...")
        if not save_current_dwg_paradigm():
            print("[错误] 保存文件失败")
            return False
        print("[成功] 文件已保存")

        print("\n[清理] 关闭文件...")
        close_current_dwg_paradigm("no_save")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        print("\n" + "=" * 60)
        print("[完成] 所有坐标标注添加成功!")
        print("=" * 60)
        print("\n标注特点:")
        print("- 使用引线(Leader)连接点和标注")
        print("- 文字高度: 默认样式")
        print("- 引线距点约14个单位")
        print("- 标注方向: 右上45度")

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
    success = add_leader_labels()
    sys.exit(0 if success else 1)
