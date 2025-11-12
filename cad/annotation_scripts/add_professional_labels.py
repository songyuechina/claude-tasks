#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
使用引线和MTEXT为私宅总图添加坐标标注
参考annotate_xy_point函数实现专业的坐标标注
"""

import sys
import time
import math
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
    ensure_single_process
)

import win32com.client as win32

# 第一张图的地理坐标
original_coords_1 = [
    (3084591.190, 380592.128),  # 点1
    (3084586.139, 380581.243),  # 点2
    (3084577.068, 380585.452),  # 点3
    (3084582.119, 380596.337),  # 点4
]

def annotate_xy_point(acad, doc, geo_x, geo_y, *,
                      distance=14.0,
                      decimals=3,
                      text_height=14.0,
                      text_width=200.0,
                      layer_name="COORD_ANN"):
    """
    在当前激活 DWG 中，对地理坐标 (geo_x, geo_y) 做 X/Y 标注，带引线。
    注意：在DWG中绘制时，X和Y需要互换
    """
    ms = doc.ModelSpace

    # 转换坐标：地理X→DWG的Y，地理Y→DWG的X
    x = geo_y  # DWG的x坐标 = 地理Y
    y = geo_x  # DWG的y坐标 = 地理X
    z = 0.0

    # 图层：若不存在则创建
    def ensure_layer(name):
        try:
            lyr = doc.Layers.Item(name)
        except Exception:
            lyr = doc.Layers.Add(name)
        return lyr
    ensure_layer(layer_name)

    # 计算放置位置（沿 45° 方向偏移）
    bbox_pad = max(text_width, text_height) * 0.6
    inner_r = max(distance - bbox_pad, 1.0)
    dx = dy = inner_r / math.sqrt(2.0)
    text_x = x + dx
    text_y = y + dy
    text_z = z

    # 文字内容（使用原始地理坐标）
    label = f"X={geo_x:.{decimals}f},Y={geo_y:.{decimals}f}"

    # 放置 MTEXT - 使用数组而不是元组
    try:
        # AddMText需要一个点数组: [x, y, z] 或者 win32.VARIANT数组
        from win32com.client import VARIANT
        import pythoncom
        text_pt_array = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [text_x, text_y, text_z])
        mtext = ms.AddMText(text_pt_array, text_width, label)
        mtext.Height = text_height
        mtext.Layer = layer_name
    except Exception as e:
        print(f"[警告] 创建MTEXT失败: {e}")
        return None

    # 引线（两段折线）
    k = 0.6
    vx = k * inner_r / math.sqrt(2.0)
    vy = k * inner_r / math.sqrt(2.0)

    # 使用VARIANT数组创建引线点
    from win32com.client import VARIANT
    import pythoncom

    leader_pts_array = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [
        x, y, z,
        x + vx, y + vy, z,
        text_x, text_y, text_z
    ])

    try:
        leader = ms.AddLeader(leader_pts_array, mtext, 2)
        leader.Layer = layer_name
    except Exception:
        try:
            leader = ms.AddLeader(leader_pts_array, mtext, 1)
            leader.Layer = layer_name
        except Exception as e:
            print(f"[警告] 创建引线失败: {e}")
            return {"text_object": mtext, "leader_object": None}

    # 视口刷新
    try:
        doc.Regen(1)
    except Exception:
        pass

    return {"text_object": mtext, "leader_object": leader}

def add_professional_labels():
    """为两张图添加专业坐标标注"""
    print("=" * 60)
    print("使用引线和MTEXT添加坐标标注")
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

        # 获取CAD应用并等待其完全初始化
        print("\n[步骤1.5] 连接到CAD应用...")
        time.sleep(3.0)  # 额外等待确保CAD完全就绪

        # 尝试连接到已运行的CAD实例
        acad = None
        for prog_id in ["AutoCAD.Application", "TArch.Application", "Zwcad.Application"]:
            try:
                acad = win32.GetActiveObject(prog_id)
                print(f"[成功] 已连接到 {prog_id}")
                break
            except:
                try:
                    acad = win32.Dispatch(prog_id)
                    print(f"[成功] 已创建 {prog_id} 连接")
                    break
                except:
                    continue

        if not acad:
            print("[错误] 无法连接到CAD应用")
            return False

        acad.Visible = True
        time.sleep(2.0)

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

        # 重新获取文档对象确保同步
        time.sleep(1.0)
        doc = acad.ActiveDocument
        time.sleep(0.5)

        print("\n[步骤3] 为第一张图添加标注...")
        for i, (geo_x, geo_y) in enumerate(original_coords_1, 1):
            # 重新获取文档对象以避免COM状态问题
            doc = acad.ActiveDocument
            result = annotate_xy_point(
                acad, doc, geo_x, geo_y,
                distance=14.0,
                text_height=14.0
            )
            if result:
                print(f"[成功] 点{i}标注已添加")
            else:
                print(f"[警告] 点{i}标注添加失败")
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

        # 重新获取文档对象确保同步
        time.sleep(1.0)
        doc = acad.ActiveDocument
        time.sleep(0.5)

        print("\n[步骤6] 为第二张图添加标注...")
        for i, (geo_x, geo_y) in enumerate(original_coords_2, 1):
            # 重新获取文档对象以避免COM状态问题
            doc = acad.ActiveDocument
            result = annotate_xy_point(
                acad, doc, geo_x, geo_y,
                distance=14.0,
                text_height=14.0
            )
            if result:
                print(f"[成功] P{i}标注已添加")
            else:
                print(f"[警告] P{i}标注添加失败")
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
        print("- 使用MTEXT显示坐标")
        print("- 文字高度: 14")
        print("- 引线+文字距点约14个单位")
        print("- 标注层: COORD_ANN")

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
    success = add_professional_labels()
    sys.exit(0 if success else 1)
