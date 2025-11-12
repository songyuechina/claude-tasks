#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""修正后的窗测试 - 遵守测试规范"""
import sys
import subprocess
import time
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_basic import close_all_cad_processes, start_applicationV9
from CAD_file_operations import new_file, draw_tarch_wall, insert_tarch_window, save_file
from CAD_coordination import ensure_single_process, wait_quiescent

print("="*60)
print("修正后的天正窗测试")
print("="*60)

# 步骤1: 启动cad_dialog_killer.py
print("\n步骤1: 启动cad_dialog_killer.py...")
dialog_killer = subprocess.Popen([sys.executable, "D:/claude-tasks/cad/scripts/cad_dialog_killer.py"])
time.sleep(2)
print("[成功] cad_dialog_killer.py已启动")

try:
    # 步骤2: 启动天正CAD
    print("\n步骤2: 启动天正CAD...")
    close_all_cad_processes()
    proc = start_applicationV9(PTH=r"C:\Tangent\TArchT20V9")
    ensure_single_process()
    wait_quiescent(min_quiet=2.0, timeout=30.0)
    print("[成功] 天正CAD已启动")

    # 步骤3: 新建天正测试文件2.dwg
    print("\n步骤3: 新建天正测试文件2.dwg...")
    new_file("D:/claude-tasks/tests/test_files/天正测试文件2.dwg")
    wait_quiescent(min_quiet=1.0, timeout=15.0)
    print("[成功] 文件已创建")

    # 步骤4: 绘制三角形墙体（厚度200）
    print("\n步骤4: 绘制三角形墙体...")
    walls = [
        ((100000, 75000, 0), (105000, 75000, 0), "底边"),
        ((105000, 75000, 0), (102500, 79330, 0), "右边"),
        ((102500, 79330, 0), (100000, 75000, 0), "左边"),
    ]

    for p1, p2, name in walls:
        print(f"  绘制{name}墙体...")
        draw_tarch_wall(p1, p2, thickness=200)
        wait_quiescent(min_quiet=1.0, timeout=10.0)
    print("[成功] 三角形墙体已绘制")

    # 步骤5: 在每堵墙中间插入窗
    print("\n步骤5: 插入窗...")
    windows = [
        ((102500, 75000, 0), 900, 3, "底边"),
        ((103750, 77165, 0), 1200, 7, "右边"),
        ((101250, 77165, 0), 2000, 9, "左边"),
    ]

    for p, width, wtype, name in windows:
        print(f"  在{name}插入窗 (宽度{width}, 类型{wtype})...")
        result = insert_tarch_window(p, width, wtype)
        if result['success']:
            print(f"  [成功] {name}窗已插入")
        else:
            print(f"  [失败] {name}窗插入失败")
        wait_quiescent(min_quiet=2.0, timeout=15.0)

    # 步骤6: 保存文件
    print("\n步骤6: 保存文件...")
    save_file()
    print("[成功] 文件已保存")

    print("\n" + "="*60)
    print("测试完成！")
    print("="*60)

except Exception as e:
    print(f"\n[错误] {e}")
    import traceback
    traceback.print_exc()

finally:
    # 关闭CAD
    close_all_cad_processes()
    # 终止dialog_killer
    dialog_killer.terminate()
    print("\n清理完成")
