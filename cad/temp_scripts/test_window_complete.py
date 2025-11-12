#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""完整的天正窗测试脚本"""
import sys, time
from pathlib import Path
sys.path.append(str(Path(__file__).parent))

from CAD_file_operations import (start_cad_session, new_file, save_file_as,
                                  draw_tarch_wall, copy_file_content_pywin32,
                                  save_file, restore_to_uncertain_state)
from CAD_coordination import wait_quiescent, send_cmd_with_sync
from CAD_basic import stc, transfer_props_by_matchprop, get_object_property, set_object_property
import win32com.client

print("="*60)
print("完整的天正窗测试")
print("="*60)

# 1. 启动天正CAD
print("\n[1/8] 启动天正CAD...")
start_cad_session()

# 2. 新建文件
print("\n[2/8] 新建文件...")
new_file()
wait_quiescent(min_quiet=1.0, timeout=10.0)

# 3. 绘制三角形
print("\n[3/8] 绘制三角形...")
p1, p2, p3 = (0, 0, 0), (15000, 0, 0), (7500, 13000, 0)
send_cmd_with_sync(f"_PLINE\n{p1[0]},{p1[1]}\n{p2[0]},{p2[1]}\n{p3[0]},{p3[1]}\nC\n", wait_after=1.0)
wait_quiescent(min_quiet=1.0, timeout=10.0)
print("[OK] 三角形已绘制")

# 4. 沿三角形画天正墙
print("\n[4/8] 沿三角形画天正墙...")
draw_tarch_wall(p1, p2)
draw_tarch_wall(p2, p3)
draw_tarch_wall(p3, p1)
wait_quiescent(min_quiet=1.0, timeout=10.0)
print("[OK] 天正墙已绘制")

# 5. 删除三角形
print("\n[5/8] 删除三角形...")
send_cmd_with_sync("_ERASE\nL\n\n", wait_after=1.0)
wait_quiescent(min_quiet=1.0, timeout=10.0)
print("[OK] 三角形已删除")

# 6. 保存文件
print("\n[6/8] 保存文件...")
save_path = "D:/claude-tasks/tests/test_files/天正窗测试文件.dwg"
save_file_as(save_path)

# 7. 插入MC_yuan.dwg
print("\n[7/8] 插入MC_yuan.dwg...")
copy_file_content_pywin32('D:/claude-tasks/cad/xitongwenjian/MC_yuan.dwg', save_path)
wait_quiescent(min_quiet=1.0, timeout=10.0)
print("[OK] MC_yuan.dwg已插入")

# 8. 插入3个窗（带重试和详细日志）
print("\n[8/8] 插入3个窗...")

def insert_window_with_retry(pos, wtype, width, max_tries=5):
    """带重试的窗插入"""
    acad = win32com.client.GetActiveObject("AutoCAD.Application")
    ms = acad.ActiveDocument.ModelSpace

    for attempt in range(1, max_tries + 1):
        print(f"\n  尝试 {attempt}/{max_tries} - {wtype} @ {pos}")
        try:
            # 插入门
            send_cmd_with_sync(f"TOpening\n{pos[0]},{pos[1]}\n\n", wait_after=2.0)
            wait_quiescent(min_quiet=1.0, timeout=10.0)
            time.sleep(1.0)

            # 获取门对象
            door = None
            for i in range(ms.Count-1, max(0, ms.Count-10), -1):
                try:
                    obj = ms.Item(i)
                    if obj.ObjectName == "TDbOpening":
                        door = obj
                        print(f"  找到门: {obj.Handle}")
                        break
                except:
                    pass

            if not door:
                print(f"  未找到门对象")
                send_cmd_with_sync("u\n", wait_after=1.0)
                continue

            # 获取窗对象
            window = stc(wtype)
            if isinstance(window, list):
                if len(window) == 0:
                    print(f"  未找到 {wtype} 窗对象")
                    send_cmd_with_sync("u\n", wait_after=1.0)
                    continue
                window = window[0]

            print(f"  找到窗: {window.Layer}")

            # 传递属性
            success = transfer_props_by_matchprop(window, door)
            if not success:
                print(f"  属性传递失败")
                send_cmd_with_sync("u\n", wait_after=1.0)
                wait_quiescent(min_quiet=1.0, timeout=10.0)
                continue

            print(f"  属性已传递，Layer={door.Layer}")

            # 设置宽度
            set_object_property(door, 'Width', width)
            actual_width = get_object_property(door, 'Width')
            print(f"  [OK] 成功！宽度={actual_width}")

            return {'success': True, 'window': door, 'width': actual_width}

        except Exception as e:
            print(f"  错误: {e}")
            try:
                send_cmd_with_sync("u\n", wait_after=1.0)
                wait_quiescent(min_quiet=1.0, timeout=10.0)
            except:
                pass
            time.sleep(2)

    print(f"  [FAIL] 失败")
    return {'success': False}

# 在墙的1/3位置插入窗
r1 = insert_window_with_retry((5000, 0, 0), "jz-tuilamen", 1200)
r2 = insert_window_with_retry((12500, 4333, 0), "jz-gaochuang", 900)
r3 = insert_window_with_retry((5000, 8667, 0), "jz-shuangmen", 1400)

# 保存
save_file()

print("\n" + "="*60)
print("测试完成！")
print(f"窗1 (jz-tuilamen, 1200): {r1['success']}")
print(f"窗2 (jz-gaochuang, 900): {r2['success']}")
print(f"窗3 (jz-shuangmen, 1400): {r3['success']}")
print("="*60)

# 恢复状态
restore_to_uncertain_state()
