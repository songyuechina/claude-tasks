#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
新的天正窗绘制函数
使用 transfer_props_by_matchprop 实现门变窗
"""

import sys
import io
import time
from pathlib import Path

# 配置编码
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
sys.path.append(str(Path(__file__).parent))

from CAD_basic import stc, transfer_props_by_matchprop, get_object_property, set_object_property
from CAD_file_operations import copy_file_content_pywin32
from CAD_coordination import send_cmd_with_sync, wait_quiescent

def insert_tarch_window_new(p, window_type, width, max_retries=3):
    """
    在指定位置插入指定类型和宽度的天正窗

    Args:
        p: 插入点坐标 (x, y, z)
        window_type: 窗类型图层名，可选值：
            "jz-menlianchuang", "jz-dong", "jz-gaochuang",
            "jz-baiyechuang", "jz-tuchuang", "jz-pingchuang",
            "jz-zimumen", "jz-juanlianmen", "jz-tuilamen", "jz-shuangmen"
        width: 窗宽度
        max_retries: 最大重试次数（如果transfer_props_by_matchprop失败）

    Returns:
        dict: {'success': bool, 'window': 窗对象, 'width': 宽度}
    """
    import win32com.client

    # 验证窗类型
    valid_types = [
        "jz-menlianchuang", "jz-dong", "jz-gaochuang",
        "jz-baiyechuang", "jz-tuchuang", "jz-pingchuang",
        "jz-zimumen", "jz-juanlianmen", "jz-tuilamen", "jz-shuangmen"
    ]
    if window_type not in valid_types:
        print(f"[错误] 无效的窗类型: {window_type}")
        return {'success': False, 'window': None, 'width': None}

    acad = win32com.client.GetActiveObject("AutoCAD.Application")
    doc = acad.ActiveDocument
    ms = doc.ModelSpace

    print(f"\n{'='*60}")
    print(f"插入天正窗: 类型={window_type}, 宽度={width}")
    print(f"{'='*60}")

    for attempt in range(1, max_retries + 1):
        try:
            print(f"\n[尝试 {attempt}/{max_retries}]")

            # 步骤1: 插入门
            print(f"[1/5] 在位置 {p} 插入门...")
            cmd = f"TOpening\n{p[0]},{p[1]}\n\n"
            send_cmd_with_sync(cmd, wait_after=2.0)
            wait_quiescent(min_quiet=1.0, timeout=10.0)
            time.sleep(1.0)

            # 获取刚插入的门
            from CAD_basic import last_obj
            door = last_obj()
            if door.ObjectName != "TDbOpening":
                print(f"[错误] 插入的对象不是门: {door.ObjectName}")
                continue

            door_handle = door.Handle
            print(f"✓ 门已插入，Handle={door_handle}")

            # 记录MC_yuan插入前的对象数量
            count_before = ms.Count

            # 步骤2: 插入MC_yuan.dwg
            print(f"[2/5] 插入 MC_yuan.dwg...")
            target_doc_name = doc.FullName
            result = copy_file_content_pywin32(
                'D:/claude-tasks/cad/xitongwenjian/MC_yuan.dwg',
                target_doc_name
            )
            if not result:
                print(f"[错误] MC_yuan.dwg 插入失败")
                # 使用u命令恢复
                send_cmd_with_sync("u\n", wait_after=1.0)
                continue

            # 重新激活目标文档（copy_file_content_pywin32可能改变了激活文档）
            for d in acad.Documents:
                if d.FullName == target_doc_name:
                    acad.ActiveDocument = d
                    doc = d
                    ms = d.ModelSpace
                    break

            wait_quiescent(min_quiet=1.0, timeout=10.0)
            print(f"✓ MC_yuan.dwg 已插入")

            # 步骤3: 获取指定类型的窗对象
            print(f"[3/5] 获取 {window_type} 图层的窗对象...")
            window = stc(window_type)
            if isinstance(window, list):
                window = window[0]
            print(f"✓ 窗对象已获取，Layer={window.Layer}")

            # 步骤4: 使用transfer_props_by_matchprop将门变为窗
            print(f"[4/5] 使用 transfer_props_by_matchprop 将门变为窗...")

            # 通过Handle重新获取门对象（因为可能被更新）
            door = None
            for i in range(ms.Count):
                try:
                    obj = ms.Item(i)
                    if obj.Handle == door_handle:
                        door = obj
                        break
                except:
                    pass

            if door is None:
                print(f"[错误] 无法找到门对象")
                # 使用u命令恢复
                send_cmd_with_sync("u\n", wait_after=1.0)
                continue

            success = transfer_props_by_matchprop(window, door)
            if not success:
                print(f"[错误] transfer_props_by_matchprop 失败")
                # 使用u命令恢复到插入MC_yuan之前
                send_cmd_with_sync("u\n", wait_after=1.0)
                wait_quiescent(min_quiet=1.0, timeout=10.0)
                continue

            print(f"✓ 门已变为窗，Layer={door.Layer}")

            # 步骤5: 设置宽度
            print(f"[5/5] 设置窗宽度为 {width}...")
            set_object_property(door, 'Width', width)
            actual_width = get_object_property(door, 'Width')
            print(f"✓ 窗宽度已设置为 {actual_width}")

            print(f"\n{'='*60}")
            print(f"✓ 成功！窗已插入，类型={door.Layer}, 宽度={actual_width}")
            print(f"{'='*60}\n")

            return {
                'success': True,
                'window': door,
                'width': actual_width,
                'layer': door.Layer
            }

        except Exception as e:
            print(f"[错误] 第 {attempt} 次尝试失败: {e}")
            import traceback
            traceback.print_exc()

            # 使用u命令恢复
            try:
                send_cmd_with_sync("u\n", wait_after=1.0)
                wait_quiescent(min_quiet=1.0, timeout=10.0)
            except:
                pass

            if attempt < max_retries:
                print(f"等待后重试...")
                time.sleep(2.0)

    print(f"\n[失败] 经过 {max_retries} 次尝试仍未成功")
    return {'success': False, 'window': None, 'width': None}

if __name__ == "__main__":
    # 测试函数
    result = insert_tarch_window_new((5000, 0, 0), "jz-pingchuang", 1200)
    print(f"\n测试结果: {result}")
