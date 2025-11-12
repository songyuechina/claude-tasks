#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
严格按照规范测试 insert_tarch_window 函数
"""

import sys
import time
from pathlib import Path

# 添加scripts目录到路径
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_basic import (
    li, stc, get_object_property, set_object_property,
    transfer_props_by_matchprop, get_acad_doc, start_applicationV9
)
from CAD_file_operations import copy_file_content_pywin32
from CAD_coordination import ensure_single_process, wait_quiescent, send_cmd_with_sync


def insert_tarch_door_strict(p, width=None, height=None):
    """在墙体上插入天正门（严格版本）"""
    try:
        # 获取插入前的对象数量
        _, doc = get_acad_doc()
        ms = doc.ModelSpace
        count_before = ms.Count
        print(f"[调试] 插入前对象数量: {count_before}")

        # 发送TOpening命令
        cmd = f"TOpening\n{p[0]},{p[1]}\n\n"
        print(f"[调试] 发送命令: TOpening")
        print(f"[调试] 坐标: {p[0]},{p[1]}")
        send_cmd_with_sync(cmd, wait_after=4.0)

        # 等待对象创建
        time.sleep(3)

        # 遍历新增对象
        door = None
        count_after = ms.Count
        print(f"[调试] 插入后对象数量: {count_after}, 新增: {count_after - count_before}")

        for i in range(count_before, count_after):
            obj = ms.Item(i)
            print(f"[调试] 对象 {i}: {obj.ObjectName}")
            if obj.ObjectName == "TDbOpening":
                door = obj
                print(f"[成功] 找到门对象")
                break

        if door is None:
            print(f"[错误] 未找到门对象")
            return {'success': False, 'door': None, 'width': None, 'height': None}

        # 设置尺寸
        if width is not None:
            set_object_property(door, 'Width', width)
        if height is not None:
            set_object_property(door, 'Height', height)

        current_width = get_object_property(door, 'Width')
        current_height = get_object_property(door, 'Height')
        print(f"[成功] 已插入门 - 宽:{current_width}, 高:{current_height}")

        return {
            'success': True,
            'door': door,
            'width': current_width,
            'height': current_height
        }

    except Exception as e:
        print(f"[错误] 插入门失败: {e}")
        import traceback
        traceback.print_exc()
        return {'success': False, 'door': None, 'width': None, 'height': None}


def insert_tarch_window_strict(p, width, height, window_type, delete_mc_yuan=False):
    """在墙体上插入天正窗（严格版本）"""
    try:
        print(f"\n{'='*60}")
        print(f"开始插入窗: 类型={window_type}, 宽={width}, 高={height}")
        print(f"位置: {p}")
        print(f"{'='*60}")

        # 1. 连接当前文件
        print("\n[步骤 1/6] 连接当前文件...")
        li()

        # 2. 检查MC_yuan
        print("\n[步骤 2/6] 检查MC_yuan...")
        lb = stc('Mc_yuan_bj')
        if len(lb) == 0:
            print("[信息] 需要插入MC_yuan.dwg")
            _, doc = get_acad_doc()
            current_file = doc.FullName
            copy_file_content_pywin32("D:/claude-tasks/cad/xitongwenjian/MC_yuan.dwg", current_file)
            print("[成功] 已插入MC_yuan.dwg")
            li()
        else:
            print(f"[信息] MC_yuan已存在 ({len(lb)}个对象)")

        # 3. 插入门
        print(f"\n[步骤 3/6] 插入门...")
        result = insert_tarch_door_strict(p, width=width, height=height)
        if not result['success']:
            print("[失败] 插入门失败")
            return {'success': False}
        m1 = result['door']

        # 4. 选择窗类型
        print(f"\n[步骤 4/6] 选择窗类型: {window_type}")
        lc = stc(window_type)
        if len(lc) == 0:
            print(f"[错误] 未找到窗类型: {window_type}")
            return {'success': False}

        window_src = lc[0]
        set_object_property(window_src, "Width", width)
        set_object_property(window_src, "Height", height)
        print(f"[成功] 已设置窗元尺寸")

        # 5. 属性匹配
        print(f"\n[步骤 5/6] 属性匹配...")
        success = False
        for attempt in range(1, 6):
            try:
                result_match = transfer_props_by_matchprop(window_src, m1, max_try=3, delay=0.4)
                if result_match:
                    m1_layer = m1.Layer
                    if m1_layer == window_type:
                        print(f"[成功] 第{attempt}次匹配成功")
                        success = True
                        break
                    else:
                        print(f"[警告] 第{attempt}次匹配后图层不正确: {m1_layer}")
            except Exception as e:
                print(f"[警告] 第{attempt}次匹配失败: {e}")
            time.sleep(0.5)

        if not success:
            print("[失败] 属性匹配失败")
            return {'success': False}

        # 6. 删除MC_yuan（可选）
        if delete_mc_yuan:
            print(f"\n[步骤 6/6] 删除MC_yuan对象...")
            deleted_count = 0
            for obj in stc("MC_yuan_qiang"):
                try:
                    obj.Delete()
                    deleted_count += 1
                except:
                    pass
            for obj in stc("MC_yuan_bj"):
                try:
                    obj.Delete()
                    deleted_count += 1
                except:
                    pass
            print(f"[成功] 已删除 {deleted_count} 个对象")

        print(f"\n[成功] 窗插入完成!")
        return {'success': True, 'window': m1, 'width': width, 'height': height}

    except Exception as e:
        print(f"\n[错误] 插入窗失败: {e}")
        import traceback
        traceback.print_exc()
        return {'success': False}


def main():
    """主测试函数 - 严格遵守规范"""
    print("="*70)
    print("insert_tarch_window 函数测试 - 严格遵守规范版本")
    print("="*70)

    # 步骤1: 关闭所有CAD进程
    print("\n[步骤 1/6] 关闭所有CAD进程...")
    ensure_single_process()
    print("[完成] CAD进程已清理")

    # 步骤2: 启动天正CAD（界面可见）
    print("\n[步骤 2/6] 启动天正CAD...")
    print("使用: start_applicationV9(PTH=r'C:\\Tangent\\TArchT20V9')")
    start_applicationV9(
        PTH=r"C:\Tangent\TArchT20V9",
        max_retries=3,
        retry_delay=2.0
    )
    wait_quiescent(min_quiet=1.0, timeout=30.0)
    print("[完成] 天正CAD已启动")

    # 步骤3: 打开测试文件（使用最简单的方式）
    print("\n[步骤 3/6] 打开测试文件...")
    test_file = "D:/claude-tasks/cad/Function_testing/insert_tarch_window-1.dwg"

    # 使用COM直接打开文件
    import win32com.client
    acad = win32com.client.GetActiveObject("AutoCAD.Application")

    # 关闭所有文档
    while acad.Documents.Count > 0:
        doc = acad.Documents.Item(0)
        doc.Close(False)
    print("[信息] 已关闭所有文档")

    # 打开测试文件
    acad.Documents.Open(test_file)
    wait_quiescent(min_quiet=1.0, timeout=15.0)
    print(f"[完成] 已打开: {test_file}")

    # 步骤4: 连接当前文件
    print("\n[步骤 4/6] 使用 li() 连接当前激活文件...")
    li()
    print("[完成] 已连接")

    # 步骤5: 执行测试
    print("\n[步骤 5/6] 执行测试...")

    # 测试1
    print("\n" + "="*70)
    print("测试1: 插入 jz-gaochuang 窗")
    print("="*70)
    result1 = insert_tarch_window_strict(
        p=(38612.86565445, 48750.63891910, 0),
        width=1200,
        height=1000,
        window_type="jz-gaochuang",
        delete_mc_yuan=False
    )

    if result1['success']:
        print("\n[OK] 测试1成功")
    else:
        print("\n[FAIL] 测试1失败")
        return False

    time.sleep(2)

    # 测试2
    print("\n" + "="*70)
    print("测试2: 插入 jz-pingchuang 窗")
    print("="*70)
    result2 = insert_tarch_window_strict(
        p=(44695.30568975, 46646.78059028, 0),
        width=2400,
        height=1800,
        window_type="jz-pingchuang",
        delete_mc_yuan=True
    )

    if result2['success']:
        print("\n[OK] 测试2成功")
    else:
        print("\n[FAIL] 测试2失败")
        return False

    # 步骤6: 保存文件
    print("\n[步骤 6/6] 保存文件...")
    acad.ActiveDocument.Save()
    wait_quiescent(min_quiet=1.0, timeout=10.0)
    print("[完成] 文件已保存")

    print("\n" + "="*70)
    print("✓ 所有测试完成!")
    print("="*70)

    return True


if __name__ == "__main__":
    try:
        success = main()
        if not success:
            print("\n测试失败")
            sys.exit(1)
    except Exception as e:
        print(f"\n测试过程中发生错误: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
