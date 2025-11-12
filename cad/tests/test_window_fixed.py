#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试 insert_tarch_window 函数 - 修正版
"""

import sys
import time
import logging
from pathlib import Path

# 添加scripts目录到路径
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_basic import (
    li, stc, get_object_property, set_object_property,
    transfer_props_by_matchprop, get_acad_doc, start_applicationV9
)
from CAD_file_operations import copy_file_content_pywin32, open_file, save_file
from CAD_coordination import ensure_single_process, wait_quiescent, send_cmd_with_sync


def insert_tarch_door_fixed(p, width=None, height=None):
    """
    在墙体上插入天正门（修正版）

    Args:
        p: 插入点坐标 (x, y, z)
        width: 门宽（可选）
        height: 门高（可选）

    Returns:
        dict: {'success': bool, 'door': 门对象, 'width': 实际宽度, 'height': 实际高度}
    """
    try:
        # 获取插入前的对象数量
        _, doc = get_acad_doc()
        ms = doc.ModelSpace
        count_before = ms.Count
        print(f"[调试] 插入前对象数量: {count_before}")

        # 发送TOpening命令插入门
        cmd = f"TOpening\n{p[0]},{p[1]}\n\n"
        print(f"[调试] 发送命令: {cmd.strip()}")
        send_cmd_with_sync(cmd, wait_after=4.0)

        # 等待对象创建完成
        time.sleep(3)

        # 遍历新增的对象，找到门对象
        door = None
        count_after = ms.Count
        print(f"[调试] 插入后对象数量: {count_after}, 新增: {count_after - count_before}")

        # 如果没有新增对象，可能是命令执行失败，尝试刷新
        if count_after == count_before:
            print("[调试] 没有新增对象，尝试刷新模型空间...")
            time.sleep(2)
            count_after = ms.Count
            print(f"[调试] 刷新后对象数量: {count_after}, 新增: {count_after - count_before}")

        for i in range(count_before, count_after):
            obj = ms.Item(i)
            print(f"[调试] 对象 {i}: {obj.ObjectName}")
            if obj.ObjectName == "TDbOpening":
                door = obj
                print(f"[调试] 找到门对象: {obj.ObjectName}")
                break

        if door is None:
            print(f"[错误] 未找到门对象，新增了 {count_after - count_before} 个对象")
            return {'success': False, 'door': None, 'width': None, 'height': None}

        # 读取当前尺寸
        current_width = get_object_property(door, 'Width')
        current_height = get_object_property(door, 'Height')
        print(f"[调试] 门当前尺寸: 宽{current_width}, 高{current_height}")

        # 设置尺寸（如果指定）
        if width is not None:
            set_object_property(door, 'Width', width)
            current_width = width

        if height is not None:
            set_object_property(door, 'Height', height)
            current_height = height

        print(f"[成功] 已插入天正门 - 宽度:{current_width}, 高度:{current_height}")

        return {
            'success': True,
            'door': door,
            'width': current_width,
            'height': current_height
        }

    except Exception as e:
        import traceback
        print(f"[错误] 插入门失败: {e}")
        traceback.print_exc()
        return {'success': False, 'door': None, 'width': None, 'height': None}


def insert_tarch_window(p, width=600, height=1000, window_type="jz-pingchuang", delete_mc_yuan=False):
    """
    在墙体上插入天正窗

    Args:
        p: 插入点坐标 (x, y, z)
        width: 窗宽度，默认600
        height: 窗高度，默认1000
        window_type: 窗类型，默认"jz-pingchuang"
        delete_mc_yuan: 是否删除MC_yuan.dwg插入的对象，默认False

    Returns:
        dict: {'success': bool, 'window': 窗对象, 'width': 宽度, 'height': 高度}
    """
    # 配置日志
    log_dir = Path("D:/claude-tasks/cad/logs")
    log_dir.mkdir(exist_ok=True)
    log_file = log_dir / f"insert_tarch_window_{time.strftime('%Y%m%d_%H%M%S')}.log"

    logger = logging.getLogger('insert_tarch_window')
    logger.setLevel(logging.INFO)
    if logger.handlers:
        logger.handlers.clear()
    fh = logging.FileHandler(str(log_file), encoding='utf-8')
    fh.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    fh.setFormatter(formatter)
    logger.addHandler(fh)

    # 允许的窗类型列表
    allowed_types = [
        "jz-menlianchuang", "jz-dong", "jz-gaochuang", "jz-baiyechuang",
        "jz-tuchuang", "jz-pingchuang", "jz-zimumen", "jz-juanlianmen",
        "jz-tuilamen", "jz-shuangmen"
    ]

    try:
        # 1. 检查窗类型
        if window_type not in allowed_types:
            logger.error(f"窗类型错误: {window_type}")
            print(f"[错误] 窗类型错误: {window_type}")
            return {'success': False, 'window': None, 'width': None, 'height': None}
        logger.info(f"窗类型检查通过: {window_type}")
        print(f"[信息] 窗类型检查通过: {window_type}")

        # 2. 连接当前激活文件
        li()
        logger.info("已连接当前激活文件")
        print("[信息] 已连接当前激活文件")

        # 3. 检查是否需要插入MC_yuan.dwg
        lb = stc('Mc_yuan_bj')
        if len(lb) == 0:
            logger.info("未找到Mc_yuan_bj图层，需要插入MC_yuan.dwg")
            print("[信息] 未找到Mc_yuan_bj图层，正在插入MC_yuan.dwg...")
            _, doc = get_acad_doc()
            current_file = doc.FullName
            logger.info(f"当前文件: {current_file}")
            copy_file_content_pywin32("D:/claude-tasks/cad/xitongwenjian/MC_yuan.dwg", current_file)
            logger.info("已插入MC_yuan.dwg")
            print("[成功] 已插入MC_yuan.dwg")
            li()
        else:
            logger.info(f"已存在Mc_yuan_bj图层 (找到{len(lb)}个对象)")
            print(f"[信息] 已存在Mc_yuan_bj图层 (找到{len(lb)}个对象)")

        # 4. 插入门（使用修正版函数）
        print(f"[信息] 正在插入门... 位置:{p}, 宽度:{width}, 高度:{height}")
        result = insert_tarch_door_fixed(p, width=width, height=height)
        if not result['success']:
            logger.error("插入门失败")
            print("[错误] 插入门失败")
            return {'success': False, 'window': None, 'width': None, 'height': None}
        m1 = result['door']
        logger.info(f"已插入门，宽度:{width}, 高度:{height}")
        print(f"[成功] 已插入门，宽度:{width}, 高度:{height}")

        # 5. 选择窗类型图层的窗元并修改尺寸
        print(f"[信息] 正在选择窗类型图层: {window_type}")
        lc = stc(window_type)
        if len(lc) == 0:
            logger.error(f"未找到窗类型图层: {window_type}")
            print(f"[错误] 未找到窗类型图层: {window_type}")
            return {'success': False, 'window': None, 'width': None, 'height': None}

        window_src = lc[0]
        set_object_property(window_src, "Width", width)
        set_object_property(window_src, "Height", height)
        logger.info(f"已设置窗元尺寸: 宽{width}, 高{height}")
        print(f"[成功] 已设置窗元尺寸: 宽{width}, 高{height}")

        # 6. 使用transfer_props_by_matchprop匹配属性
        print("[信息] 正在进行属性匹配...")
        success = False
        for attempt in range(1, 6):
            try:
                result_match = transfer_props_by_matchprop(window_src, m1, max_try=3, delay=0.4)
                if result_match:
                    m1_layer = m1.Layer
                    if m1_layer == window_type:
                        logger.info(f"第{attempt}次匹配成功，门已转换为窗，图层:{m1_layer}")
                        print(f"[成功] 第{attempt}次匹配成功，门已转换为窗，图层:{m1_layer}")
                        success = True
                        break
                    else:
                        logger.warning(f"第{attempt}次匹配后图层不正确: {m1_layer}")
                        print(f"[警告] 第{attempt}次匹配后图层不正确: {m1_layer}")
            except Exception as e:
                logger.warning(f"第{attempt}次匹配失败: {e}")
                print(f"[警告] 第{attempt}次匹配失败: {e}")
            time.sleep(0.5)

        if not success:
            logger.error("属性匹配5次仍然失败")
            print("[错误] 属性匹配5次仍然失败")
            return {'success': False, 'window': None, 'width': None, 'height': None}

        # 7. 可选：删除MC_yuan对象
        if delete_mc_yuan:
            logger.info("正在删除MC_yuan对象...")
            print("[信息] 正在删除MC_yuan对象...")
            try:
                deleted_count = 0
                for obj in stc("MC_yuan_qiang"):
                    try:
                        obj.Delete()
                        deleted_count += 1
                    except Exception as e:
                        logger.warning(f"删除对象失败: {e}")
                for obj in stc("MC_yuan_bj"):
                    try:
                        obj.Delete()
                        deleted_count += 1
                    except Exception as e:
                        logger.warning(f"删除对象失败: {e}")
                logger.info(f"已删除 {deleted_count} 个MC_yuan对象")
                print(f"[成功] 已删除 {deleted_count} 个MC_yuan对象")
            except Exception as e:
                logger.warning(f"删除MC_yuan对象失败: {e}")
                print(f"[警告] 删除MC_yuan对象失败: {e}")

        logger.info(f"成功插入天正窗 - 位置:{p}, 宽度:{width}, 高度:{height}, 类型:{window_type}")
        print(f"[成功] 已插入天正窗 - 宽度:{width}, 高度:{height}, 类型:{window_type}")
        return {'success': True, 'window': m1, 'width': width, 'height': height}

    except Exception as e:
        logger.error(f"插入窗失败: {e}", exc_info=True)
        print(f"[错误] 插入窗失败: {e}")
        import traceback
        traceback.print_exc()
        return {'success': False, 'window': None, 'width': None, 'height': None}


def main():
    """主测试函数"""
    print("="*60)
    print("insert_tarch_window 函数测试 - 修正版")
    print("="*60)

    # 1. 确保单进程
    print("\n[步骤 1/7] 关闭所有CAD进程...")
    ensure_single_process()
    print("[完成] CAD进程已清理")

    # 2. 启动天正CAD
    print("\n[步骤 2/7] 启动天正CAD...")
    start_applicationV9(
        PTH=r"C:\Tangent\TArchT20V9",
        max_retries=3,
        retry_delay=2.0
    )
    wait_quiescent(min_quiet=1.0, timeout=30.0)
    print("[完成] 天正CAD已启动")

    # 关闭所有已打开的文档
    print("\n[步骤 2.5/7] 关闭所有已打开的文档...")
    import win32com.client
    acad = win32com.client.GetActiveObject("AutoCAD.Application")
    while acad.Documents.Count > 0:
        try:
            doc = acad.Documents.Item(0)
            doc_name = doc.Name
            doc.Close(False)  # False = 不保存
            print(f"[信息] 已关闭文档: {doc_name}")
        except:
            break
    print("[完成] 所有文档已关闭")

    # 3. 打开测试文件
    print("\n[步骤 3/7] 打开测试文件 insert_tarch_window-2.dwg...")
    test_file = "D:/claude-tasks/cad/Function_testing/insert_tarch_window-2.dwg"
    open_file(test_file)
    wait_quiescent(min_quiet=1.0, timeout=15.0)
    print("[完成] 测试文件已打开")

    # 激活测试文件
    print("[信息] 激活测试文件...")
    import win32com.client
    acad = win32com.client.GetActiveObject("AutoCAD.Application")
    target_path_lower = str(Path(test_file).resolve()).lower()
    print(f"[调试] 目标文件路径: {target_path_lower}")
    print(f"[调试] 当前打开的文档数量: {acad.Documents.Count}")

    found = False
    for doc in acad.Documents:
        doc_path_lower = str(Path(doc.FullName).resolve()).lower()
        print(f"[调试] 文档: {doc.Name}, 路径: {doc_path_lower}")
        if doc_path_lower == target_path_lower:
            acad.ActiveDocument = doc
            print(f"[成功] 已激活文件: {doc.Name}")
            found = True
            break

    if not found:
        print("[警告] 未找到匹配的文档，使用当前激活的文档")
        print(f"[信息] 当前激活文档: {acad.ActiveDocument.Name}")

    # 4. 连接当前文件
    print("\n[步骤 4/7] 使用 li() 连接当前激活文件...")
    li()
    print("[完成] 已连接到当前激活文件")

    # 5. 测试1
    print("\n[步骤 5/7] 测试1: 插入 jz-gaochuang 窗...")
    print("  位置: (38612.86565445, 48750.63891910, 0)")
    print("  宽度: 1200, 高度: 1000")
    print("  类型: jz-gaochuang")
    print("  delete_mc_yuan: False")

    result1 = insert_tarch_window(
        p=(38612.86565445, 48750.63891910, 0),
        width=1200,
        height=1000,
        window_type="jz-gaochuang",
        delete_mc_yuan=False
    )

    if result1['success']:
        print(f"[成功] 测试1完成 - 窗宽度:{result1['width']}, 高度:{result1['height']}")
    else:
        print("[失败] 测试1失败")
        return False

    time.sleep(2)

    # 6. 测试2
    print("\n[步骤 6/7] 测试2: 插入 jz-pingchuang 窗...")
    print("  位置: (44695.30568975, 46646.78059028, 0)")
    print("  宽度: 2400, 高度: 1800")
    print("  类型: jz-pingchuang")
    print("  delete_mc_yuan: True")

    result2 = insert_tarch_window(
        p=(44695.30568975, 46646.78059028, 0),
        width=2400,
        height=1800,
        window_type="jz-pingchuang",
        delete_mc_yuan=True
    )

    if result2['success']:
        print(f"[成功] 测试2完成 - 窗宽度:{result2['width']}, 高度:{result2['height']}")
    else:
        print("[失败] 测试2失败")
        return False

    # 7. 保存文件
    print("\n[步骤 7/7] 保存测试文件...")
    save_file()
    print("[完成] 文件已保存")

    print("\n" + "="*60)
    print("✓ 测试完成！所有测试均成功！")
    print("="*60)
    print(f"\n测试文件已保存到: {test_file}")
    print("\n测试总结:")
    print("  - 测试1: jz-gaochuang (1200x1000) ✓")
    print("  - 测试2: jz-pingchuang (2400x1800) ✓")

    return True


if __name__ == "__main__":
    try:
        success = main()
        if not success:
            print("\n测试失败，请检查日志")
            sys.exit(1)
    except Exception as e:
        print(f"\n测试过程中发生错误: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
