#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试 insert_tarch_window 函数
"""

import sys
import time
import logging
from pathlib import Path

# 添加scripts目录到路径
sys.path.append(str(Path(__file__).parent))

from CAD_basic import (
    li, stc, last_obj, get_object_property, set_object_property,
    transfer_props_by_matchprop, get_acad_doc, start_applicationV9
)
from CAD_file_operations import copy_file_content_pywin32, insert_tarch_door, open_file, save_file
from CAD_coordination import ensure_single_process, wait_quiescent


def insert_tarch_window(p, width=600, height=1000, window_type="jz-pingchuang", delete_mc_yuan=False):
    """
    在墙体上插入天正窗

    Args:
        p: 插入点坐标 (x, y, z)
        width: 窗宽度，默认600
        height: 窗高度，默认1000
        window_type: 窗类型，默认"jz-pingchuang"，允许的类型:
            "jz-menlianchuang", "jz-dong", "jz-gaochuang", "jz-baiyechuang",
            "jz-tuchuang", "jz-pingchuang", "jz-zimumen", "jz-juanlianmen",
            "jz-tuilamen", "jz-shuangmen"
        delete_mc_yuan: 是否删除MC_yuan.dwg插入的对象，默认False不删除

    Returns:
        dict: {'success': bool, 'window': 窗对象, 'width': 宽度, 'height': 高度}
    """
    # 配置日志
    log_dir = Path("D:/claude-tasks/cad/logs")
    log_dir.mkdir(exist_ok=True)
    log_file = log_dir / f"insert_tarch_window_{time.strftime('%Y%m%d_%H%M%S')}.log"

    # 配置日志处理器
    logger = logging.getLogger('insert_tarch_window')
    logger.setLevel(logging.INFO)
    # 清除已有的处理器
    if logger.handlers:
        logger.handlers.clear()
    # 添加文件处理器
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
            logger.error(f"窗类型错误: {window_type}, 允许的类型: {allowed_types}")
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
            # 获取当前文件路径
            _, doc = get_acad_doc()
            current_file = doc.FullName
            logger.info(f"当前文件: {current_file}")
            copy_file_content_pywin32("D:/claude-tasks/cad/xitongwenjian/MC_yuan.dwg", current_file)
            logger.info("已插入MC_yuan.dwg")
            print("[成功] 已插入MC_yuan.dwg")
            # 重新连接
            li()
        else:
            logger.info(f"已存在Mc_yuan_bj图层，无需插入MC_yuan.dwg (找到{len(lb)}个对象)")
            print(f"[信息] 已存在Mc_yuan_bj图层 (找到{len(lb)}个对象)")

        # 4. 插入门
        print(f"[信息] 正在插入门... 位置:{p}, 宽度:{width}, 高度:{height}")
        result = insert_tarch_door(p, width=width, height=height)
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

        # 6. 使用transfer_props_by_matchprop匹配属性，最多5次
        print("[信息] 正在进行属性匹配...")
        success = False
        for attempt in range(1, 6):
            try:
                result_match = transfer_props_by_matchprop(window_src, m1, max_try=3, delay=0.4)
                if result_match:
                    # 检查图层是否改变
                    m1_layer = m1.Layer
                    if m1_layer == window_type:
                        logger.info(f"第{attempt}次匹配成功，门已转换为窗，图层:{m1_layer}")
                        print(f"[成功] 第{attempt}次匹配成功，门已转换为窗，图层:{m1_layer}")
                        success = True
                        break
                    else:
                        logger.warning(f"第{attempt}次匹配后图层不正确: {m1_layer}, 期望:{window_type}")
                        print(f"[警告] 第{attempt}次匹配后图层不正确: {m1_layer}, 期望:{window_type}")
            except Exception as e:
                logger.warning(f"第{attempt}次匹配失败: {e}")
                print(f"[警告] 第{attempt}次匹配失败: {e}")
            time.sleep(0.5)

        if not success:
            logger.error("transfer_props_by_matchprop匹配5次仍然失败")
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
        return {'success': False, 'window': None, 'width': None, 'height': None}


def main():
    """主测试函数"""
    print("="*60)
    print("insert_tarch_window 函数测试")
    print("="*60)

    # 1. 关闭所有CAD进程
    print("\n[1/5] 关闭所有CAD进程...")
    ensure_single_process()

    # 2. 启动天正CAD
    print("\n[2/5] 启动天正CAD...")
    start_applicationV9(PTH=r"C:\Tangent\TArchT20V9", max_retries=3, retry_delay=2.0)
    wait_quiescent(min_quiet=1.0, timeout=30.0)

    # 3. 打开测试文件
    print("\n[3/5] 打开测试文件...")
    test_file = "D:/claude-tasks/cad/Function_testing/insert_tarch_window-1.dwg"
    open_file(test_file)
    wait_quiescent(min_quiet=1.0, timeout=15.0)

    # 连接当前文件
    li()
    print(f"[成功] 已打开并连接测试文件: {test_file}")

    # 4. 执行测试
    print("\n[4/5] 执行插入窗测试...")

    # 测试1: 在指定位置插入 jz-gaochuang，宽1200，高1000
    print("\n--- 测试1: 插入 jz-gaochuang ---")
    result1 = insert_tarch_window(
        p=(38612.86565445, 48750.63891910, 0),
        width=1200,
        height=1000,
        window_type="jz-gaochuang"
    )
    if result1['success']:
        print(f"[OK] 测试1成功")
    else:
        print(f"[FAIL] 测试1失败: {result1}")

    time.sleep(2)

    # 测试2: 在指定位置插入 jz-pingchuang，宽2400，高1800，删除MC_yuan对象
    print("\n--- 测试2: 插入 jz-pingchuang (delete_mc_yuan=True) ---")
    result2 = insert_tarch_window(
        p=(44695.30568975, 46646.78059028, 0),
        width=2400,
        height=1800,
        window_type="jz-pingchuang",
        delete_mc_yuan=True
    )
    if result2['success']:
        print(f"[OK] 测试2成功")
    else:
        print(f"[FAIL] 测试2失败: {result2}")

    # 5. 保存文件
    print("\n[5/5] 保存文件...")
    save_file()
    print("[成功] 文件已保存")

    print("\n" + "="*60)
    print("测试完成!")
    print("="*60)


##if __name__ == "__main__":
##    main()
