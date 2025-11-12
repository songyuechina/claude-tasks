#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
insert_tarch_window 函数 V3 - 严格按照1-9步骤，移动窗元到门附近
"""

import sys
import time
import logging
from pathlib import Path

sys.path.append(str(Path(__file__).parent))

from CAD_basic import (
    li, stc, last_obj, get_object_property, set_object_property,
    transfer_props_by_matchprop, get_acad_doc
)
from CAD_file_operations import copy_file_content_pywin32, insert_tarch_door


def insert_tarch_window(p, width=600, height=1000, window_type="jz-pingchuang"):
    """
    在墙体上插入天正窗 V3 - 严格按照步骤1-9

    Args:
        p: 插入点坐标 (x, y, z)
        width: 窗宽度，默认600
        height: 窗高度，默认1000
        window_type: 窗类型，默认"jz-pingchuang"

    Returns:
        dict: {'success': bool, 'window': 窗对象, 'width': 宽度, 'height': 高度}
    """
    # 配置日志
    log_dir = Path("D:/claude-tasks/cad/logs")
    log_dir.mkdir(exist_ok=True)
    log_file = log_dir / f"insert_tarch_window_v3_{time.strftime('%Y%m%d_%H%M%S')}.log"

    logger = logging.getLogger('insert_tarch_window_v3')
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
        # 步骤1: 检查窗类型
        if window_type not in allowed_types:
            logger.error(f"窗类型错误: {window_type}")
            print(f"[错误] 窗类型错误: {window_type}")
            return {'success': False, 'window': None, 'width': None, 'height': None}
        logger.info(f"步骤1: 窗类型检查通过: {window_type}")
        print(f"[步骤1] 窗类型: {window_type}")

        # 步骤2: 连接当前激活文件
        li()
        logger.info("步骤2: 已连接当前激活文件")
        print("[步骤2] 已连接当前激活文件")

        # 步骤3: 检查是否需要插入MC_yuan.dwg
        lb = stc('Mc_yuan_bj')
        if len(lb) == 0:
            logger.info("步骤3: 未找到Mc_yuan_bj图层，需要插入MC_yuan.dwg")
            print("[步骤3] 正在插入MC_yuan.dwg...")
            _, doc = get_acad_doc()
            current_file = doc.FullName
            copy_file_content_pywin32("D:/claude-tasks/cad/xitongwenjian/MC_yuan.dwg", current_file)
            logger.info("已插入MC_yuan.dwg")
            print("[成功] 已插入MC_yuan.dwg")
            li()
        else:
            logger.info(f"步骤3: 已存在Mc_yuan_bj图层 (找到{len(lb)}个对象)")
            print(f"[步骤3] 已存在Mc_yuan_bj图层")

        # 步骤4: 插入门
        print(f"[步骤4] 正在插入门... 位置:{p}, 宽度:{width}, 高度:{height}")
        result = insert_tarch_door(p, width=width, height=height)
        if not result['success']:
            logger.error("步骤4: 插入门失败")
            print("[错误] 插入门失败")
            return {'success': False, 'window': None, 'width': None, 'height': None}
        m1 = result['door']
        logger.info(f"步骤4: 已插入门，宽度:{width}, 高度:{height}")
        print(f"[成功] 已插入门")

        # 步骤5: 选择窗类型图层的窗元并修改尺寸
        print(f"[步骤5] 正在选择窗类型图层: {window_type}")
        lc = stc(window_type)
        if len(lc) == 0:
            logger.error(f"步骤5: 未找到窗类型图层: {window_type}")
            print(f"[错误] 未找到窗类型图层: {window_type}")
            return {'success': False, 'window': None, 'width': None, 'height': None}

        window_src = lc[0]

        # 获取窗元的原始位置
        try:
            src_bbox = window_src.GetBoundingBox()
            src_center_x = (src_bbox[0][0] + src_bbox[1][0]) / 2
            src_center_y = (src_bbox[0][1] + src_bbox[1][1]) / 2
            logger.info(f"窗元原始位置: ({src_center_x}, {src_center_y})")

            # 移动窗元到门附近（门的位置偏移一点）
            offset_x = p[0] + 2000 - src_center_x  # 偏移2000单位
            offset_y = p[1] - src_center_y
            window_src.Move(src_bbox[0], (src_bbox[0][0] + offset_x, src_bbox[0][1] + offset_y, src_bbox[0][2]))
            logger.info(f"已移动窗元到门附近，偏移: ({offset_x}, {offset_y})")
            print(f"[信息] 已移动窗元到门附近")
        except Exception as e:
            logger.warning(f"移动窗元失败: {e}")
            print(f"[警告] 移动窗元失败，继续尝试")

        set_object_property(window_src, "Width", width)
        set_object_property(window_src, "Height", height)
        logger.info(f"步骤5: 已设置窗元尺寸: 宽{width}, 高{height}")
        print(f"[成功] 已设置窗元尺寸: 宽{width}, 高{height}")

        # 步骤6: 使用transfer_props_by_matchprop匹配属性，最多5次
        print("[步骤6] 正在进行属性匹配...")
        success = False
        for attempt in range(1, 6):
            try:
                result_match = transfer_props_by_matchprop(window_src, m1, max_try=3, delay=0.4)
                if result_match:
                    # 检查图层是否改变
                    m1_layer = m1.Layer
                    if m1_layer == window_type:
                        logger.info(f"第{attempt}次匹配成功，门已转换为窗，图层:{m1_layer}")
                        print(f"[成功] 第{attempt}次匹配成功，门已转换为窗")
                        success = True
                        break
                    else:
                        logger.warning(f"第{attempt}次匹配后图层不正确: {m1_layer}, 期望:{window_type}")
                        print(f"[警告] 第{attempt}次匹配后图层不正确")
            except Exception as e:
                logger.warning(f"第{attempt}次匹配失败: {e}")
                print(f"[警告] 第{attempt}次匹配失败")
            time.sleep(0.5)

        if not success:
            logger.error("步骤6: transfer_props_by_matchprop匹配5次仍然失败")
            print("[错误] 属性匹配5次仍然失败")
            return {'success': False, 'window': None, 'width': None, 'height': None}

        logger.info(f"成功插入天正窗 - 位置:{p}, 宽度:{width}, 高度:{height}, 类型:{window_type}")
        print(f"[成功] 已插入天正窗 - 宽度:{width}, 高度:{height}, 类型:{window_type}")
        return {'success': True, 'window': m1, 'width': width, 'height': height}

    except Exception as e:
        logger.error(f"插入窗失败: {e}", exc_info=True)
        print(f"[错误] 插入窗失败: {e}")
        return {'success': False, 'window': None, 'width': None, 'height': None}


if __name__ == "__main__":
    print("insert_tarch_window V3 - 严格按照1-9步骤，移动窗元到门附近")
