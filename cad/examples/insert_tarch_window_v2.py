#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
insert_tarch_window 函数 V2 - 直接修改属性版本
不使用 transfer_props_by_matchprop，直接修改门对象的Layer属性
"""

import sys
import time
import logging
from pathlib import Path

sys.path.append(str(Path(__file__).parent))

from CAD_basic import (
    li, stc, last_obj, get_object_property, set_object_property, get_acad_doc
)
from CAD_file_operations import copy_file_content_pywin32, insert_tarch_door


def insert_tarch_window(p, width=600, height=1000, window_type="jz-pingchuang"):
    """
    在墙体上插入天正窗 V2

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
    log_file = log_dir / f"insert_tarch_window_v2_{time.strftime('%Y%m%d_%H%M%S')}.log"

    logger = logging.getLogger('insert_tarch_window_v2')
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
        print(f"[信息] 窗类型: {window_type}")

        # 2. 连接当前激活文件
        li()
        logger.info("已连接当前激活文件")

        # 3. 检查是否需要插入MC_yuan.dwg
        lb = stc('Mc_yuan_bj')
        if len(lb) == 0:
            logger.info("未找到Mc_yuan_bj图层，需要插入MC_yuan.dwg")
            print("[信息] 正在插入MC_yuan.dwg...")
            _, doc = get_acad_doc()
            current_file = doc.FullName
            copy_file_content_pywin32("D:/claude-tasks/cad/xitongwenjian/MC_yuan.dwg", current_file)
            logger.info("已插入MC_yuan.dwg")
            print("[成功] 已插入MC_yuan.dwg")
            li()
        else:
            logger.info(f"已存在Mc_yuan_bj图层 (找到{len(lb)}个对象)")

        # 4. 插入门
        print(f"[信息] 正在插入门... 位置:{p}, 宽度:{width}, 高度:{height}")
        result = insert_tarch_door(p, width=width, height=height)
        if not result['success']:
            logger.error("插入门失败")
            print("[错误] 插入门失败")
            return {'success': False, 'window': None, 'width': None, 'height': None}
        m1 = result['door']
        logger.info(f"已插入门，宽度:{width}, 高度:{height}")
        print(f"[成功] 已插入门")

        # 5. 直接修改门的Layer属性为窗类型
        print(f"[信息] 正在修改门的图层为: {window_type}")
        try:
            m1.Layer = window_type
            logger.info(f"已修改门的图层为: {window_type}")
            print(f"[成功] 已修改门的图层为: {window_type}")
        except Exception as e:
            logger.error(f"修改图层失败: {e}")
            print(f"[错误] 修改图层失败: {e}")
            return {'success': False, 'window': None, 'width': None, 'height': None}

        # 6. 验证图层是否修改成功
        current_layer = m1.Layer
        if current_layer == window_type:
            logger.info(f"成功插入天正窗 - 位置:{p}, 宽度:{width}, 高度:{height}, 类型:{window_type}")
            print(f"[成功] 已插入天正窗 - 宽度:{width}, 高度:{height}, 类型:{window_type}")
            return {'success': True, 'window': m1, 'width': width, 'height': height}
        else:
            logger.error(f"图层修改失败，当前图层:{current_layer}, 期望:{window_type}")
            print(f"[错误] 图层修改失败")
            return {'success': False, 'window': None, 'width': None, 'height': None}

    except Exception as e:
        logger.error(f"插入窗失败: {e}", exc_info=True)
        print(f"[错误] 插入窗失败: {e}")
        return {'success': False, 'window': None, 'width': None, 'height': None}


if __name__ == "__main__":
    print("insert_tarch_window V2 - 直接修改属性版本")
