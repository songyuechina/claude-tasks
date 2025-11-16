#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
创建用于插入测试的源DWG文件
使用天正墙和天正门创建一个简单的房间
"""

import sys
import time
from pathlib import Path

# 添加脚本路径
SCRIPT_DIR = Path(__file__).parent.parent / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from CAD_basic import start_applicationV9, close_all_cad_processes
from CAD_file_operations import (
    new_file,
    draw_tarch_wall,
    insert_tarch_door,
    save_file,
    close_file,
    cad_zt_oneb
)
from CAD_coordination import wait_quiescent, send_cmd_with_sync

def create_test_insertion_source():
    """创建用于插入测试的源文件"""
    print("=" * 60)
    print("创建测试插入源文件")
    print("=" * 60)

    try:
        # 1. 确保CAD处于单文件不确定状态
        print("\n[步骤1] 准备CAD环境...")
        cad_zt_oneb()
        wait_quiescent(min_quiet=1.0, timeout=30.0)
        print("[成功] CAD环境准备完成")

        # 2. 创建新文件
        test_dir = Path(__file__).parent / "test_files"
        test_dir.mkdir(exist_ok=True)

        output_file = test_dir / "test_insertion_source.dwg"

        print(f"\n[步骤2] 创建新DWG文件: {output_file}")
        if not new_file(str(output_file)):
            print("[错误] 创建文件失败")
            return False

        wait_quiescent(min_quiet=1.0, timeout=15.0)
        print("[成功] 文件创建成功")

        # 3. 绘制一个简单的矩形房间（使用天正墙）
        print("\n[步骤3] 绘制矩形房间...")

        # 房间尺寸: 6000mm x 4000mm
        # 墙厚: 240mm
        room_width = 6000
        room_height = 4000
        wall_thickness = 240

        # 底边墙 (0,0) -> (6000,0)
        print("  绘制底边墙...")
        if not draw_tarch_wall((0, 0, 0), (room_width, 0, 0), wall_thickness):
            print("[错误] 绘制底边墙失败")
            return False
        wait_quiescent(min_quiet=0.5, timeout=10.0)

        # 右边墙 (6000,0) -> (6000,4000)
        print("  绘制右边墙...")
        if not draw_tarch_wall((room_width, 0, 0), (room_width, room_height, 0), wall_thickness):
            print("[错误] 绘制右边墙失败")
            return False
        wait_quiescent(min_quiet=0.5, timeout=10.0)

        # 顶边墙 (6000,4000) -> (0,4000)
        print("  绘制顶边墙...")
        if not draw_tarch_wall((room_width, room_height, 0), (0, room_height, 0), wall_thickness):
            print("[错误] 绘制顶边墙失败")
            return False
        wait_quiescent(min_quiet=0.5, timeout=10.0)

        # 左边墙 (0,4000) -> (0,0)
        print("  绘制左边墙...")
        if not draw_tarch_wall((0, room_height, 0), (0, 0, 0), wall_thickness):
            print("[错误] 绘制左边墙失败")
            return False
        wait_quiescent(min_quiet=0.5, timeout=10.0)

        print("[成功] 矩形房间绘制完成")

        # 4. 在墙上插入门
        print("\n[步骤4] 插入天正门...")

        # 在底边墙中间插入门 (3000, 0)
        door_width = 900
        door_height = 2100

        print(f"  插入门: 位置(3000, 0), 宽度{door_width}, 高度{door_height}...")
        door_result = insert_tarch_door((3000, 0, 0), width=door_width, height=door_height)

        if not door_result['success']:
            print("[警告] 插入门失败，继续执行")
        else:
            print(f"[成功] 门已插入，宽度:{door_result['width']}, 高度:{door_result['height']}")

        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 5. 缩放到全图
        print("\n[步骤5] 调整视图...")
        send_cmd_with_sync("_ZOOM\\n_E\\n", wait_after=1.0)
        wait_quiescent(min_quiet=0.5, timeout=15.0)

        # 6. 保存文件
        print(f"\n[步骤6] 保存文件: {output_file}")
        if not save_file():
            print("[错误] 保存文件失败")
            return False

        print("[成功] 文件已保存")

        # 7. 验证文件
        if output_file.exists():
            file_size = output_file.stat().st_size
            print(f"[成功] 文件创建成功")
            print(f"        文件路径: {output_file}")
            print(f"        文件大小: {file_size} 字节")
        else:
            print("[错误] 文件未创建")
            return False

        # 8. 关闭文件
        print("\n[步骤7] 关闭文件...")
        close_file("no_save")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        print("\n" + "=" * 60)
        print("[完成] 测试插入源文件创建成功!")
        print("=" * 60)
        print(f"文件位置: {output_file}")
        print(f"房间尺寸: {room_width}mm x {room_height}mm")
        print(f"墙体厚度: {wall_thickness}mm")
        print(f"门尺寸: {door_width}mm x {door_height}mm")

        return True

    except Exception as e:
        print(f"\n[错误] 执行异常: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        # 恢复到单文件不确定状态
        print("\n[清理] 恢复到单文件不确定状态...")
        try:
            cad_zt_oneb()
            print("[成功] CAD状态已恢复")
        except:
            pass

if __name__ == "__main__":
    success = create_test_insertion_source()
    sys.exit(0 if success else 1)
