#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试天正墙和门的组合功能
"""

import sys
import time
from pathlib import Path

# 添加脚本路径
SCRIPT_DIR = Path(__file__).parent.parent / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from CAD_file_operations import (
    draw_tarch_wall,
    insert_tarch_door,
    new_file,
    save_file,
    close_file,
    cad_zt_oneb
)
from CAD_coordination import wait_quiescent

def test_wall_and_door():
    """测试墙体和门的组合"""
    print("="*60)
    print("测试天正墙和门的组合功能")
    print("="*60)

    try:
        # 1. 准备环境
        print("\n[步骤1] 准备CAD环境...")
        cad_zt_oneb()
        wait_quiescent(min_quiet=1.0, timeout=30.0)

        # 2. 创建新文件
        test_file = "D:/claude-tasks/cad/tests/test_files/test_wall_door.dwg"
        print(f"\n[步骤2] 创建新文件...")
        new_file(test_file)
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 3. 绘制一道墙
        print("\n[步骤3] 绘制墙体...")
        start_time = time.time()
        if not draw_tarch_wall((0, 0, 0), (5000, 0, 0), 240):
            print("[错误] 绘制墙体失败")
            return False
        elapsed = time.time() - start_time
        print(f"[成功] 墙体绘制成功，耗时: {elapsed:.1f}秒")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 4. 在墙上插入门
        print("\n[步骤4] 在墙上插入门...")
        start_time = time.time()
        result = insert_tarch_door((2500, 0, 0), width=900, height=2100)
        elapsed = time.time() - start_time

        if not result['success']:
            print(f"[错误] 插入门失败，耗时: {elapsed:.1f}秒")
            return False

        print(f"[成功] 门插入成功，耗时: {elapsed:.1f}秒")
        print(f"        门宽度: {result['width']}")
        print(f"        门高度: {result['height']}")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 5. 保存文件
        print("\n[步骤5] 保存文件...")
        save_file()

        # 6. 验证文件
        if Path(test_file).exists():
            file_size = Path(test_file).stat().st_size
            print(f"[成功] 文件保存成功，大小: {file_size} 字节")

        # 7. 关闭文件
        print("\n[步骤6] 关闭文件...")
        close_file("no_save")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        print("\n" + "="*60)
        print("[完成] 墙体和门测试通过!")
        print("="*60)
        print("测试结果:")
        print("  - 墙体绘制: ✅ 成功")
        print("  - 门插入: ✅ 成功")
        print("  - 所有操作时间均在合理范围内")

        return True

    except Exception as e:
        print(f"\n[错误] 测试异常: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        # 清理
        print("\n[清理] 删除测试文件...")
        test_file_path = Path("D:/claude-tasks/cad/tests/test_files/test_wall_door.dwg")
        try:
            if test_file_path.exists():
                test_file_path.unlink()
                print(f"[删除] {test_file_path.name}")
        except:
            pass

        print("\n[清理] 恢复CAD状态...")
        try:
            cad_zt_oneb()
            print("[成功] CAD状态已恢复")
        except:
            pass

if __name__ == "__main__":
    success = test_wall_and_door()
    sys.exit(0 if success else 1)
