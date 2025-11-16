#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试 insert_tarch_door() - 三角形墙体+三种门
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

def test_insert_tarch_door_triangle():
    """测试在三角形墙体上插入三种不同尺寸的门"""
    print("="*60)
    print("测试 insert_tarch_door() - 三角形墙体+三种门")
    print("="*60)

    try:
        # 1. 准备环境
        print("\n[步骤1] 准备CAD环境...")
        cad_zt_oneb()
        wait_quiescent(min_quiet=2.0, timeout=30.0)

        # 2. 创建新文件
        test_file = "D:/claude-tasks/cad/tests/test_files/test_door_triangle.dwg"
        print(f"\n[步骤2] 创建新文件: {test_file}")
        new_file(test_file)
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 3. 绘制三角形墙体
        # 三角形顶点坐标
        p1 = (0, 0, 0)
        p2 = (6000, 0, 0)
        p3 = (3000, 5196, 0)  # 等边三角形高度 = 6000 * sqrt(3)/2 ≈ 5196

        print("\n[步骤3] 绘制三角形墙体...")

        # 底边
        print("  绘制底边...")
        start_time = time.time()
        if not draw_tarch_wall(p1, p2, 240):
            print("[错误] 绘制底边失败")
            return False
        print(f"  [成功] 底边绘制完成，耗时: {time.time() - start_time:.1f}秒")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 右边
        print("  绘制右边...")
        start_time = time.time()
        if not draw_tarch_wall(p2, p3, 240):
            print("[错误] 绘制右边失败")
            return False
        print(f"  [成功] 右边绘制完成，耗时: {time.time() - start_time:.1f}秒")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 左边
        print("  绘制左边...")
        start_time = time.time()
        if not draw_tarch_wall(p3, p1, 240):
            print("[错误] 绘制左边失败")
            return False
        print(f"  [成功] 左边绘制完成，耗时: {time.time() - start_time:.1f}秒")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 4. 在每堵墙的中间插入不同尺寸的门
        print("\n[步骤4] 在墙体上插入门...")

        # 底边中间：宽1000，高2100
        print("  底边插入门 (1000x2100)...")
        start_time = time.time()
        result1 = insert_tarch_door((3000, 0, 0), width=1000, height=2100)
        if not result1['success']:
            print(f"  [错误] 插入门失败，耗时: {time.time() - start_time:.1f}秒")
            return False
        print(f"  [成功] 门插入成功，耗时: {time.time() - start_time:.1f}秒")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 右边中间：宽1200，高2100
        print("  右边插入门 (1200x2100)...")
        start_time = time.time()
        result2 = insert_tarch_door((4500, 2598, 0), width=1200, height=2100)
        if not result2['success']:
            print(f"  [错误] 插入门失败，耗时: {time.time() - start_time:.1f}秒")
            return False
        print(f"  [成功] 门插入成功，耗时: {time.time() - start_time:.1f}秒")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 左边中间：宽1500，高2400
        print("  左边插入门 (1500x2400)...")
        start_time = time.time()
        result3 = insert_tarch_door((1500, 2598, 0), width=1500, height=2400)
        if not result3['success']:
            print(f"  [错误] 插入门失败，耗时: {time.time() - start_time:.1f}秒")
            return False
        print(f"  [成功] 门插入成功，耗时: {time.time() - start_time:.1f}秒")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 5. 保存文件
        print("\n[步骤5] 保存文件...")
        save_file()
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 6. 验证文件
        if Path(test_file).exists():
            file_size = Path(test_file).stat().st_size
            print(f"[成功] 文件保存成功，大小: {file_size} 字节")

        # 7. 关闭文件
        print("\n[步骤6] 关闭文件...")
        close_file("no_save")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        print("\n" + "="*60)
        print("[完成] insert_tarch_door() 测试通过!")
        print("="*60)
        print("测试结果:")
        print("  - 三角形墙体: [成功]")
        print("  - 门1 (1000x2100): [成功]")
        print("  - 门2 (1200x2100): [成功]")
        print("  - 门3 (1500x2400): [成功]")

        return True

    except Exception as e:
        print(f"\n[错误] 测试异常: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        print("\n[清理] 恢复CAD状态...")
        try:
            cad_zt_oneb()
            print("[成功] CAD状态已恢复")
        except:
            pass

if __name__ == "__main__":
    success = test_insert_tarch_door_triangle()
    sys.exit(0 if success else 1)
