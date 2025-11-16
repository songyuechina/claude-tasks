#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试 insert_tarch_window() - 矩形墙体+三种窗
"""

import sys
import time
from pathlib import Path

# 添加脚本路径
SCRIPT_DIR = Path(__file__).parent.parent / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from CAD_file_operations import (
    draw_tarch_wall,
    insert_tarch_window,
    new_file,
    save_file,
    close_file,
    cad_zt_oneb
)
from CAD_coordination import wait_quiescent

def test_insert_tarch_window_rectangle():
    """测试在矩形墙体上插入三种不同类型的窗"""
    print("="*60)
    print("测试 insert_tarch_window() - 矩形墙体+三种窗")
    print("="*60)

    try:
        # 1. 准备环境
        print("\n[步骤1] 准备CAD环境...")
        cad_zt_oneb()
        wait_quiescent(min_quiet=2.0, timeout=30.0)

        # 2. 创建新文件
        test_file = "D:/claude-tasks/cad/tests/test_files/test_window_rectangle.dwg"
        print(f"\n[步骤2] 创建新文件: {test_file}")
        new_file(test_file)
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 3. 绘制矩形墙体 (8000mm x 6000mm)
        print("\n[步骤3] 绘制矩形墙体...")

        # 底边
        print("  绘制底边...")
        start_time = time.time()
        if not draw_tarch_wall((0, 0, 0), (8000, 0, 0), 240):
            print("[错误] 绘制底边失败")
            return False
        print(f"  [成功] 底边绘制完成，耗时: {time.time() - start_time:.1f}秒")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 右边
        print("  绘制右边...")
        start_time = time.time()
        if not draw_tarch_wall((8000, 0, 0), (8000, 6000, 0), 240):
            print("[错误] 绘制右边失败")
            return False
        print(f"  [成功] 右边绘制完成，耗时: {time.time() - start_time:.1f}秒")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 顶边
        print("  绘制顶边...")
        start_time = time.time()
        if not draw_tarch_wall((8000, 6000, 0), (0, 6000, 0), 240):
            print("[错误] 绘制顶边失败")
            return False
        print(f"  [成功] 顶边绘制完成，耗时: {time.time() - start_time:.1f}秒")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 左边
        print("  绘制左边...")
        start_time = time.time()
        if not draw_tarch_wall((0, 6000, 0), (0, 0, 0), 240):
            print("[错误] 绘制左边失败")
            return False
        print(f"  [成功] 左边绘制完成，耗时: {time.time() - start_time:.1f}秒")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 4. 在每堵墙的1/4和3/4处插入不同类型的窗
        print("\n[步骤4] 在墙体上插入窗...")

        # 底边：1/4处和3/4处
        print("  底边插入窗...")

        # 1/4处：jz-pingchuang (1000x2100)
        print("    位置1/4: jz-pingchuang (1000x2100)...")
        start_time = time.time()
        result1 = insert_tarch_window((2000, 0, 0), width=1000, height=2100,
                                      window_type="jz-pingchuang", delete_mc_yuan=False)
        if not result1['success']:
            print(f"    [错误] 插入窗失败，耗时: {time.time() - start_time:.1f}秒")
            return False
        print(f"    [成功] 窗插入成功，耗时: {time.time() - start_time:.1f}秒")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 3/4处：jz-juanlianmen (1200x2100)
        print("    位置3/4: jz-juanlianmen (1200x2100)...")
        start_time = time.time()
        result2 = insert_tarch_window((6000, 0, 0), width=1200, height=2100,
                                      window_type="jz-juanlianmen", delete_mc_yuan=False)
        if not result2['success']:
            print(f"    [错误] 插入窗失败，耗时: {time.time() - start_time:.1f}秒")
            return False
        print(f"    [成功] 窗插入成功，耗时: {time.time() - start_time:.1f}秒")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 右边：1/4处和3/4处
        print("  右边插入窗...")

        # 1/4处：jz-tuilamen (1500x2400)
        print("    位置1/4: jz-tuilamen (1500x2400)...")
        start_time = time.time()
        result3 = insert_tarch_window((8000, 1500, 0), width=1500, height=2400,
                                      window_type="jz-tuilamen", delete_mc_yuan=False)
        if not result3['success']:
            print(f"    [错误] 插入窗失败，耗时: {time.time() - start_time:.1f}秒")
            return False
        print(f"    [成功] 窗插入成功，耗时: {time.time() - start_time:.1f}秒")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 3/4处：jz-pingchuang (1000x2100)
        print("    位置3/4: jz-pingchuang (1000x2100)...")
        start_time = time.time()
        result4 = insert_tarch_window((8000, 4500, 0), width=1000, height=2100,
                                      window_type="jz-pingchuang", delete_mc_yuan=False)
        if not result4['success']:
            print(f"    [错误] 插入窗失败，耗时: {time.time() - start_time:.1f}秒")
            return False
        print(f"    [成功] 窗插入成功，耗时: {time.time() - start_time:.1f}秒")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 顶边：1/4处和3/4处
        print("  顶边插入窗...")

        # 1/4处：jz-juanlianmen (1200x2100)
        print("    位置1/4: jz-juanlianmen (1200x2100)...")
        start_time = time.time()
        result5 = insert_tarch_window((6000, 6000, 0), width=1200, height=2100,
                                      window_type="jz-juanlianmen", delete_mc_yuan=False)
        if not result5['success']:
            print(f"    [错误] 插入窗失败，耗时: {time.time() - start_time:.1f}秒")
            return False
        print(f"    [成功] 窗插入成功，耗时: {time.time() - start_time:.1f}秒")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 3/4处：jz-tuilamen (1500x2400)
        print("    位置3/4: jz-tuilamen (1500x2400)...")
        start_time = time.time()
        result6 = insert_tarch_window((2000, 6000, 0), width=1500, height=2400,
                                      window_type="jz-tuilamen", delete_mc_yuan=False)
        if not result6['success']:
            print(f"    [错误] 插入窗失败，耗时: {time.time() - start_time:.1f}秒")
            return False
        print(f"    [成功] 窗插入成功，耗时: {time.time() - start_time:.1f}秒")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 左边：1/4处和3/4处
        print("  左边插入窗...")

        # 1/4处：jz-pingchuang (1000x2100)
        print("    位置1/4: jz-pingchuang (1000x2100)...")
        start_time = time.time()
        result7 = insert_tarch_window((0, 4500, 0), width=1000, height=2100,
                                      window_type="jz-pingchuang", delete_mc_yuan=False)
        if not result7['success']:
            print(f"    [错误] 插入窗失败，耗时: {time.time() - start_time:.1f}秒")
            return False
        print(f"    [成功] 窗插入成功，耗时: {time.time() - start_time:.1f}秒")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 3/4处：jz-juanlianmen (1200x2100)
        print("    位置3/4: jz-juanlianmen (1200x2100)...")
        start_time = time.time()
        result8 = insert_tarch_window((0, 1500, 0), width=1200, height=2100,
                                      window_type="jz-juanlianmen", delete_mc_yuan=False)
        if not result8['success']:
            print(f"    [错误] 插入窗失败，耗时: {time.time() - start_time:.1f}秒")
            return False
        print(f"    [成功] 窗插入成功，耗时: {time.time() - start_time:.1f}秒")
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
        print("[完成] insert_tarch_window() 测试通过!")
        print("="*60)
        print("测试结果:")
        print("  - 矩形墙体: [成功]")
        print("  - 窗类型1 (jz-pingchuang): [成功] x3")
        print("  - 窗类型2 (jz-juanlianmen): [成功] x3")
        print("  - 窗类型3 (jz-tuilamen): [成功] x2")
        print("  - 共插入8个窗")

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
    success = test_insert_tarch_window_rectangle()
    sys.exit(0 if success else 1)
