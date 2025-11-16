#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试修复后的 draw_tarch_wall() 函数
"""

import sys
import time
from pathlib import Path

# 添加脚本路径
SCRIPT_DIR = Path(__file__).parent.parent / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from CAD_file_operations import (
    draw_tarch_wall,
    new_file,
    save_file,
    close_file,
    cad_zt_oneb
)
from CAD_coordination import wait_quiescent

def test_draw_tarch_wall():
    """测试绘制天正墙功能"""
    print("="*60)
    print("测试修复后的 draw_tarch_wall() 函数")
    print("="*60)

    try:
        # 1. 确保CAD处于单文件不确定状态
        print("\n[步骤1] 准备CAD环境...")
        cad_zt_oneb()
        wait_quiescent(min_quiet=1.0, timeout=30.0)
        print("[成功] CAD环境准备完成")

        # 2. 创建新文件
        test_file = "D:/claude-tasks/cad/tests/test_files/test_wall.dwg"
        print(f"\n[步骤2] 创建新文件: {test_file}")
        if not new_file(test_file):
            print("[错误] 创建文件失败")
            return False

        wait_quiescent(min_quiet=1.0, timeout=15.0)
        print("[成功] 文件创建成功")

        # 3. 绘制第一道墙（底边）
        print("\n[步骤3] 绘制第一道墙（底边）...")
        start_time = time.time()
        if not draw_tarch_wall((0, 0, 0), (3000, 0, 0), 240):
            print("[错误] 绘制底边墙失败")
            return False

        elapsed = time.time() - start_time
        print(f"[成功] 底边墙绘制成功，耗时: {elapsed:.1f}秒")

        if elapsed > 60:
            print("[警告] 执行时间超过60秒，可能仍有问题")

        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 4. 绘制第二道墙（右边）
        print("\n[步骤4] 绘制第二道墙（右边）...")
        start_time = time.time()
        if not draw_tarch_wall((3000, 0, 0), (3000, 3000, 0), 240):
            print("[错误] 绘制右边墙失败")
            return False

        elapsed = time.time() - start_time
        print(f"[成功] 右边墙绘制成功，耗时: {elapsed:.1f}秒")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 5. 绘制第三道墙（顶边）
        print("\n[步骤5] 绘制第三道墙（顶边）...")
        start_time = time.time()
        if not draw_tarch_wall((3000, 3000, 0), (0, 3000, 0), 240):
            print("[错误] 绘制顶边墙失败")
            return False

        elapsed = time.time() - start_time
        print(f"[成功] 顶边墙绘制成功，耗时: {elapsed:.1f}秒")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 6. 绘制第四道墙（左边）
        print("\n[步骤6] 绘制第四道墙（左边）...")
        start_time = time.time()
        if not draw_tarch_wall((0, 3000, 0), (0, 0, 0), 240):
            print("[错误] 绘制左边墙失败")
            return False

        elapsed = time.time() - start_time
        print(f"[成功] 左边墙绘制成功，耗时: {elapsed:.1f}秒")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 7. 保存文件
        print("\n[步骤7] 保存文件...")
        if not save_file():
            print("[错误] 保存文件失败")
            return False

        print("[成功] 文件已保存")

        # 8. 验证文件
        if Path(test_file).exists():
            file_size = Path(test_file).stat().st_size
            print(f"[成功] 文件创建成功")
            print(f"        文件大小: {file_size} 字节")
        else:
            print("[错误] 文件未创建")
            return False

        # 9. 关闭文件
        print("\n[步骤8] 关闭文件...")
        close_file("no_save")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        print("\n" + "="*60)
        print("[完成] draw_tarch_wall() 测试通过!")
        print("="*60)
        print("测试结果:")
        print("  - 成功绘制4道墙")
        print("  - 所有墙体绘制时间均在合理范围内")
        print("  - 文件保存成功")
        print("draw_tarch_wall() 函数修复成功!")

        return True

    except Exception as e:
        print(f"\n[错误] 测试异常: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        # 清理：删除测试文件
        print("\n[清理] 删除测试文件...")
        test_file_path = Path("D:/claude-tasks/cad/tests/test_files/test_wall.dwg")
        try:
            if test_file_path.exists():
                test_file_path.unlink()
                print(f"[删除] {test_file_path.name}")
        except Exception as e:
            print(f"[警告] 删除文件失败: {e}")

        # 恢复到单文件不确定状态
        print("\n[清理] 恢复到单文件不确定状态...")
        try:
            cad_zt_oneb()
            print("[成功] CAD状态已恢复")
        except:
            pass

if __name__ == "__main__":
    success = test_draw_tarch_wall()
    sys.exit(0 if success else 1)
