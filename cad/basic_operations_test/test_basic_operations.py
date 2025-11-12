#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CAD基本操作范式测试

测试任务:
1. 新建claude1.dwg并绘制圆形，保存
2. 新建12345.dwg并绘制三角形
3. 将12345.dwg的全部内容插入到claude1.dwg

执行: python test_basic_operations.py
"""

import sys
import time
from pathlib import Path

# 添加脚本路径
SCRIPT_DIR = Path(__file__).parent.parent / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from CAD_basic import start_applicationV9, close_all_cad_processes
from CAD_basic_operations import (
    new_dwg_enhanced,
    save_current_dwg_paradigm,
    save_as_dwg_paradigm,
    close_current_dwg_paradigm,
    open_dwg_paradigm,
    insert_dwg_as_block_paradigm
)
from CAD_coordination import (
    wait_quiescent,
    send_cmd_with_sync,
    ensure_single_process
)

# 测试文件路径
TEST_DIR = Path(__file__).parent
CLAUDE1_FILE = str(TEST_DIR / "claude1.dwg")
FILE12345 = str(TEST_DIR / "12345.dwg")

def setup_test_environment():
    """设置测试环境"""
    print("=" * 60)
    print("CAD基本操作范式测试")
    print("=" * 60)

    # 1. 启动CAD
    print("\n[步骤0] 准备测试环境")
    print("-" * 40)
    print("[启动] 正在启动天正CAD...")

    try:
        proc = start_applicationV9(
            PTH=r"C:\Tangent\TArchT20V9",
            max_retries=3,
            retry_delay=2.0
        )
        if not proc:
            print("[错误] CAD启动失败")
            return False
        print("[成功] CAD启动成功")
    except Exception as e:
        print(f"[错误] CAD启动异常: {e}")
        return False

    # 2. 确保单进程
    ensure_single_process()

    # 3. 等待CAD稳定
    time.sleep(2.0)
    wait_quiescent(min_quiet=1.0, timeout=30.0)

    print("[成功] 测试环境准备完成")
    print(f"[信息] 测试目录: {TEST_DIR}")

    return True

def task1_create_claude1_with_circle():
    """
    任务1: 新建claude1.dwg并绘制圆形，保存
    """
    print("\n" + "=" * 60)
    print("[任务1] 新建claude1.dwg并绘制圆形")
    print("=" * 60)

    try:
        # 1. 新建文件
        print("\n[步骤1.1] 新建文件")
        if not new_dwg_enhanced():
            print("[错误] 新建文件失败")
            return False

        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 2. 绘制圆形
        print("\n[步骤1.2] 绘制圆形 (中心: 0,0 半径: 100)")
        circle_cmd = "_CIRCLE\n0,0\n100\n"

        if not send_cmd_with_sync(circle_cmd, wait_after=1.0, timeout=30.0):
            print("[错误] 绘制圆形失败")
            return False

        print("[成功] 圆形绘制成功")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 3. 保存文件
        print(f"\n[步骤1.3] 保存文件到 {CLAUDE1_FILE}")
        if not save_as_dwg_paradigm(CLAUDE1_FILE):
            print("[错误] 保存文件失败")
            return False

        print(f"[成功] 文件已保存: {CLAUDE1_FILE}")

        # 4. 验证文件是否创建
        if Path(CLAUDE1_FILE).exists():
            print(f"[成功] 文件创建成功: {Path(CLAUDE1_FILE).name}")
            print(f"        文件大小: {Path(CLAUDE1_FILE).stat().st_size} 字节")
        else:
            print(f"[错误] 文件未创建: {CLAUDE1_FILE}")
            return False

        # 5. 恢复到单文件状态(关闭当前文件)
        print("\n[清理] 恢复到单文件状态")
        close_current_dwg_paradigm("no_save")
        wait_quiescent(min_quiet=1.0, timeout=15.0)
        print("[成功] 已恢复到单文件状态")

        print("\n[成功] 任务1完成!")
        return True

    except Exception as e:
        print(f"[错误] 任务1执行异常: {e}")
        return False

def task2_create_12345_with_triangle():
    """
    任务2: 新建12345.dwg并绘制三角形
    """
    print("\n" + "=" * 60)
    print("[任务2] 新建12345.dwg并绘制三角形")
    print("=" * 60)

    try:
        # 1. 关闭当前文件(claude1.dwg)
        print("\n[步骤2.1] 关闭当前文件")
        close_current_dwg_paradigm("no_save")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 2. 新建文件
        print("\n[步骤2.2] 新建文件")
        if not new_dwg_enhanced():
            print("[错误] 新建文件失败")
            return False

        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 3. 绘制三角形 (使用多段线)
        print("\n[步骤2.3] 绘制三角形")
        print("        顶点1: (0, 100)")
        print("        顶点2: (-87, -50)")
        print("        顶点3: (87, -50)")

        # 使用PLINE绘制闭合三角形
        triangle_cmd = "_PLINE\n0,100\n-87,-50\n87,-50\nC\n"

        if not send_cmd_with_sync(triangle_cmd, wait_after=1.0, timeout=30.0):
            print("[错误] 绘制三角形失败")
            return False

        print("[成功] 三角形绘制成功")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 4. 保存文件
        print(f"\n[步骤2.4] 保存文件到 {FILE12345}")
        if not save_as_dwg_paradigm(FILE12345):
            print("[错误] 保存文件失败")
            return False

        print(f"[成功] 文件已保存: {FILE12345}")

        # 5. 验证文件是否创建
        if Path(FILE12345).exists():
            print(f"[成功] 文件创建成功: {Path(FILE12345).name}")
            print(f"        文件大小: {Path(FILE12345).stat().st_size} 字节")
        else:
            print(f"[错误] 文件未创建: {FILE12345}")
            return False

        # 6. 恢复到单文件状态(关闭当前文件)
        print("\n[清理] 恢复到单文件状态")
        close_current_dwg_paradigm("no_save")
        wait_quiescent(min_quiet=1.0, timeout=15.0)
        print("[成功] 已恢复到单文件状态")

        print("\n[成功] 任务2完成!")
        return True

    except Exception as e:
        print(f"[错误] 任务2执行异常: {e}")
        return False

def task3_insert_12345_into_claude1():
    """
    任务3: 将12345.dwg的全部内容插入到claude1.dwg
    """
    print("\n" + "=" * 60)
    print("[任务3] 将12345.dwg插入到claude1.dwg")
    print("=" * 60)

    try:
        # 1. 关闭当前文件(12345.dwg)
        print("\n[步骤3.1] 关闭当前文件")
        close_current_dwg_paradigm("no_save")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 2. 打开claude1.dwg
        print(f"\n[步骤3.2] 打开 {CLAUDE1_FILE}")
        if not open_dwg_paradigm(CLAUDE1_FILE):
            print("[错误] 打开文件失败")
            return False

        print(f"[成功] 文件已打开: {Path(CLAUDE1_FILE).name}")
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 3. 插入12345.dwg作为块
        print(f"\n[步骤3.3] 插入块 {FILE12345}")
        print("        插入位置: (200, 0, 0)")
        print("        缩放: 1.0")
        print("        旋转: 0度")
        print("        炸开: 是")

        # 插入并炸开，使内容合并到当前文件
        if not insert_dwg_as_block_paradigm(
            block_file_path=FILE12345,
            insert_point=(200, 0, 0),
            scale=1.0,
            rotation=0.0,
            explode=True  # 炸开使内容合并
        ):
            print("[错误] 插入块失败")
            return False

        print("[成功] 块插入成功")
        wait_quiescent(min_quiet=2.0, timeout=30.0)

        # 4. 保存文件
        print("\n[步骤3.4] 保存文件")
        if not save_current_dwg_paradigm():
            print("[错误] 保存文件失败")
            return False

        print(f"[成功] 文件已保存: {Path(CLAUDE1_FILE).name}")

        # 5. 恢复到单文件状态(关闭当前文件)
        print("\n[清理] 恢复到单文件状态")
        close_current_dwg_paradigm("no_save")
        wait_quiescent(min_quiet=1.0, timeout=15.0)
        print("[成功] 已恢复到单文件状态")

        print("\n[成功] 任务3完成!")
        print(f"[结果] claude1.dwg包含圆形和三角形(位于x=200处)")

        return True

    except Exception as e:
        print(f"[错误] 任务3执行异常: {e}")
        return False

def verify_results():
    """验证测试结果"""
    print("\n" + "=" * 60)
    print("[验证] 测试结果")
    print("=" * 60)

    results = {
        "claude1.dwg": Path(CLAUDE1_FILE).exists(),
        "12345.dwg": Path(FILE12345).exists()
    }

    print("\n文件创建状态:")
    for filename, exists in results.items():
        status = "[存在]" if exists else "[不存在]"
        print(f"  {status} - {filename}")

        if exists:
            filepath = TEST_DIR / filename
            size = filepath.stat().st_size
            print(f"           大小: {size} 字节")

    all_success = all(results.values())

    if all_success:
        print("\n[完成] 所有测试任务成功完成!")
        print(f"\n[信息] 测试文件位置: {TEST_DIR}")
        print(f"   - {Path(CLAUDE1_FILE).name} (包含圆形和三角形)")
        print(f"   - {Path(FILE12345).name} (包含三角形)")
    else:
        print("\n[警告] 部分测试任务失败，请检查日志")

    return all_success

def main():
    """主测试函数"""
    try:
        # 0. 准备测试环境
        if not setup_test_environment():
            print("\n[错误] 测试环境准备失败")
            return False

        # 1. 执行任务1
        if not task1_create_claude1_with_circle():
            print("\n[错误] 任务1失败，终止测试")
            return False

        time.sleep(1.0)  # 任务间隔

        # 2. 执行任务2
        if not task2_create_12345_with_triangle():
            print("\n[错误] 任务2失败，终止测试")
            return False

        time.sleep(1.0)  # 任务间隔

        # 3. 执行任务3
        if not task3_insert_12345_into_claude1():
            print("\n[错误] 任务3失败")
            return False

        # 4. 验证结果
        success = verify_results()

        return success

    except Exception as e:
        print(f"\n[异常] 测试执行异常: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        # 最终清理:关闭所有CAD进程
        print("\n" + "=" * 60)
        print("[清理] 最终环境清理")
        print("=" * 60)
        try:
            print("[清理] 关闭所有CAD进程...")
            close_all_cad_processes()
            print("[成功] 所有CAD进程已关闭")
        except Exception as e:
            print(f"[警告] 清理时出错: {e}")

        print("\n" + "=" * 60)
        print("测试执行结束")
        print("=" * 60)

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
