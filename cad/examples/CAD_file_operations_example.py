#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CAD文件操作使用示例
"""

import sys
from pathlib import Path

sys.path.append(str(Path(__file__).parent))

from CAD_file_operations import (
    new_file,
    open_file,
    save_file,
    save_file_as,
    close_file,
    insert_file_as_block,
    insert_file_exploded,
    copy_file_content
)
from CAD_basic import start_applicationV9
from CAD_coordination import send_cmd_with_sync, wait_quiescent, ensure_single_process

def example_1_create_and_save():
    """示例1: 新建文件并保存"""
    print("\n示例1: 新建文件并保存")
    print("-" * 40)

    # 新建空白文件
    new_file()
    wait_quiescent(min_quiet=1.0, timeout=15.0)

    # 绘制内容
    send_cmd_with_sync("_CIRCLE\n0,0\n50\n", wait_after=1.0)

    # 另存为
    save_file_as("D:/test/example1.dwg")

    # 关闭
    close_file("no_save")

def example_2_open_and_edit():
    """示例2: 打开文件并编辑"""
    print("\n示例2: 打开文件并编辑")
    print("-" * 40)

    # 打开文件
    open_file("D:/test/example1.dwg")
    wait_quiescent(min_quiet=1.0, timeout=15.0)

    # 添加内容
    send_cmd_with_sync("_RECTANG\n100,0\n200,100\n", wait_after=1.0)

    # 保存
    save_file()

    # 关闭
    close_file("no_save")

def example_3_insert_as_block():
    """示例3: 插入文件为块"""
    print("\n示例3: 插入文件为块")
    print("-" * 40)

    # 新建文件
    new_file()
    wait_quiescent(min_quiet=1.0, timeout=15.0)

    # 插入为块
    insert_file_as_block(
        source_file="D:/test/example1.dwg",
        x=0, y=0, z=0,
        scale=1.0,
        rotation=0.0
    )

    # 保存
    save_file_as("D:/test/example3.dwg")
    close_file("no_save")

def example_4_insert_exploded():
    """示例4: 插入文件并炸开"""
    print("\n示例4: 插入文件并炸开")
    print("-" * 40)

    # 新建文件
    new_file()
    wait_quiescent(min_quiet=1.0, timeout=15.0)

    # 插入并炸开
    insert_file_exploded(
        source_file="D:/test/example1.dwg",
        x=0, y=0, z=0,
        scale=1.0
    )

    # 保存
    save_file_as("D:/test/example4.dwg")
    close_file("no_save")

def example_5_copy_content():
    """示例5: 拷贝文件内容"""
    print("\n示例5: 拷贝文件内容")
    print("-" * 40)

    # 拷贝为块
    copy_file_content(
        source_file="D:/test/example1.dwg",
        target_file="D:/test/example5_block.dwg",
        explode=False,
        x=0, y=0
    )
    close_file("no_save")

    # 拷贝并炸开
    copy_file_content(
        source_file="D:/test/example1.dwg",
        target_file="D:/test/example5_exploded.dwg",
        explode=True,
        x=0, y=0
    )
    close_file("no_save")

def main():
    """运行所有示例"""
    print("="*60)
    print("CAD文件操作使用示例")
    print("="*60)

    # 启动CAD
    print("\n启动CAD...")
    start_applicationV9(PTH=r"C:\Tangent\TArchT20V9")
    ensure_single_process()
    wait_quiescent(min_quiet=2.0, timeout=30.0)

    # 运行示例
    example_1_create_and_save()
    example_2_open_and_edit()
    example_3_insert_as_block()
    example_4_insert_exploded()
    example_5_copy_content()

    print("\n" + "="*60)
    print("所有示例完成!")
    print("="*60)

if __name__ == "__main__":
    main()
