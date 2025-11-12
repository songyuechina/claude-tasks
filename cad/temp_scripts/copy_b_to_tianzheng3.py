#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
将 B.dwg 的全部内容以非块方式原位插入天正测试文件3.dwg
使用 copy_file_content_pywin32 函数
"""

import sys
from pathlib import Path

# 添加脚本路径
sys.path.append(str(Path(__file__).parent))

from CAD_file_operations import copy_file_content_pywin32, start_cad_session

def main():
    """主函数"""
    print("=" * 70)
    print("复制 B.dwg 内容到天正测试文件3.dwg")
    print("=" * 70)

    # 文件路径
    source_file = "D:/claude-tasks/tests/test_files/B.dwg"
    target_file = "D:/claude-tasks/tests/test_files/天正测试文件3.dwg"

    print(f"\n源文件: {source_file}")
    print(f"目标文件: {target_file}")
    print(f"模式: 非块方式（自动炸开）")
    print(f"位置: 原位插入\n")

    # 启动 CAD 会话
    print("步骤 1: 启动 CAD 会话...")
    if not start_cad_session():
        print("[错误] CAD 会话启动失败")
        return False

    print("\n步骤 2: 执行复制操作...")
    # 调用 copy_file_content_pywin32 函数
    # 该函数会自动：
    # 1. 打开源文件并获取所有对象
    # 2. 打开目标文件
    # 3. 复制对象到目标文件
    # 4. 自动检查并炸开块引用（非块方式）
    # 5. 保存目标文件
    result = copy_file_content_pywin32(source_file, target_file)

    if result:
        print("\n" + "=" * 70)
        print("[成功] 操作完成！")
        print("=" * 70)
        print(f"B.dwg 的所有内容已以非块方式原位插入到天正测试文件3.dwg")
        print(f"目标文件已保存")
    else:
        print("\n" + "=" * 70)
        print("[失败] 操作未完成")
        print("=" * 70)

    return result

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
