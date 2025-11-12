#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_file_operations import start_cad_session, copy_file_content, restore_to_uncertain_state

if __name__ == "__main__":
    start_cad_session()

    try:
        # 测试：将B.dwg内容复制到天正测试文件3.dwg，炸开
        success = copy_file_content(
            source_file="D:/claude-tasks/tests/test_files/B.dwg",
            target_file="D:/claude-tasks/tests/test_files/天正测试文件3.dwg",
            explode=True,
            x=0,
            y=0
        )

        if success:
            print("\n[测试成功] copy_file_content 函数工作正常")
        else:
            print("\n[测试失败] copy_file_content 函数执行失败")

    finally:
        restore_to_uncertain_state()
