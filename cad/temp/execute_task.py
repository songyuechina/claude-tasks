#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_file_operations import start_cad_session, copy_file_content_pywin32, restore_to_uncertain_state

if __name__ == "__main__":
    start_cad_session()

    try:
        copy_file_content_pywin32(
            source_file="D:/claude-tasks/tests/test_files/B.dwg",
            target_file="D:/claude-tasks/tests/test_files/天正测试文件3.dwg"
        )
        print("\n[完成] 操作成功")
    finally:
        restore_to_uncertain_state()
