#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_file_operations import start_cad_session, new_file, restore_to_uncertain_state

if __name__ == "__main__":
    start_cad_session()

    try:
        new_file("D:/claude-tasks/tests/test_files/天正测试文件3.dwg")
        print("[成功] 文件创建完成")
    finally:
        restore_to_uncertain_state()
