#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_file_operations import start_cad_session, open_file, restore_to_uncertain_state

if __name__ == "__main__":
    start_cad_session()

    try:
        open_file("D:/claude-tasks/tests/test_files/B.dwg")
        print("[成功] 文件已打开")
    finally:
        restore_to_uncertain_state()
