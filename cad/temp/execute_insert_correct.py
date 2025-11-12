#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_file_operations import (
    start_cad_session,
    restore_to_uncertain_state,
    open_file,
    insert_file_exploded,
    save_file
)

if __name__ == "__main__":
    start_cad_session()

    # 先恢复到单文件不确定状态
    restore_to_uncertain_state()

    try:
        # 打开目标文件
        open_file("D:/claude-tasks/tests/test_files/天正测试文件3.dwg")

        # 将B.dwg以炸开方式原位插入（0,0,0）
        insert_file_exploded(
            source_file="D:/claude-tasks/tests/test_files/B.dwg",
            x=0, y=0, z=0,
            scale=1.0
        )

        # 保存文件
        save_file()

        print("[成功] B.dwg内容已炸开插入到天正测试文件3.dwg并保存")

    finally:
        # 任务完成后恢复到单文件不确定状态
        restore_to_uncertain_state()
