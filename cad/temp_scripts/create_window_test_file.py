#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys, time
from pathlib import Path
sys.path.append(str(Path(__file__).parent))

from CAD_file_operations import start_cad_session, new_file, save_file_as, draw_tarch_wall, copy_file_content_pywin32, save_file, restore_to_uncertain_state
from insert_tarch_window_simple import insert_tarch_window_simple
from CAD_coordination import wait_quiescent, send_cmd_with_sync

# 0. 启动天正CAD会话
start_cad_session()

# 1. 新建文件
new_file()
wait_quiescent(min_quiet=1.0, timeout=10.0)

# 2. 绘制三角形墙
p1, p2, p3 = (0, 0, 0), (15000, 0, 0), (7500, 13000, 0)
draw_tarch_wall(p1, p2)
draw_tarch_wall(p2, p3)
draw_tarch_wall(p3, p1)
wait_quiescent(min_quiet=1.0, timeout=10.0)

# 3. 保存文件
save_path = "D:/claude-tasks/tests/test_files/天正窗测试文件.dwg"
save_file_as(save_path)

# 4. 插入MC_yuan.dwg
copy_file_content_pywin32('D:/claude-tasks/cad/xitongwenjian/MC_yuan.dwg', save_path)
wait_quiescent(min_quiet=1.0, timeout=10.0)

# 5. 插入3个窗（带重试）
def insert_with_retry(pos, wtype, width, max_tries=3):
    for i in range(max_tries):
        try:
            r = insert_tarch_window_simple(pos, wtype, width)
            if r['success']:
                return r
            time.sleep(1)
        except:
            send_cmd_with_sync("u\n", wait_after=1.0)
            time.sleep(2)
    return {'success': False}

r1 = insert_with_retry((5000, 0, 0), "jz-tuilamen", 1200)
r2 = insert_with_retry((12500, 4333, 0), "jz-gaochuang", 900)
r3 = insert_with_retry((5000, 8667, 0), "jz-shuangmen", 1400)

save_file()
print(f"\n结果: 窗1={r1['success']}, 窗2={r2['success']}, 窗3={r3['success']}")

# 6. 恢复到不确定状态
restore_to_uncertain_state()
