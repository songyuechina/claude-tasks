#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

import pyautogui
import time

print("取消CAD当前操作...")

# 按ESC键多次取消当前操作
for i in range(5):
    pyautogui.press('esc')
    time.sleep(0.2)

print("等待3秒...")
time.sleep(3)

print("\n现在测试墙体绘制...")

from CAD_basic import get_acad_doc
from CAD_coordination import send_cmd_with_sync

try:
    acad, doc = get_acad_doc()
    print(f"成功获取文档: {doc.Name}")

    p1 = (88800.42585138, 77306.33788321, 0)
    p2 = (94193.69907482, 82695.99449027, 0)

    cmd = f"tgwall\n{p1[0]},{p1[1]}\n{p2[0]},{p2[1]}\n\n"
    send_cmd_with_sync(cmd, wait_after=3.0)

    print("墙体绘制完成!")

except Exception as e:
    print(f"错误: {e}")
