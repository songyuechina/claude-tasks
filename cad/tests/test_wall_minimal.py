#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_basic import get_acad_doc
from CAD_coordination import send_cmd_with_sync
import time

print("测试墙体绘制")

try:
    # 测试获取CAD文档
    print("\n1. 获取CAD文档...")
    acad, doc = get_acad_doc()
    print(f"   成功: {doc.Name}")

    # 测试发送命令
    print("\n2. 发送墙体命令...")
    p1 = (88800.42585138, 77306.33788321, 0)
    p2 = (94193.69907482, 82695.99449027, 0)

    cmd = f"tgwall\n{p1[0]},{p1[1]}\n{p2[0]},{p2[1]}\n\n"
    print(f"   命令: {cmd[:50]}...")

    send_cmd_with_sync(cmd, wait_after=3.0)
    print("   墙体绘制完成")

    print("\n测试成功!")

except Exception as e:
    print(f"\n错误: {e}")
    import traceback
    traceback.print_exc()
