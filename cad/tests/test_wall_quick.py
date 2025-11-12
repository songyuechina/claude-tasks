#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_file_operations import draw_tarch_wall, save_file

p1 = (88800.42585138, 77306.33788321, 0)
p2 = (94193.69907482, 82695.99449027, 0)

result = draw_tarch_wall(p1, p2, thickness=320)
if result:
    save_file()
    print("\n测试完成，文件已保存")
