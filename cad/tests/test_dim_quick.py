#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_basic import dim_by_points

p1 = (113237.38580577, 58536.68585148, 0)
p2 = (141299.19187486, 71150.16060377, 0)
p3 = (120686.37419487, 60053.91084175, 0)

result = dim_by_points(p1, p2, p3)
print(f"\n标注结果: {'成功' if result else '失败'}")
