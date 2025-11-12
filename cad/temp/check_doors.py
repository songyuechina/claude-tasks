#!/usr/bin/env python3
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent / "scripts"))

from CAD_basic import close_all_cad_processes, start_applicationV9, get_acad_doc
from CAD_file_operations import open_file
from CAD_coordination import ensure_single_process, wait_quiescent

close_all_cad_processes()
proc = start_applicationV9(PTH=r"C:\Tangent\TArchT20V9")
ensure_single_process()
wait_quiescent(min_quiet=2.0, timeout=30.0)

open_file("D:/claude-tasks/tests/test_files/天正测试文件2.dwg")
wait_quiescent(min_quiet=1.0, timeout=10.0)

_, doc = get_acad_doc()
ms = doc.ModelSpace

print(f"总对象数: {ms.Count}")

doors = []
for i in range(ms.Count):
    try:
        obj = ms.Item(i)
        if obj.ObjectName == "TDbOpening":
            bbox = obj.GetBoundingBox()
            p = bbox[0]
            doors.append((p[0], p[1]))
    except:
        pass

print(f"门窗对象数: {len(doors)}")
for i, (x, y) in enumerate(doors):
    print(f"  门{i+1}: ({x:.0f}, {y:.0f})")

close_all_cad_processes()
