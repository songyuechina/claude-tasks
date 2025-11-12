import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, r"D:\claude-tasks\cad\scripts")

from CAD_basic import start_applicationV9, stc, transfer_props_by_matchprop
from CAD_file_operations import open_file

start_applicationV9(PTH=r"C:\Tangent\TArchT20V9", max_retries=3, retry_delay=2.0)
open_file(r"D:\claude-tasks\tests\test_files\天正测试文件2.dwg")

lb = stc("WINDOW")
print(f"选中 {len(lb)} 个对象")

objs = []
for obj in lb:
    minpt, maxpt = obj.GetBoundingBox()
    y = (minpt[1] + maxpt[1]) / 2
    objs.append((obj, y))
    print(f"Y坐标: {y}")

objs.sort(key=lambda x: x[1], reverse=True)
window, door = objs[0][0], objs[1][0]

print(f"\n窗Y={objs[0][1]:.2f}, 门Y={objs[1][1]:.2f}")
print("开始属性传递...")
transfer_props_by_matchprop(window, door)
print("完成！")
