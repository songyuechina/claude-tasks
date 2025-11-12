import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, r"D:\claude-tasks\cad\scripts")

from CAD_enhanced_functions import start_cad_session_with_coordination
from CAD_file_operations import new_file, draw_tarch_wall, insert_tarch_window
import math

start_cad_session_with_coordination()
new_file(r"D:\claude-tasks\tests\test_files\天正测试文件3.dwg")

# 检查对象数量
from CAD_basic import get_acad_doc
_, doc = get_acad_doc()
ms = doc.ModelSpace
print(f"绘制前对象数: {ms.Count}")

# 绘制一条墙
print("绘制墙体...")
draw_tarch_wall((0, 0, 0), (15000, 0, 0), thickness=240)

print(f"绘制后对象数: {ms.Count}")

# 检查最后几个对象
for i in range(max(0, ms.Count-3), ms.Count):
    try:
        obj = ms.Item(i)
        print(f"对象{i}: {obj.ObjectName}")
    except:
        pass

# 插入一个门测试
print("\n插入门测试...")
from CAD_file_operations import insert_tarch_door
result = insert_tarch_door((7500, 0, 0), width=1000)
print(f"门插入结果: {result}")

print("完成！")
