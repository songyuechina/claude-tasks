import sys, io, math
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, r"D:\claude-tasks\cad\scripts")

from CAD_enhanced_functions import start_cad_session_with_coordination
from CAD_file_operations import new_file, draw_tarch_wall
from CAD_basic import get_acad_doc

start_cad_session_with_coordination()
new_file(r"D:\claude-tasks\tests\test_files\墙体测试.dwg")

_, doc = get_acad_doc()
ms = doc.ModelSpace

# 三角形顶点
L = 15000
pts = [(0, 0, 0), (L, 0, 0), (L/2, L*math.sqrt(3)/2, 0)]

print(f"绘制前对象数: {ms.Count}")

# 绘制三角形墙体
for i in range(3):
    print(f"\n墙{i+1}: {pts[i][:2]} -> {pts[(i+1)%3][:2]}")
    result = draw_tarch_wall(pts[i], pts[(i+1)%3], thickness=240)
    print(f"结果: {result}")

print(f"\n绘制后对象数: {ms.Count}")

# 检查对象
for i in range(ms.Count):
    try:
        obj = ms.Item(i)
        print(f"对象{i}: {obj.ObjectName}")
    except:
        pass

print("\n完成！")
