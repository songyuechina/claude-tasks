import sys, io, math
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, r"D:\claude-tasks\cad\scripts")

from CAD_enhanced_functions import start_cad_session_with_coordination
from CAD_file_operations import new_file, draw_tarch_wall, insert_tarch_window

start_cad_session_with_coordination()
new_file(r"D:\claude-tasks\tests\test_files\天正测试文件3.dwg")

# 三角形顶点（边长>12000）
L = 15000
pts = [(0, 0, 0), (L, 0, 0), (L/2, L*math.sqrt(3)/2, 0)]

# 绘制三角形墙体
print("绘制三角形墙体...")
for i in range(3):
    print(f"墙{i+1}: {pts[i][:2]} -> {pts[(i+1)%3][:2]}")
    draw_tarch_wall(pts[i], pts[(i+1)%3], thickness=240)

# 计算中点并插入窗
windows = [
    ((pts[0][0]+pts[1][0])/2, (pts[0][1]+pts[1][1])/2, 0, 1000, "jz-menlianchuang"),
    ((pts[1][0]+pts[2][0])/2, (pts[1][1]+pts[2][1])/2, 0, 1200, "jz-baiyechuang"),
    ((pts[2][0]+pts[0][0])/2, (pts[2][1]+pts[0][1])/2, 0, 1800, "jz-tuilamen")
]

print("\n插入窗...")
for x, y, z, w, layer in windows:
    print(f"位置:({x:.0f},{y:.0f}), 宽度:{w}, 图层:{layer}")
    result = insert_tarch_window((x, y, z), w, layer)
    print(f"结果: {result}")

print("\n完成！")
