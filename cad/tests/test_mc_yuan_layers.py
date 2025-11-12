import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, r"D:\claude-tasks\cad\scripts")

from CAD_enhanced_functions import start_cad_session_with_coordination
from CAD_file_operations import open_file
from CAD_basic import get_acad_doc

start_cad_session_with_coordination()
open_file(r"D:/claude-tasks/cad/xitongwenjian/MC_yuan.dwg")

_, doc = get_acad_doc()
ms = doc.ModelSpace
layers = doc.Layers

print(f"\n文件中的对象数: {ms.Count}")

# 列出所有图层
print(f"\n文件中的图层数: {layers.Count}")
print("\n所有图层:")
for i in range(layers.Count):
    try:
        layer = layers.Item(i)
        print(f"  {i}: {layer.Name}")
    except:
        pass

# 列出所有对象及其图层
print("\n所有对象及其图层:")
target_layers = ["jz-menlianchuang", "jz-dong", "jz-gaochuang", "jz-baiyechuang",
                 "jz-tuchuang", "jz-pingchuang", "jz-zimumen", "jz-juanlianmen",
                 "jz-tuilamen", "jz-shuangmen"]

layer_counts = {layer: 0 for layer in target_layers}

for i in range(ms.Count):
    try:
        obj = ms.Item(i)
        obj_layer = obj.Layer
        obj_name = obj.ObjectName
        print(f"  对象{i}: {obj_name}, 图层: {obj_layer}")

        if obj_layer in target_layers:
            layer_counts[obj_layer] += 1
    except Exception as e:
        print(f"  对象{i}: 错误 - {e}")

print("\n目标图层统计:")
for layer, count in layer_counts.items():
    if count > 0:
        print(f"  {layer}: {count}个对象")

print("\n完成！")
