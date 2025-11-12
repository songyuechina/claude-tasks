import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, r"D:\claude-tasks\cad\scripts")

from CAD_enhanced_functions import start_cad_session_with_coordination
from CAD_file_operations import new_file
from CAD_coordination import send_cmd_with_sync, wait_quiescent
from CAD_basic import get_acad_doc, last_obj
import time

start_cad_session_with_coordination()
new_file(r"D:\claude-tasks\tests\test_files\墙体测试.dwg")

_, doc = get_acad_doc()
ms = doc.ModelSpace
print(f"绘制前对象数: {ms.Count}")

# 直接发送tgwall命令
print("发送tgwall命令...")
send_cmd_with_sync("tgwall\n", wait_after=0.5)
send_cmd_with_sync("0,0\n", wait_after=0.5)
send_cmd_with_sync("10000,0\n", wait_after=0.5)
send_cmd_with_sync("\n", wait_after=1.0)
wait_quiescent(min_quiet=1.0, timeout=10.0)

print(f"绘制后对象数: {ms.Count}")

# 检查最后几个对象
for i in range(max(0, ms.Count-5), ms.Count):
    try:
        obj = ms.Item(i)
        print(f"对象{i}: {obj.ObjectName}")
    except:
        pass

print("完成！")
