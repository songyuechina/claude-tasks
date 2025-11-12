"""创建空白文件：天正测试文件3.dwg"""
import sys
sys.path.insert(0, 'D:/claude-tasks/cad/scripts')

from CAD_file_operations import new_file
import time

print("正在启动天正CAD并创建空白文件...")
print("请稍候，CAD窗口即将打开...")

# 创建空白文件
new_file('D:/claude-tasks/tests/test_files/天正测试文件3.dwg')

print("\n=== 创建完成 ===")
print("文件路径: D:/claude-tasks/tests/test_files/天正测试文件3.dwg")
print("CAD窗口将保持打开状态，请检查...")

# 保持脚本运行，让用户查看
input("\n按回车键退出...")
