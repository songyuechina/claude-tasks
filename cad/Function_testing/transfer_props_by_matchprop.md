1 测试记录

= RESTART: D:\claude-tasks\scripts\CAD_basic.py
[成功] CAD协同机制模块已加载
__________________  CAD基本操作开始运行 _________________________
li()
当前桌面文件： transfer_props_by_matchprop.dwg
win32已经连接正常—CAD基本操作
True
c1=HandleToObject('331')
m1=HandleToObject('331')
set_object_property(c1, "Width", 2000)
True
c1=HandleToObject('36B')
set_object_property(c1, "Width", 1200)
True
set_object_property(m1, "Width", 900)
True
transfer_props_by_matchprop(c1,m1,max_try=3, delay=0.4)
[OK] 已高亮选择区域 (45389.06653471782,16752.089372565126) ~ (46832.04180942132,17557.113422994844) 的对象
[OK] 已高亮目标对象
[OK] 第 1 次匹配成功，Layer 改为 WINDOW
True


2 使用方法
通过D:\claude-tasks\scripts\CAD_basic.py的li()连接当前激活文件。
通过HandleToObject('331')获取该文件中的窗c1，通过m1=HandleToObject('331')获取该文件中的门c1，调用
D:\claude-tasks\scripts\CAD_basic.py的函数transfer_props_by_matchprop，将门变成了窗。