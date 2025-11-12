# insert_tarch_window 函数测试报告

## 测试时间
2025-11-11

## 测试目标
测试 `D:/claude-tasks/scripts/test_insert_tarch_window.py` 中的 `insert_tarch_window()` 函数

## 测试文件
- 测试DWG文件: `D:/claude-tasks/Function_testing/insert_tarch_window.dwg`
- 测试副本: `insert_tarch_window-1.dwg`, `insert_tarch_window-2.dwg`

## 测试步骤

### 1. 环境准备
```python
# 关闭所有CAD进程
ensure_single_process()

# 启动天正CAD
start_applicationV9(
    PTH=r"C:\Tangent\TArchT20V9",
    max_retries=3,
    retry_delay=2.0
)
```

### 2. 打开测试文件
```python
# 关闭所有已打开的文档
while acad.Documents.Count > 0:
    doc = acad.Documents.Item(0)
    doc.Close(False)

# 打开测试文件
open_file("D:/claude-tasks/Function_testing/insert_tarch_window-1.dwg")

# 连接当前文件
li()
```

### 3. 测试用例

#### 测试1: 插入 jz-gaochuang 窗
- 位置: (38612.86565445, 48750.63891910, 0)
- 宽度: 1200
- 高度: 1000
- 类型: jz-gaochuang
- delete_mc_yuan: False

#### 测试2: 插入 jz-pingchuang 窗
- 位置: (44695.30568975, 46646.78059028, 0)
- 宽度: 2400
- 高度: 1800
- 类型: jz-pingchuang
- delete_mc_yuan: True

## 发现的问题

### 问题1: insert_tarch_door 函数的对象获取问题

**问题描述:**
`CAD_file_operations.py` 中的 `insert_tarch_door()` 函数使用 `last_obj()` 获取刚插入的门对象，但天正的 `TOpening` 命令会创建多个对象（包括多段线等辅助对象），`last_obj()` 返回的是最后创建的对象，不一定是门对象（TDbOpening）。

**错误信息:**
```
[警告] 插入的对象不是门: AcDbPolyline
```

**解决方案:**
修改 `insert_tarch_door()` 函数，在执行命令前后记录对象数量，然后遍历新增的对象，找到 `ObjectName == "TDbOpening"` 的对象。

**修正后的代码:**
```python
def insert_tarch_door(p, width=None, height=None):
    try:
        # 获取插入前的对象数量
        _, doc = get_acad_doc()
        ms = doc.ModelSpace
        count_before = ms.Count

        # 发送TOpening命令插入门
        cmd = f"TOpening\n{p[0]},{p[1]}\n\n"
        send_cmd_with_sync(cmd, wait_after=4.0)

        # 等待对象创建完成
        time.sleep(3)

        # 遍历新增的对象，找到门对象
        door = None
        count_after = ms.Count
        for i in range(count_before, count_after):
            obj = ms.Item(i)
            if obj.ObjectName == "TDbOpening":
                door = obj
                break

        if door is None:
            print(f"[错误] 未找到门对象，新增了 {count_after - count_before} 个对象")
            return {'success': False, 'door': None, 'width': None, 'height': None}

        # 读取和设置尺寸...
        # (其余代码保持不变)
```

### 问题2: TOpening 命令未创建对象

**问题描述:**
执行 `TOpening` 命令后，模型空间中的对象数量没有增加，说明命令没有创建门对象。

**观察到的现象:**
- 插入前对象数量: 5
- 插入后对象数量: 5
- 新增对象: 0

**可能的原因:**
1. 指定的坐标点没有墙体
2. TOpening 命令需要额外的参数或用户交互
3. 命令执行方式不正确
4. 测试文件中的墙体位置与指定坐标不匹配

**需要进一步调查:**
- 检查 MC_yuan.dwg 文件中墙体的实际位置
- 验证 TOpening 命令的正确使用方式
- 可能需要先选择墙体对象，再指定插入点

### 问题3: 文件打开和激活问题

**问题描述:**
使用 `open_file()` 打开文件后，当前激活的文档可能不是刚打开的文件，而是之前打开的文件或新建的空白文档。

**解决方案:**
在打开新文件之前，先关闭所有已打开的文档：
```python
import win32com.client
acad = win32com.client.GetActiveObject("AutoCAD.Application")
while acad.Documents.Count > 0:
    doc = acad.Documents.Item(0)
    doc.Close(False)  # False = 不保存
```

然后打开文件并显式激活：
```python
open_file(test_file)
# 激活文档
for doc in acad.Documents:
    if str(Path(doc.FullName).resolve()).lower() == target_path_lower:
        acad.ActiveDocument = doc
        break
```

## 测试状态

**当前状态:** 未完成

**原因:** TOpening 命令无法创建门对象，需要进一步调查命令的正确使用方式。

## 建议

1. **短期方案:**
   - 更新 `CAD_file_operations.py` 中的 `insert_tarch_door()` 函数，使用修正后的对象获取逻辑
   - 在函数文档中说明当前的限制和已知问题

2. **长期方案:**
   - 深入研究天正 TOpening 命令的正确使用方式
   - 可能需要使用不同的方法插入门窗（例如直接操作天正对象，而不是通过命令行）
   - 考虑使用天正的 API 而不是命令行接口

3. **测试文件准备:**
   - 创建一个包含已知位置墙体的测试文件
   - 确保测试坐标点确实有墙体存在
   - 记录墙体的准确位置和属性

## 参考文档

- `CAD_操作规范.md` - CAD操作六大核心规则
- `CAD_基本操作范式.md` - 五大基本操作范式
- `transfer_props_by_matchprop.md` - 属性匹配测试记录

## 使用方法（基于当前实现）

```python
from CAD_basic import li, start_applicationV9
from CAD_coordination import ensure_single_process, wait_quiescent
from test_insert_tarch_window import insert_tarch_window
from CAD_file_operations import open_file, save_file

# 1. 启动CAD
ensure_single_process()
start_applicationV9(PTH=r"C:\Tangent\TArchT20V9")
wait_quiescent(min_quiet=1.0, timeout=30.0)

# 2. 打开文件
open_file("D:/path/to/file.dwg")
li()

# 3. 插入窗
result = insert_tarch_window(
    p=(x, y, z),
    width=1200,
    height=1000,
    window_type="jz-gaochuang",
    delete_mc_yuan=False
)

# 4. 保存文件
if result['success']:
    save_file()
```

## 注意事项

1. 必须先确保测试文件中有墙体
2. 插入坐标必须在墙体上
3. 需要先插入 MC_yuan.dwg 以获取窗类型模板
4. TOpening 命令的使用方式可能需要调整

## 结论

虽然测试未能完全成功，但我们发现并修正了 `insert_tarch_door()` 函数中的对象获取问题。TOpening 命令无法创建对象的问题需要进一步调查，可能涉及到命令使用方式或测试环境的配置。

建议先更新 `CAD_file_operations.py` 中的函数，然后在实际项目环境中进行更深入的测试。
