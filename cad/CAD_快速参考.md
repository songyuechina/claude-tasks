# CAD 开发快速参考指南

**新对话必读文档** - 所有CAD相关任务开始前必须阅读本文档

---

## 🚨 核心原则（必须遵守）

### 1. CAD启动方式
- ❌ **禁止**直接启动CAD
- ✅ **必须**通过天正软件启动：`C:/Tangent/TArchT20V9/TGStart.exe`
- ✅ 使用函数：`start_cad_with_dialog_killer()` （包含弹窗治理）

### 2. 协同机制（最重要！）
**CAD是物理驱动的软件，命令执行需要时间**

- ✅ **必须**等待每个命令完全执行完成后才能执行下一个
- ✅ **必须**使用协同机制函数，确保命令顺序执行
- ❌ **禁止**快速连续发送多个命令

### 3. 弹窗治理
- ✅ 每次启动CAD后**必须**同时启动弹窗治理脚本
- 脚本位置：`D:/claude-tasks/scripts/cad_dialog_killer.py`
- 自动功能：每15秒检测并关闭干扰弹窗

---

## 📚 文档结构

### 核心规范文档
1. **CAD_操作规范.md** - 完整的六大核心规则（必读）
2. **CAD 协同机制实现报告.md** - 详细实现说明
3. **即时对话.txt** - 原始需求文档

### 代码模块
1. **CAD-basic.py** - 基础CAD操作函数（已集成协同机制）
2. **CAD_coordination.py** - 协同机制核心模块
3. **CAD_enhanced_functions.py** - 增强功能模块（推荐使用）
4. **cad_dialog_killer.py** - 弹窗治理脚本

### 资料参考
- **ziliao/20251010-0143-CAD开发/** - 完整的CAD开发规范和示例

---

## 🎯 CAD四个核心状态

| 状态名称 | 定义 | 使用场景 |
|---------|------|---------|
| 单文件不确定状态 | 单进程+1张未保存空白图 | 测试前归位、异常恢复 |
| 单文件确定状态 | 单进程+1张指定DWG | 单文件精确操作 |
| 双文件确定状态 | 单进程+2张指定DWG | 文件对比、跨图操作 |
| 多文件状态 | 单进程+多个DWG | 批量处理（禁止重复打开同一文件） |

---

## 🔧 常用操作函数（推荐使用）

### 1. 启动CAD会话
```python
from CAD_enhanced_functions import start_cad_session_with_coordination

# 启动完整的CAD会话（包含弹窗治理、单进程确保、空闲等待）
start_cad_session_with_coordination()
```

### 2. 🏗 CAD基本操作范式(必须使用)
```python
from CAD_basic_operations import (
    new_dwg_enhanced,           # 新建文件范式
    open_dwg_paradigm,         # 打开文件范式
    close_current_dwg_paradigm, # 关闭文件范式
    save_current_dwg_paradigm,  # 保存文件范式
    insert_dwg_as_block_paradigm # 插入块范式
)

# 🆕 新建文件范式
new_dwg_enhanced("D:/output/new_file.dwg")  # 新建并保存
new_dwg_enhanced()                           # 新建未保存文件

# 📂 打开文件范式(推荐使用)
open_dwg_paradigm("D:/path/to/file.dwg")    # 单文件打开
open_multiple_files_paradigm([files])        # 多文件顺序打开

# 💾 保存文件范式
save_current_dwg_paradigm()                  # 保存当前文件
save_as_dwg_paradigm("D:/output/save.dwg")   # 另存为

# 🔄 关闭文件范式
close_current_dwg_paradigm("auto_save")      # 关闭并自动保存
close_current_dwg_paradigm("no_save")        # 关闭不保存
close_all_dwg_paradigm()                     # 关闭所有文件

# 📦 插入块范式
insert_dwg_as_block_paradigm(
    "D:/blocks/item.dwg",
    insert_point=(100, 50, 0),
    scale=1.0,
    rotation=45.0,
    explode=False
)

# 🔄 标准工作流范式(推荐)
from CAD_basic_operations import standard_workflow_paradigm

# 完整工作流:打开 → 插入块 → 保存 → 关闭
standard_workflow_paradigm(
    source_file="D:/source/base.dwg",
    block_files=[
        {
            'path': 'D:/blocks/furniture.dwg',
            'point': (100, 100, 0),
            'scale': 1.0,
            'explode': False
        }
    ],
    output_file="D:/output/result.dwg"
)
```

### 2.1 📂 传统打开文件方式(备选)
```python
from CAD_enhanced_functions import open_dwg_sync

# 推荐：同步打开文件（最安全）
success = open_dwg_sync("D:/path/to/file.dwg", visible=True)

# 或者：需要返回acad和doc对象
from CAD_enhanced_functions import open_dwg_enhanced
acad, doc = open_dwg_enhanced("D:/path/to/file.dwg")
```

### 3. 发送CAD命令
```python
from CAD_coordination import send_cmd_with_sync

# 推荐：同步发送命令（确保执行完成）
success = send_cmd_with_sync("_.LINE\n0,0\n100,100\n\n", wait_after=1.0)

# 不推荐：旧方法（不等待完成）
# send_cmd("_.LINE\n0,0\n100,100\n\n")
```

### 4. 等待CAD空闲
```python
from CAD_coordination import wait_quiescent

# 等待CAD进入空闲状态
wait_quiescent(min_quiet=0.5, timeout=30.0)
```

### 5. 确保单进程
```python
from CAD_coordination import ensure_single_process

# 确保只有一个CAD进程运行
ensure_single_process()
```

---

## 📋 标准操作流程

### 流程1：启动CAD并打开文件
```python
from CAD_enhanced_functions import (
    start_cad_session_with_coordination,
    open_dwg_sync
)

# 1. 启动CAD会话
start_cad_session_with_coordination()

# 2. 打开DWG文件
open_dwg_sync("D:/path/to/file.dwg")
```

### 流程2：发送命令并等待
```python
from CAD_coordination import send_cmd_with_sync, wait_quiescent

# 1. 发送命令
send_cmd_with_sync("_.LINE\n0,0\n100,100\n\n")

# 2. 等待空闲（如需要）
wait_quiescent()

# 3. 发送下一个命令
send_cmd_with_sync("_.CIRCLE\n50,50\n25\n")
```

### 流程3：完整的CAD操作任务
```python
from CAD_enhanced_functions import *

# 1. 启动CAD会话（包含所有初始化）
if not start_cad_session_with_coordination():
    print("CAD启动失败")
    exit(1)

# 2. 打开文件
if not open_dwg_sync("D:/test.dwg"):
    print("文件打开失败")
    exit(1)

# 3. 执行操作
send_cmd_with_sync("_.ZOOM\n_E\n")
send_cmd_with_sync("_.LINE\n0,0\n100,100\n\n")

# 4. 等待完成
wait_quiescent(min_quiet=1.0)

print("操作完成")
```

---

## ⚠️ 常见错误和解决方案

### 错误1：命令执行失败
**原因**：命令发送太快，CAD还在处理上一个命令

**解决**：
```python
# ❌ 错误做法
send_cmd("_.LINE\n0,0\n100,100\n\n")
send_cmd("_.CIRCLE\n50,50\n25\n")  # 太快！

# ✅ 正确做法
send_cmd_with_sync("_.LINE\n0,0\n100,100\n\n", wait_after=1.0)
send_cmd_with_sync("_.CIRCLE\n50,50\n25\n", wait_after=1.0)
```

### 错误2：文件打开失败
**原因**：多个CAD进程或文件已被打开

**解决**：
```python
# ✅ 先确保单进程
ensure_single_process()

# ✅ 使用同步打开函数
open_dwg_sync("D:/file.dwg")
```

### 错误3：弹窗导致卡住
**原因**：未启动弹窗治理脚本

**解决**：
```python
# ✅ 使用包含弹窗治理的启动函数
start_cad_session_with_coordination()
```

---

## 🔍 调试和日志

### 日志位置
- CAD协同机制日志：控制台输出（带✅❌⚠等标记）
- 弹窗治理日志：`D:/claude-tasks/scripts/cad_dialog_killer.log`

### 调试技巧
```python
# 1. 测试协同机制
from CAD_enhanced_functions import test_cad_coordination
test_cad_coordination()

# 2. 检查CAD进程数
from CAD_coordination import ensure_single_process
ensure_single_process()

# 3. 测试等待空闲
from CAD_coordination import wait_quiescent
success = wait_quiescent(min_quiet=0.5, timeout=15.0)
print(f"空闲检测: {success}")
```

---

## 📖 深入学习路径

### 新对话开始时的学习顺序：
1. **本文档** - 快速参考（5分钟）
2. **CAD_操作规范.md** - 详细规范（10分钟）
3. **CAD 协同机制实现报告.md** - 实现细节（按需）
4. **ziliao/20251010-0143-CAD开发/docs/INDEX.md** - 完整规范（深入学习）

---

## ✅ 检查清单（每次CAD任务前）

- [ ] 已阅读本快速参考指南
- [ ] 理解协同机制的重要性（命令必须等待完成）
- [ ] 知道使用增强函数模块（CAD_enhanced_functions.py）
- [ ] 知道如何启动CAD会话（start_cad_session_with_coordination）
- [ ] 知道如何打开文件（open_dwg_sync）
- [ ] 知道如何发送命令（send_cmd_with_sync）
- [ ] 知道弹窗治理会自动启动

---

## 🎯 快速答案

**Q: 我要打开一个dwg文件，用什么函数？**
A: `open_dwg_sync("D:/path/to/file.dwg")`

**Q: 我要发送CAD命令，用什么函数？**
A: `send_cmd_with_sync("命令内容", wait_after=1.0)`

**Q: 我要启动CAD，用什么函数？**
A: `start_cad_session_with_coordination()`

**Q: 为什么命令执行失败？**
A: 可能没有等待上一个命令完成，使用带`_sync`后缀的函数

**Q: 在哪里找完整的规范？**
A: `CAD_操作规范.md` 和 `ziliao/20251010-0143-CAD开发/docs/`

---

**最后提醒**：所有CAD操作必须遵循协同机制，确保命令顺序执行！