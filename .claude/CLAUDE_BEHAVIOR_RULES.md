# Claude 行为准则 - CAD 开发工作站

本文档定义 Claude 在 `D:\claude-tasks` 工作站的强制性行为准则。

## 🎯 工作站定位

这是一个专门用于 **CAD 自动化开发和测试** 的工作环境，所有与 CAD 相关的操作都必须遵循严格的规范体系。

## 🚨 强制性行为规则

### 规则 1：CAD 操作前必读文档

在执行任何 CAD 相关操作前，你**必须**先阅读以下文档：

1. `CAD_操作规范.md` - 了解六大核心规则
2. `CAD_基本操作范式.md` - 了解五大基本操作范式

**触发条件：**
- 用户提到"CAD"、"天正"、"DWG"、"打开文件"、"插入块"等关键词
- 用户要求进行 CAD 测试
- 用户要求创建或修改 CAD 相关脚本

**执行动作：**
```
在执行操作前，我需要先阅读相关规范文档，确保按照正确的流程操作。

正在阅读：
- CAD_操作规范.md
- CAD_基本操作范式.md
[使用 Read 工具读取文档]

文档阅读完成，现在按照规范执行操作...
```

### 规则 2：必须使用范式函数

所有 CAD 基本操作（新建、打开、关闭、保存、插入）**必须**使用范式函数。

**禁止行为：**
- ❌ 直接使用 `acad.Documents.Open()`
- ❌ 直接使用 `doc.Save()`
- ❌ 直接发送 CAD 命令而不等待

**强制要求：**
- ✅ 使用 `open_dwg_paradigm()`
- ✅ 使用 `save_current_dwg_paradigm()`
- ✅ 使用 `send_cmd_with_sync()` + `wait_quiescent()`

**代码模板：**
```python
# 正确的方式
from CAD_basic_operations import open_dwg_paradigm
result = open_dwg_paradigm("D:/test.dwg")

# 错误的方式（禁止使用）
import win32com.client
acad = win32com.client.GetActiveObject("AutoCAD.Application")
acad.Documents.Open("D:/test.dwg")  # ❌ 禁止
```

### 规则 3：协同机制强制执行

CAD 是物理驱动的软件，**每个命令后必须等待完成**。

**强制模式：**
```python
# 每个操作的标准模式
from CAD_coordination import send_cmd_with_sync, wait_quiescent

# 1. 等待CAD空闲
wait_quiescent(min_quiet=0.5, timeout=15.0)

# 2. 发送同步命令
send_cmd_with_sync("_COMMAND\n", wait_after=1.0, timeout=30.0)

# 3. 再次等待完成
wait_quiescent(min_quiet=1.0, timeout=30.0)
```

**禁止模式：**
```python
# ❌ 错误：快速连续发送命令
send_command("_LINE\n")
send_command("0,0\n")
send_command("100,100\n")
send_command("\n")
```

### 规则 4：启动 CAD 的标准流程

启动 CAD 时，必须按照以下顺序：

```python
# 标准启动流程
from CAD_enhanced_functions import start_cad_session_with_coordination

# 这个函数会自动：
# 1. 通过天正软件启动 CAD（不直接启动 CAD）
# 2. 启动弹窗治理脚本
# 3. 确保单进程
# 4. 等待 CAD 进入空闲状态
start_cad_session_with_coordination()
```

**禁止：**
- ❌ 直接启动 `acad.exe`
- ❌ 不启动弹窗治理脚本
- ❌ 不确保单进程

### 规则 5：测试任务的标准流程

如果是测试任务，必须：

**1. 阅读测试规范**
```
正在阅读 CAD_测试规范.md...
```

**2. 使用测试框架**
```python
from CAD_test_framework import CADTestFramework

# 使用框架而不是随意编写测试
```

**3. 遵循测试八大原则**
- 测试前归位到"单文件不确定状态"
- 每个测试独立运行
- 详细记录测试日志
- 异常时自动恢复环境
- 等等...

### 规则 6：错误处理标准

出现错误时的标准处理流程：

```python
try:
    # CAD 操作
    result = open_dwg_paradigm(file_path)
    if not result:
        print("❌ 操作失败")
        # 尝试恢复状态
        from CAD_coordination import ensure_single_process
        ensure_single_process()
        return False

except Exception as e:
    print(f"❌ 操作异常: {e}")
    # 记录详细日志
    import traceback
    traceback.print_exc()
    # 尝试恢复到安全状态
    return False
```

### 规则 7：代码创建规范

创建 CAD 操作脚本时：

**必须包含的部分：**
```python
#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
脚本功能描述
"""

import sys
from pathlib import Path

# 添加 scripts 目录到路径
sys.path.insert(0, str(Path(__file__).parent / "scripts"))

from CAD_basic_operations import (
    # 导入需要的范式函数
)

from CAD_coordination import (
    wait_quiescent,
    send_cmd_with_sync,
    ensure_single_process
)

def main():
    """主函数"""
    print("="*60)
    print("脚本名称")
    print("="*60)

    # 1. 确保 CAD 环境就绪
    if not ensure_single_process():
        print("❌ CAD 环境初始化失败")
        return False

    # 2. 等待空闲
    wait_quiescent(min_quiet=0.5, timeout=15.0)

    # 3. 执行操作（使用范式函数）
    # ...

    print("\n✅ 脚本执行完成")
    return True

if __name__ == "__main__":
    try:
        success = main()
        sys.exit(0 if success else 1)
    except Exception as e:
        print(f"❌ 脚本执行异常: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
```

### 规则 8：日志输出规范

所有操作必须有清晰的日志输出：

```python
# 操作开始
print("🔄 正在打开文件: test.dwg")

# 关键步骤
print("  - 检查文件是否存在...")
print("  - 等待 CAD 空闲...")
print("  - 执行打开命令...")

# 操作结果
print("✅ 文件打开成功")
# 或
print("❌ 文件打开失败: 原因")
```

**日志图标规范：**
- 🔄 正在执行
- ✅ 成功
- ❌ 失败
- ⚠️  警告
- 📂 文件操作
- 💾 保存操作
- 🚀 启动操作
- 📊 统计信息

### 规则 9：文件路径规范

处理文件路径时：

```python
from pathlib import Path

# 1. 使用 Path 对象
file_path = Path("D:/test/file.dwg")

# 2. 检查文件存在
if not file_path.exists():
    print(f"❌ 文件不存在: {file_path}")
    return False

# 3. 转换为字符串传递
result = open_dwg_paradigm(str(file_path))

# 4. 处理中文路径（使用短路径）
from CAD_basic_operations import _get_short_path
short_path = _get_short_path(str(file_path))
```

### 规则 10：对象属性访问规范

访问 CAD 对象属性时，必须根据对象类型选择正确的方法：

**CAD 标准对象**（多段线、圆、直线、标注等）：

必须先使用 `_maybe_cast()` 转换到专用接口，然后才能访问属性：

```python
from CAD_basic import _maybe_cast, last_obj

# 获取对象
ob = last_obj()

# 必须先转换！
ob = _maybe_cast(ob)

# 现在可以访问属性了
coords = ob.Coordinates  # 多段线坐标
start = ob.StartPoint    # 直线起点
center = ob.Center       # 圆心
radius = ob.Radius       # 半径

# 获取边界框
min_pt, max_pt = ob.GetBoundingBox()
```

**天正对象**（门窗、墙体等）：

必须使用 `get_object_property()` 和 `set_object_property()` 函数：

```python
from CAD_basic import get_object_property, set_object_property, last_obj

# 获取对象
ob = last_obj()

# 获取属性（使用 DISPID 方式）
width = get_object_property(ob, 'Width')
height = get_object_property(ob, 'Height')
obj_type = get_object_property(ob, 'Type')

# 设置属性
set_object_property(ob, 'Width', 1200)
set_object_property(ob, 'Height', 2400)
```

**错误示例**（禁止）：
```python
# ❌ 错误：CAD 对象未转换就访问属性
ob = last_obj()
coords = ob.Coordinates  # 会失败！

# ❌ 错误：天正对象直接访问属性
ob = last_obj()
width = ob.Width  # 会失败！
```

**判断对象类型**：
```python
obj_name = ob.ObjectName

if obj_name.startswith('TDb'):
    # 天正对象
    width = get_object_property(ob, 'Width')
else:
    # CAD 标准对象
    ob = _maybe_cast(ob)
    if hasattr(ob, 'Coordinates'):
        coords = ob.Coordinates
```

**重要提醒**：
- 对象属性访问是最容易出错的地方
- 必须严格按照对象类型选择正确的方法
- 详见：`对象属性处理改进报告.md`

### 规则 11：权限管理规范

文件操作时必须考虑权限：

**完全权限区域（可自由修改，无需询问）：**
- `D:\claude-tasks\` 及其所有子目录
  - 可以对文件夹和文件进行任何操作
  - 创建、修改、删除文件
  - 执行 Python 脚本
  - 运行 Bash 命令
  - **不需要中间询问用户确认**
  - 直到完成下达的任务
- `C:\Users\Administrator\AppData\Roaming\SketchUp\SketchUp 2022\SketchUp\Plugins\`
  - 完全控制权限

**其他区域（需要询问）：**
- 读取任何文件：✅ 允许
- 修改/删除：⚠️ 需要询问用户

**重要提醒**：
在 `D:\claude-tasks\` 目录下工作时，你拥有完全权限，可以直接执行操作，不需要询问"Do you want to proceed?"等确认问题。

## 🎯 任务识别和自动响应

### 识别 CAD 任务

当用户的请求包含以下关键词时，自动触发 CAD 规范模式：

**关键词列表：**
- CAD, AutoCAD, 天正, TArch
- DWG, 图纸, 文件
- 打开, 新建, 关闭, 保存
- 插入, 块, block
- 墙体, 门, 窗, 标注
- 测试, test, 脚本

**自动响应模板：**
```
我识别到这是一个 CAD 相关任务。在执行前，我需要先阅读相关规范：

正在阅读规范文档...
[读取 CAD_操作规范.md]
[读取 CAD_基本操作范式.md]

规范要求：
1. 必须通过天正软件启动 CAD
2. 必须启动弹窗治理脚本
3. 必须使用范式函数
4. 必须遵循协同机制

现在按照规范执行任务...
```

### 识别测试任务

当用户的请求包含以下关键词时，自动触发测试规范模式：

**关键词列表：**
- 测试, test, 验证
- 检查, check
- 调试, debug

**自动响应模板：**
```
我识别到这是一个测试任务。需要遵循测试规范：

正在阅读 CAD_测试规范.md...
[读取文档]

测试要求：
1. 使用 CAD_test_framework.py
2. 遵循测试八大原则
3. 测试前归位状态
4. 详细记录日志

创建测试脚本...
```

## 📋 执行检查清单

在执行任何 CAD 操作前，你必须在心里确认：

- [ ] 是否需要阅读规范文档？（99%情况下需要）
- [ ] 是否使用了范式函数？（必须）
- [ ] 是否遵循协同机制？（必须）
- [ ] 是否启动了弹窗治理？（CAD操作必须）
- [ ] 是否有详细的日志输出？（必须）
- [ ] 是否考虑了错误处理？（必须）
- [ ] 路径是否正确处理？（必须）

## 🔄 规范文档更新机制

当发现规范需要改进时：

1. **记录问题**：在操作过程中发现的问题
2. **建议改进**：提出规范的改进建议
3. **等待确认**：询问用户是否更新规范
4. **更新文档**：获得确认后更新相关文档

**更新模板：**
```
我注意到当前规范在 XXX 方面可能需要改进：

问题：[描述问题]
建议：[改进建议]
影响：[影响范围]

是否需要更新规范文档？
```

## 💡 最佳实践示例

### 示例 1：打开文件任务

```
用户：打开 D:/test/sample.dwg

Claude：
我识别到这是 CAD 文件操作任务。

第一步：阅读操作规范
[读取 CAD_操作规范.md 中的启动方式和物理驱动等待机制部分]
[读取 CAD_基本操作范式.md 中的打开文件范式部分]

第二步：创建操作脚本
按照范式要求，我会使用 open_dwg_paradigm() 函数...

[创建脚本，包含所有必需的步骤]

第三步：执行脚本
[执行并输出详细日志]
```

### 示例 2：测试任务

```
用户：测试 CAD 打开和关闭文件功能

Claude：
我识别到这是测试任务。

第一步：阅读测试规范
[读取 CAD_测试规范.md]

第二步：使用测试框架
我会使用 CAD_test_framework.py 创建标准测试...

[创建符合测试八大原则的测试脚本]

第三步：执行测试
[执行并记录详细日志]
```

## ✅ 规则确认

每次启动会话时，如果用户使用 `/cad` 命令或提到 CAD 相关任务，你应该：

1. **确认规范**：表示你已理解并会遵循所有规范
2. **询问任务**：询问用户的具体需求
3. **读取文档**：在执行前读取相关规范文档
4. **严格执行**：按照规范执行，不得偏离

---

**本文档是强制性的。违反这些规则会导致 CAD 操作失败、数据丢失或系统不稳定。**

**版本：1.0**
**更新日期：2025-11-10**
