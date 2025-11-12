# CAD自动化脚本说明文档

本目录包含天正建筑CAD自动化操作的核心脚本和功能模块。

## 📋 目录结构

### 🔴 核心脚本（必需）

#### 1. **CAD_basic.py**
- **功能**：CAD基础操作函数库
- **说明**：包含所有CAD基本操作函数，如对象选择、属性获取/设置、图层操作、绘图函数等
- **依赖**：pywin32, pywinauto
- **重要性**：★★★★★ 核心基础库，其他所有脚本都依赖此文件
- **主要功能**：
  - 对象选择和操作（stc, last_obj等）
  - 属性传递（transfer_props_by_matchprop）
  - 图层管理
  - 绘图函数（绘制多段线、矩形等）
  - 窗口管理
  - 文件操作辅助函数

#### 2. **CAD_file_operations.py**
- **功能**：CAD文件操作统一接口
- **说明**：提供DWG文件的打开、新建、关闭、保存、另存、插入等高级操作
- **依赖**：CAD_basic.py, CAD_basic_operations.py, CAD_coordination.py
- **重要性**：★★★★★ 核心文件操作库
- **主要功能**：
  - 文件新建与打开（new_file, open_file）
  - 文件保存（save_file, save_file_as）
  - 文件关闭（close_file, close_all_files）
  - 文件插入（insert_file_as_block, copy_file_content_pywin32）
  - 天正墙绘制（draw_tarch_wall）
  - 天正门插入（insert_tarch_door）
  - 天正窗插入（insert_tarch_window）
  - CAD会话管理（start_cad_session, restore_to_uncertain_state）

#### 3. **CAD_basic_operations.py**
- **功能**：CAD基础操作范式
- **说明**：提供文件操作的底层实现函数
- **依赖**：CAD_coordination.py
- **重要性**：★★★★☆ 核心依赖库
- **主要功能**：
  - 新建DWG文件（new_dwg_enhanced）
  - 打开DWG文件（open_dwg_paradigm）
  - 保存DWG文件（save_current_dwg_paradigm, save_as_dwg_paradigm）
  - 关闭DWG文件（close_current_dwg_paradigm, close_all_dwg_paradigm）
  - 插入DWG文件（insert_dwg_as_block_paradigm, insert_and_explode_paradigm）

#### 4. **CAD_coordination.py**
- **功能**：CAD协同机制
- **说明**：提供CAD命令发送、等待、进程管理等协同功能
- **依赖**：无
- **重要性**：★★★★☆ 核心依赖库
- **主要功能**：
  - 命令发送（send_cmd_with_sync）
  - 等待空闲（wait_quiescent）
  - 进程管理（ensure_single_process）
  - 窗口管理

#### 5. **cad_dialog_killer.py**
- **功能**：CAD弹窗治理脚本
- **说明**：后台运行，自动关闭CAD的各种弹窗（缺少SHX文件、错误报告等）
- **依赖**：pywinauto
- **重要性**：★★★★☆ 自动化必需
- **使用方式**：由start_cad_session()自动启动

### 🟢 功能模块

#### 6. **insert_tarch_window_simple.py**
- **功能**：简化的天正窗绘制函数
- **说明**：假设MC_yuan.dwg已在文件中，通过插入门→查找窗→传递属性→设置宽度的方式插入天正窗
- **依赖**：CAD_basic.py, CAD_coordination.py
- **重要性**：★★★☆☆ 功能模块
- **主要函数**：
  - `insert_tarch_window_simple(p, window_type, width, tolerance=10)`
  - 支持的窗类型：jz-menlianchuang, jz-dong, jz-gaochuang, jz-baiyechuang, jz-tuchuang, jz-pingchuang, jz-zimumen, jz-juanlianmen, jz-tuilamen, jz-shuangmen

### 🟡 辅助脚本

#### 7. **CAD_enhanced_functions.py**
- **功能**：CAD增强功能
- **说明**：提供额外的CAD操作增强功能
- **依赖**：CAD_coordination.py
- **重要性**：★★☆☆☆

#### 8. **CAD_file_operations_example.py**
- **功能**：文件操作示例
- **说明**：演示如何使用CAD_file_operations.py中的函数
- **重要性**：★☆☆☆☆ 示例代码

#### 9. **CAD_object_properties.py**
- **功能**：CAD对象属性操作
- **说明**：提供对象属性的读取和设置功能
- **重要性**：★★☆☆☆

#### 10. **CAD_tarch_properties.py**
- **功能**：天正对象属性操作
- **说明**：专门处理天正CAD对象的属性
- **重要性**：★★☆☆☆

#### 11. **CAD_test_framework.py**
- **功能**：CAD测试框架
- **说明**：提供自动化测试的框架和工具
- **重要性**：★★☆☆☆

## 🔧 使用方法

### 基本使用流程

```python
from CAD_file_operations import (
    start_cad_session,
    new_file,
    draw_tarch_wall,
    insert_tarch_door,
    save_file_as,
    restore_to_uncertain_state
)

# 1. 启动CAD会话
start_cad_session()

# 2. 新建文件
new_file()

# 3. 绘制墙体
draw_tarch_wall((0, 0, 0), (10000, 0, 0), thickness=240)

# 4. 插入门
result = insert_tarch_door((5000, 0, 0), width=900, height=2100)

# 5. 保存文件
save_file_as("D:/test.dwg")

# 6. 恢复到不确定状态
restore_to_uncertain_state()
```

### 天正窗插入示例

```python
from CAD_file_operations import (
    start_cad_session,
    new_file,
    draw_tarch_wall,
    save_file_as,
    copy_file_content_pywin32
)
from insert_tarch_window_simple import insert_tarch_window_simple

# 1. 启动并创建文件
start_cad_session()
new_file()

# 2. 绘制墙
draw_tarch_wall((0, 0, 0), (10000, 0, 0))

# 3. 保存文件
save_file_as("D:/test.dwg")

# 4. 插入MC_yuan.dwg（包含窗对象）
copy_file_content_pywin32('D:/claude-tasks/xitongwenjian/MC_yuan.dwg', "D:/test.dwg")

# 5. 插入窗
result = insert_tarch_window_simple((5000, 0, 0), 'jz-pingchuang', 1200)
print(f"插入结果: {result}")
```

## 📝 重要规则

### CAD四个核心状态

| 状态名称           | 定义                      | 使用场景                 |
|-------------------|---------------------------|-------------------------|
| 单文件不确定状态   | 单进程+1张未保存空白图     | 测试前归位、异常恢复     |
| 单文件确定状态     | 单进程+1张指定DWG         | 单文件精确操作           |
| 双文件确定状态     | 单进程+2张指定DWG         | 文件对比、跨图操作       |
| 多文件状态         | 单进程+多个DWG            | 批量处理                 |

**重要规则**：
1. 每个任务后必须恢复到单文件不确定状态（使用`restore_to_uncertain_state()`）
2. 卡住了就关掉进程重启（使用 `taskkill /F /IM acad.exe`）

### 多文件操作规范

- 最多同时打开两个文件
- 使用 `CAD_basic.li()` 连接新打开的文件

```python
from CAD_basic import li
from CAD_file_operations import open_file

open_file("D:/file1.dwg")
open_file("D:/file2.dwg")
li()  # 连接到新打开的文件
```

## ⚠️ 注意事项

1. **emoji编码问题已修复**：CAD_basic.py中的142处emoji已替换为普通文本（[OK]、[FAIL]等），避免Windows gbk编码错误
2. **必须使用start_cad_session()启动CAD**：确保单进程、启动弹窗治理、正确初始化
3. **属性传递**：`transfer_props_by_matchprop`函数使用MATCHPROP命令传递属性，在某些情况下可能失败，需要重试机制
4. **文件路径**：所有路径使用正斜杠（/）或双反斜杠（\\\\）

## 📂 相关目录

- **D:/claude-tasks/temp_scripts/** - 临时测试脚本（已移出核心目录）
- **D:/claude-tasks/xitongwenjian/** - 系统文件（如MC_yuan.dwg）
- **D:/claude-tasks/tests/test_files/** - 测试文件输出目录

## 🔄 更新日志

- **2025-11-11**：修复emoji编码问题（142处）
- **2025-11-11**：创建insert_tarch_window_simple.py简化窗函数
- **2025-11-11**：整理脚本目录，移动临时脚本到temp_scripts/

---

**维护者**：Claude Code
**最后更新**：2025-11-11
