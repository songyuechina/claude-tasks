# CAD 测试规范

本文档定义CAD自动化测试的完整规范,所有测试任务必须严格遵守。

---

## 📋 测试规范八大原则

### 原则1: 测试任务定义
**测试任务是指**:对编写的脚本进行DWG文件的操作验证,包括:
- 文件基本操作:打开、关闭、保存、另存为
- 文件插入操作:块插入、图形插入
- 对象操作:创建、修改、删除、查询

### 原则2: 测试任务编号
- 每个测试任务必须编号:测试任务1、测试任务2、...
- **顺序执行**:完成当前任务才能执行下一个任务
- 编号格式:`TEST_TASK_001`、`TEST_TASK_002`等

### 原则3: 弹窗治理检查
**每个测试任务执行前**,必须检查:
- [ ] `cad_dialog_killer.py`是否正在运行
- [ ] 如未运行,必须先启动弹窗治理脚本
- [ ] 确认脚本正常工作(检查日志文件)

### 原则4: 窗口监测机制
**每个测试任务开始前**:
- [ ] 分析当前所有窗口名称列表
- [ ] 记录窗口列表到日志文件
- [ ] 如发现异常窗口,先关闭后再开始测试

**测试过程中出现错误或弹窗**:
- [ ] 分析当前多出的窗口
- [ ] 记录新增窗口信息
- [ ] 关闭不能被`cad_dialog_killer.py`清除的窗口
- [ ] 更新弹窗数据库

### 原则5: CAD状态检查和调整
**每个测试任务开始前**:
- [ ] 检查CAD进程数量(必须为1)
- [ ] 检查当前打开的DWG文件数量
- [ ] 将CAD调整为以下状态之一:
  - 单文件不确定状态(默认)
  - 单文件确定状态(指定文件)

### 原则6: 进程和文件约束
**测试过程中必须遵守**:
- [ ] CAD进程只能有一个
- [ ] 同时打开的DWG文件最多两个
- [ ] 禁止打开同一文件多次
- [ ] 一般只使用:单文件确定、单文件不确定、双文件确定状态

### 原则7: 测试结束后恢复
**每个测试任务结束后**:
- [ ] 关闭测试过程中打开的文件
- [ ] 将CAD调整为单文件不确定状态或单文件确定状态
- [ ] 记录测试过程中产生的弹窗
- [ ] 清理临时文件

### 原则8: 弹窗记录和分析
**持续记录**:
- [ ] 软件刚运行时的弹窗
- [ ] 每次测试完毕时的弹窗
- [ ] 测试中途被阻塞时的弹窗
- [ ] 重启天正后检查产生的弹窗
- [ ] 更新弹窗知识库

---

## 🔧 标准测试流程

### 阶段1: 测试准备
```python
# 1. 检查弹窗治理脚本
check_dialog_killer_running()

# 2. 记录当前窗口列表
initial_windows = capture_window_list()
log_windows(initial_windows, "测试开始前")

# 3. 检查CAD状态
check_cad_process_count()  # 必须为1
check_dwg_file_count()     # 记录当前文件数

# 4. 调整为初始状态
set_to_single_unsaved_state()
# 或 set_to_single_definite_state("path/to/file.dwg")
```

### 阶段2: 执行测试
```python
# 5. 执行测试操作
try:
    perform_test_operations()
except Exception as e:
    # 6. 错误处理
    current_windows = capture_window_list()
    new_windows = find_new_windows(initial_windows, current_windows)
    log_error_windows(new_windows)
    close_error_windows(new_windows)
    raise
```

### 阶段3: 测试清理
```python
# 7. 记录测试结束时的窗口
final_windows = capture_window_list()
log_windows(final_windows, "测试结束后")

# 8. 恢复CAD状态
set_to_single_unsaved_state()

# 9. 记录弹窗信息
record_dialogs_to_database()
```

---

## 📝 基本操作范式

### 1. 打开文件
```python
def test_open_file(file_path: str) -> bool:
    """
    测试任务范式:打开单个DWG文件

    前置条件:
    - CAD进程数为1
    - 当前状态为单文件不确定状态

    后置条件:
    - 文件成功打开
    - 状态变为单文件确定状态
    """
    # 前置检查
    assert get_cad_process_count() == 1
    assert is_single_unsaved_state()

    # 执行操作
    success = open_dwg_sync(file_path)

    # 后置检查
    assert success
    assert is_file_opened(file_path)
    assert get_open_file_count() == 1

    return success
```

### 2. 关闭文件
```python
def test_close_file(file_name: str) -> bool:
    """
    测试任务范式:关闭指定文件

    前置条件:
    - 文件已打开

    后置条件:
    - 文件已关闭
    - 状态恢复为单文件不确定状态
    """
    # 前置检查
    assert is_file_opened(file_name)

    # 执行操作
    success = close_file_by_name(file_name)

    # 后置检查
    assert success
    assert not is_file_opened(file_name)

    return success
```

### 3. 保存文件
```python
def test_save_file() -> bool:
    """
    测试任务范式:保存当前文件

    前置条件:
    - 有文件打开
    - 文件有未保存的更改

    后置条件:
    - 文件已保存
    - 文件状态为已保存
    """
    # 前置检查
    assert get_open_file_count() > 0

    # 执行操作
    success = save_current_file()

    # 后置检查
    assert success
    assert not has_unsaved_changes()

    return success
```

### 4. 另存为
```python
def test_save_as(new_path: str) -> bool:
    """
    测试任务范式:另存为新文件

    前置条件:
    - 有文件打开

    后置条件:
    - 新文件已创建
    - 当前激活文件为新文件
    """
    # 前置检查
    assert get_open_file_count() > 0

    # 执行操作
    success = save_as_file(new_path)

    # 后置检查
    assert success
    assert os.path.exists(new_path)
    assert get_active_file_path() == new_path

    return success
```

### 5. 插入块
```python
def test_insert_block(block_path: str, insert_point: tuple) -> bool:
    """
    测试任务范式:插入块

    前置条件:
    - 有文件打开
    - 块文件存在

    后置条件:
    - 块已成功插入
    - 文件有未保存的更改
    """
    # 前置检查
    assert get_open_file_count() > 0
    assert os.path.exists(block_path)

    # 执行操作
    success = insert_block_at_point(block_path, insert_point)

    # 后置检查
    assert success
    assert has_unsaved_changes()

    return success
```

### 6. 双文件操作
```python
def test_two_files_operation(file_a: str, file_b: str) -> bool:
    """
    测试任务范式:双文件对比操作

    前置条件:
    - CAD进程数为1
    - 当前状态为单文件不确定状态

    后置条件:
    - 两个文件都已打开
    - 状态为双文件确定状态
    """
    # 前置检查
    assert get_cad_process_count() == 1
    assert get_open_file_count() == 0

    # 执行操作
    success = open_two_files_sync(file_a, file_b)

    # 后置检查
    assert success
    assert get_open_file_count() == 2
    assert is_file_opened(file_a)
    assert is_file_opened(file_b)

    return success
```

---

## 🧪 标准测试模板

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试任务编号: TEST_TASK_XXX
测试名称: [具体测试内容]
创建日期: YYYY-MM-DD
"""

import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent))

from CAD_test_framework import TestTask, TestResult
from CAD_enhanced_functions import *
from CAD_coordination import *

class TestTaskXXX(TestTask):
    def __init__(self):
        super().__init__(
            task_id="TEST_TASK_XXX",
            task_name="[测试名称]",
            description="[测试描述]"
        )

    def setup(self) -> bool:
        """测试准备"""
        # 1. 检查弹窗治理
        if not self.check_dialog_killer():
            return False

        # 2. 记录初始窗口
        self.capture_initial_windows()

        # 3. 检查CAD状态
        if not self.check_cad_state():
            return False

        # 4. 调整为初始状态
        return self.set_initial_state()

    def execute(self) -> TestResult:
        """执行测试"""
        try:
            # 执行具体测试操作
            result = self.perform_test()

            return TestResult(
                success=True,
                message="测试成功",
                data=result
            )
        except Exception as e:
            return TestResult(
                success=False,
                message=f"测试失败: {e}",
                error=e
            )

    def teardown(self) -> bool:
        """测试清理"""
        # 1. 记录最终窗口
        self.capture_final_windows()

        # 2. 恢复CAD状态
        if not self.restore_cad_state():
            return False

        # 3. 记录弹窗信息
        self.record_dialogs()

        return True

    def perform_test(self):
        """具体测试逻辑 - 子类实现"""
        raise NotImplementedError

if __name__ == "__main__":
    test = TestTaskXXX()
    result = test.run()
    print(result)
```

---

## 📊 测试状态转换图

```
初始状态
    ↓
[检查弹窗治理脚本]
    ↓
[记录当前窗口列表]
    ↓
[检查CAD进程和文件]
    ↓
[调整为测试初始状态]
    ↓
[执行测试操作] ←→ [监测窗口变化]
    ↓
[记录测试结果]
    ↓
[恢复CAD状态]
    ↓
[记录弹窗信息]
    ↓
结束状态
```

---

## 🛡 弹窗处理机制

### 弹窗分类
1. **可自动关闭**: 由`cad_dialog_killer.py`处理
2. **需要记录**: 新出现的未知弹窗
3. **需要手动处理**: 关键错误弹窗

### 弹窗记录格式
```json
{
    "timestamp": "2025-11-06 15:30:00",
    "test_task": "TEST_TASK_001",
    "phase": "执行测试",
    "window_title": "AutoCAD错误报告",
    "window_class": "#32770",
    "action": "自动关闭",
    "handled_by": "cad_dialog_killer"
}
```

### 弹窗数据库
位置: `D:\claude-tasks\tests\dialog_records.json`
- 记录所有出现过的弹窗
- 统计弹窗出现频率
- 分析弹窗触发条件

---

## 📁 测试目录结构

```
D:\claude-tasks\tests\
├── TEST_TASK_001_open_file.py
├── TEST_TASK_002_close_file.py
├── TEST_TASK_003_save_file.py
├── TEST_TASK_004_save_as.py
├── TEST_TASK_005_insert_block.py
├── TEST_TASK_006_two_files.py
├── ...
├── logs/
│   ├── TEST_TASK_001/
│   ├── TEST_TASK_002/
│   └── ...
├── dialog_records.json
└── test_results.json
```

---

## ✅ 测试任务检查清单

### 测试开始前
- [ ] `cad_dialog_killer.py`正在运行
- [ ] 记录了当前窗口列表
- [ ] CAD进程数为1
- [ ] CAD状态已调整为初始状态

### 测试执行中
- [ ] 监测窗口变化
- [ ] 记录操作日志
- [ ] 处理异常情况

### 测试结束后
- [ ] 记录了最终窗口列表
- [ ] 恢复了CAD状态
- [ ] 记录了弹窗信息
- [ ] 清理了临时文件

---

**重要提醒**:所有测试必须严格遵循本规范,确保测试的可靠性和可重复性!