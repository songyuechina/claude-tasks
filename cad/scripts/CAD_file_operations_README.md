# CAD文件操作统一接口

## 📋 简介

`CAD_file_operations.py` 提供了DWG文件操作的统一接口，将所有常用操作封装为简单易用的函数。

## 🚀 快速开始

```python
from CAD_file_operations import *

# 新建文件
new_file("D:/output/new.dwg")

# 打开文件
open_file("D:/path/to/file.dwg")

# 保存文件
save_file()

# 插入文件为块
insert_file_as_block("D:/source.dwg", x=100, y=100)

# 插入文件并炸开
insert_file_exploded("D:/source.dwg", x=200, y=200)

# 关闭文件
close_file("no_save")
```

## 📚 函数列表

### 文件新建与打开

| 函数 | 说明 | 参数 |
|------|------|------|
| `new_file(path)` | 新建文件 | path: 保存路径(可选) |
| `open_file(path)` | 打开文件 | path: 文件路径 |

### 文件保存

| 函数 | 说明 | 参数 |
|------|------|------|
| `save_file()` | 保存当前文件 | 无 |
| `save_file_as(path)` | 另存为 | path: 保存路径 |

### 文件关闭

| 函数 | 说明 | 参数 |
|------|------|------|
| `close_file(option)` | 关闭当前文件 | option: "prompt"/"auto_save"/"no_save" |
| `close_all_files()` | 关闭所有文件 | 无 |

### 文件插入

| 函数 | 说明 | 参数 |
|------|------|------|
| `insert_file_as_block()` | 插入为块 | source_file, x, y, z, scale, rotation |
| `insert_file_exploded()` | 插入并炸开 | source_file, x, y, z, scale |

### 完整工作流

| 函数 | 说明 | 参数 |
|------|------|------|
| `copy_file_content()` | 拷贝文件内容 | source_file, target_file, explode, x, y |

## 💡 使用示例

### 示例1: 新建文件并绘制

```python
from CAD_file_operations import new_file, save_file_as, close_file
from CAD_coordination import send_cmd_with_sync

# 新建空白文件
new_file()

# 绘制圆形
send_cmd_with_sync("_CIRCLE\n0,0\n100\n", wait_after=1.0)

# 保存
save_file_as("D:/output/circle.dwg")

# 关闭
close_file("no_save")
```

### 示例2: 打开文件并编辑

```python
from CAD_file_operations import open_file, save_file, close_file
from CAD_coordination import send_cmd_with_sync

# 打开文件
open_file("D:/input/drawing.dwg")

# 添加矩形
send_cmd_with_sync("_RECTANG\n0,0\n100,50\n", wait_after=1.0)

# 保存
save_file()

# 关闭
close_file("no_save")
```

### 示例3: 插入文件为块

```python
from CAD_file_operations import new_file, insert_file_as_block, save_file_as, close_file

# 新建文件
new_file()

# 插入为块
insert_file_as_block(
    source_file="D:/blocks/furniture.dwg",
    x=100, y=100, z=0,
    scale=1.0,
    rotation=45.0
)

# 保存
save_file_as("D:/output/result.dwg")
close_file("no_save")
```

### 示例4: 插入文件并炸开

```python
from CAD_file_operations import new_file, insert_file_exploded, save_file_as, close_file

# 新建文件
new_file()

# 插入并炸开
insert_file_exploded(
    source_file="D:/source/base.dwg",
    x=0, y=0, z=0,
    scale=1.0
)

# 保存
save_file_as("D:/output/merged.dwg")
close_file("no_save")
```

### 示例5: 拷贝文件内容

```python
from CAD_file_operations import copy_file_content, close_file

# 拷贝为块
copy_file_content(
    source_file="D:/source/A.dwg",
    target_file="D:/output/B_with_block.dwg",
    explode=False,
    x=0, y=0
)
close_file("no_save")

# 拷贝并炸开
copy_file_content(
    source_file="D:/source/A.dwg",
    target_file="D:/output/B_exploded.dwg",
    explode=True,
    x=0, y=0
)
close_file("no_save")
```

## 🔑 关键特性

1. **简单易用**: 函数名直观，参数清晰
2. **协同机制**: 所有函数都集成了协同机制，确保命令顺序执行
3. **统一接口**: 所有文件操作集中在一个模块中
4. **两种插入模式**: 支持块模式和炸开模式

## 📝 注意事项

1. 使用前需要先启动CAD: `start_applicationV9()`
2. 所有函数都会自动等待操作完成
3. 插入操作会自动处理中文路径
4. 建议在操作间使用 `wait_quiescent()` 确保稳定

## 🔗 相关文档

- `CAD_基本操作范式.md` - 详细的操作范式说明
- `CAD_快速参考.md` - 快速参考指南
- `CAD_file_operations_example.py` - 完整使用示例
