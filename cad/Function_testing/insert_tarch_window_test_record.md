# insert_tarch_window 函数测试记录

## 1 测试过程

### 测试环境
```python
# 测试文件
测试文件: D:/claude-tasks/Function_testing/insert_tarch_window - 3.dwg

# CAD启动
from CAD_basic import start_applicationV9
start_applicationV9(PTH=r"C:\Tangent\TArchT20V9", max_retries=3, retry_delay=2.0)

# 文件连接
from CAD_basic import li
li()
```

### 测试1: jz-gaochuang窗
```python
from CAD_file_operations import insert_tarch_window

result = insert_tarch_window(
    p=(38612.86565445, 48750.63891910, 0),
    width=1200,
    height=1000,
    window_type="jz-gaochuang",
    delete_mc_yuan=False
)

# 输出结果
[信息] 窗类型检查通过: jz-gaochuang
[信息] 已连接当前激活文件
[信息] 未找到Mc_yuan_bj图层，正在插入MC_yuan.dwg...
[成功] 已插入MC_yuan.dwg
[成功] 已插入天正门 - 宽度:1200, 高度:1000
[成功] 已设置窗元尺寸: 宽1200, 高1000
[成功] 第1次匹配成功，门已转换为窗，图层:jz-gaochuang
[成功] 已插入天正窗 - 宽度:1200, 高度:1000, 类型:jz-gaochuang

result['success'] = True
```

### 测试2: jz-pingchuang窗（删除MC_yuan对象）
```python
result = insert_tarch_window(
    p=(44695.30568975, 46646.78059028, 0),
    width=2400,
    height=1800,
    window_type="jz-pingchuang",
    delete_mc_yuan=True
)

# 输出结果
[信息] 窗类型检查通过: jz-pingchuang
[信息] 已连接当前激活文件
[信息] 已存在Mc_yuan_bj图层 (找到1个对象)
[成功] 已插入天正门 - 宽度:2400, 高度:1800
[成功] 已设置窗元尺寸: 宽2400, 高1800
[成功] 第1次匹配成功，门已转换为窗，图层:jz-pingchuang
[信息] 正在删除MC_yuan对象...
[成功] 已删除 11 个MC_yuan对象
[成功] 已插入天正窗 - 宽度:2400, 高度:1800, 类型:jz-pingchuang

result['success'] = True
```

### 保存文件
```python
import win32com.client
acad = win32com.client.GetActiveObject("AutoCAD.Application")
acad.ActiveDocument.Save()
```

## 2 使用方法

### 基本用法

#### 2.1 启动CAD和打开文件
```python
import sys
sys.path.append("D:/claude-tasks/scripts")

from CAD_basic import start_applicationV9, li
from CAD_coordination import ensure_single_process, wait_quiescent
from CAD_file_operations import insert_tarch_window
import win32com.client

# 1. 关闭所有CAD
ensure_single_process()

# 2. 启动CAD
start_applicationV9(PTH=r"C:\Tangent\TArchT20V9", max_retries=3, retry_delay=2.0)
wait_quiescent(min_quiet=1.0, timeout=30.0)

# 3. 打开文件
acad = win32com.client.GetActiveObject("AutoCAD.Application")
acad.Documents.Open("D:/path/to/your/file.dwg")
wait_quiescent(min_quiet=1.0, timeout=15.0)

# 4. 连接文件
li()
```

#### 2.2 插入窗
```python
# 插入jz-gaochuang窗（高窗）
result = insert_tarch_window(
    p=(x, y, z),              # 插入点坐标
    width=1200,               # 窗宽度
    height=1000,              # 窗高度
    window_type="jz-gaochuang",  # 窗类型
    delete_mc_yuan=False      # 不删除MC_yuan对象
)

if result['success']:
    print(f"成功插入窗: 宽{result['width']}, 高{result['height']}")
```

#### 2.3 支持的窗类型
- `"jz-menlianchuang"` - 门联窗
- `"jz-dong"` - 洞
- `"jz-gaochuang"` - 高窗
- `"jz-baiyechuang"` - 百叶窗
- `"jz-tuchuang"` - 凸窗
- `"jz-pingchuang"` - 平开窗（默认）
- `"jz-zimumen"` - 子母门
- `"jz-juanlianmen"` - 卷帘门
- `"jz-tuilamen"` - 推拉门
- `"jz-shuangmen"` - 双门

#### 2.4 参数说明
```python
insert_tarch_window(
    p,                    # 必需：插入点坐标元组 (x, y, z)
    width=600,            # 可选：窗宽度（默认600）
    height=1000,          # 可选：窗高度（默认1000）
    window_type="jz-pingchuang",  # 可选：窗类型（默认平开窗）
    delete_mc_yuan=False  # 可选：是否删除MC_yuan对象（默认False）
)
```

#### 2.5 返回值
```python
{
    'success': True/False,    # 是否成功
    'window': 窗对象,          # 插入的窗对象（COM对象）
    'width': 实际宽度值,       # 实际窗宽度
    'height': 实际高度值       # 实际窗高度
}
```

### 完整示例

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
sys.path.append("D:/claude-tasks/scripts")

from CAD_basic import start_applicationV9, li
from CAD_coordination import ensure_single_process, wait_quiescent
from CAD_file_operations import insert_tarch_window
import win32com.client

def main():
    # 1. 关闭所有CAD
    ensure_single_process()

    # 2. 启动CAD
    start_applicationV9(PTH=r"C:\Tangent\TArchT20V9", max_retries=3, retry_delay=2.0)
    wait_quiescent(min_quiet=1.0, timeout=30.0)

    # 3. 打开文件
    acad = win32com.client.GetActiveObject("AutoCAD.Application")
    acad.Documents.Open("D:/test/my_drawing.dwg")
    wait_quiescent(min_quiet=1.0, timeout=15.0)

    # 4. 连接文件
    li()

    # 5. 插入高窗
    result1 = insert_tarch_window(
        p=(1000, 2000, 0),
        width=1200,
        height=1000,
        window_type="jz-gaochuang",
        delete_mc_yuan=False
    )
    print(f"测试1结果: {result1['success']}")

    # 6. 插入平开窗并删除MC_yuan对象
    result2 = insert_tarch_window(
        p=(3000, 2000, 0),
        width=2400,
        height=1800,
        window_type="jz-pingchuang",
        delete_mc_yuan=True
    )
    print(f"测试2结果: {result2['success']}")

    # 7. 保存
    acad.ActiveDocument.Save()
    print("文件已保存")

if __name__ == "__main__":
    main()
```

## 3 注意事项

1. **CAD启动方式**
   - 必须使用 `start_applicationV9()` 启动CAD
   - 不能使用其他方式启动
   - CAD界面必须可见（不能后台运行）

2. **文件连接**
   - 打开文件后必须使用 `li()` 连接当前激活文件
   - 连接后才能使用插入函数

3. **MC_yuan.dwg**
   - 函数会自动检查并插入MC_yuan.dwg（如果需要）
   - MC_yuan.dwg包含窗类型模板
   - `delete_mc_yuan=True` 会在完成后删除MC_yuan相关对象

4. **日志记录**
   - 每次调用都会在 `D:/claude-tasks/logs/` 生成日志文件
   - 日志文件名格式: `insert_tarch_window_YYYYMMDD_HHMMSS.log`

5. **错误处理**
   - 函数会进行完整的错误检查
   - 如果失败，返回 `{'success': False, ...}`
   - 检查返回值的 `success` 字段判断是否成功

6. **墙体要求**
   - 插入点必须在墙体上
   - 墙体必须是有效的天正墙体对象
   - 如果墙体无效（如AcDbZombieEntity），插入会失败

## 4 测试结果

✅ **所有测试通过**

- 测试1: jz-gaochuang (1200x1000) - ✅ 成功
- 测试2: jz-pingchuang (2400x1800, delete_mc_yuan=True) - ✅ 成功
- 文件保存: ✅ 成功

测试文件: `D:/claude-tasks/Function_testing/insert_tarch_window - 3.dwg`

## 5 更新说明

函数已成功更新到 `D:/claude-tasks/scripts/CAD_file_operations.py`

**新函数签名:**
```python
def insert_tarch_window(p, width=600, height=1000, window_type="jz-pingchuang", delete_mc_yuan=False):
```

**主要改进:**
- 支持指定窗高度
- 支持10种窗类型
- 自动MC_yuan.dwg管理
- 详细的日志记录
- 完整的错误处理
- 可选的对象清理

测试日期: 2025-11-11
