# insert_tarch_window 函数测试文档

## 1 函数说明

`insert_tarch_window` 函数用于在天正CAD中的墙体上插入天正窗。该函数通过先插入门，再通过属性匹配的方式将门转换为指定类型的窗。

## 2 函数签名

```python
def insert_tarch_window(p, width=600, height=1000, window_type="jz-pingchuang", delete_mc_yuan=False)
```

### 参数说明

- `p`: 插入点坐标，格式为 `(x, y, z)`
- `width`: 窗宽度，默认 600
- `height`: 窗高度，默认 1000
- `window_type`: 窗类型，默认 "jz-pingchuang"
- `delete_mc_yuan`: 是否删除MC_yuan.dwg插入的对象，默认 False（不删除）。连续插入多个窗时建议设为False

### 允许的窗类型

函数只允许以下窗类型之一：
- "jz-menlianchuang" (门联窗)
- "jz-dong" (洞)
- "jz-gaochuang" (高窗)
- "jz-baiyechuang" (百叶窗)
- "jz-tuchuang" (凸窗)
- "jz-pingchuang" (平窗)
- "jz-zimumen" (子母门)
- "jz-juanlianmen" (卷帘门)
- "jz-tuilamen" (推拉门)
- "jz-shuangmen" (双门)

### 返回值

返回字典类型：
```python
{
    'success': bool,    # 是否成功
    'window': 窗对象,    # 成功时返回窗对象，失败时为None
    'width': 宽度,       # 成功时返回实际宽度，失败时为None
    'height': 高度       # 成功时返回实际高度，失败时为None
}
```

## 3 函数原理

函数的工作流程如下：

1. **检查窗类型**：验证 `window_type` 是否在允许的类型列表中
2. **连接当前文件**：通过 `li()` 连接当前激活的CAD文件
3. **检查MC_yuan.dwg**：检查是否已插入系统文件MC_yuan.dwg（通过图层'Mc_yuan_bj'判断）
   - 如果未插入，则调用 `copy_file_content_pywin32` 插入该文件
4. **插入门**：使用 `insert_tarch_door` 在指定位置插入门
5. **选择窗元**：通过 `stc(window_type)` 选择窗类型图层的窗元，并设置其宽度和高度
6. **属性匹配**：使用 `transfer_props_by_matchprop` 将窗元的属性匹配给门，最多尝试5次
   - 匹配成功的标志是门的图层变为窗类型图层
7. **返回结果**：返回成功信息和窗对象

## 4 使用方法

### 前置准备

1. 启动天正CAD并连接
2. 打开或新建一个包含墙体的DWG文件
3. 确保墙体已正确绘制

### 示例代码

```python
from test_insert_tarch_window import insert_tarch_window
from CAD_basic import li

# 连接当前激活文件
li()

# 示例1: 插入高窗（宽1200，高1000）
result1 = insert_tarch_window(
    p=(38612.86565445, 48750.63891910, 0),
    width=1200,
    height=1000,
    window_type="jz-gaochuang"
)

if result1['success']:
    print(f"成功插入高窗，宽度:{result1['width']}, 高度:{result1['height']}")
else:
    print("插入失败")

# 示例2: 插入平窗（宽2400，高1800）
result2 = insert_tarch_window(
    p=(44695.30568975, 46646.78059028, 0),
    width=2400,
    height=1800,
    window_type="jz-pingchuang"
)

if result2['success']:
    print(f"成功插入平窗，宽度:{result2['width']}, 高度:{result2['height']}")
else:
    print("插入失败")
```

## 5 测试记录

### 测试环境
- 测试文件：`D:/claude-tasks/Function_testing/insert_tarch_window-1.dwg`
- 天正版本：TArchT20V9
- 测试时间：2025-11-11

### 测试用例

#### 测试1：插入高窗
- 位置：(38612.86565445, 48750.63891910, 0)
- 窗类型：jz-gaochuang
- 宽度：1200
- 高度：1000
- 预期结果：成功插入高窗

#### 测试2：插入平窗
- 位置：(44695.30568975, 46646.78059028, 0)
- 窗类型：jz-pingchuang
- 宽度：2400
- 高度：1800
- 预期结果：成功插入平窗

### 测试结果

**测试状态：✅ 全部通过**

#### 测试1：插入高窗
- 结果：✅ 成功
- 第1次属性匹配成功
- 门已成功转换为窗，图层改为 jz-gaochuang
- 实际宽度：1200，实际高度：1000

#### 测试2：插入平窗
- 结果：✅ 成功
- 第1次属性匹配成功
- 门已成功转换为窗，图层改为 jz-pingchuang
- 实际宽度：2400，实际高度：1800

**测试总结**：
- 两个测试用例均成功执行
- transfer_props_by_matchprop 函数工作正常
- 门成功转换为指定类型的窗
- 宽度和高度设置正确
- 文件已成功保存

**关键点**：
- MC_yuan.dwg 中的窗元必须位于原点附近，这样 transfer_props_by_matchprop 才能正确工作
- 函数严格按照1-9步骤实现，使用 transfer_props_by_matchprop 进行属性匹配
- 仅修改 Layer 属性无法实现门转窗，必须使用 transfer_props_by_matchprop

## 6 注意事项

1. **窗类型限制**：窗类型必须是允许列表中的之一，否则函数会直接返回失败
2. **墙体要求**：插入位置必须有墙体，否则门无法插入，函数会失败
3. **MC_yuan.dwg**：首次使用时会自动插入系统文件MC_yuan.dwg，该文件包含各种窗类型的模板
4. **属性匹配**：函数最多尝试5次属性匹配，如果5次都失败会返回失败
5. **日志记录**：函数会在 `D:/claude-tasks/logs` 目录下生成详细的日志文件

## 7 错误处理

函数采用 try-except 方式编写，以下情况会导致函数失败：
- 窗类型不在允许列表中
- 未找到墙体，门插入失败
- 未找到对应的窗类型图层
- 属性匹配5次仍然失败
- 其他异常情况

所有错误都会记录在日志文件中，便于调试和问题排查。

## 8 相关函数

- `insert_tarch_door`: 插入天正门
- `transfer_props_by_matchprop`: 属性匹配功能
- `copy_file_content_pywin32`: 复制文件内容
- `stc`: 按图层选择对象
- `li`: 连接当前激活文件
