# CAD函数测试报告说明

## 文件夹结构

```
test_reports/
├── README.md                              # 本说明文件
├── CAD函数测试报告_YYYYMMDD_HHMMSS.xlsx   # Excel格式测试报告
└── test_dwg_files/                        # 测试DWG文件存档
    ├── 0.dwg                              # 系统测试文件
    ├── 1.dwg                              # 系统测试文件
    ├── 2.dwg                              # 系统测试文件
    ├── MC_yuan.dwg                        # 窗类型源文件
    ├── test_door_triangle.dwg             # 三角形墙体+门测试文件
    └── test_window_rectangle.dwg          # 矩形墙体+窗测试文件
```

## Excel报告内容

Excel测试报告包含4个工作表：

### 1. 测试详情
- **函数名**: 被测试的函数名称
- **测试状态**: 通过/失败
- **执行时间**: 函数执行耗时
- **测试脚本**: 使用的测试脚本文件名
- **测试DWG文件**: 使用的测试DWG文件
- **文件位置**: DWG文件所在路径
- **说明**: 测试说明和备注

### 2. 测试汇总
- 总函数数
- 测试通过数
- 测试覆盖率
- 实际可用率
- 测试日期
- 测试环境

### 3. Bug修复记录
- Bug编号
- 文件名
- 行号
- 问题描述
- 修复方案
- 修复状态

### 4. 测试文件清单
- 文件名
- 文件类型（系统文件/测试文件）
- 用途说明
- 文件位置

## 如何复测

### 前提条件
1. 安装天正建筑 T20V9
2. 安装Python 3.11或更高版本
3. 安装必要的Python包：
   ```bash
   pip install pywin32 pandas openpyxl
   ```

### 复测步骤

#### 1. 准备测试文件
将 `test_dwg_files/` 文件夹中的DWG文件复制到对应位置：
- 系统文件（0.dwg, 1.dwg, 2.dwg, MC_yuan.dwg）-> `D:/claude-tasks/cad/xitongwenjian/`
- 测试文件（test_*.dwg）-> `D:/claude-tasks/cad/tests/test_files/`

#### 2. 运行测试脚本
根据Excel报告中的"测试脚本"列，运行相应的测试：

```bash
# 测试三角形墙体+门
python D:\claude-tasks\cad\tests\test_insert_tarch_door_triangle.py

# 测试矩形墙体+窗
python D:\claude-tasks\cad\tests\test_insert_tarch_window_rectangle.py

# 测试墙体绘制
python D:\claude-tasks\cad\tests\test_draw_tarch_wall_fixed.py
```

#### 3. 验证测试结果
- 检查测试脚本的输出，确认"[成功]"标记
- 检查生成的DWG文件是否正常
- 对比执行时间是否在合理范围内

## 测试文件说明

### 系统测试文件
- **0.dwg**: 空白DWG文件，用于基础文件操作测试
- **1.dwg**: 空白DWG文件，用于多文件操作测试
- **2.dwg**: 空白DWG文件，用于多文件操作测试
- **MC_yuan.dwg**: 包含窗类型定义的源文件（21个对象）

### 功能测试文件
- **test_door_triangle.dwg** (32,279字节)
  - 包含：3堵墙（三角形布局）+ 3个门
  - 门尺寸：1000x2100, 1200x2100, 1500x2400
  - 用于测试：insert_tarch_door()

- **test_window_rectangle.dwg** (59,802字节)
  - 包含：4堵墙（矩形布局，8000x6000mm）+ 8个窗
  - 窗类型：jz-pingchuang (平窗) x3
           jz-juanlianmen (卷帘门) x3
           jz-tuilamen (推拉门) x2
  - 用于测试：insert_tarch_window()

## 重要测试原则

### 1分钟原则
**任何命令运行时间超过1分钟就必须关掉进程重启，因为命令被卡住了。**

### 性能基准
- 墙体绘制：3.4-4.0秒/墙
- 门插入：2.8秒/门
- 窗插入：
  - 首次（需插入MC_yuan.dwg）：13.5秒
  - 后续：7.9-8.1秒/窗

## 生成新报告

运行以下脚本生成新的Excel测试报告：
```bash
python D:\claude-tasks\cad\tests\generate_excel_report.py
```

报告文件名格式：`CAD函数测试报告_YYYYMMDD_HHMMSS.xlsx`

## 联系信息

- 测试日期：2025-11-13
- 测试环境：Windows + 天正建筑 T20V9 + Python 3.11
- 测试覆盖率：100% (24/24函数)
- 测试通过率：100%

## 版本历史

### v1.0 (2025-11-13)
- 完成所有24个函数的测试
- 修复8个关键bug
- 建立完整测试框架
- 生成Excel格式测试报告
