# Claude CAD/SU 开发工作站

这是一个为 Claude 准备的开发环境,专注于建筑工程设计软件的开发和测试。

## 🎯 工作目标
- CAD 软件自动化开发测试
- SketchUp 插件开发验证
- 图像处理工具开发
- 建筑工程相关软件

## ⚡ 快速启动

### 方式一：初始化环境（推荐）
**新会话开始时输入:**
```
/init
```
这将自动加载所有权限设置和工作环境。

### 方式二：CAD 任务模式
**执行 CAD 任务时输入:**
```
/cad
```
这将启动 CAD 开发模式，Claude 会自动：
- 阅读 CAD 操作规范文档
- 强制使用范式函数
- 遵循协同机制
- 详细记录操作日志

### 配置系统
`.claude/` 目录包含 Claude 的行为配置：
- **CLAUDE_BEHAVIOR_RULES.md** - Claude 的强制性行为准则
- **commands/cad.md** - `/cad` 命令定义
- **commands/init.md** - `/init` 命令定义
- **claude_config.json** - 配置参数

详见：`.claude/README.md`

## 🔧 CAD开发专用

**📋 CAD任务前必读:**
- **CAD_快速参考.md** - 新对话必读!包含所有协同机制规范
- **CAD_操作规范.md** - 完整的六大核心规则
- **CAD_基本操作范式.md** - 【必读】五大基本操作范式(新建/打开/关闭/保存/插入)
- **CAD_测试规范.md** - 完整的测试规范和八大原则
- **即时对话.txt** - 原始需求文档

**🚀 CAD开发流程:**
1. 阅读 `CAD_快速参考.md`
2. 使用 `CAD_enhanced_functions.py` 模块
3. 遵循协同机制(命令必须等待完成)
4. 通过天正软件启动CAD

**🧪 CAD测试流程:**
1. 阅读 `CAD_测试规范.md`
2. 使用 `CAD_test_framework.py` 测试框架
3. 遵循测试八大原则
4. 使用 `run_all_tests.py` 批量测试

**⚠ 重要提醒**:
- CAD操作必须严格遵循协同机制,确保命令顺序执行!
- CAD测试必须严格遵循测试八大原则!

## 📋 权限配置

### 完全权限区域（无需询问确认）
1. `D:\claude-tasks\` - 主工作目录
   - 可以对文件夹和文件进行任何操作
   - 创建、修改、删除文件
   - 执行 Python 脚本和 Bash 命令
   - **不需要中间询问确认，直到完成任务**
2. `C:\Users\Administrator\AppData\Roaming\SketchUp\SketchUp 2022\SketchUp\Plugins\` - SU插件目录
   - 完全控制权限

### 其他区域
- 只读查看任何文件
- 修改/删除需要询问用户

## 🛠 软件环境

- SketchUp 2018, 2022
- Autodesk 相关软件
- Python 3.11.5
- Node.js, Git

## 📁 项目结构

```
D:\claude-tasks\
├── .claude/                        # Claude 配置
├── README.md                       # 本文件
├── PERMISSIONS.md                  # 详细权限说明
├── start-work.md                   # 启动工作指南
├── 文档索引.md                      # 文档导航索引
├── 即时对话.txt                     # 原始需求文档
├── CAD_快速参考.md                  # 【新对话必读】CAD快速参考
├── CAD_操作规范.md                  # CAD六大核心规则
├── CAD_测试规范.md                  # 【测试必读】测试八大原则
├── CAD 协同机制实现报告.md          # 协同机制实现报告
├── CAD 测试机制实现报告.md          # 测试机制实现报告
├── projects/                       # 开发项目
├── tests/                         # 测试文件
│   ├── TEST_TASK_001_open_file.py # 打开文件测试
│   ├── TEST_TASK_002_close_file.py # 关闭文件测试
│   ├── run_all_tests.py           # 批量测试运行器
│   ├── logs/                      # 测试日志目录
│   ├── test_files/                # 测试文件目录
│   ├── dialog_records.json        # 弹窗记录数据库
│   └── test_report_*.json         # 测试报告文件
├── scripts/                       # 自动化脚本
│   ├── CAD-basic.py               # 基础CAD操作(已集成协同机制)
│   ├── CAD_coordination.py        # 协同机制核心模块
│   ├── CAD_enhanced_functions.py  # 增强功能模块(推荐使用)
│   ├── CAD_test_framework.py      # 测试框架核心模块
│   └── cad_dialog_killer.py       # 弹窗治理脚本
└── ziliao/                        # 资料文件夹
    └── 20251010-0143-CAD开发/     # 完整CAD开发规范
```

---

**使用方法**: 新会话时输入 `/init` 即可开始工作!