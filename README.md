# Claude Tasks

多领域自动化开发项目集合

## 项目结构

```
claude-tasks/
├── cad/          # CAD自动化开发 (天正建筑)
├── su/           # SketchUp自动化开发
├── wenben/       # 文本处理自动化
├── video/        # 视频处理自动化
└── zonghe/       # 跨领域综合脚本
```

## 子项目说明

### cad/
CAD自动化开发项目，主要针对天正建筑CAD软件的自动化操作。

**核心功能:**
- 文件操作 (新建、打开、保存、关闭)
- 天正构件插入 (墙体、门、窗)
- 对象属性操作
- 协同机制管理

**核心脚本:**
- `scripts/CAD_basic.py` - 基础操作库
- `scripts/CAD_file_operations.py` - 文件操作统一接口

详见: [cad/README.md](cad/README.md)

### su/
SketchUp自动化开发项目 (待开发)

### wenben/
文本处理自动化项目 (待开发)

### video/
视频处理自动化项目 (待开发)

### zonghe/
跨领域综合脚本，涉及多个子项目的集成功能

## 开发环境

- Python 3.x
- Windows 10/11
- 各子项目有独立的依赖要求

## 使用说明

每个子项目都有独立的开发环境和文档，请参考各子项目的README文件。

## 贡献

本项目使用Claude Code进行AI辅助开发。

## 许可

MIT License
