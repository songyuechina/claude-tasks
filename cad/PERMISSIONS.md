# Claude 工作权限配置

**生效日期**: 2025-11-05
**电脑**: Windows 工作站 (专为 Claude 准备)

## ✅ 完全权限区域 (可读写删改)

### 1. Claude 工作目录
```
D:\claude-tasks\
```
- **权限**: 完全控制
- **用途**: 所有开发、测试、临时文件
- **无需询问**: 任何文件操作
- **重要**: 可以对文件夹和文件进行任何操作，不需要中间询问用户确认（如"Do you want to proceed?"），直到完成下达的任务

### 2. SketchUp 2022 插件目录
```
C:\Users\Administrator\AppData\Roaming\SketchUp\SketchUp 2022\SketchUp\Plugins\
```
- **权限**: 完全控制
- **用途**: 开发和测试 SketchUp 插件
- **无需询问**: 任何插件文件的增删改

## 📖 只读权限区域

### 3. 系统其他位置
```
C:\Program Files\
C:\Program Files (x86)\
其他所有目录
```
- **权限**: 只读查看
- **限制**: 写入、修改、删除需要明确询问并获得许可
- **例外**: 上述1、2区域

## 🎯 工作原则

1. **主工作目录**: `D:\claude-tasks\` - 自由操作
2. **SketchUp 插件**: `Plugins\` 目录 - 自由操作
3. **其他位置**: 先查看，需要修改时询问
4. **软件操作**: 可通过脚本/API启动和控制CAD、SU等软件进行测试

## 📝 软件环境

- SketchUp 2018, 2022
- Autodesk 相关软件
- Python 3.11.5
- Git, Node.js

---
**注意**: 此权限配置在每次会话中生效，Claude 会遵守这些规则。
