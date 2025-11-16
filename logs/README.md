# Claude Code 日志归档目录

## 📁 目录说明

此目录用于存储Claude Code的历史日志归档。

## 📋 文件命名规则

- **当前日志**: `D:\claude-tasks\ai_console.log`
- **归档日志**: `archive\ai_console_YYYY-MM-DD.log`

## 🔄 自动管理机制

### 归档规则
- 每次启动Claude时，如果发现 `ai_console.log` 不是今天的日志，会自动归档
- 归档文件命名格式：`ai_console_2025-11-14.log`
- 同一天的多次启动会追加到同一个归档文件

### 清理规则
- 自动删除7天前的归档日志
- 保留最近7天的所有日志记录

## 📊 日志内容

日志包含：
- Claude Code的所有输出信息
- 命令执行记录
- 错误和警告信息
- 系统提示信息

## 🔍 查看日志

### 查看当前日志
```powershell
Get-Content D:\claude-tasks\ai_console.log -Tail 50
```

### 查看归档日志
```powershell
Get-Content D:\claude-tasks\logs\archive\ai_console_2025-11-14.log
```

### 搜索日志内容
```powershell
Select-String -Path "D:\claude-tasks\logs\archive\*.log" -Pattern "关键词"
```

## 📈 日志统计

### 查看归档文件列表
```powershell
Get-ChildItem D:\claude-tasks\logs\archive -Filter "*.log" |
    Select-Object Name, Length, LastWriteTime |
    Sort-Object LastWriteTime -Descending
```

### 计算总大小
```powershell
$totalSize = (Get-ChildItem D:\claude-tasks\logs\archive -Filter "*.log" |
    Measure-Object -Property Length -Sum).Sum / 1MB
Write-Host "归档总大小: $([math]::Round($totalSize, 2)) MB"
```

## ⚙️ 自定义设置

如需修改归档保留天数，编辑 `start_claude_with_log.ps1`：

```powershell
# 修改这一行的数字（默认7天）
$cutoffDate = (Get-Date).AddDays(-7)
```

## 🚀 使用方法

使用启动脚本启动Claude：
```powershell
powershell -ExecutionPolicy Bypass -File D:\claude-tasks\start_claude_with_log.ps1
```

或者创建快捷方式，目标设置为：
```
powershell.exe -ExecutionPolicy Bypass -File "D:\claude-tasks\start_claude_with_log.ps1"
```

---

**最后更新**: 2025-11-14
