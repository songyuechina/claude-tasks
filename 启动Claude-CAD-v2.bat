@echo off
chcp 65001 >nul
title ClaudeLauncher_v2

REM 0) 最小化所有窗口，避免干扰
powershell -NoLogo -NoProfile -Command "(New-Object -ComObject Shell.Application).MinimizeAll()"

REM 1) 新开 PowerShell 窗口启动 Claude 控制台（带 Transcript）
start "" powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File "D:\claude-tasks\ClaudeConsole.ps1"

REM 2) 等 Claude 启动（可按需要调整 3/5 秒）
timeout /T 5 /NOBREAK >nul

REM 3) 用 v2 的 Python 脚本发送 /cad
python "D:\claude-tasks\send_cad_command_v2.py"

pause
