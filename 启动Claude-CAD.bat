@echo off
chcp 65001 >nul
title ClaudeLauncher

REM 0) minimize all windows to avoid interference
powershell -NoLogo -NoProfile -Command "(New-Object -ComObject Shell.Application).MinimizeAll()"

REM 1) start Claude console (with transcript) in a new PowerShell window
start "" powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File "D:\claude-tasks\ClaudeConsole.ps1"

REM 2) wait a bit for Claude to start (you可以按需要调 3/5 秒)
timeout /T 5 /NOBREAK >nul

REM 3) use Python + pyautogui to send /cad + Enter
python "D:\claude-tasks\send_cad_command.py"

pause
