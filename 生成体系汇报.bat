@echo off
chcp 65001 >nul
echo ======================================
echo   D:/claude-tasks 体系汇报生成器
echo ======================================
echo.

python "D:\claude-tasks\cad\system_visualizer.py"

if errorlevel 1 (
    echo.
    echo 错误：程序执行失败
    pause
)
