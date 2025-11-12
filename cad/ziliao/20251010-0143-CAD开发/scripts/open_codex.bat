@echo off
setlocal enabledelayedexpansion
rem @script id: open_codex_bat
rem @script name: Open Codex (Batch)
rem @script description: 在当前任务目录启动 Codex（优先 codex，其次 npx，最后 docker）。
rem @script usage: scripts\open_codex.bat [-Npx] [-Docker] [-Workdir <path>]

chcp 65001 >nul

set "ARG_DOCKER=0"
set "ARG_NPX=0"
set "ARG_WORKDIR="

:parse
if "%~1"=="" goto after_parse
if /I "%~1"=="-Docker" set "ARG_DOCKER=1" & shift & goto parse
if /I "%~1"=="-Npx"    set "ARG_NPX=1"    & shift & goto parse
if /I "%~1"=="-Workdir" set "ARG_WORKDIR=%~2" & shift & shift & goto parse
shift
goto parse

:after_parse
if "%ARG_WORKDIR%"=="" (
  pushd "%~dp0\.."
) else (
  pushd "%ARG_WORKDIR%"
)

set "WD=%CD%"
echo [codex] working directory: "%WD%"

if "%ARG_DOCKER%"=="1" goto run_docker
if "%ARG_NPX%"=="1"    goto run_npx

where codex >nul 2>nul
if %ERRORLEVEL%==0 goto run_codex

where npx >nul 2>nul
if %ERRORLEVEL%==0 goto run_npx

where docker >nul 2>nul
if %ERRORLEVEL%==0 goto run_docker

echo [codex] 未找到 codex、npx 或 docker，请先安装其中之一。
goto end

:run_codex
echo [codex] 使用全局 codex
codex
goto end

:run_npx
echo [codex] 使用 npx @openai/codex-cli
npx -y @openai/codex-cli
goto end

:run_docker
echo [codex] 使用 docker 模式
docker run -it --rm -v "%WD%:%WD%" -w "%WD%" ghcr.io/openai/codex-cli:latest
goto end

:end
popd
endlocal

