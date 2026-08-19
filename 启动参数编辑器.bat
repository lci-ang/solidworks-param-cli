@echo off
chcp 65001 >nul
title SolidWorks 参数编辑器
cd /d "%~dp0"
echo ============================================
echo   SolidWorks 参数编辑器
echo ============================================
echo.

REM 查找 Python：环境变量 SW_PYTHON 优先，然后 py launcher，然后 PATH
set "PY=%SW_PYTHON%"

if not defined PY (
    where py >nul 2>nul
    if not errorlevel 1 (
        for /f "delims=" %%i in ('py -3 -c "import sys; print(sys.executable)" 2^>nul') do set "PY=%%i"
    )
)

if not defined PY (
    where python >nul 2>nul
    if not errorlevel 1 set "PY=python"
)

if not defined PY (
    echo [错误] 未找到 Python，请先安装 Python 3.8+ 并勾选 Add to PATH
    echo 下载地址: https://www.python.org/downloads/
    echo 或者设置环境变量 SW_PYTHON 指向 python.exe 完整路径
    pause
    exit /b 1
)

REM 检查 pywin32
"%PY%" -c "import win32com.client" 2>nul
if errorlevel 1 (
    echo [提示] 首次使用需要安装 pywin32 库...
    "%PY%" -m pip install pywin32
    echo.
)

echo 正在启动（SolidWorks 需要处于运行状态）...
echo.
"%PY%" "%~dp0scripts\sw_cli.py" %*
echo.
pause
