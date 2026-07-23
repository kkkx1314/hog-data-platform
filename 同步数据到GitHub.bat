@echo off
cd /d "C:\Users\CC\test-claude"

:: 尝试多种方式找到 Python
set PYTHON=
where python >nul 2>&1 && set PYTHON=python
if "%PYTHON%"=="" where py >nul 2>&1 && set PYTHON=py
if "%PYTHON%"=="" where python3 >nul 2>&1 && set PYTHON=python3
if "%PYTHON%"=="" (
    echo [ERROR] 未找到 Python，请确认已安装并添加到 PATH
    pause
    exit /b 1
)

%PYTHON% sync_data.py
pause
