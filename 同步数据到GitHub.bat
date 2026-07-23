@echo off
cd /d "C:\Users\CC\test-claude"

:: 使用 miniconda 的 Python（确保双击也能找到）
set PYTHON=D:\miniconda\python.exe
if not exist "%PYTHON%" set PYTHON=python
if not exist "%PYTHON%" set PYTHON=py

"%PYTHON%" sync_data.py
pause
