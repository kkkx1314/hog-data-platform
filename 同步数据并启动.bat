@echo off
chcp 65001 >nul
echo 🐷 同步数据到云端 + 启动平台...
echo.
cd /d "C:\Users\CC\test-claude"
python sync_data.py
pause
