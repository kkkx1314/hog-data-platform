@echo off
chcp 65001 >nul
echo 🐖 更新云南日报数据（同步桌面涌益日度数据 + 重新生成日报 + 推送 GitHub）...
echo.
cd /d "C:\Users\CC\test-claude"
D:\miniconda\python.exe "C:\Users\CC\test-claude\update_yunnan_data.py"
pause
