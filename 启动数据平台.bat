@echo off
chcp 65001 >nul
echo 🐷 生猪数据平台 - 启动中...
echo.
cd /d "C:\Users\CC\test-claude"
streamlit run excel_dashboard001.py --server.port 8502
pause
