@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Starting dashboard at http://localhost:8501 ...
"D:\miniconda\python.exe" -m streamlit run "%~dp0app.py" --server.port 8501 --server.headless false
pause
