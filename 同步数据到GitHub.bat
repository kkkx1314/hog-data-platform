@echo off
chcp 65001 >nul
echo 🐷 同步平台数据到 GitHub...
echo.

cd /d "C:\Users\CC\test-claude"

:: 1. 复制平台数据到 data/
echo [1/4] 复制平台数据...
python -c "import shutil; from pathlib import Path; src=Path(r'D:\CC\Desktop\平台数据'); dst=Path(r'data'); dst.mkdir(exist_ok=True); [shutil.copy2(f, dst/f.name) or print(f'  {f.name}') for f in src.glob('*.xlsx')]; print('  复制完成')"

:: 2. 删除 data/ 中平台目录已不存在的旧文件
echo.
echo [2/4] 清理旧文件...
python -c "from pathlib import Path; src=Path(r'D:\CC\Desktop\平台数据'); dst=Path(r'data'); src_names={f.name for f in src.glob('*.xlsx')}; [print(f'  删除 {f.name}') or f.unlink() for f in dst.glob('*.xlsx') if f.name not in src_names and not f.name.startswith('~$')]"

:: 3. Git 提交
echo.
echo [3/4] 提交到 Git...
git add data/
git commit -m "data: 同步平台数据 %date%"

:: 4. 推送到 GitHub
echo.
echo [4/4] 推送到 GitHub...
git push

echo.
echo ✅ 同步完成！
pause
