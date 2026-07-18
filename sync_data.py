"""一键同步：复制平台数据到data/ → 推送到GitHub → 启动Streamlit"""
import shutil, subprocess, sys
from pathlib import Path

SRC = Path(r"D:\CC\Desktop\平台数据")
DST = Path(__file__).parent / "data"
DST.mkdir(exist_ok=True)

# 1. 复制Excel文件
print("Copying files...")
copied = 0
for f in SRC.glob("*.xlsx"):
    shutil.copy2(f, DST / f.name)
    print(f"  {f.name}")
    copied += 1
if not copied:
    print("  No Excel files found!")
    sys.exit(1)

# 2. Git提交推送
print("\nCommitting to GitHub...")
subprocess.run(["git", "add", "data/"], cwd=Path(__file__).parent, check=True)
subprocess.run(["git", "commit", "-m", "data: sync platform data"], cwd=Path(__file__).parent)
subprocess.run(["git", "push"], cwd=Path(__file__).parent, check=True)

# 3. 启动Streamlit
print("\nStarting Streamlit...")
subprocess.run([sys.executable, "-m", "streamlit", "run", "excel_dashboard001.py", "--server.port", "8502"], cwd=Path(__file__).parent)
