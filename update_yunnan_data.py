# -*- coding: utf-8 -*-
"""
更新云南日报数据：把桌面最新「涌益咨询日度数据」同步到 data/ 并重新生成两个日报，
最后提交推送到 GitHub。
双击 更新云南日报数据.bat 即可运行。
"""
import shutil
import subprocess
from pathlib import Path

import price_reports as pr

BASE = Path(__file__).parent
DATA = BASE / "data"
DATA.mkdir(exist_ok=True)

print("=" * 55)
print("        更新云南日报数据")
print("=" * 55)

# [1/4] 定位桌面最新涌益日度数据
path, latest_date = pr.latest_file()
src = Path(path)
print(f"[1/4] 最新数据文件: {src.name}")

# [2/4] 复制到 data/
dst = DATA / src.name
shutil.copy2(src, dst)
print(f"[2/4] 已复制到 data/{src.name}")

# [3/4] 重新生成两个日报
print("[3/4] 重新生成日报 HTML ...")
province_data, p_latest = pr.load_provinces(path)
yunnan = pr.load_yunnan(path)
yn_latest = yunnan["dates"][-1] if yunnan["dates"] else latest_date

with open(BASE / "价差日报.html", "w", encoding="utf-8") as f:
    f.write(pr.build_price_report_html(province_data, p_latest))
with open(BASE / "云南日报.html", "w", encoding="utf-8") as f:
    f.write(pr.build_yunnan_report_html(yunnan, yn_latest))
print(f"  价差日报.html / 云南日报.html 已更新（数据截止 {yn_latest}）")

# [4/4] 提交并推送
print("[4/4] 提交并推送到 GitHub ...")
subprocess.run(["git", "add", "data/", "价差日报.html", "云南日报.html"],
               cwd=str(BASE), check=True)
diff = subprocess.run(["git", "diff", "--cached", "--name-only"],
                      cwd=str(BASE), capture_output=True, text=True)
changed = [l.strip() for l in diff.stdout.strip().split("\n") if l.strip()] if diff.stdout.strip() else []
if changed:
    subprocess.run(["git", "commit", "-m", f"data: 更新涌益日度数据 {latest_date}"], cwd=str(BASE), check=True)
    print("  提交成功")
else:
    print("  数据无变化，跳过提交")

subprocess.run(["git", "push"], cwd=str(BASE), check=True)
print()
print("完成 ✅")
