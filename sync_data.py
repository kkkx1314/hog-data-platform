"""
一键同步：平台数据 + 调运数据 -> data/ -> GitHub
双击 同步数据到GitHub.bat 即可运行
"""
import shutil, re, subprocess
from pathlib import Path

BASE = Path(__file__).parent
PLATFORM = Path(r"D:\CC\Desktop\平台数据")
TRANSPORT_DIR = Path(r"C:\Users\CC")
DATA = BASE / "data"
DATA.mkdir(exist_ok=True)

# 收集摘要信息
summary_platform = []
summary_transport = None
summary_deleted = []

# [1/5] 复制平台数据
print("[1/5] 复制平台数据...")
copied = 0
if PLATFORM.exists():
    for f in sorted(PLATFORM.glob("*.xlsx")):
        if f.name.startswith("~$"):
            continue
        shutil.copy2(f, DATA / f.name)
        print(f"  {f.name}")
        copied += 1
        # 从文件名提取日期范围便于摘要
        dates = re.findall(r"(\d{4}\.\d{1,2}\.\d{1,2}|\d{4}年\d{1,2}月\d{1,2}日|\d{8})", f.name)
        date_tag = ""
        if len(dates) >= 2:
            date_tag = f"  [{dates[0]} ~ {dates[-1]}]"
        elif len(dates) == 1:
            date_tag = f"  [{dates[0]}]"
        summary_platform.append(f"  {f.name}{date_tag}")
    print(f"  复制完成 ({copied} 个文件)")
else:
    print(f"  [WARN] 平台数据目录不存在: {PLATFORM}")

# [2/5] 扫描最新调运数据
print()
print("[2/5] 扫描最新调运数据...")
pattern = re.compile(r"(\d{8})-(\d{8}).*二次去重版\.xlsx")
transport_files = []
for f in TRANSPORT_DIR.glob("*.xlsx"):
    if f.name.startswith("~$"):
        continue
    m = pattern.search(f.name)
    if m:
        transport_files.append((m.group(2), m.group(1), f))

if transport_files:
    transport_files.sort(key=lambda x: (x[0], x[1]), reverse=True)
    latest_end, latest_start, latest_file = transport_files[0]
    sd = f"{latest_start[:4]}.{latest_start[4:6]}.{latest_start[6:]}"
    ed = f"{latest_end[:4]}.{latest_end[4:6]}.{latest_end[6:]}"
    print(f"  最新调运数据: {sd} - {ed}")
    print(f"  文件: {latest_file.name}")
    summary_transport = f"  {latest_file.name}  [{sd} ~ {ed}]"

    if PLATFORM.exists():
        shutil.copy2(latest_file, PLATFORM / latest_file.name)
        print(f"  -> 已复制到平台数据目录")
    shutil.copy2(latest_file, DATA / latest_file.name)
    print(f"  -> 已复制到 data/")

    if len(transport_files) > 1:
        old = transport_files[1][0]
        print(f"  (调运目录共 {len(transport_files)} 个日期段，上次截止 {old[:4]}.{old[4:6]}.{old[6:]})")
else:
    print(f"  [WARN] 未在 {TRANSPORT_DIR} 找到调运分析文件")

# [3/5] 清理旧文件
print()
print("[3/5] 清理旧文件...")
if PLATFORM.exists():
    src_names = {f.name for f in PLATFORM.glob("*.xlsx")}
    src_names.update(f.name for f in TRANSPORT_DIR.glob("*.xlsx") if not f.name.startswith("~$"))
    for f in DATA.glob("*.xlsx"):
        if f.name not in src_names and not f.name.startswith("~$"):
            print(f"  删除 {f.name}")
            summary_deleted.append(f"  {f.name}")
            f.unlink()
if not summary_deleted:
    print("  无需清理")
else:
    print("  清理完成")

# [4/5] Git 提交
print()
print("[4/5] 提交到 Git...")
subprocess.run(["git", "add", "data/"], cwd=str(BASE), check=True)
diff = subprocess.run(
    ["git", "diff", "--cached", "--name-only"],
    cwd=str(BASE), capture_output=True, text=True
)
changed_files = [l.strip() for l in diff.stdout.strip().split("\n") if l.strip()] if diff.stdout.strip() else []
if changed_files:
    subprocess.run(["git", "commit", "-m", "data: 同步平台数据+调运数据"], cwd=str(BASE))
    print("  提交成功")
else:
    print("  数据无变化，跳过提交")

# [5/5] 推送
print()
print("[5/5] 推送到 GitHub...")
subprocess.run(["git", "push"], cwd=str(BASE), check=True)

# ==================== 同步摘要 ====================
print()
print("=" * 55)
print("  同步摘要")
print("=" * 55)

if summary_platform:
    print(f"  [平台数据] 共 {len(summary_platform)} 个文件:")
    for s in summary_platform:
        print(s)

if summary_transport:
    print()
    print(f"  [调运数据]")
    print(summary_transport)

if summary_deleted:
    print()
    print(f"  [已清理] 共 {len(summary_deleted)} 个旧文件:")
    for d in summary_deleted:
        print(d)

if changed_files:
    print()
    print(f"  [Git] 已提交 {len(changed_files)} 个变更并推送到 GitHub")
else:
    print()
    print(f"  [Git] 无变更，已是最新")

print("=" * 55)
