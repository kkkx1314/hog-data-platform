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

# 收集同步信息
synced = {}          # {类别: [(日期标签, 文件名)]}
transport_info = None
deleted_files = []


def classify_file(name: str):
    """根据文件名识别数据类别"""
    if "日度" in name:
        return "涌益日度数据"
    if "周度" in name:
        return "涌益周度数据"
    if "调运" in name or "二次去重" in name:
        return "猪只调运数据"
    if "鲜品" in name:
        return "神农肉业-鲜品价格"
    if "冻品" in name:
        return "神农肉业-冻品价格"
    return "其他数据"


def extract_date_tag(name: str) -> str:
    """从文件名提取日期标签"""
    # 匹配 2026年7月22日 这种
    m = re.search(r"(\d{4})年(\d{1,2})月(\d{1,2})日", name)
    if m:
        return f"{m.group(1)}.{int(m.group(2)):02d}.{int(m.group(3)):02d}"
    # 匹配 2026.7.10-2026.7.16 这种
    m = re.search(r"(\d{4}\.\d{1,2}\.\d{1,2}).*?(\d{4}\.\d{1,2}\.\d{1,2})", name)
    if m:
        return f"{m.group(1)} ~ {m.group(2)}"
    # 匹配 20260331-20260716 这种(8位数字)
    m = re.search(r"(\d{8})-(\d{8})", name)
    if m:
        a, b = m.group(1), m.group(2)
        return f"{a[:4]}.{a[4:6]}.{a[6:]} ~ {b[:4]}.{b[4:6]}.{b[6:]}"
    # 单个日期
    m = re.search(r"(\d{4}\.\d{1,2}\.\d{1,2})", name)
    if m:
        return m.group(1)
    return ""


# [1/5] 复制平台数据
print("[1/5] 复制平台数据...")
copied = 0
if PLATFORM.exists():
    for f in sorted(PLATFORM.glob("*.xlsx")):
        if f.name.startswith("~$"):
            continue
        shutil.copy2(f, DATA / f.name)
        cat = classify_file(f.name)
        tag = extract_date_tag(f.name)
        synced.setdefault(cat, []).append((tag, f.name))
        print(f"  {f.name}")
        copied += 1
    print(f"  复制完成 ({copied} 个文件)")
else:
    print(f"  [WARN] 平台数据目录不存在: {PLATFORM}")

# [2/5] 扫描最新调运数据
print()
print("[2/5] 扫描最新调运数据...")
pattern = re.compile(r"(\d{8})-(\d{8}).*二次去重版.*\.xlsx")
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
    print(f"  最新调运数据: {sd} ~ {ed}")
    print(f"  文件: {latest_file.name}")
    transport_info = (f"{sd} ~ {ed}", latest_file.name)

    if PLATFORM.exists():
        shutil.copy2(latest_file, PLATFORM / latest_file.name)
        print(f"  -> 已复制到平台数据目录")
    shutil.copy2(latest_file, DATA / latest_file.name)
    print(f"  -> 已复制到 data/")

    # 合并到 synced
    synced.setdefault("猪只调运数据", []).append((f"{sd} ~ {ed}", latest_file.name))

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
            deleted_files.append(f.name)
            f.unlink()
if not deleted_files:
    print("  无需清理")
else:
    print(f"  清理完成 ({len(deleted_files)} 个旧文件)")

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
print("        本次同步数据清单")
print("=" * 55)

# 按类别有序展示
order = ["涌益日度数据", "涌益周度数据", "猪只调运数据", "神农肉业-鲜品价格", "神农肉业-冻品价格"]
shown = set()
for cat in order:
    if cat in synced:
        items = synced[cat]
        print(f"  [{cat}]")
        for tag, fname in items:
            if tag:
                print(f"    {tag}")
            else:
                print(f"    {fname}")
        shown.add(cat)

# 其他未分类
for cat, items in synced.items():
    if cat not in shown:
        print(f"  [{cat}]")
        for tag, fname in items:
            print(f"    {fname}")

if transport_info:
    ttag, tfname = transport_info
    # 调运数据已在上面显示

if deleted_files:
    print()
    print(f"  [已清理旧文件] {len(deleted_files)} 个")

if changed_files:
    print()
    print(f"  [Git] 已提交 {len(changed_files)} 个变更，已推送到 GitHub")
else:
    print()
    print(f"  [Git] 无变更")

print("=" * 55)
