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

# 收集同步信息（用文件名去重）
synced = {}          # {类别: [(日期标签, 文件名)]}
seen_files = set()   # 已记录的文件名，避免重复
transport_info = None
deleted_files = []   # 已删除的文件名列表（用于展示）
deleted_names = set()  # 已删除的文件名集合（用于去重摘要）


def classify_file(name: str) -> str:
    if "日度" in name:
        return "涌益日度数据"
    if "周度" in name:
        return "涌益周度数据"
    if "调运" in name:
        return "猪只调运数据"
    if "鲜品" in name:
        return "神农肉业-鲜品价格"
    if "冻品" in name:
        return "神农肉业-冻品价格"
    return "其他"


def extract_date_tag(name: str) -> str:
    m = re.search(r"(\d{4})年(\d{1,2})月(\d{1,2})日", name)
    if m:
        return f"{m.group(1)}.{int(m.group(2)):02d}.{int(m.group(3)):02d}"
    m = re.search(r"(\d{4}\.\d{1,2}\.\d{1,2}).*?(\d{4}\.\d{1,2}\.\d{1,2})", name)
    if m:
        return f"{m.group(1)} ~ {m.group(2)}"
    m = re.search(r"(\d{8})-(\d{8})", name)
    if m:
        a, b = m.group(1), m.group(2)
        return f"{a[:4]}.{a[4:6]}.{a[6:]} ~ {b[:4]}.{b[4:6]}.{b[6:]}"
    m = re.search(r"(\d{4}\.\d{1,2}\.\d{1,2})", name)
    if m:
        return m.group(1)
    return ""


def add_synced(cat: str, tag: str, fname: str):
    """添加同步记录，按文件名去重"""
    if fname not in seen_files:
        seen_files.add(fname)
        synced.setdefault(cat, []).append((tag, fname))


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
        add_synced(cat, tag, f.name)
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

    # 复制到平台数据目录
    if PLATFORM.exists():
        shutil.copy2(latest_file, PLATFORM / latest_file.name)
        print(f"  -> 已复制到平台数据目录")

        # 清理平台数据目录中被取代的旧调运文件
        for old in PLATFORM.glob("*调运*.xlsx"):
            if old.name.startswith("~$") or old.name == latest_file.name:
                continue
            if "全量合并" in old.name or "二次去重" in old.name:
                print(f"  清理平台旧调运: {old.name}")
                old.unlink()
                deleted_files.append(f"[平台] {old.name}")
                deleted_names.add(old.name)

    # 复制到 data/
    shutil.copy2(latest_file, DATA / latest_file.name)
    print(f"  -> 已复制到 data/")

    # 清理 data/ 中被取代的旧调运文件
    for old in DATA.glob("*调运*.xlsx"):
        if old.name.startswith("~$") or old.name == latest_file.name:
            continue
        if "全量合并" in old.name or "二次去重" in old.name:
            print(f"  清理 data 旧调运: {old.name}")
            old.unlink()
            deleted_files.append(f"[data] {old.name}")
            deleted_names.add(old.name)

    # 从摘要中清除被删除的旧调运记录
    for cat in list(synced.keys()):
        synced[cat] = [(t, n) for t, n in synced[cat] if n not in deleted_names]
        if not synced[cat]:
            del synced[cat]
    seen_files.difference_update(deleted_names)

    # 记录最新调运到摘要
    add_synced("猪只调运数据", f"{sd} ~ {ed}", latest_file.name)

    if len(transport_files) > 1:
        old = transport_files[1][0]
        print(f"  (调运目录共 {len(transport_files)} 个日期段，上次截止 {old[:4]}.{old[4:6]}.{old[6:]})")
else:
    print(f"  [WARN] 未在 {TRANSPORT_DIR} 找到调运分析文件")

# [3/5] 清理 data/ 中平台目录已不存在的旧文件
print()
print("[3/5] 清理旧文件...")
if PLATFORM.exists():
    src_names = {f.name for f in PLATFORM.glob("*.xlsx")}
    src_names.update(f.name for f in TRANSPORT_DIR.glob("*.xlsx") if not f.name.startswith("~$"))
    for f in DATA.glob("*.xlsx"):
        if f.name not in src_names and not f.name.startswith("~$"):
            print(f"  删除 {f.name}")
            deleted_files.append(f.name)
            deleted_names.add(f.name)
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

for cat, items in synced.items():
    if cat not in shown:
        print(f"  [{cat}]")
        for tag, fname in items:
            print(f"    {fname}")

if deleted_files:
    print()
    print(f"  [已清理] {len(deleted_files)} 个旧文件")

if changed_files:
    print()
    print(f"  [Git] 已提交并推送到 GitHub")
else:
    print()
    print(f"  [Git] 无变更")

print("=" * 55)
