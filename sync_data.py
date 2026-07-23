"""
一键同步：平台数据 + 调运数据 -> data/ -> GitHub
双击 同步数据到GitHub.bat 即可运行

调运数据策略：找到最新的二次去重文件，将其中新日期追加到全量合并文件，
生成一个日期范围更大的统一全量文件。
"""
import shutil, re, subprocess
from pathlib import Path
import pandas as pd

BASE = Path(__file__).parent
PLATFORM = Path(r"D:\CC\Desktop\平台数据")
TRANSPORT_DIR = Path(r"C:\Users\CC")
DATA = BASE / "data"
DATA.mkdir(exist_ok=True)

synced = {}
seen_files = set()
deleted_files = []
transport_new_dates = None  # 追加了多少天新数据


def classify_file(name: str) -> str:
    if "日度" in name: return "涌益日度数据"
    if "周度" in name: return "涌益周度数据"
    if "调运" in name: return "猪只调运数据"
    if "鲜品" in name: return "神农肉业-鲜品价格"
    if "冻品" in name: return "神农肉业-冻品价格"
    return "其他"


def extract_date_tag(name: str) -> str:
    m = re.search(r"(\d{4})年(\d{1,2})月(\d{1,2})日", name)
    if m: return f"{m.group(1)}.{int(m.group(2)):02d}.{int(m.group(3)):02d}"
    m = re.search(r"(\d{4}\.\d{1,2}\.\d{1,2}).*?(\d{4}\.\d{1,2}\.\d{1,2})", name)
    if m: return f"{m.group(1)} ~ {m.group(2)}"
    m = re.search(r"(\d{8})-(\d{8})", name)
    if m:
        a, b = m.group(1), m.group(2)
        return f"{a[:4]}.{a[4:6]}.{a[6:]} ~ {b[:4]}.{b[4:6]}.{b[6:]}"
    m = re.search(r"(\d{4}\.\d{1,2}\.\d{1,2})", name)
    if m: return m.group(1)
    return ""


def add_synced(cat: str, tag: str, fname: str):
    if fname not in seen_files:
        seen_files.add(fname)
        synced.setdefault(cat, []).append((tag, fname))


# [1/5] 复制平台数据到 data/（排除调运文件，调运在 [2/5] 单独处理）
print("[1/5] 复制平台数据...")
copied = 0
if PLATFORM.exists():
    for f in sorted(PLATFORM.glob("*.xlsx")):
        if f.name.startswith("~$") or "调运" in f.name:
            continue
        shutil.copy2(f, DATA / f.name)
        add_synced(classify_file(f.name), extract_date_tag(f.name), f.name)
        print(f"  {f.name}")
        copied += 1
    print(f"  复制完成 ({copied} 个文件)")
else:
    print(f"  [WARN] 平台数据目录不存在: {PLATFORM}")

# [2/5] 调运数据：追加新日期到全量文件
print()
print("[2/5] 调运数据：合并追加新日期...")

# 2a. 找到现有的全量合并文件
existing_full = None
for d in [DATA, PLATFORM] if PLATFORM.exists() else [DATA]:
    for f in sorted(d.glob("*全量合并*.xlsx"), reverse=True):
        if not f.name.startswith("~$"):
            existing_full = f
            break
    if existing_full:
        break

# 2b. 找到最新的二次去重文件
pattern = re.compile(r"(\d{8})-(\d{8}).*二次去重版.*\.xlsx")
transport_files = []
for f in TRANSPORT_DIR.glob("*.xlsx"):
    if f.name.startswith("~$"): continue
    m = pattern.search(f.name)
    if m:
        transport_files.append((m.group(2), m.group(1), f))

if not transport_files:
    print(f"  [WARN] 未在 {TRANSPORT_DIR} 找到二次去重文件")
else:
    transport_files.sort(key=lambda x: (x[0], x[1]), reverse=True)
    latest_end, latest_start, latest_file = transport_files[0]
    sd = f"{latest_start[:4]}.{latest_start[4:6]}.{latest_start[6:]}"
    ed = f"{latest_end[:4]}.{latest_end[4:6]}.{latest_end[6:]}"
    print(f"  最新二次去重: {sd} ~ {ed}  ({latest_file.name})")

    if existing_full and existing_full.exists():
        # 加载现有全量文件
        df_full = pd.read_excel(existing_full)
        full_max_date = pd.to_datetime(df_full.iloc[:, 6] if df_full.shape[1] > 6 else df_full.iloc[:, 0], errors="coerce").max()
        print(f"  现有全量文件: {existing_full.name}")
        print(f"  全量截止日期: {full_max_date.date()}")

        # 加载二次去重文件
        df_new = pd.read_excel(latest_file)
        new_date_col = df_new.iloc[:, 6] if df_new.shape[1] > 6 else df_new.iloc[:, 0]
        df_new["_date_parsed"] = pd.to_datetime(new_date_col, errors="coerce")

        # 只保留全量截止日期之后的新数据
        df_really_new = df_new[df_new["_date_parsed"] > full_max_date].copy()
        df_really_new = df_really_new.drop(columns=["_date_parsed"])

        if len(df_really_new) > 0:
            new_min = df_new[df_new["_date_parsed"] > full_max_date]["_date_parsed"].min()
            new_max = df_new[df_new["_date_parsed"] > full_max_date]["_date_parsed"].max()
            print(f"  新增 {len(df_really_new)} 条记录 ({new_min.date()} ~ {new_max.date()})")

            # 合并
            merged = pd.concat([df_full, df_really_new], ignore_index=True)
            all_dates = pd.to_datetime(merged.iloc[:, 6] if merged.shape[1] > 6 else merged.iloc[:, 0], errors="coerce")
            new_start = all_dates.min().strftime("%Y%m%d")
            new_end = all_dates.max().strftime("%Y%m%d")
            new_name = f"{new_start}-{new_end}猪只调运智能分析结果（全量合并去重版）.xlsx"

            # 保存新全量文件
            new_path_data = DATA / new_name
            merged.to_excel(new_path_data, index=False)
            print(f"  -> 生成新全量文件: {new_name}")

            new_path_platform = PLATFORM / new_name if PLATFORM.exists() else None
            if new_path_platform:
                merged.to_excel(new_path_platform, index=False)
                print(f"  -> 已复制到平台数据目录")

            # 删除旧全量文件
            for old in list(DATA.glob("*全量合并*.xlsx")) + list(PLATFORM.glob("*全量合并*.xlsx") if PLATFORM.exists() else []):
                if old.name != new_name and not old.name.startswith("~$"):
                    print(f"  清理旧全量: {old.name}")
                    old.unlink()
                    deleted_files.append(old.name)

            transport_new_dates = f"{new_min.date()} ~ {new_max.date()}"
            add_synced("猪只调运数据",
                       f"{new_start[:4]}.{new_start[4:6]}.{new_start[6:]} ~ {new_end[:4]}.{new_end[4:6]}.{new_end[6:]}",
                       new_name)
        else:
            print(f"  二次去重文件没有比全量更新的数据，跳过")
            add_synced("猪只调运数据",
                       f"{existing_full.name[:8]}-{existing_full.name[9:17]}",
                       existing_full.name)
    else:
        # 没有现有全量文件：直接把二次去重作为初始全量
        df_new = pd.read_excel(latest_file)
        date_col = df_new.iloc[:, 6] if df_new.shape[1] > 6 else df_new.iloc[:, 0]
        dates = pd.to_datetime(date_col, errors="coerce")
        start = dates.min().strftime("%Y%m%d")
        end = dates.max().strftime("%Y%m%d")
        new_name = f"{start}-{end}猪只调运智能分析结果（全量合并去重版）.xlsx"

        new_path_data = DATA / new_name
        df_new.to_excel(new_path_data, index=False)
        print(f"  -> 生成初始全量文件: {new_name}")

        if PLATFORM.exists():
            df_new.to_excel(PLATFORM / new_name, index=False)
            print(f"  -> 已复制到平台数据目录")

        add_synced("猪只调运数据", f"{start[:4]}.{start[4:6]}.{start[6:]} ~ {end[:4]}.{end[4:6]}.{end[6:]}", new_name)

    # 清理 data/ 中的旧二次去重文件（已合并到全量）
    for old in DATA.glob("*二次去重*.xlsx"):
        if not old.name.startswith("~$"):
            print(f"  清理 data 二次去重: {old.name}")
            old.unlink()
            deleted_files.append(old.name)

    if len(transport_files) > 1:
        old = transport_files[1][0]
        print(f"  (调运目录共 {len(transport_files)} 个日期段，上次截止 {old[:4]}.{old[4:6]}.{old[6:]})")

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

if transport_new_dates:
    print()
    print(f"  [调运追加] 新增 {transport_new_dates} 的数据")

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
