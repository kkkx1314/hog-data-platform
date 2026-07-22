@echo off
chcp 65001 >nul
echo 🐷 同步平台数据 + 调运数据到 GitHub...
echo.
cd /d "C:\Users\CC\test-claude"

python -c "
import shutil, re, subprocess
from pathlib import Path

BASE = Path(r'C:\Users\CC\test-claude')
PLATFORM = Path(r'D:\CC\Desktop\平台数据')
TRANSPORT_DIR = Path(r'C:\Users\CC')
DATA = BASE / 'data'
DATA.mkdir(exist_ok=True)

# ── [1/5] 复制平台数据到 data/ ──
print('[1/5] 复制平台数据...')
copied = 0
if PLATFORM.exists():
    for f in PLATFORM.glob('*.xlsx'):
        if f.name.startswith('~$'): continue
        shutil.copy2(f, DATA / f.name)
        print(f'  {f.name}')
        copied += 1
    print(f'  复制完成 ({copied} 个文件)')
else:
    print(f'  ⚠️ 平台数据目录不存在: {PLATFORM}')

# ── [2/5] 查找最新调运分析文件 ──
print()
print('[2/5] 扫描最新调运数据...')
pattern = re.compile(r'(\d{8})-(\d{8}).*二次去重版\.xlsx')
transport_files = []
for f in TRANSPORT_DIR.glob('*.xlsx'):
    if f.name.startswith('~$'): continue
    m = pattern.search(f.name)
    if m:
        transport_files.append((m.group(2), m.group(1), f))  # (end_date, start_date, path)

if transport_files:
    transport_files.sort(key=lambda x: (x[0], x[1]), reverse=True)
    latest_end, latest_start, latest_file = transport_files[0]
    date_label = f'{latest_start[:4]}.{latest_start[4:6]}.{latest_start[6:]}-{latest_end[:4]}.{latest_end[4:6]}.{latest_end[6:]}'
    print(f'  最新调运数据: {date_label}')
    print(f'  文件: {latest_file.name}')

    # 复制到平台数据目录
    if PLATFORM.exists():
        shutil.copy2(latest_file, PLATFORM / latest_file.name)
        print(f'  → 已复制到 {PLATFORM / latest_file.name}')

    # 复制到 data/
    shutil.copy2(latest_file, DATA / latest_file.name)
    print(f'  → 已复制到 {DATA / latest_file.name}')

    if len(transport_files) > 1:
        old_end = transport_files[1][0]
        print(f'  (调运目录共 {len(transport_files)} 个日期段，上次截止 {old_end[:4]}.{old_end[4:6]}.{old_end[6:]})')
else:
    print(f'  ⚠️ 未在 {TRANSPORT_DIR} 找到调运分析文件')

# ── [3/5] 清理 data/ 中旧文件 ──
print()
print('[3/5] 清理旧文件...')
if PLATFORM.exists():
    src_names = {f.name for f in PLATFORM.glob('*.xlsx')}
    # 也保留调运文件
    src_names.update(f.name for f in TRANSPORT_DIR.glob('*.xlsx') if not f.name.startswith('~$'))
    for f in DATA.glob('*.xlsx'):
        if f.name not in src_names and not f.name.startswith('~$'):
            print(f'  删除 {f.name}')
            f.unlink()
print('  清理完成')

# ── [4/5] Git 提交 ──
print()
print('[4/5] 提交到 Git...')
subprocess.run(['git', 'add', 'data/'], cwd=str(BASE), check=True)

# 检查是否有变更
diff = subprocess.run(['git', 'diff', '--cached', '--name-only'], cwd=str(BASE), capture_output=True, text=True)
if diff.stdout.strip():
    subprocess.run(['git', 'commit', '-m', 'data: 同步平台数据+调运数据'], cwd=str(BASE))
    print('  提交成功')
else:
    print('  数据无变化，跳过提交')

# ── [5/5] 推送 ──
print()
print('[5/5] 推送到 GitHub...')
subprocess.run(['git', 'push'], cwd=str(BASE), check=True)

print()
print('✅ 同步完成！')
"

pause
