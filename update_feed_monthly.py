# -*- coding: utf-8 -*-
"""把「2026年饲料工业生产情况表(本级汇总).xlsx」的最新月度数据更新到四个图表版文件。

与旧版（只追加当月）相比，这一版会：
- 用源文件覆盖所有已存在月份的数值（协会常修订上月，如本期 2026-07/08 上调）；
- 追加源文件里新增的月份（本期 2026-09）；
- 由「月度数据」整体重算「月度环比 / 月度同比」（历史值与旧表逐格复核 0 差异）；
- 同步三张隐藏图表辅助表的 2026 年列，并下延条件格式；
- 保留原图表部件（清 numCache，Excel 打开时重绘）。

用法：python update_feed_monthly.py [源文件路径] [--no-dashboard]
默认在更新完四份 Excel 后，顺带重建线上看板快照 data/feed_dashboard.json。
"""
import copy
import datetime
import os
import re
import shutil
import sys
import time
import warnings
import zipfile

import pandas as pd
import openpyxl
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.utils import get_column_letter

warnings.filterwarnings('ignore')

_ARGS = [a for a in sys.argv[1:] if a != '--no-dashboard']
SRC = (_ARGS[0] if _ARGS else
       r'D:\CC\Desktop\2026年饲料工业生产情况表(本级汇总).xlsx')
DST_DIR = r'D:/CC/Desktop/'
STAMP = datetime.datetime.now().strftime('%Y%m%d')

CATS9 = ['仔猪饲料', '哺乳仔猪饲料', '保育饲料', '母猪饲料', '后备母猪饲料',
         '其他母猪饲料', '生长肥育猪饲料', '其他猪饲料', '猪料']
CATS8 = ['仔猪饲料', '哺乳仔猪饲料', '保育饲料', '母猪饲料', '后备母猪饲料',
         '其他母猪饲料', '生长肥育猪饲料', '猪料']

# 源表列号（0-based iloc）
COL_HP = {'猪料': 50, '仔猪饲料': 51, '哺乳仔猪饲料': 52, '保育饲料': 53,
          '母猪饲料': 54, '后备母猪饲料': 55, '其他母猪饲料': 56,
          '生长肥育猪饲料': 57, '其他猪饲料': 58}
COL_NS = {'猪料': 83, '仔猪饲料': 84, '哺乳仔猪饲料': 85, '保育饲料': 86,
          '母猪饲料': 87, '后备母猪饲料': 88, '其他母猪饲料': 89,
          '生长肥育猪饲料': 90, '其他猪饲料': 91}
COL_TJJ = {'猪料': 116, '仔猪饲料': 117, '哺乳仔猪饲料': 122, '保育饲料': 127,
           '母猪饲料': 132, '后备母猪饲料': 137, '其他母猪饲料': 142,
           '生长肥育猪饲料': 147, '其他猪饲料': 152}
COL_ALL = {'猪料': 21, '仔猪饲料': 27, '哺乳仔猪饲料': 28, '保育饲料': 29,
           '母猪饲料': 30, '后备母猪饲料': 31, '其他母猪饲料': 32,
           '生长肥育猪饲料': 33}

TARGETS = [
    ('饲料_配合饲料_猪料整理_图表版.xlsx', CATS9, COL_HP),
    ('饲料_浓缩饲料_猪料整理_图表版.xlsx', CATS9, COL_NS),
    ('饲料_添加剂预混合饲料_猪料整理_图表版.xlsx', CATS9, COL_TJJ),
    ('饲料_猪料各品类整理_图表版.xlsx', CATS8, COL_ALL),
]


def fnum(v):
    if isinstance(v, (int, float)) and not pd.isna(v):
        return float(v)
    return None


def load_src(months_needed=None):
    """返回 {'202609': {cat: value, ...}, ...}；没有数据的月份不入表。"""
    df = pd.read_excel(SRC, sheet_name='云南饲基31表', header=0)
    out = {}
    for _, row in df.iterrows():
        m = str(row.iloc[1]).strip()
        if not re.fullmatch(r'20\d{4}', m):
            continue
        if months_needed and m not in months_needed:
            continue
        out[m] = row
    return out


def month_key(d):
    return f'{d.year}{d.month:02d}'


def row_date(ws, row):
    v = ws.cell(row, 1).value
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, (int, float)):
        from openpyxl.utils.datetime import from_excel
        try:
            return from_excel(v).date()
        except Exception:
            return None
    return None


def find_rows(ws):
    """返回 {(year, month): row_number}"""
    out = {}
    for r in range(2, ws.max_row + 1):
        d = row_date(ws, r)
        if d:
            out[(d.year, d.month)] = r
    return out


def add_row_like(ws, src_row, new_row, date_value):
    """在 new_row 追加一行，样式/行高照抄 src_row。"""
    ws.cell(new_row, 1).value = date_value
    for c in range(1, ws.max_column + 1):
        s = ws.cell(src_row, c)
        t = ws.cell(new_row, c)
        if c == 1:
            t.number_format = s.number_format
        elif s.has_style:
            t._style = copy.copy(s._style)
    h = ws.row_dimensions[src_row].height
    if h is not None:
        ws.row_dimensions[new_row].height = h


def extend_cf(ws, old_last, n=1):
    """把以 old_last 结尾的 CF 范围下延 n 行。"""
    new_cfs = []
    for cf in ws.conditional_formatting:
        parts = []
        for rng in cf.sqref.ranges:
            c1 = f'{get_column_letter(rng.min_col)}{rng.min_row}'
            end = rng.max_row + (n if rng.max_row == old_last else 0)
            c2 = f'{get_column_letter(rng.max_col)}{end}'
            parts.append(f'{c1}:{c2}' if c1 != c2 else c1)
        new_cfs.append((' '.join(parts), list(cf.rules)))
    ws.conditional_formatting = ConditionalFormattingList()
    for rngstr, rules in new_cfs:
        for rule in rules:
            ws.conditional_formatting.add(rngstr, rule)


def strip_chart_cache(txt):
    for tag in ('numCache', 'strCache', 'multiLvlStrCache'):
        txt = re.sub(rf'<c:{tag}>.*?</c:{tag}>', '', txt, flags=re.S)
    return txt


def swap_chart_parts(saved_path, bak_path):
    """恢复备份中的原生图表部件，并清除缓存。"""
    with zipfile.ZipFile(bak_path, 'r') as zb:
        bak_names = set(zb.namelist())
        bak_charts = sorted(n for n in bak_names if re.match(r'xl/charts/chart\d+\.xml$', n))
        bak_drawings = sorted(n for n in bak_names if re.match(r'xl/drawings/drawing\d+\.xml$', n))
        bak_drels = sorted(n for n in bak_names if re.match(r'xl/drawings/_rels/drawing\d+\.xml\.rels$', n))
        orig = {n: zb.read(n) for n in bak_charts + bak_drawings + bak_drels}

    with zipfile.ZipFile(saved_path, 'r') as zs:
        names = zs.namelist()
        items = {n: zs.read(n) for n in names}
    saved_charts = sorted(n for n in names if re.match(r'xl/charts/chart\d+\.xml$', n))

    same = (saved_charts == bak_charts)
    if same:
        for n in bak_charts:
            items[n] = strip_chart_cache(orig[n].decode('utf-8')).encode('utf-8')
        for n in bak_drawings + bak_drels:
            items[n] = orig[n]
    else:
        for n in saved_charts:
            items[n] = strip_chart_cache(items[n].decode('utf-8')).encode('utf-8')

    tmp = saved_path + '.tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
        for n, data in items.items():
            zo.writestr(n, data)
    for _ in range(5):
        try:
            os.replace(tmp, saved_path)
            break
        except PermissionError:
            time.sleep(2)
    else:
        raise SystemExit(f'无法替换 {saved_path}（文件被占用，请关闭 Excel 后重试；'
                         f'更新数据在 {tmp}）')
    return same


def block_category(ws, start, series_2025):
    """用 2025 年列匹配，判断该辅助表块属于哪个品类。"""
    ycol = None
    for c in range(1, ws.max_column + 1):
        if ws.cell(start, c).value in ('2025年', 2025):
            ycol = c
            break
    if not ycol:
        return None
    vals = [ws.cell(start + m, ycol).value for m in range(1, 13)]
    best, score = None, -1
    for cat, series in series_2025.items():
        s = sum(1 for a, b in zip(vals, series)
                if a is not None and b is not None and abs(float(a) - float(b)) < 0.011)
        if s > score:
            best, score = cat, s
    return best if score >= 10 else None


def update_helper(ws_h, ws_src, cats):
    """把辅助表各品类块的 2026 年列，按 ws_src 重写 1-12 月。"""
    starts = [r for r in range(1, ws_h.max_row + 1) if ws_h.cell(r, 1).value == '月份']
    src_hdr = [ws_src.cell(1, c).value for c in range(1, ws_src.max_column + 1)]
    src_rows = find_rows(ws_src)
    series_2025 = {
        cat: [ws_src.cell(src_rows[(2025, m)], src_hdr.index(cat) + 1).value
              if (2025, m) in src_rows else None for m in range(1, 13)]
        for cat in cats
    }
    # 2026 年目标列
    cols = {}
    if starts:
        for c in range(1, ws_h.max_column + 1):
            v = ws_h.cell(starts[0], c).value
            if v in ('2026年', 2026):
                cols['2026'] = c
    if '2026' not in cols:
        raise SystemExit(f'{ws_h.title}: 找不到 2026 年列')
    written = 0
    for st in starts:
        cat = block_category(ws_h, st, series_2025)
        if not cat:
            raise SystemExit(f'{ws_h.title}: 第 {st} 行块无法匹配品类')
        col = src_hdr.index(cat) + 1
        for m in range(1, 13):
            r = st + m
            if (2026, m) in src_rows:
                val = ws_src.cell(src_rows[(2026, m)], col).value
                ws_h.cell(r, cols['2026']).value = (
                    round(float(val), 2) if isinstance(val, (int, float)) else None)
                written += 1
            else:
                ws_h.cell(r, cols['2026']).value = None
    return written


def process(fname, cats, colmap, src):
    path = os.path.join(DST_DIR, fname)
    lock = os.path.join(DST_DIR, '~$' + fname)
    locked = os.path.exists(lock)
    out_path = path.replace('.xlsx', '_NEW.xlsx') if locked else path
    bak_path = path.replace('.xlsx', f'.bak_{STAMP}.xlsx')
    if not locked and not os.path.exists(bak_path):
        shutil.copy2(path, bak_path)

    wb = openpyxl.load_workbook(path)
    md = wb['月度数据']
    mom = wb['月度环比']
    yoy = wb['月度同比']
    h_val = wb['图表数据']
    h_mom = wb['月度环比图数据']
    h_yoy = wb['月度同比图数据']

    md_hdr = [md.cell(1, c).value for c in range(1, md.max_column + 1)]
    missing = [c for c in cats if c not in md_hdr[1:]]
    if missing:
        raise SystemExit(f'{fname}: 月度数据缺列 {missing}')

    # ---- 1) 月度数据：覆盖修订 + 追加新月份 ----
    md_rows = find_rows(md)
    last_date = max(md_rows)
    last_row = md_rows[last_date]
    added = 0
    revised = 0
    for m, row in src.items():
        vals = {c: fnum(row.iloc[colmap[c]]) for c in cats}
        if any(v is None for v in vals.values()):
            continue  # 该月尚未填报（如 2026-10）
        y, mo = int(m[:4]), int(m[4:])
        key = (y, mo)
        if key in md_rows:
            r = md_rows[key]
            for cat in cats:
                col = md_hdr.index(cat) + 1
                old = md.cell(r, col).value
                new = vals[cat]
                if old is None or abs(float(old) - new) > 0.001:
                    md.cell(r, col).value = new
                    revised += 1
        elif key > last_date:
            last_row += 1
            add_row_like(md, last_row - 1, last_row, datetime.datetime(y, mo, 1))
            for cat in cats:
                col = md_hdr.index(cat) + 1
                md.cell(last_row, col).value = vals[cat]
            last_date = key
            added += 1

    # 重新建立完整序列（供环比/同比/辅助表使用）
    md_rows = find_rows(md)
    series = {}
    for (y, mo), r in md_rows.items():
        series[(y, mo)] = {cat: md.cell(r, md_hdr.index(cat) + 1).value for cat in cats}

    # ---- 2) 月度环比：整体重算 + 追加 + CF 下延 ----
    def rebuild(ws, kind):
        rows = find_rows(ws)
        hdr = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        n_add = 0
        last = max(rows)
        last_r = rows[last]
        for key in sorted(series):
            y, mo = key
            if kind == 'mom':
                py, pm = (y - 1, 12) if mo == 1 else (y, mo - 1)
            else:
                py, pm = y - 1, mo
            if (py, pm) not in series:
                continue
            if key in rows:
                r = rows[key]
            else:
                if key <= last:
                    continue
                last_r += 1
                add_row_like(ws, last_r - 1, last_r, datetime.datetime(y, mo, 1))
                rows[key] = last_r
                n_add += 1
                r = last_r
            for cat in cats:
                cur = series[key].get(cat)
                prev = series[(py, pm)].get(cat)
                if cur is None or prev in (None, 0):
                    v = None
                else:
                    v = round((float(cur) / float(prev) - 1) * 100, 2)
                ws.cell(r, hdr.index(cat) + 1).value = v
        if n_add:
            extend_cf(ws, last_r - n_add, n_add)
        return n_add

    add_mom = rebuild(mom, 'mom')
    add_yoy = rebuild(yoy, 'yoy')

    # ---- 3) 辅助表：2026 年列 ----
    w1 = update_helper(h_val, md, cats)
    w2 = update_helper(h_mom, mom, cats)
    w3 = update_helper(h_yoy, yoy, cats)

    wb.save(out_path)
    same = swap_chart_parts(out_path, bak_path)
    print(f'  [完成] {os.path.basename(out_path)}  修订 {revised} 格 / 新增 {added} 月 '
          f'(环比+{add_mom}, 同比+{add_yoy}) / 辅助表 {w1 + w2 + w3} 格 / 图表原样={same}')
    return out_path


if __name__ == '__main__':
    src_all = load_src()
    print(f'源文件: {SRC}')
    print(f'可用月份: {sorted(src_all)[-6:]}')
    for fname, cats, colmap in TARGETS:
        print(f'== {fname}')
        process(fname, cats, colmap, src_all)
    print('全部完成')

    # 月度补齐后，顺带补齐已满 3 个月的季度（幂等，已是最新则跳过）
    import subprocess
    here = os.path.dirname(os.path.abspath(__file__))
    quarter_script = os.path.join(here, 'update_feed_quarterly.py')
    print('\n== 补齐季度数据 ==')
    rq = subprocess.run([sys.executable, quarter_script], cwd=here)
    if rq.returncode:
        print('[WARN] 季度数据补齐失败；可稍后单独运行 update_feed_quarterly.py')

    if '--no-dashboard' not in sys.argv:
        script = os.path.join(here, 'build_feed_dashboard_data.py')
        print('\n== 生成线上看板快照 ==')
        r = subprocess.run([sys.executable, script], cwd=here)
        if r.returncode:
            print('[WARN] 看板快照生成失败（Excel 已更新）；'
                  '可稍后单独运行 build_feed_dashboard_data.py')
        else:
            site = os.path.join(here, 'build_feed_public_site.py')
            print('\n== 生成公开页 docs/feed/index.html ==')
            r2 = subprocess.run([sys.executable, site], cwd=here)
            if r2.returncode:
                print('[WARN] 公开页生成失败；可稍后单独运行 build_feed_public_site.py')
