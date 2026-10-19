# -*- coding: utf-8 -*-
"""给四个「饲料_*_猪料整理_图表版.xlsx」补 2026 年第 3 季度，并修复新月度行格式。

做三件事：
1. 季度数据 / 季度环比 / 季度同比：追加 2026年第3季度
   - 季度数据 = 2026 年 7/8/9 月之和（保留 2 位）
   - 季度环比 = 本季 ÷ 上季 − 1
   - 季度同比 = 本季 ÷ 上年同季 − 1
2. 新行样式完整照抄上一行（宋体11加粗、居中、细边框、原数字格式），
   并把 季度环比/季度同比 的条件格式范围下延到新行。
3. 修复 2026-09 月度行 A 列：此前追加行只复制了日期数字格式，
   丢了字体/边框/居中；这里按历史行的统一样式补齐。
   （月度数据 / 月度环比 / 月度同比 三张表同样处理）

保存后沿用 update_feed_monthly 的图表部件恢复逻辑，保留原生图表。

用法：python update_feed_quarterly.py [--dry-run]
"""
import copy
import datetime
import os
import re
import shutil
import sys
import warnings

import openpyxl

from update_feed_monthly import swap_chart_parts

warnings.filterwarnings('ignore')

DST_DIR = r'D:/CC/Desktop/'
STAMP = datetime.datetime.now().strftime('%Y%m%d')

FILES = [
    ('饲料_配合饲料_猪料整理_图表版.xlsx', 9),
    ('饲料_浓缩饲料_猪料整理_图表版.xlsx', 9),
    ('饲料_添加剂预混合饲料_猪料整理_图表版.xlsx', 9),
    ('饲料_猪料各品类整理_图表版.xlsx', 8),
]

RATE_SHEETS = ['季度环比', '季度同比']


def last_data_row(ws):
    last = ws.max_row
    while last > 1 and all(ws.cell(last, c).value is None
                           for c in range(1, ws.max_column + 1)):
        last -= 1
    return last


def col_map(ws):
    """表头名 -> 列号。"""
    return {str(ws.cell(1, c).value).strip(): c
            for c in range(1, ws.max_column + 1)
            if ws.cell(1, c).value is not None}


def quarter_label(q):
    return f'{q[0]}年第{q[1]}季度'


def next_quarter(q):
    return (q[0] + 1, 1) if q[1] == 4 else (q[0], q[1] + 1)


def latest_complete_quarter(monthly, cats):
    """月度数据里最后一个 3 个月齐全的季度。"""
    if not monthly:
        return None
    y, m = max(monthly)
    q = (y, (m - 1) // 3 + 1)
    floor = min(yy for yy, _ in monthly)
    while q[0] >= floor:
        if all(quarter_value(monthly, q, c) is not None for c in cats):
            return q
        q = (q[0] - 1, 4) if q[1] == 1 else (q[0], q[1] - 1)
    return None


def read_monthly(wb):
    """{(year, month): {category: value}}"""
    ws = wb['月度数据']
    headers = col_map(ws)
    out = {}
    for r in range(2, last_data_row(ws) + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, datetime.datetime):
            y, m = v.year, v.month
        elif isinstance(v, (int, float)):
            from openpyxl.utils.datetime import from_excel
            d = from_excel(v)
            y, m = d.year, d.month
        else:
            continue
        out[(y, m)] = {name: ws.cell(r, c).value
                       for name, c in headers.items() if name != '月份'}
    return out


def read_quarterly(wb):
    """{(year, quarter): {category: value}}"""
    ws = wb['季度数据']
    headers = col_map(ws)
    out = {}
    for r in range(2, last_data_row(ws) + 1):
        m = re.match(r'(\d{4})年第(\d)季度', str(ws.cell(r, 1).value or ''))
        if not m:
            continue
        out[(int(m.group(1)), int(m.group(2)))] = {
            name: ws.cell(r, c).value
            for name, c in headers.items() if name != '季度'}
    return out


def quarter_value(monthly, q, cat):
    months = [f'{q[0]}-{((q[1] - 1) * 3 + i):02d}' for i in (1, 2, 3)]
    vals = [monthly.get((int(m[:4]), int(m[5:])), {}).get(cat) for m in months]
    if any(v is None for v in vals):
        return None
    return round(sum(float(v) for v in vals), 2)


def pct(cur, base):
    if cur is None or base in (None, 0):
        return None
    return round((float(cur) / float(base) - 1) * 100, 2)


def dominant_a_style(ws):
    """历史数据行里出现次数最多的 A 列样式（返回代表行号）。"""
    counts = {}
    reps = {}
    for r in range(2, last_data_row(ws) + 1):
        if ws.cell(r, 1).value is None:
            continue
        key = ws.cell(r, 1)._style
        sig = tuple(sorted(vars(key).items())) if hasattr(key, '__dict__') else str(key)
        counts[sig] = counts.get(sig, 0) + 1
        reps.setdefault(sig, r)
    if not counts:
        return None
    return reps[max(counts, key=counts.get)]


def fix_month_a_styles(wb):
    """把 A 列样式偏离历史主流样式的数据行修正过来（保留值和数字格式）。"""
    fixed = 0
    for name in ['月度数据', '月度环比', '月度同比']:
        ws = wb[name]
        rep = dominant_a_style(ws)
        if rep is None:
            continue
        ref_style = copy.copy(ws.cell(rep, 1)._style)
        ref_nf = ws.cell(rep, 1).number_format
        for r in range(2, last_data_row(ws) + 1):
            cell = ws.cell(r, 1)
            if cell.value is None:
                continue
            sig = tuple(sorted(vars(cell._style).items())) if hasattr(cell._style, '__dict__') else str(cell._style)
            ref_sig = tuple(sorted(vars(ref_style).items())) if hasattr(ref_style, '__dict__') else str(ref_style)
            if sig != ref_sig:
                cell._style = copy.copy(ref_style)
                cell.number_format = ref_nf
                fixed += 1
    return fixed


def style_copy_row(ws, src_row, dst_row):
    """整行样式照抄（含 A 列），并复制行高。"""
    for c in range(1, ws.max_column + 1):
        ws.cell(dst_row, c)._style = copy.copy(ws.cell(src_row, c)._style)
    h = ws.row_dimensions[src_row].height
    if h is not None:
        ws.row_dimensions[dst_row].height = h


def extend_cf_to(ws, old_last, new_last):
    """把以 old_last 结尾的条件格式范围下延到 new_last。"""
    from openpyxl.formatting.formatting import ConditionalFormattingList

    entries = []
    for cf in ws.conditional_formatting:
        ranges = []
        for rng in cf.sqref.ranges:
            end = new_last if rng.max_row == old_last else rng.max_row
            ranges.append((rng.min_col, rng.min_row, rng.max_col, end))
        entries.append((ranges, list(cf.rules)))
    ws.conditional_formatting = ConditionalFormattingList()
    from openpyxl.utils import get_column_letter
    for ranges, rules in entries:
        parts = []
        for c1, r1, c2, r2 in ranges:
            a = f'{get_column_letter(c1)}{r1}'
            b = f'{get_column_letter(c2)}{r2}'
            parts.append(a if a == b else f'{a}:{b}')
        for rule in rules:
            ws.conditional_formatting.add(' '.join(parts), rule)


def append_quarter(wb, q, monthly, quarterly):
    """追加季度数据/环比/同比三行；返回 (摘要, 是否实际追加)。"""
    cats = [k for k in col_map(wb['季度数据']).keys() if k != '季度']
    q_now = {c: quarter_value(monthly, q, c) for c in cats}
    if any(v is None for v in q_now.values()):
        raise SystemExit(f'季度 {quarter_label(q)} 有品类缺月，无法计算：'
                         f'{[c for c, v in q_now.items() if v is None]}')

    prev_q = (q[0], q[1] - 1) if q[1] > 1 else (q[0] - 1, 4)
    yoy_q = (q[0] - 1, q[1])
    rates_mom = {c: pct(q_now[c], quarterly[prev_q][c]) for c in cats}
    rates_yoy = {c: pct(q_now[c], quarterly[yoy_q][c]) for c in cats}

    summary = {'q': q_now, 'mom': rates_mom, 'yoy': rates_yoy}
    appended = 0

    for name, values in [('季度数据', q_now), ('季度环比', rates_mom), ('季度同比', rates_yoy)]:
        ws = wb[name]
        headers = col_map(ws)
        last = last_data_row(ws)
        # 幂等：已经存在该季度则只校验不重复追加
        existing = None
        for r in range(2, last + 1):
            if str(ws.cell(r, 1).value or '').strip() == quarter_label(q):
                existing = r
                break
        if existing is not None:
            continue
        new = last + 1
        style_copy_row(ws, last, new)
        ws.cell(new, 1).value = quarter_label(q)
        for cat, val in values.items():
            ws.cell(new, headers[cat]).value = val
        if name in RATE_SHEETS:
            extend_cf_to(ws, last, new)
        summary[name + '_row'] = new
        appended += 1
    return summary, appended


def process(fname, dry_run=False):
    path = os.path.join(DST_DIR, fname)
    lock = os.path.join(DST_DIR, '~$' + fname)
    locked = os.path.exists(lock)
    out_path = path.replace('.xlsx', '_NEW.xlsx') if locked else path
    bak_path = path.replace('.xlsx', f'.bak_{STAMP}_quarterly.xlsx')
    if not locked and not os.path.exists(bak_path):
        shutil.copy2(path, bak_path)

    wb = openpyxl.load_workbook(path)
    monthly = read_monthly(wb)
    quarterly = read_quarterly(wb)

    # 历史校验：季度值应等于 3 个月之和（抽查最近 4 个季度）
    recent = sorted(quarterly)[-4:]
    for q in recent:
        for cat, v in quarterly[q].items():
            calc = quarter_value(monthly, q, cat)
            if calc is not None and abs(calc - float(v)) > 0.02:
                raise SystemExit(
                    f'{fname}: {quarter_label(q)} {cat} 季度值与月度和不符 '
                    f'({v} vs {calc})，请先核对月度数据')

    fixed = fix_month_a_styles(wb)

    cats = [k for k in col_map(wb['季度数据']).keys() if k != '季度']
    target = latest_complete_quarter(monthly, cats)
    if target is None:
        wb.close()
        print('  [跳过] 还没有 3 个月齐全的季度')
        return out_path, None

    # 从已有最后一个季度往后补（不补历史缺口，如 2023Q1）
    q = next_quarter(max(quarterly))
    summary = None
    added = []
    appended = 0
    while q <= target:
        summary, n = append_quarter(wb, q, monthly, quarterly)
        quarterly[q] = summary['q']
        added.append(q)
        appended += n
        q = next_quarter(q)

    if dry_run:
        wb.close()
        if summary is None:
            print(f'  [dry-run] A列修复 {fixed} 行；季度数据已是最新（{quarter_label(target)}）')
        else:
            print(f'  [dry-run] A列修复 {fixed} 行；拟补 '
                  f'{"、".join(quarter_label(x) for x in added)}: '
                  f'猪料={summary["q"]["猪料"]:,.2f}, 环比={summary["mom"]["猪料"]}%, '
                  f'同比={summary["yoy"]["猪料"]}%')
        return out_path, summary

    if appended == 0 and fixed == 0:
        wb.close()
        print(f'  [跳过] 季度数据已是最新（{quarter_label(target)}）')
        return out_path, summary

    wb.save(out_path)
    wb.close()
    same = swap_chart_parts(out_path, bak_path)
    if summary is None:
        print(f'  [完成] {os.path.basename(out_path)}  A列修复 {fixed} 行  '
              f'（季度数据本已是最新）  图表原样={same}')
    else:
        print(f'  [完成] {os.path.basename(out_path)}  '
              f'A列修复 {fixed} 行 / 补 {"、".join(quarter_label(x) for x in added)}  '
              f'猪料={summary["q"]["猪料"]:,.2f}  '
              f'环比={summary["mom"]["猪料"]:+.2f}%  同比={summary["yoy"]["猪料"]:+.2f}%  '
              f'图表原样={same}')
    return out_path, summary


if __name__ == '__main__':
    dry = '--dry-run' in sys.argv
    results = {}
    for fname, _n in FILES:
        print(f'== {fname}')
        results[fname] = process(fname, dry_run=dry)
    print('全部完成' if not dry else '（dry-run，未写文件）')
