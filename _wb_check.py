# -*- coding: utf-8 -*-
"""
独立复算：不用 _wb_pipeline 的聚合函数，直接用 pandas 重算所有展示字段，
与 _wb_data.json 逐项对比，列出不一致项。
"""
import datetime as dt
import glob, json, os, sys
import numpy as np
import pandas as pd

SRCS = sorted(glob.glob(r'D:\CC\Desktop\车辆过磅单\车辆过磅单-*.xlsx'))
J = json.load(open(r'C:\Users\CC\test-claude\_wb_data.json', encoding='utf-8'))

RETAIL_MAX_HEAD = 60
BIG_AVG, BIG_MIN_HEAD = 150.0, 70
BINS = [0, 90, 120, 150, 175, 200, 10**9]
BIN_LABELS = ['<90kg', '90-120kg', '120-150kg', '150-175kg', '175-200kg', '>200kg']

print('源: ' + ' + '.join(os.path.basename(s) for s in SRCS))
# 与管道一致: 全量快照 + 每日增量一起读, 做并集后再去重
df = pd.concat([pd.read_excel(s, dtype=str) for s in SRCS], ignore_index=True)
df.columns = [c.strip() for c in df.columns]
for c in ['毛重(KG)', '皮重(KG)', '净重(KG)', '猪只数量', '均重(KG)']:
    df[c] = pd.to_numeric(df[c].astype(str).str.replace(',', '', regex=False).str.strip(),
                         errors='coerce')
df['_mod'] = pd.to_datetime(df['最后修改时间'], errors='coerce')
n_raw = len(df)
df = df.sort_values('_mod', na_position='first').drop_duplicates('单据编号', keep='last')
st = df['单据状态'].astype(str).str.strip()
df = df[st.isin(['已保存']) | st.str.contains('已保存|正常|生效', na=False)]
blob = (df['车牌号'].astype(str).fillna('') + '|' + df['车辆规格'].astype(str).fillna('')
        + '|' + df['客户'].astype(str).fillna(''))
df = df[~blob.str.contains('车间|猪头|连肝肉|分割|副产|消泡剂', na=False)]
df = df[df['净重(KG)'] > 0]
df = df[~(df['毛重(KG)'].notna() & df['皮重(KG)'].notna()
          & (df['毛重(KG)'] <= df['皮重(KG)']))]
df = df[df['猪只数量'] > 0]
df = df[~((df['猪只数量'] == 1) & (df['净重(KG)'] > 1000))]
df['单车均重'] = df['净重(KG)'] / df['猪只数量']
df = df[df['单车均重'] >= 30]   # 无体重上界
df['业务日期'] = pd.to_datetime(df['业务日期'], errors='coerce')
df = df[df['业务日期'].notna()]
df['头数'] = df['猪只数量']
df['净重'] = df['净重(KG)']
df['日'] = df['业务日期'].dt.strftime('%Y-%m-%d')
iso = df['业务日期'].dt.isocalendar()
df['周'] = iso['year'].astype(str) + '-W' + iso['week'].astype(str).str.zfill(2)
df['月'] = df['业务日期'].dt.strftime('%Y-%m')
df['散户'] = df['头数'] < RETAIL_MAX_HEAD
df['大猪'] = df['单车均重'] >= BIG_AVG
df['外调'] = df['大猪'] & (df['头数'] > BIG_MIN_HEAD)
df['体重段'] = pd.cut(df['单车均重'], bins=BINS, labels=BIN_LABELS, right=False).astype(str)

bad = []
def _nz(v):
    """NaN 与 None 等价: pandas 建 DataFrame 时会把 None 变成 NaN, 看板端 scrub 会还原成 null"""
    if v is None: return None
    if isinstance(v, float) and np.isnan(v): return None
    return v

NCHK = [0]


def chk(tag, got, exp, tol=0.011):
    NCHK[0] += 1
    got, exp = _nz(got), _nz(exp)
    if exp is None and got == 0.0 and '占比' in tag:   # 看板约定: 占比无样本记 0
        return
    if exp is None or (isinstance(exp, float) and np.isnan(exp)):
        ok = got is None
        if not ok: bad.append((tag, got, exp))
        return
    if got is None:
        bad.append((tag, got, exp)); return
    if isinstance(exp, (int, float, np.number)):
        if abs(float(got) - float(exp)) > tol:
            bad.append((tag, got, exp))
    elif got != exp:
        bad.append((tag, got, exp))


# ---------- 总量 ----------
chk('meta.cars', J['meta']['cars'], len(df))
chk('meta.heads', J['meta']['heads'], int(df['头数'].sum()))
chk('meta.net', J['meta']['net'], round(float(df['净重'].sum()), 2))

# ---------- 各粒度 ----------
from lunardate import LunarDate as _LD
_ld = [_LD.fromSolarDate(d.year, d.month, d.day) for d in df['业务日期']]
df['农历年'] = [x.year for x in _ld]
df['农历月号'] = [x.month for x in _ld]
df['农历日'] = [x.day for x in _ld]
df['是否闰月'] = [bool(x.isLeapMonth) for x in _ld]
# 月键带闰月标记: 2025 年的【六月】与【闰六月】必须分成两桶, 合成一个桶的话
# 那期自然日数翻倍, 趋势图上会是根假高柱
df['农历月'] = [f'{x.year}-{x.month:02d}{"闰" if x.isLeapMonth else ""}' for x in _ld]
# 农历周 = 农历月内每 7 天一段(初一~初七 = 第 1 周)
_cn = ['正', '二', '三', '四', '五', '六', '七', '八', '九', '十', '冬', '腊']
df['农历周序'] = ((df['农历日'] - 1) // 7 + 1).astype(int)
df['农历周键'] = [f'{y}-{m:02d}{"闰" if lp else ""}-W{n}'
               for y, m, n, lp in zip(df['农历年'], df['农历月号'],
                                      df['农历周序'], df['是否闰月'])]
KEYS = {'日': '日', '周': '周', '月': '月', '农历周': '农历周键', '农历月': '农历月'}
for gname, key in KEYS.items():
    P = {r['period']: r for r in J['periods'][gname]}
    got_keys = set(df[key].unique()); exp_keys = set(P.keys())
    if got_keys != exp_keys:
        bad.append((f'{gname} 期集合', sorted(got_keys ^ exp_keys), '对称差'))
    for k, g in df.groupby(key):
        r = P.get(str(k))
        if r is None:
            bad.append((f'{gname} {k} 缺失', None, None)); continue
        # 看板按「截止日」筛期就是靠 d0/d1 判断哪几期落在窗口内, 错了会漏期或多期
        chk(f'{gname} {k} d0', r['d0'], g['日'].min())
        chk(f'{gname} {k} d1', r['d1'], g['日'].max())
        h = g['头数'].sum(); n = g['净重'].sum()
        ret = g[g['散户']]; big = g[g['大猪']]; ext = g[g['外调']]
        chk(f'{gname} {k} 车次数', r['车次数'], len(g))
        chk(f'{gname} {k} 总头数', r['总头数'], int(h))
        chk(f'{gname} {k} 总净重', r['总净重'], round(float(n), 2))
        chk(f'{gname} {k} 到场总均重', r['到场总均重'], round(float(n / h), 2))
        chk(f'{gname} {k} 散户车次数', r['散户车次数'], len(ret))
        chk(f'{gname} {k} 散户头数', r['散户头数'], int(ret['头数'].sum()))
        chk(f'{gname} {k} 散户头数占比', r['散户头数占比'],
            round(ret['头数'].sum() / h * 100, 2))
        chk(f'{gname} {k} 散户均重', r['散户均重'],
            round(ret['净重'].sum() / ret['头数'].sum(), 2) if len(ret) else None)
        chk(f'{gname} {k} 规模场头数', r['规模场头数'], int(h - ret['头数'].sum()))
        chk(f'{gname} {k} 规模场均重', r['规模场均重'],
            round((n - ret['净重'].sum()) / (h - ret['头数'].sum()), 2))
        chk(f'{gname} {k} 大猪150头数', r['大猪150头数'], int(big['头数'].sum()))
        chk(f'{gname} {k} 大猪150占比', r['大猪150占比'],
            round(big['头数'].sum() / h * 100, 2))
        chk(f'{gname} {k} 外调车次数', r['外调车次数'], len(ext))
        chk(f'{gname} {k} 外调头数', r['外调头数'], int(ext['头数'].sum()))
        chk(f'{gname} {k} 外调头数占比', r['外调头数占比'],
            round(ext['头数'].sum() / h * 100, 2) if len(ext) else None)
        chk(f'{gname} {k} 单车均重均值', r['单车均重均值'],
            round(float(g['单车均重'].mean()), 2))
        for lb in BIN_LABELS:
            s = g[g['体重段'] == lb]
            chk(f'{gname} {k} bin {lb}', r['bin_' + lb], int(s['头数'].sum()))
            chk(f'{gname} {k} binw {lb}', r['binw_' + lb], round(float(s['净重'].sum()), 2))

# ---------- 环比 ----------
for gname in KEYS:
    arr = J['periods'][gname]
    for i, r in enumerate(arr):
        for c in ['总头数', '总净重', '到场总均重', '散户头数占比', '大猪150占比', '外调车次数']:
            exp = None if i == 0 else round(arr[i][c] - arr[i-1][c], 2)
            chk(f'{gname}[{i}] d_{c}', r['d_' + c], exp, tol=0.02)
        prev_h = arr[i-1]['总头数'] if i else 0
        exp = None if i == 0 or prev_h == 0 else round(
            (arr[i]['总头数'] / prev_h - 1) * 100, 2)
        chk(f'{gname}[{i}] pct_总头数', r['pct_总头数'], exp, tol=0.02)

# ---------- 头数档 × 均重 (口径诊断) ----------
HB = J['head_band']
hb = pd.cut(df['头数'], [0, 20, 40, 60, 80, 100, 150, 10**9],
            labels=['<20', '20-40', '40-60', '60-80', '80-100', '100-150', '>150'])
for i, (k, g) in enumerate(df.groupby(hb, observed=True)):
    r = HB[i] if i < len(HB) else None
    if r is None or r['档'] != str(k):
        bad.append((f'head_band[{i}] 档序', r and r['档'], str(k))); continue
    chk(f'head_band {k} 车次', r['车次'], len(g))
    chk(f'head_band {k} 头数合计', r['头数合计'], int(g['头数'].sum()))
    chk(f'head_band {k} 均重', r['均重'], round(float(g['单车均重'].mean()), 2))
    chk(f'head_band {k} 车次占比', r['车次占比'], round(len(g) / len(df) * 100, 2))
    chk(f'head_band {k} 头数占比', r['头数占比'],
        round(g['头数'].sum() / df['头数'].sum() * 100, 2))
    # 期望: 该档内 150kg+ 的【头数】占该档头数
    chk(f'head_band {k} 大猪占该档%', r['大猪占该档%'],
        round(float(g[g['大猪']]['头数'].sum() / g['头数'].sum() * 100), 2))

# ---------- diag ----------
D = J['diag']
chk('diag.pearson', D['pearson'], round(float(df['头数'].corr(df['单车均重'])), 4))
chk('diag.spearman', D['spearman'],
    round(float(df['头数'].corr(df['单车均重'], method='spearman')), 4))
chk('diag.retail_big_share', D['retail_big_share'],
    round(float(df[df['散户'] & df['大猪']]['头数'].sum() / df[df['散户']]['头数'].sum() * 100), 2))
chk('diag.scale_big_share', D['scale_big_share'],
    round(float(df[~df['散户'] & df['大猪']]['头数'].sum() / df[~df['散户']]['头数'].sum() * 100), 2))

# ---------- bin_group ----------
BG = {(r['群体'], r['体重段']): r for r in J['bin_group']}
df['群体'] = np.where(df['散户'], '散户(<60头)', '规模场(>=60头)')
tot = 0
for (gp, lb), g in df.groupby(['群体', '体重段'], observed=True):
    r = BG.get((gp, str(lb)))
    if r is None:
        bad.append((f'bin_group 缺 {gp}/{lb}', None, None)); continue
    chk(f'bin_group {gp} {lb} 车次', r['车次数'], len(g))
    chk(f'bin_group {gp} {lb} 头数', r['头数'], int(g['头数'].sum()))
    chk(f'bin_group {gp} {lb} 净重', r['净重'], round(float(g['净重'].sum()), 2))
    chk(f'bin_group {gp} {lb} 均重', r['均重'], round(float(g['单车均重'].mean()), 2))
    tot += int(g['头数'].sum())
chk('bin_group 头数合计=总头数', tot, int(df['头数'].sum()))

# ---------- 季节同比 ----------
Y = J['yoy']
YSLOTS = {'solar': 365, 'lunar': 360}


def _yoy_expect(g):
    h = float(g['头数'].sum()); n = float(g['净重'].sum())
    if not h:
        return {}
    return {'总头数': int(h), '到场总均重': round(n / h, 2),
            '散户头数占比': round(float(g[g['散户']]['头数'].sum()) / h * 100, 2),
            '大猪150占比': round(float(g[g['大猪']]['头数'].sum()) / h * 100, 2)}


for mode in ('solar', 'lunar'):
    M = Y['modes'][mode]
    ns = YSLOTS[mode]
    chk(f'yoy.{mode} 槽位数', len(M['labels']), ns)
    chk(f'yoy.{mode} tips 长度', len(M['tips']), ns)
    if mode == 'solar':
        # 槽位 i -> 平年的第 i 天
        slot_of = {str(dt.date(2025, 1, 1) + dt.timedelta(days=i))[5:]: i for i in range(ns)}
        yr_rows = {y: df[df['业务日期'].dt.year == y] for y in (2025, 2026)}
        keyf = lambda g: g['业务日期'].dt.strftime('%m-%d')
    else:
        slot_of = {f'{i // 30 + 1:02d}-{i % 30 + 1:02d}': i for i in range(ns)}
        # 闰月剔除, 两年才严格对齐
        dpl = df[~df['是否闰月']]
        yr_rows = {y: dpl[dpl['农历年'] == y] for y in (2025, 2026)}
        keyf = lambda g: (g['农历月号'].map('{:02d}'.format) + '-'
                          + g['农历日'].map('{:02d}'.format))
    for y in (2025, 2026):
        sub = yr_rows[y]
        if sub.empty:
            continue
        for kk, g in sub.groupby(keyf(sub)):
            exp = _yoy_expect(g)
            j = slot_of[kk]
            for mk in ['总头数', '到场总均重', '散户头数占比', '大猪150占比']:
                chk(f'yoy.{mode} {y} 槽{j} {mk}', M['data'][str(y)][mk][j], exp.get(mk))
    # 月度对比表
    for y in (2025, 2026):
        sub = yr_rows[y]
        if sub.empty:
            continue
        grp = sub.groupby(sub['业务日期'].dt.month) if mode == 'solar' else sub.groupby(sub['农历月号'])
        agg = {k2: _yoy_expect(g) for k2, g in grp}
        for mk in ['总头数', '到场总均重', '散户头数占比', '大猪150占比']:
            tbl = M['table'][mk][0 if y == 2025 else 1]
            for mo in range(1, 13):
                exp = (agg.get(mo) or {}).get(mk)
                got = tbl.get(str(mo), tbl.get(mo))
                if exp is None and got is None:
                    continue
                chk(f'yoy.{mode} 表 {y}-{mo} {mk}', got, exp)
    # 颜色/线型约定: 2025 绿虚线, 2026 红实线
    ser = {s['key']: s for s in M['series']}
    chk(f'yoy.{mode} 2025色', ser['2025']['color'], '--y25')
    chk(f'yoy.{mode} 2026色', ser['2026']['color'], '--y26')
    if bool(ser['2025'].get('dash')) == bool(ser['2026'].get('dash')):
        bad.append((f'yoy.{mode} 线型未区分', ser['2025'].get('dash'), ser['2026'].get('dash')))

print(f'\n对比项完成，共 {NCHK[0]:,} 项，不一致 {len(bad)} 项')
import re as _re
from collections import Counter
def cls(t):
    return _re.sub(r'^(日|周|月) \S+ ', r'\1 ', t)
cnt = Counter(cls(t) for t, _, _ in bad)
for k, v in cnt.most_common():
    print(f'  {v:4d}  {k}')
print('\n--- 明细 ---')
for t, g, e in bad:
    print(f'  X {t}\n      实际(JSON)={g}   独立复算={e}')   # ✗ 编不进 GBK：输出重定向到文件时会崩
