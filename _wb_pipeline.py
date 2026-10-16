# -*- coding: utf-8 -*-
"""
车辆过磅单 -> 社会供应压力看板  数据管道
源: D:\\CC\\Desktop\\车辆过磅单\\车辆过磅单-*.xlsx
出: _wb_data.json  (看板用)  +  _wb_有效车次表.xlsx (存档)
"""
import datetime as dt
import fnmatch
import glob, json, os, sys
import numpy as np
import pandas as pd
from lunardate import LunarDate

SRC_DIR = r'D:\CC\Desktop\车辆过磅单'
OUT_JSON = r'C:\Users\CC\test-claude\_wb_data.json'
OUT_XLSX = r'C:\Users\CC\test-claude\_wb_有效车次表.xlsx'

# ---------- 参数 ----------
RETAIL_MAX_HEAD = 60          # 散户: 头数 < 60
BIG_AVG_KG = 150.0            # 大猪: 单车均重 >= 150
BIG_MIN_HEAD = 70             # 外调大猪: 头数 > 70
AVG_LO = 30.0                 # 单车均重下限: 低于此值 = 非生猪/空车/皮重异常
QC_HI = 300.0                 # 存疑阈值: 超过只标记不剔除 (疑头数少记/皮重缺失)
BINS = [0, 90, 120, 150, 175, 200, 10**9]
BIN_LABELS = ['<90kg', '90-120kg', '120-150kg', '150-175kg', '175-200kg', '>200kg']
BAD_KW = ['车间', '猪头', '连肝肉', '分割', '副产', '消泡剂']

AUDIT = {}


def num(s):
    return pd.to_numeric(
        s.astype(str).str.replace(',', '', regex=False).str.strip(),
        errors='coerce')


def load_raw():
    """目录里会同时躺着【全量快照】和【只含某几天的增量导出】:
    例如 车辆过磅单-2026-09-20 是全量(数据到 09-19), 车辆过磅单-2026-09-21 只有 09-20 那一天。
    —— 只按文件名取"最新"会挑中那份几十行的增量, 把看板打成空壳。
    所以全部读进来做并集, 重复的单据交给 clean() 按【最后修改时间】去重。
    并集对"增量导出"和"又发了一次全量"两种情况都成立。"""
    pat = os.path.join(SRC_DIR, '车辆过磅单-*.xlsx')
    files = sorted(glob.glob(pat))
    # 只认 车辆过磅单-*.xlsx。名字不对的文件会被静默忽略(导出的文件常被浏览器
    # 改名成 车辆过磅单(1).xlsx 之类), 那样数据不进看板而用户毫无察觉 —— 必须报出来。
    others = [os.path.basename(p) for p in glob.glob(os.path.join(SRC_DIR, '*.xlsx'))
              if not fnmatch.fnmatch(p, pat) and not os.path.basename(p).startswith('~$')]
    if others:
        print('!! 以下 .xlsx 名字不匹配「车辆过磅单-*.xlsx」, 不会被读进来:')
        for o in others:
            print(f'     {o}')
    if not files:
        sys.exit('未找到源文件')

    frames, used, skipped = [], [], []
    for p in files:
        name = os.path.basename(p)
        try:
            d = pd.read_excel(p, dtype=str)
        except Exception as e:
            skipped.append(f'{name} ({e})')
            continue
        d.columns = [c.strip() for c in d.columns]
        frames.append(d)
        used.append((name, len(d)))
    if not frames:
        sys.exit('所有源文件都读不出来: ' + '; '.join(skipped))
    df = pd.concat(frames, ignore_index=True)
    print(f'源文件 {len(frames)} 个 -> 合并 {len(df):,} 行')
    for name, n in used:
        print(f'    {name}  {n:,} 行')
    for t in skipped:
        print(f'    跳过(读不出) {t}')
    AUDIT['原始行数'] = len(df)
    AUDIT['源文件数'] = len(frames)
    AUDIT['源文件'] = ' + '.join(name for name, _ in used)
    if len(used) > 1:
        AUDIT['源文件明细'] = '; '.join(f'{n} {c:,}行' for n, c in used)
    return df


def clean(df):
    n0 = len(df)
    # 数值化
    for c in ['毛重(KG)', '皮重(KG)', '净重(KG)', '猪只数量', '均重(KG)']:
        df[c] = num(df[c])

    # 0 未清洗全样本 (对账用: 与源文件导出「合计」行可比)
    _h, _n = df['猪只数量'].sum(), df['净重(KG)'].sum()
    AUDIT['全样本-车次'] = int(df['单据编号'].notna().sum())
    AUDIT['全样本-头数'] = int(_h)
    AUDIT['全样本-净重'] = round(float(_n), 2)
    AUDIT['全样本-到场总均重'] = round(float(_n / _h), 2) if _h else None

    # 1 去重: 单据编号, 保留最后修改时间最新
    if '最后修改时间' in df.columns:
        df['_mod'] = pd.to_datetime(df['最后修改时间'], errors='coerce')
    else:
        df['_mod'] = pd.NaT
    before = len(df)
    df = (df.sort_values('_mod', na_position='first')
            .drop_duplicates('单据编号', keep='last'))
    AUDIT['去重-单据编号'] = before - len(df)

    # 2 单据状态: 只留有效
    st = df['单据状态'].astype(str).str.strip()
    keep = st.isin(['已保存']) | st.str.contains('已保存|正常|生效', na=False)
    AUDIT['剔除-作废/删除状态'] = int((~keep).sum())
    df = df[keep]

    # 3 非生猪记录 (车牌/车辆规格/客户 含关键字)
    blob = (df['车牌号'].astype(str).fillna('') + '|' +
            df['车辆规格'].astype(str).fillna('') + '|' +
            df['客户'].astype(str).fillna(''))
    bad = blob.str.contains('|'.join(BAD_KW), na=False)
    AUDIT['剔除-非生猪(' + '/'.join(BAD_KW) + ')'] = int(bad.sum())
    df = df[~bad]

    # 4 净重异常
    m = df['净重(KG)'] <= 0
    AUDIT['剔除-净重<=0'] = int(m.sum())
    df = df[~m]
    m = (df['毛重(KG)'].notna() & df['皮重(KG)'].notna()
         & (df['毛重(KG)'] <= df['皮重(KG)']))
    AUDIT['剔除-毛重<=皮重'] = int(m.sum())
    df = df[~m]

    # 5 头数异常
    m = df['猪只数量'] <= 0
    AUDIT['剔除-头数<=0'] = int(m.sum())
    df = df[~m]
    m = (df['猪只数量'] == 1) & (df['净重(KG)'] > 1000)
    AUDIT['剔除-头数=1且净重>1000'] = int(m.sum())
    df = df[~m]

    # 6 单车均重下限: 低于 30kg 的不是商品猪 (空车/副产/皮重异常/头数填错)
    #   不设上界 —— 上界会误剔真实的大猪车次, 导致看板与源文件「合计」行对不上
    df['单车均重'] = df['净重(KG)'] / df['猪只数量']
    m = df['单车均重'] < AVG_LO
    AUDIT[f'剔除-单车均重<{AVG_LO:.0f}kg(非生猪/异常)'] = int(m.sum())
    df = df[~m]

    # 业务日期
    df['业务日期'] = pd.to_datetime(df['业务日期'], errors='coerce')
    m = df['业务日期'].isna()
    AUDIT['剔除-业务日期空'] = int(m.sum())
    df = df[~m]

    AUDIT['清洗后有效车次'] = len(df)
    AUDIT['清洗保留率'] = f'{len(df)/n0*100:.2f}%'
    _h2, _n2 = df['猪只数量'].sum(), df['净重(KG)'].sum()
    AUDIT['清洗后-头数'] = int(_h2)
    AUDIT['清洗后-净重'] = round(float(_n2), 2)
    AUDIT['清洗后-到场总均重'] = round(float(_n2 / _h2), 2) if _h2 else None
    return df


def enrich(df):
    df = df.copy()
    df['头数'] = df['猪只数量'].astype(int)
    df['净重'] = df['净重(KG)'].astype(float)
    df['是否散户'] = df['头数'] < RETAIL_MAX_HEAD
    df['是否大猪'] = df['单车均重'] >= BIG_AVG_KG
    df['是否外调'] = df['是否大猪'] & (df['头数'] > BIG_MIN_HEAD)
    df['体重段'] = pd.cut(df['单车均重'], bins=BINS, labels=BIN_LABELS,
                          right=False).astype(str)

    # 时间维度
    iso = df['业务日期'].dt.isocalendar()
    df['日'] = df['业务日期'].dt.strftime('%Y-%m-%d')
    df['周'] = iso['year'].astype(str) + '-W' + iso['week'].astype(str).str.zfill(2)
    df['月'] = df['业务日期'].dt.strftime('%Y-%m')

    ld = [LunarDate.fromSolarDate(d.year, d.month, d.day) for d in df['业务日期']]
    df['农历年'] = [x.year for x in ld]
    df['农历月号'] = [x.month for x in ld]
    df['农历日'] = [x.day for x in ld]
    df['是否闰月'] = [bool(x.isLeapMonth) for x in ld]
    cn = ['正', '二', '三', '四', '五', '六', '七', '八', '九', '十', '冬', '腊']
    # 月键带闰月标记 —— 否则 2025 年【六月】和【闰六月】会并成一个桶,
    # 那个桶的自然日数是别人的两倍, 趋势图上凭空冒出一根高柱。
    df['农历月'] = [f'{y}-{m:02d}{"闰" if lp else ""}'
                  for y, m, lp in zip(df['农历年'], df['农历月号'], df['是否闰月'])]
    df['农历月名'] = [
        f'{y}年{"闰" if lp else ""}{cn[m-1] if 1 <= m <= 12 else m}月'
        for y, m, lp in zip(df['农历年'], df['农历月号'], df['是否闰月'])]
    df['农历日序'] = [x.day for x in ld]

    # 农历日: 一个自然日一桶, 与「日」同粒度, 只是键与标签走农历
    # (键含闰月标记: 2025-06-05 与 2025-06闰-05 是两个不同的日子)
    df['农历日名'] = [f'{"闰" if lp else ""}{cn[m-1] if 1 <= m <= 12 else m}月' + _lunar_day_name(d)
                    for m, d, lp in zip(df['农历月号'], df['农历日'], df['是否闰月'])]
    df['农历日键'] = [f'{y}-{m:02d}{"闰" if lp else ""}-{d:02d}'
                   for y, m, d, lp in zip(df['农历年'], df['农历月号'],
                                          df['农历日'], df['是否闰月'])]

    # 农历周: 农历月内每 7 天切一段 (初一~初七 = 第 1 周), 跨月归零
    wk = ((df['农历日'] - 1) // 7 + 1).astype(int)
    df['农历周序'] = wk
    df['农历周键'] = [f'{y}-{m:02d}{"闰" if lp else ""}-W{n}'
                   for y, m, n, lp in zip(df['农历年'], df['农历月号'], wk,
                                          df['是否闰月'])]
    df['农历周名'] = [f'{mn}第{n}周' for mn, n in zip(df['农历月名'], wk)]
    return df


def agg_period(df, key, first_day=None, days=None, label_col=None):
    """按 key 聚合。days 用于外调车次的日均阈值折算。

    label_col 给每期加一个人读标签(农历日名/农历周名/农历月名);
    d0/d1 是该期的首末业务日期 —— 看板按「截止日期」筛选时要靠它判断哪几期落在窗口内。
    """
    rows = []
    for k, g in df.groupby(key, sort=True):
        head = g['头数'].sum()
        net = g['净重'].sum()
        ret = g[g['是否散户']]
        big = g[g['是否大猪']]
        ext = g[g['是否外调']]
        n_ret, n_big, n_ext = len(ret), len(big), len(ext)
        d = {
            'period': str(k),
            'd0': g['日'].min(),
            'd1': g['日'].max(),
            '自然日数': int(g['日'].nunique()),
            '车次数': int(len(g)),
            '总头数': int(head),
            '总净重': round(float(net), 2),
            '到场总均重': round(float(net / head), 2) if head else None,
            '散户车次数': int(n_ret),
            '散户车次占比': round(n_ret / len(g) * 100, 2) if len(g) else None,
            '散户头数': int(ret['头数'].sum()),
            '散户头数占比': round(ret['头数'].sum() / head * 100, 2) if head else None,
            '散户净重': round(float(ret['净重'].sum()), 2),
            '散户均重': round(float(ret['净重'].sum() / ret['头数'].sum()), 2) if n_ret else None,
            '规模场车次数': int(len(g) - n_ret),
            '规模场头数': int(head - ret['头数'].sum()),
            '规模场均重': (round(float((net - ret['净重'].sum()) / (head - ret['头数'].sum())), 2)
                       if head - ret['头数'].sum() else None),
            '大猪150头数': int(big['头数'].sum()),
            '大猪150占比': round(big['头数'].sum() / head * 100, 2) if head else None,
            '外调车次数': int(n_ext),
            '外调头数': int(ext['头数'].sum()) if n_ext else 0,
            '外调头数占比': round(ext['头数'].sum() / head * 100, 2) if head else None,
            '外调车次占比': round(n_ext / len(g) * 100, 2) if len(g) else None,
            '外调均重': round(float(ext['净重'].sum() / ext['头数'].sum()), 2) if n_ext else None,
            '单车均重均值': round(float(g['单车均重'].mean()), 2),
        }
        for lb in BIN_LABELS:
            s = g[g['体重段'] == lb]
            d['bin_' + lb] = int(s['头数'].sum())
            d['binw_' + lb] = round(float(s['净重'].sum()), 2)
        if label_col is not None:
            d['label'] = str(g.sort_values('业务日期')[label_col].iloc[0])
        rows.append(d)
    out = pd.DataFrame(rows)
    if first_day is not None:
        out['days'] = days
    return out


def add_mom(df):
    """环比: 相对上一期。"""
    cols = ['到场总均重', '散户头数占比', '大猪150占比', '外调车次数',
            '总头数', '总净重', '散户车次数', '总车次数']
    for c in cols:
        if c in df.columns:
            df['d_' + c] = df[c].diff()
            df['pct_' + c] = (df[c].pct_change() * 100).round(2)
    return df


# ============================================================
#  季节同比 —— 两年叠在同一张图上比同期
#  公历对齐: 横轴 01-01..12-31, 两条线按同一日历日对齐
#  农历对齐: 横轴 正月初一..腊月三十, 按【农历月+日】对齐
#            (不能按"农历年内第 N 天"——乙巳蛇年有闰六月共 384 天,
#             丙午马年没有共 354 天, 按天序会被闰月整体错开 29 天)
# ============================================================
LUNAR_MONTH_CN = ['正月', '二月', '三月', '四月', '五月', '六月',
                  '七月', '八月', '九月', '十月', '冬月', '腊月']
CN_NUM = '一二三四五六七八九十'
GAN = '甲乙丙丁戊己庚辛壬癸'
ZHI = '子丑寅卯辰巳午未申酉戌亥'

# 同比图比哪两年。顺序 = 月度对比表的行序(第 0 行 2025, 第 1 行 2026), 看板端按它取数
YOY_YEARS = (2025, 2026)

# (df 列名, 展示标题, 说明, 是否百分比)
YOY_METRICS = [
    ('总头数',        '到场总头数',        '到场生猪头数',                                 False),
    ('到场总均重',    '到场总均重',        'Σ净重 ÷ Σ头数，按头数加权（kg）',               False),
    ('散户头数占比',  '散户头数占比',      '头数 < 60 的车次头数 ÷ 总头数（%）',            True),
    ('大猪150占比',   '150kg 以上头数占比', '单车均重 ≥ 150kg 的车次头数 ÷ 总头数（%）',    True),
]


def _lunar_day_name(d):
    if d <= 10:
        return '初' + CN_NUM[d - 1]
    if d < 20:
        return '十' + CN_NUM[d - 11]
    if d == 20:
        return '二十'
    if d < 30:
        return '廿' + CN_NUM[d - 21]
    return '三十'


def _yoy_agg(g):
    """把一个时间桶聚成同比图需要的 4 个指标。"""
    h = float(g['头数'].sum())
    n = float(g['净重'].sum())
    if not h:
        return {k: None for k, *_ in YOY_METRICS}
    return {
        '总头数': int(h),
        '到场总均重': round(n / h, 2),
        '散户头数占比': round(float(g[g['是否散户']]['头数'].sum()) / h * 100, 2),
        '大猪150占比': round(float(g[g['是否大猪']]['头数'].sum()) / h * 100, 2),
    }


def build_yoy(df):
    """生成公历/农历两套对齐的同比较据。"""
    modes = {}

    # ---------- 公历对齐 ----------
    base = dt.date(2025, 1, 1)
    s_slots = [base + dt.timedelta(days=i) for i in range(365)]   # 2025/2026 均为平年
    s_labels, s_ticks, s_tips = [], [], []
    for d in s_slots:
        s_tips.append(d.strftime('%m-%d'))
        s_labels.append(f'{d.month}月' if d.day == 1 else '')
        if d.day == 1:
            s_ticks.append(len(s_labels) - 1)
    s_data, s_dates = {}, {}
    for y in YOY_YEARS:
        sub = df[df['业务日期'].dt.year == y]
        by = {str(k): v for k, v in sub.groupby(sub['业务日期'].dt.strftime('%m-%d'))}
        vals = {k: [] for k, *_ in YOY_METRICS}
        # 槽位对应的真实日期 —— 看板按「截止日期」截断曲线时靠它逐点判断,
        # 比让前端反推"第 i 个槽是几月几号"稳(公历/农历两套槽位规则不同)
        s_dates[str(y)] = [dt.date(y, d.month, d.day).isoformat() for d in s_slots]
        for d in s_slots:
            row = _yoy_agg(by[d.strftime('%m-%d')]) if d.strftime('%m-%d') in by else {}
            for k, *_ in YOY_METRICS:
                vals[k].append(row.get(k))
        s_data[str(y)] = vals
    # 月度对比表: 公历月
    s_tbl = {k: [] for k, *_ in YOY_METRICS}
    for y in YOY_YEARS:
        sub = df[df['业务日期'].dt.year == y]
        agg = {kk: _yoy_agg(g) for kk, g in sub.groupby(sub['业务日期'].dt.month)}
        for k, *_ in YOY_METRICS:
            s_tbl[k].append({m: (agg.get(m) or {}).get(k) for m in range(1, 13)})
    modes['solar'] = {
        'labels': s_labels, 'tickAt': s_ticks, 'tips': s_tips, 'data': s_data,
        'dates': s_dates,
        'table': s_tbl, 'tableLab': [f'{m}月' for m in range(1, 13)],
        'series': [{'key': '2025', 'name': '2025年', 'color': '--y25', 'dash': '6 4'},
                   {'key': '2026', 'name': '2026年', 'color': '--y26', 'dash': None}],
        'hint': '横轴 01-01 ~ 12-31，两年按同一日历日对齐；2026 只到 %s，之后为空。'
                % df['日'].max()[5:],
    }

    # ---------- 农历对齐 ----------
    # 槽位 = 农历月号 1..12 × 日 1..30，共 360 格；闰月剔除（两年才严格可比）
    l_labels, l_ticks, l_tips = [], [], []
    for i in range(360):
        mo, dy = i // 30 + 1, i % 30 + 1
        l_tips.append(LUNAR_MONTH_CN[mo - 1] + _lunar_day_name(dy))
        l_labels.append(LUNAR_MONTH_CN[mo - 1] if dy == 1 else '')
        if dy == 1:
            l_ticks.append(i)
    l_data, l_names, l_spans, l_dates = {}, {}, {}, {}
    dp = df[~df['是否闰月']]
    for y in YOY_YEARS:
        sub = dp[dp['农历年'] == y]
        if sub.empty:
            continue
        idx = (sub['农历月号'] - 1) * 30 + (sub['农历日'] - 1)
        vals = {k: [None] * 360 for k, *_ in YOY_METRICS}
        dates = [None] * 360
        for pos, g in sub.groupby(idx):
            row = _yoy_agg(g)
            for k, *_ in YOY_METRICS:
                vals[k][int(pos)] = row[k]
            dates[int(pos)] = g['日'].iloc[0]
        l_data[str(y)] = vals
        l_dates[str(y)] = dates
        yz = GAN[(y - 4) % 10] + ZHI[(y - 4) % 12]
        animal = '鼠牛虎兔龙蛇马羊猴鸡狗猪'[(y - 4) % 12]
        l_names[str(y)] = f'{y} {yz}{animal}年'
        l_spans[str(y)] = f"{sub['日'].min()} ~ {sub['日'].max()}"
    l_tbl = {k: [] for k, *_ in YOY_METRICS}
    for y in YOY_YEARS:
        sub = dp[dp['农历年'] == y]
        if sub.empty:
            for k, *_ in YOY_METRICS:
                l_tbl[k].append({m: None for m in range(1, 13)})
            continue
        agg = {kk: _yoy_agg(g) for kk, g in sub.groupby(sub['农历月号'])}
        for k, *_ in YOY_METRICS:
            l_tbl[k].append({m: (agg.get(m) or {}).get(k) for m in range(1, 13)})
    modes['lunar'] = {
        'labels': l_labels, 'tickAt': l_ticks, 'tips': l_tips, 'data': l_data,
        'dates': l_dates,
        'table': l_tbl, 'tableLab': LUNAR_MONTH_CN,
        'series': [{'key': '2025', 'name': l_names.get('2025', '2025'),
                    'color': '--y25', 'dash': '6 4'},
                   {'key': '2026', 'name': l_names.get('2026', '2026'),
                    'color': '--y26', 'dash': None}],
        'hint': '横轴 正月初一 ~ 腊月三十，按农历日期对齐；闰月已剔除；'
                + '　'.join(f'{l_names[k]} {v}' for k, v in l_spans.items()),
    }

    return {
        'metrics': [{'key': k, 'title': t, 'note': n, 'pct': p} for k, t, n, p in YOY_METRICS],
        'years': [str(y) for y in YOY_YEARS],   # table 里第 i 行固定对应哪一年
        'modes': modes,
    }


def main():
    raw = load_raw()
    df = clean(raw)
    df = enrich(df)

    # ---- 有效车次表存档 ----
    keep_cols = ['业务日期', '日', '周', '月', '农历月', '农历月名', '农历日',
                 '单据编号', '单据状态', '车牌号', '车辆规格', '毛重(KG)', '皮重(KG)',
                 '净重', '头数', '单车均重', '客户', '创建人',
                 '是否散户', '是否大猪', '是否外调', '体重段']
    df[keep_cols].sort_values(['业务日期', '单据编号']).to_excel(
        OUT_XLSX, index=False, sheet_name='有效过磅车次表')
    print(f'有效车次表已存 -> {OUT_XLSX}')

    # ---- 各粒度聚合 ----
    periods = {}
    # 「农历日」不单独出一份 —— 它和「日」是完全相同的桶(一个自然日一个),
    # 只有标签不同, 所以看板端用 lunar_of_day 换标签即可, 省掉一整份 650 期的重复数据
    for name, key, lab in [('日', '日', None), ('周', '周', None), ('月', '月', None),
                           ('农历周', '农历周键', '农历周名'),
                           ('农历月', '农历月', '农历月名')]:
        g = agg_period(df, key, label_col=lab)
        g = add_mom(g)
        periods[name] = g

    # 农历月加中文名
    nm = df.groupby('农历月')['农历月名'].first()
    periods['农历月']['农历月名'] = periods['农历月']['period'].map(nm)

    # 公历日期 -> 农历日名: 看板的日期选择器要同时显示两个历法,
    # 而「历法=农历 + 粒度=日」只是把这批标签换到 x 轴上, 桶本身与「日」相同, 不必再存一份
    lunar_of_day = (df.sort_values('业务日期').groupby('日')['农历日名'].first().to_dict())

    # ---- 存疑记录 (只标记不剔除, 供人工核单据) ----
    qm = (df['单车均重'] > QC_HI) | (df['皮重(KG)'] <= 0)
    AUDIT[f'存疑-单车均重>{QC_HI:.0f}kg或皮重为0'] = int(qm.sum())
    AUDIT['存疑说明'] = '仅标记未剔除, 疑头数少记或皮重缺失'
    qc = df[qm][['日', '单据编号', '车牌号', '净重', '头数', '单车均重', '皮重(KG)']].copy()
    qc.columns = ['业务日期', '单据编号', '车牌号', '净重', '头数', '单车均重', '皮重']
    qc['单车均重'] = qc['单车均重'].round(2)
    qc = qc.sort_values('单车均重', ascending=False)

    # 散户/规模场 分体重段
    df['群体'] = np.where(df['是否散户'], '散户(<60头)', '规模场(>=60头)')
    bin_grp = df.groupby(['群体', '体重段'], observed=True).agg(
        车次数=('净重', 'size'), 头数=('头数', 'sum'), 净重=('净重', 'sum'),
        均重=('单车均重', 'mean')).reset_index()
    bin_grp['均重'] = bin_grp['均重'].round(2)
    bin_grp['净重'] = bin_grp['净重'].round(2)

    # ---- 外调大猪明细 ----
    ext = df[df['是否外调']].copy()
    ext_out = ext[['日', '单据编号', '车牌号', '车辆规格', '头数', '单车均重',
                   '净重', '客户', '创建人']].copy()
    ext_out.columns = ['业务日期', '单据编号', '车牌号', '车辆规格', '头数',
                       '均重', '净重', '客户', '创建人']
    ext_out['均重'] = ext_out['均重'].round(2)
    ext_out = ext_out.sort_values('业务日期', ascending=False)

    # ---- 头数档 × 均重 (口径共线性诊断) ----
    hb = pd.cut(df['头数'], [0, 20, 40, 60, 80, 100, 150, 10 ** 9],
                labels=['<20', '20-40', '40-60', '60-80', '80-100', '100-150', '>150'])
    # 大猪头数必须是【头数】口径 —— 早先写成 '是否大猪'.sum() 得到的是车次数,
    # 再除以该档头数 → 单位串了, 严重低估共线性
    df['_大猪头数'] = np.where(df['是否大猪'], df['头数'], 0)
    ht = df.groupby(hb, observed=True).agg(
        车次=('净重', 'size'), 头数合计=('头数', 'sum'), 均重=('单车均重', 'mean'),
        大猪头数=('_大猪头数', 'sum')).reset_index()
    ht.columns = ['档', '车次', '头数合计', '均重', '大猪头数']
    ht['车次占比'] = (ht['车次'] / len(df) * 100).round(2)
    ht['头数占比'] = (ht['头数合计'] / df['头数'].sum() * 100).round(2)
    ht['大猪占该档%'] = (ht['大猪头数'] / ht['头数合计'] * 100).round(2)
    ht['均重'] = ht['均重'].round(2)

    diag = {
        'pearson': round(float(df['头数'].corr(df['单车均重'])), 4),
        'spearman': round(float(df['头数'].corr(df['单车均重'], method='spearman')), 4),
        'retail_big_share': round(float(
            df[df['是否散户'] & df['是否大猪']]['头数'].sum()
            / df[df['是否散户']]['头数'].sum() * 100), 2),
        'scale_big_share': round(float(
            df[~df['是否散户'] & df['是否大猪']]['头数'].sum()
            / df[~df['是否散户']]['头数'].sum() * 100), 2),
        'retail_avg_w': round(float(df[df['是否散户']]['单车均重'].median()), 1),
        'scale_avg_w': round(float(df[~df['是否散户']]['单车均重'].median()), 1),
    }
    # 备选散户口径 (头数<40) 的实测值 —— 供口径警示卡直接引用, 不再写死
    _r40 = df[df['头数'] < 40]
    diag['retail40_car_pct'] = round(len(_r40) / len(df) * 100, 1)
    diag['retail40_head_pct'] = round(_r40['头数'].sum() / df['头数'].sum() * 100, 1)
    diag['retail40_avg_w'] = round(_r40['净重'].sum() / _r40['头数'].sum(), 1)

    # ---- 季节同比 ----
    yoy = build_yoy(df)

    payload = {
        'yoy': yoy,
        'audit': AUDIT,
        'head_band': ht.to_dict('records'),
        'diag': diag,
        'periods': {k: v.to_dict('records') for k, v in periods.items()},
        'bin_group': bin_grp.to_dict('records'),
        'ext_detail': ext_out.head(400).to_dict('records'),
        'ext_total': int(len(ext_out)),
        'qc_rows': qc.to_dict('records'),
        'lunar_of_day': lunar_of_day,
        'bin_labels': BIN_LABELS,
        'meta': {
            'date_min': df['日'].min(), 'date_max': df['日'].max(),
            'rows': int(len(df)), 'cars': int(len(df)),
            'heads': int(df['头数'].sum()), 'net': round(float(df['净重'].sum()), 2),
            'n_lunar_months': int(df['农历月'].nunique()),
            'date_min_lunar': lunar_of_day.get(df['日'].min(), ''),
            'date_max_lunar': lunar_of_day.get(df['日'].max(), ''),
            'retail_max_head': RETAIL_MAX_HEAD, 'big_avg': BIG_AVG_KG,
            'big_min_head': BIG_MIN_HEAD, 'qc_hi': QC_HI, 'avg_lo': AVG_LO,
        },
    }
    with open(OUT_JSON, 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, default=str)

    print('\n=== 清洗审计 ===')
    for k, v in AUDIT.items():
        print(f'  {k:38s} {v}')
    print('\n=== 总览 ===')
    print(f"  区间 {df['日'].min()} ~ {df['日'].max()}  自然日 {df['日'].nunique()}")
    print(f"  有效车次 {len(df):,}  总头数 {df['头数'].sum():,}  总净重 {df['净重'].sum():,.0f}kg")
    print(f"  到场总均重 {df['净重'].sum()/df['头数'].sum():.2f}kg")
    print(f"  散户头数占比 {df[df['是否散户']]['头数'].sum()/df['头数'].sum()*100:.2f}%")
    print(f"  150+占比 {df[df['是否大猪']]['头数'].sum()/df['头数'].sum()*100:.2f}%")
    print(f"  外调大猪车次 {df['是否外调'].sum()}  头数 {df[df['是否外调']]['头数'].sum():,}")
    print(f"\n  体重段分布:")
    for lb in BIN_LABELS:
        s = df[df['体重段'] == lb]
        print(f"    {lb:12s} 车次{len(s):6,}  头数{s['头数'].sum():9,}  "
              f"占比{s['头数'].sum()/df['头数'].sum()*100:6.2f}%  均重{s['单车均重'].mean() if len(s) else 0:7.2f}")
    print(f"\n  各粒度期数: " + '  '.join(f'{k}={len(v)}' for k, v in periods.items()))
    print(f"\n  季节同比:")
    for mk, mv in yoy['modes'].items():
        nn = {s['key']: sum(1 for v in mv['data'][s['key']][yoy['metrics'][0]['key']] if v is not None)
              for s in mv['series'] if s['key'] in mv['data']}
        print(f"    {mk:6s} 槽位{len(mv['labels']):4d}  有数据槽位 {nn}  ({mv['hint']})")
    print(f'\nJSON -> {OUT_JSON}  ({os.path.getsize(OUT_JSON)/1024/1024:.2f} MB)')


if __name__ == '__main__':
    main()
