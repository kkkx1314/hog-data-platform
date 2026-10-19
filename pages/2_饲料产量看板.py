# -*- coding: utf-8 -*-
"""饲料产量看板（Streamlit 页）

交互：
  · 饲料口径：全部饲料 / 配合饲料 / 浓缩饲料 / 添加剂预混料
  · 频率：月度 / 季度；指标：产量 / 环比 / 同比
  · 品类：猪料合计 + 8 个分项（嵌套口径）
  · 对比年份：默认近三年，可多选

图表：
  1. 同期对比折线（2026 红、2025 绿、2024 黑，更早年份为灰色参照）；
  2. 最新年份的月度/季度柱图（环比/同比用红涨蓝跌，带正负号标签）；
  3. 月度/季度明细表 + CSV 下载。

数据层见 feed_dashboard.py；云端快照 data/feed_dashboard.json。
"""
import html
import os
import re

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import feed_dashboard as fd

st.set_page_config(page_title="饲料产量看板", page_icon="🌾", layout="wide")

FONT = 'system-ui, -apple-system, "Segoe UI", "Microsoft YaHei", sans-serif'

# 年份固定色：2026 红 / 2025 绿 / 2024 黑（与社会供应压力看板口径一致），
# 更早年份用中性灰阶做背景参照。颜色跟随年份实体，不随筛选重排。
# 深色模式下“黑色”改为近白，否则在深底上不可见。
THEMES = {
    'light': {
        'bg': '#ffffff', 'text': '#0b0b0b', 'secondary': '#52514e',
        'muted': '#898781', 'grid': '#e1e0d9', 'axis': '#c3c2b7',
        'years': {2026: '#d2352f', 2025: '#0f8f45', 2024: '#0b0b0b',
                  2023: '#6b6a64', 2022: '#898781', 2021: '#a8a69c', 2020: '#c3c2b7'},
        'bar_level': '#d66b66', 'bar_level_latest': '#8f1712',
        'pos': '#d2352f', 'neg': '#2a78d6',
        'up_text': '#b3261e', 'down_text': '#1d4ed8',
        'kpi_bg': '#f9f9f7', 'kpi_bd': 'rgba(11,11,11,0.08)',
        'hover_bg': '#ffffff',
    },
    'dark': {
        'bg': '#0e1117', 'text': '#ffffff', 'secondary': '#c3c2b7',
        'muted': '#898781', 'grid': '#2c2c2a', 'axis': '#383835',
        'years': {2026: '#ef5b52', 2025: '#35c46e', 2024: '#f0f0f0',
                  2023: '#6b6a64', 2022: '#898781', 2021: '#a8a69c', 2020: '#c3c2b7'},
        'bar_level': '#b8534e', 'bar_level_latest': '#ff8a80',
        'pos': '#ef5b52', 'neg': '#3987e5',
        'up_text': '#ff8a80', 'down_text': '#82b1ff',
        'kpi_bg': '#1a1a19', 'kpi_bd': 'rgba(255,255,255,0.10)',
        'hover_bg': '#1a1a19',
    },
}

FREQS = {'月度': 'M', '季度': 'Q'}
METRICS = {'产量': 'level', '环比': 'mom', '同比': 'yoy'}
METRIC_UNIT = {'level': '吨', 'mom': '%', 'yoy': '%'}

st.markdown("""
<style>
  :root{--accent:#2563eb;--muted:#64748b;}
  .block-container{padding-top:1.4rem;padding-bottom:2.2rem;max-width:1500px;}
  #MainMenu,footer{visibility:hidden;}
  .app-head{border-bottom:2px solid var(--accent);padding-bottom:10px;margin-bottom:8px;}
  .app-head h1{color:inherit;font-size:26px;font-weight:800;margin:0;letter-spacing:.5px;}
  .app-head .sub{color:var(--muted);font-size:13px;margin-top:4px;}
  .kpi-row{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin:2px 0 14px;}
  .kpi{border:1px solid;border-radius:12px;padding:12px 14px;min-height:92px;}
  .kpi .k{font-size:12px;font-weight:600;letter-spacing:.2px;}
  .kpi .v{font-size:26px;font-weight:800;line-height:1.25;margin-top:2px;}
  .kpi .u{font-size:13px;font-weight:600;margin-left:3px;}
  .kpi .d{font-size:12px;font-weight:600;margin-top:3px;line-height:1.35;}
  .kpi .n{font-size:12px;font-weight:500;margin-top:3px;line-height:1.35;}
  @media (max-width:900px){.kpi-row{grid-template-columns:repeat(2,minmax(0,1fr));}}
</style>
""", unsafe_allow_html=True)


def _theme():
    try:
        t = st.context.theme.type
        if t in THEMES:
            return THEMES[t]
    except Exception:
        pass
    try:
        b = st.get_option('theme.base')
        if b in THEMES:
            return THEMES[b]
    except Exception:
        pass
    return THEMES['light']


@st.cache_data(ttl=600, show_spinner="读取饲料数据…")
def _load(source):
    return fd.load(source)


def _fmt_level(v):
    return f'{v:,.2f}' if v is not None else '—'


def _fmt_pct(v):
    return f'{v:+.2f}%' if v is not None else '—'


def _sign_color(v, theme):
    if v is None:
        return theme['muted']
    if v > 0:
        return theme['up_text']
    if v < 0:
        return theme['down_text']
    return theme['muted']


def _delta_line(v, prefix, theme):
    if v is None:
        return f'<span style="color:{theme["muted"]}">—</span>'
    if v > 0:
        arrow, color = '▲', theme['up_text']
    elif v < 0:
        arrow, color = '▼', theme['down_text']
    else:
        arrow, color = '■', theme['muted']
    return f'<span style="color:{color}">{arrow} {prefix} {v:+.2f}%</span>'


def _card(label, value, unit, line, theme, value_color=None, line_is_note=False):
    vcolor = value_color or theme['text']
    line_color = theme['muted'] if line_is_note else None
    line_style = f'color:{line_color};' if line_color else ''
    return f"""<div class="kpi" style="background:{theme['kpi_bg']};border-color:{theme['kpi_bd']}">
      <div class="k" style="color:{theme['muted']}">{html.escape(label)}</div>
      <div class="v" style="color:{vcolor}">{value}<span class="u" style="color:{theme['muted']}">{unit}</span></div>
      <div class="{'n' if line_is_note else 'd'}" style="{line_style}">{line}</div>
    </div>"""


def _year_style(year):
    # 2025 绿 / 2026 红在红绿色盲下区分度有限，用线型 + 数据点 + 直接标注
    # 做第二重编码，颜色不是唯一线索。
    return {
        2026: (3.0, True, 'solid'),
        2025: (2.4, True, 'dash'),
        2024: (2.2, True, 'dot'),
        2023: (1.4, False, 'solid'),
        2022: (1.4, False, 'dash'),
        2021: (1.4, False, 'dot'),
        2020: (1.4, False, 'dashdot'),
    }.get(year, (1.4, False, 'solid'))


def _period_keys(freq):
    """返回 (数据键后缀, 显示标签)：月度 12 个，季度 4 个。"""
    if freq == 'Q':
        return [f'Q{q}' for q in range(1, 5)], [f'Q{q}' for q in range(1, 5)]
    return [f'{m:02d}' for m in range(1, 13)], [f'{m}月' for m in range(1, 13)]


def _period_key(freq, year, idx):
    """idx 为 0-based：月度 -> 2026-03；季度 -> 2026-Q2。"""
    if freq == 'Q':
        return f'{year}-Q{idx + 1}'
    return f'{year}-{idx + 1:02d}'


def _latest_quarter_idx(scope, year):
    """该年最新完整季度的 0-based 下标；没有完整季度返回 -1。"""
    latest = scope.get('latest') or ''
    if not latest[:4].isdigit():
        return -1
    ly, lm = int(latest[:4]), int(latest[5:])
    if year < ly:
        return 3
    if year > ly:
        return -1
    return lm // 3 - 1


def _latest_period(scope, freq):
    latest = scope.get('latest')
    if not latest:
        return None
    if freq == 'M':
        return latest
    idx = _latest_quarter_idx(scope, int(latest[:4]))
    return _period_key('Q', int(latest[:4]), idx) if idx >= 0 else None


def _shift_period(period, freq, delta):
    if freq == 'Q':
        return fd.shift_quarter(period, delta)
    return fd.shift_month(period, delta)


def _period_no(period, freq):
    """期间序号（月度 1-12，季度 1-4），用于累计和当年期数。"""
    return int(period.split('-Q')[1]) if freq == 'Q' else int(period[5:])


def _year_color(year, theme):
    return theme['years'].get(year, theme['muted'])


def _label_years(data_by_year, years, idx):
    """选出要直接标注的年份：最新年份优先，旧年份与已标注值太近就不标。

    所有标签都放在同一个 idx（最新年份的最新月），这样读者看到的是同月对比，
    而不是各自折线的末端。
    """
    points = [(y, data_by_year[y][idx]) for y in years
              if data_by_year[y][idx] is not None]
    if not points:
        return set()
    vals = [v for _, v in points]
    span = max(vals) - min(vals)
    threshold = max(span * 0.08, 0.0)
    latest_year = max(years)
    kept, picked = [], set()
    for y in sorted([y for y, _ in points],
                    key=lambda y: (y != latest_year, -y)):
        v = dict(points)[y]
        if all(abs(v - kv) > threshold for _, kv in kept):
            picked.add(y)
            kept.append((y, v))
    return picked


def fig_seasonal(scope, cat, metric, years, theme, freq='M'):
    years = sorted(set(years))
    _, x_labels = _period_keys(freq)
    data_by_year = {
        y: [fd.metric_value(scope, _period_key(freq, y, i), cat, metric)
            for i in range(len(x_labels))]
        for y in years
    }
    focus_year = max(years)
    latest = scope.get('latest') or ''
    if freq == 'Q':
        label_idx = (_latest_quarter_idx(scope, focus_year)
                     if latest[:4].isdigit() and int(latest[:4]) == focus_year else 3)
    else:
        label_idx = (int(latest[5:]) - 1
                     if latest[:4].isdigit() and int(latest[:4]) == focus_year else 11)
    label_idx = max(0, label_idx)
    label_years = _label_years(data_by_year, years, label_idx)
    fig = go.Figure()

    for y in years:
        vals = data_by_year[y]
        width, focus, dash = _year_style(y)
        color = _year_color(y, theme)
        text = [''] * len(x_labels)
        if focus and y in label_years and vals[label_idx] is not None:
            v = vals[label_idx]
            text[label_idx] = f'{v:,.0f}' if metric == 'level' else f'{v:+.1f}%'
        hover = ('%{y:,.2f} 吨<extra></extra>' if metric == 'level'
                 else '%{y:+.2f}%<extra></extra>')
        fig.add_trace(go.Scatter(
            x=x_labels, y=vals, name=f'{y}年',
            mode='lines+markers+text' if focus else 'lines',
            line=dict(color=color, width=width, dash=dash),
            marker=(dict(size=8, color=color, line=dict(width=2, color=theme['bg']))
                    if focus else dict(size=0)),
            text=text, textposition='middle right',
            textfont=dict(size=11, color=theme['text']),
            cliponaxis=False, connectgaps=False,
            hovertemplate=hover,
        ))

    fig.update_layout(
        height=410,
        margin=dict(l=8, r=82, t=46, b=8),
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
        font=dict(family=FONT, size=12, color=theme['secondary']),
        hovermode='x unified',
        hoverlabel=dict(bgcolor=theme['hover_bg'], bordercolor=theme['axis'],
                        font=dict(color=theme['text'], size=12)),
        legend=dict(orientation='h', yanchor='bottom', y=1.01, x=0,
                    bgcolor='rgba(0,0,0,0)',
                    font=dict(color=theme['secondary'], size=12)),
        showlegend=len(years) > 1,
    )
    fig.update_xaxes(type='category', showgrid=False, zeroline=False,
                     linecolor=theme['axis'], ticks='',
                     tickfont=dict(color=theme['muted'], size=12))
    fig.update_yaxes(showgrid=True, gridcolor=theme['grid'], gridwidth=1,
                     zeroline=False, linecolor=theme['axis'],
                     tickfont=dict(color=theme['muted'], size=11),
                     separatethousands=True,
                     title=dict(text=METRIC_UNIT[metric],
                                font=dict(color=theme['muted'], size=11)))
    if metric == 'level':
        fig.update_yaxes(tickformat=',.0f', rangemode='tozero')
    else:
        vals = [v for y in years for v in data_by_year[y] if v is not None] + [0]
        lo, hi = min(vals), max(vals)
        pad = (hi - lo) * 0.12 or 1
        fig.update_yaxes(tickformat='+.0f', ticksuffix='%', range=[lo - pad, hi + pad])
        fig.add_hline(y=0, line_color=theme['axis'], line_width=1)
    return fig


def fig_period_bars(scope, cat, metric, year, theme, freq='M'):
    if freq == 'Q':
        n = _latest_quarter_idx(scope, year) + 1
        periods = [_period_key('Q', year, i) for i in range(n)]
        x = [f'Q{i + 1}' for i in range(n)]
    else:
        latest = scope.get('latest')
        n = int(latest[5:]) if latest and int(latest[:4]) == year else 12
        periods = [_period_key('M', year, i) for i in range(n)]
        x = [f'{i + 1}月' for i in range(n)]
    vals = [fd.metric_value(scope, p, cat, metric) for p in periods]

    if metric == 'level':
        colors = [theme['bar_level']] * len(vals)
        valid = [i for i, v in enumerate(vals) if v is not None]
        if valid:
            colors[valid[-1]] = theme['bar_level_latest']
        text = [''] * len(vals)
        if valid:
            text[valid[-1]] = f'{vals[valid[-1]]:,.0f}'
        hover = '%{y:,.2f} 吨<extra></extra>'
    else:
        colors = [theme['muted'] if v == 0 else (theme['pos'] if v > 0 else theme['neg'])
                  for v in vals]
        text = [f'{v:+.1f}%' if v is not None else '' for v in vals]
        hover = '%{y:+.2f}%<extra></extra>'

    fig = go.Figure(go.Bar(
        x=x, y=vals,
        marker=dict(color=colors, line=dict(width=0)),
        text=text, textposition='outside', cliponaxis=False,
        textfont=dict(size=10, color=theme['text']),
        hovertemplate=hover,
    ))
    fig.update_layout(
        height=300,
        margin=dict(l=8, r=18, t=16, b=8),
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
        font=dict(family=FONT, size=12, color=theme['secondary']),
        showlegend=False,
        bargap=0.3,
        hoverlabel=dict(bgcolor=theme['hover_bg'], bordercolor=theme['axis'],
                        font=dict(color=theme['text'], size=12)),
    )
    fig.update_xaxes(type='category', showgrid=False, zeroline=False,
                     linecolor=theme['axis'], ticks='',
                     tickfont=dict(color=theme['muted'], size=12))
    fig.update_yaxes(showgrid=True, gridcolor=theme['grid'], gridwidth=1,
                     zeroline=False, linecolor=theme['axis'],
                     tickfont=dict(color=theme['muted'], size=11),
                     separatethousands=True,
                     title=dict(text=METRIC_UNIT[metric],
                                font=dict(color=theme['muted'], size=11)))
    if metric == 'level':
        top = max([v for v in vals if v is not None] or [0])
        fig.update_yaxes(tickformat=',.0f', range=[0, top * 1.12 if top > 0 else 1])
    else:
        valid = [v for v in vals if v is not None] + [0]
        lo, hi = min(valid), max(valid)
        pad = (hi - lo) * 0.18 or 1
        fig.update_yaxes(tickformat='+.0f', ticksuffix='%', range=[lo - pad, hi + pad])
        fig.add_hline(y=0, line_color=theme['axis'], line_width=1)
    return fig


# ---------- 数据 ----------
try:
    source = fd.find_source()
    data = _load(source)
except Exception as e:
    st.error(f'读取饲料数据失败：{e}')
    st.markdown("""
**本地**：确认桌面四份 `饲料_*_图表版.xlsx` 都在；
**云端**：先在桌面运行 `python build_feed_dashboard_data.py`，再同步到 GitHub。
""")
    st.stop()

meta = data['meta']
scopes = data['scopes']
latest_global = meta.get('latest') or '—'
theme = _theme()

st.markdown(f"""
<div class="app-head">
  <h1>🌾 云南饲料产量看板</h1>
  <div class="sub">月度 / 季度产量 · 环比 · 同比 ｜ 数据截至 {latest_global}
    ｜ 来源：{html.escape(meta.get('source', ''))}</div>
</div>
""", unsafe_allow_html=True)

# ---------- 控件 ----------
with st.container(border=True):
    scope_labels = [fd.SCOPE_LABELS[k] for k in data['scope_order']]
    scope_label = st.segmented_control(
        '饲料口径', scope_labels, default=scope_labels[0], key='feed_scope',
        help='全部饲料 = 配合饲料 + 浓缩饲料 + 添加剂预混合饲料（逐月逐品类成立）',
    ) or scope_labels[0]
    scope_key = fd.LABEL_TO_KEY[scope_label]
    scope = scopes[scope_key]

    c1, c2, c3, c4 = st.columns([0.9, 0.9, 1.5, 1.4])
    freq_label = c1.segmented_control(
        '频率', list(FREQS), default='月度', key='feed_freq',
        help='季度产量 = 该季度 3 个月产量之和；只对完整季度计算环比/同比',
    ) or '月度'
    freq = FREQS[freq_label]
    metric_label = c2.segmented_control(
        '指标', list(METRICS), default='环比', key='feed_metric_kind',
    ) or '环比'
    metric = METRICS[metric_label]

    cat_raw = scope['categories']
    cat_labels = [fd.CATEGORY_LABELS.get(c, c) for c in cat_raw]
    cat_label = c3.selectbox(
        '品类', cat_labels, index=0, key=f'feed_cat_{scope_key}',
        help='仔猪=哺乳+保育，母猪=后备+其他母猪；分项不可相加',
    )
    cat = cat_raw[cat_labels.index(cat_label)]

    all_years = sorted({int(m[:4]) for m in scope['months']}, reverse=True)
    default_years = [y for y in (2026, 2025, 2024) if y in all_years] or all_years[:3]
    years = c4.multiselect(
        '对比年份', all_years, default=default_years, key=f'feed_years_{scope_key}',
        max_selections=6,
        help='默认近三年；2026 红色、2025 绿色、2024 黑色，更早年份灰色',
    )
    if not years:
        years = [all_years[0]]
    years = sorted(years)

# ---------- KPI ----------
latest_p = _latest_period(scope, freq)
if latest_p is None:
    st.error('该口径暂无完整季度数据。')
    st.stop()
unit_word = '月度' if freq == 'M' else '季度'
same_word = '同月' if freq == 'M' else '同季'
cur = fd.period_value(scope, latest_p, cat)
mom = fd.metric_value(scope, latest_p, cat, 'mom')
yoy = fd.metric_value(scope, latest_p, cat, 'yoy')
prev_p = _shift_period(latest_p, freq, -1)
prev_y_p = _shift_period(latest_p, freq, -12 if freq == 'M' else -4)
prev_lv = fd.period_value(scope, prev_p, cat)
prev_y_lv = fd.period_value(scope, prev_y_p, cat)

y = int(latest_p[:4])
n_periods = _period_no(latest_p, freq)
cur_vals = [fd.period_value(scope, _period_key(freq, y, i), cat)
            for i in range(n_periods)]
prev_vals = [fd.period_value(scope, _period_key(freq, y - 1, i), cat)
             for i in range(n_periods)]
if all(v is not None for v in cur_vals) and all(v is not None for v in prev_vals):
    ytd_cur, ytd_prev = sum(cur_vals), sum(prev_vals)
    ytd_yoy = fd.pct_change(ytd_cur, ytd_prev)
else:
    ytd_cur = ytd_prev = ytd_yoy = None

ytd_label = (f'{y}年1-{n_periods}月累计' if freq == 'M'
             else f'{y}年Q1-Q{n_periods}累计')

cards = [
    _card(f'最新{unit_word}产量 · {latest_p}', _fmt_level(cur), '吨',
          _delta_line(mom, '环比', theme), theme),
    _card(f'最新{unit_word}环比', _fmt_pct(mom), '',
          f'上期 {prev_p}：{_fmt_level(prev_lv)} 吨', theme,
          value_color=_sign_color(mom, theme), line_is_note=True),
    _card(f'最新{unit_word}同比', _fmt_pct(yoy), '',
          f'上年同期 {prev_y_p}：{_fmt_level(prev_y_lv)} 吨', theme,
          value_color=_sign_color(yoy, theme), line_is_note=True),
    _card(ytd_label, _fmt_level(ytd_cur), '吨',
          _delta_line(ytd_yoy, '累计同比', theme), theme),
]
st.markdown('<div class="kpi-row">' + ''.join(cards) + '</div>', unsafe_allow_html=True)

# ---------- 图 1：同期对比 ----------
st.markdown(f'#### {same_word}对比 · {freq_label}{metric_label}')
_focus_year = max(years)
_latest_year = int(scope['latest'][:4])
if freq == 'Q':
    _focus_idx = (_latest_quarter_idx(scope, _focus_year)
                  if _focus_year == _latest_year else 3)
    _focus_p = f'Q{max(_focus_idx, 0) + 1}'
    _only_word = '只画到已公布季度'
else:
    _focus_month = int(scope['latest'][5:]) if _focus_year == _latest_year else 12
    _focus_p = f'{_focus_month}月'
    _only_word = '只画到已公布月份'
if _focus_year == _latest_year:
    st.caption(f'标签为 {_focus_year} 年 {_focus_p} 的{same_word}值；'
               f'{_focus_year} 年折线{_only_word}。悬停可看所有年份同{same_word[1]}数值。')
else:
    st.caption(f'标签为 {_focus_year} 年 {_focus_p} 的{same_word}值；'
               f'悬停可看所有年份同{same_word[1]}数值。')
st.plotly_chart(fig_seasonal(scope, cat, metric, years, theme, freq),
                width='stretch',
                config={'displayModeBar': False, 'responsive': True})

# ---------- 图 2：最新年份月度/季度 ----------
st.markdown(f'#### {y} 年{freq_label}{metric_label}')
st.caption('红色 = 环比/同比上升，蓝色 = 下降；柱顶标签带正负号，颜色不是唯一线索。'
           if metric != 'level' else '柱顶只标最新一期，其余看悬停或下方明细表。')
st.plotly_chart(fig_period_bars(scope, cat, metric, y, theme, freq),
                width='stretch',
                config={'displayModeBar': False, 'responsive': True})

# ---------- 明细表 ----------
period_col = '月份' if freq == 'M' else '季度'
prev_col = '上年同月（吨）' if freq == 'M' else '上年同季（吨）'
st.markdown(f'#### {y} 年{freq_label}明细')
rows = []
for i in range(n_periods):
    period = _period_key(freq, y, i)
    lv = fd.period_value(scope, period, cat)
    if lv is None:
        continue
    py_period = _period_key(freq, y - 1, i)
    py_lv = fd.period_value(scope, py_period, cat)
    rows.append({
        period_col: period,
        '产量（吨）': lv,
        '环比（%）': fd.metric_value(scope, period, cat, 'mom'),
        '同比（%）': fd.metric_value(scope, period, cat, 'yoy'),
        prev_col: py_lv,
        '同比增减（吨）': (lv - py_lv) if py_lv is not None else None,
    })
df = pd.DataFrame(rows)


def _color_num(v):
    if v is None or pd.isna(v):
        return ''
    if v > 0:
        return f'color:{theme["up_text"]};font-weight:600'
    if v < 0:
        return f'color:{theme["down_text"]};font-weight:600'
    return f'color:{theme["muted"]}'


fmt = {'产量（吨）': '{:,.2f}', '环比（%）': '{:+.2f}%',
       '同比（%）': '{:+.2f}%', prev_col: '{:,.2f}',
       '同比增减（吨）': '{:+,.2f}'}
styled = (df.style.format(fmt, na_rep='—')
          .map(_color_num, subset=['环比（%）', '同比（%）', '同比增减（吨）']))
st.dataframe(styled, width='stretch', hide_index=True)

# 全历史 CSV
if freq == 'M':
    all_periods = list(scope['months'])
else:
    years_all = sorted({int(m[:4]) for m in scope['months']})
    all_periods = [f'{yy}-Q{q}' for yy in years_all for q in range(1, 5)]
all_rows = []
for period in all_periods:
    lv = fd.period_value(scope, period, cat)
    if lv is None:
        continue
    all_rows.append({
        period_col: period,
        '产量（吨）': lv,
        '环比（%）': fd.metric_value(scope, period, cat, 'mom'),
        '同比（%）': fd.metric_value(scope, period, cat, 'yoy'),
    })
csv = pd.DataFrame(all_rows).to_csv(index=False).encode('utf-8-sig')
safe_name = re.sub(r'[\\/:*?"<>|]', '', f'{scope_label}_{cat_label}')
st.download_button(
    f'⬇️ 下载当前口径 CSV（全部{period_col}）', csv,
    file_name=f'饲料_{safe_name}_{freq_label}.csv', mime='text/csv',
)

if meta.get('warnings'):
    st.caption('数据注记：' + '；'.join(meta['warnings']) + '。各口径按源表原值展示。')

with st.expander('口径说明与更新方式'):
    st.markdown(fd.NOTE_MD)
    st.markdown('**更新流程**')
    st.code('python update_feed_monthly.py\n'
            '同步数据到GitHub.bat', language='bat')
    st.caption('update_feed_monthly.py 会依次更新四份 Excel、重建 JSON 快照和公开页。')
    st.caption(f'快照生成时间：{meta.get("generated_at", "—")}'
               f'｜文件：{os.path.basename(fd.JSON_PATH)}')
