# -*- coding: utf-8 -*-
"""饲料产量看板（Streamlit 页）

交互：
  · 饲料口径：全部饲料 / 配合饲料 / 浓缩饲料 / 添加剂预混料
  · 指标：月度产量 / 月度环比 / 月度同比
  · 品类：猪料合计 + 8 个分项（嵌套口径）
  · 对比年份：默认近三年，可多选

图表：
  1. 同月对比折线（2026 高亮，历史年份为灰色参照）；
  2. 最新年份 1-9 月的月度柱图（环比/同比用红涨蓝跌，带正负号标签）；
  3. 月度明细表 + CSV 下载。

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
MONTH_LABELS = [f'{m}月' for m in range(1, 13)]

# 参考 dataviz 调色板（浅/深色分别取步进值）；2026/2025/2024 用固定色，
# 更早年份用中性灰阶做背景参照。颜色跟随年份实体，不随筛选重排。
THEMES = {
    'light': {
        'bg': '#ffffff', 'text': '#0b0b0b', 'secondary': '#52514e',
        'muted': '#898781', 'grid': '#e1e0d9', 'axis': '#c3c2b7',
        'years': {2026: '#2a78d6', 2025: '#eb6834', 2024: '#1baf7a',
                  2023: '#6b6a64', 2022: '#898781', 2021: '#a8a69c', 2020: '#c3c2b7'},
        'blue': '#2a78d6', 'blue_light': '#5598e7', 'blue_dark': '#184f95',
        'pos': '#e34948', 'neg': '#2a78d6',
        'up_text': '#b3261e', 'down_text': '#1d4ed8',
        'kpi_bg': '#f9f9f7', 'kpi_bd': 'rgba(11,11,11,0.08)',
        'hover_bg': '#ffffff',
    },
    'dark': {
        'bg': '#0e1117', 'text': '#ffffff', 'secondary': '#c3c2b7',
        'muted': '#898781', 'grid': '#2c2c2a', 'axis': '#383835',
        'years': {2026: '#3987e5', 2025: '#d95926', 2024: '#199e70',
                  2023: '#6b6a64', 2022: '#898781', 2021: '#a8a69c', 2020: '#c3c2b7'},
        'blue': '#3987e5', 'blue_light': '#2a78d6', 'blue_dark': '#86b6ef',
        'pos': '#e66767', 'neg': '#3987e5',
        'up_text': '#ff8a80', 'down_text': '#82b1ff',
        'kpi_bg': '#1a1a19', 'kpi_bd': 'rgba(255,255,255,0.10)',
        'hover_bg': '#1a1a19',
    },
}

METRICS = {'月度产量': 'level', '月度环比': 'mom', '月度同比': 'yoy'}
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
    return {
        2026: (2.8, True, 'solid'),
        2025: (2.2, True, 'solid'),
        2024: (2.0, True, 'solid'),
        2023: (1.4, False, 'solid'),
        2022: (1.4, False, 'dash'),
        2021: (1.4, False, 'dot'),
        2020: (1.4, False, 'dashdot'),
    }.get(year, (1.4, False, 'solid'))


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


def fig_seasonal(scope, cat, metric, years, theme):
    years = sorted(set(years))
    data_by_year = {
        y: [fd.metric_value(scope, f'{y}-{m:02d}', cat, metric) for m in range(1, 13)]
        for y in years
    }
    focus_year = max(years)
    latest = scope.get('latest') or ''
    label_idx = (int(latest[5:]) - 1
                 if latest[:4].isdigit() and int(latest[:4]) == focus_year else 11)
    labels = _label_years(data_by_year, years, label_idx)
    fig = go.Figure()

    for y in years:
        vals = data_by_year[y]
        width, focus, dash = _year_style(y)
        color = _year_color(y, theme)
        text = [''] * 12
        if focus and y in labels and vals[label_idx] is not None:
            v = vals[label_idx]
            text[label_idx] = f'{v:,.0f}' if metric == 'level' else f'{v:+.1f}%'
        hover = ('%{y:,.2f} 吨<extra></extra>' if metric == 'level'
                 else '%{y:+.2f}%<extra></extra>')
        fig.add_trace(go.Scatter(
            x=MONTH_LABELS, y=vals, name=f'{y}年',
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


def fig_monthly_bars(scope, cat, metric, year, theme):
    latest = scope.get('latest')
    max_m = int(latest[5:]) if latest and int(latest[:4]) == year else 12
    months = list(range(1, max_m + 1))
    vals = [fd.metric_value(scope, f'{year}-{m:02d}', cat, metric) for m in months]
    x = [f'{m}月' for m in months]

    if metric == 'level':
        colors = [theme['blue_light']] * len(vals)
        valid = [i for i, v in enumerate(vals) if v is not None]
        if valid:
            colors[valid[-1]] = theme['blue_dark']
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
  <div class="sub">月度产量 · 环比 · 同比 ｜ 数据截至 {latest_global}
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

    c1, c2, c3 = st.columns([1.15, 1.25, 1.6])
    metric_label = c1.segmented_control(
        '指标', list(METRICS), default='月度环比', key='feed_metric',
    ) or '月度环比'
    metric = METRICS[metric_label]

    cat_raw = scope['categories']
    cat_labels = [fd.CATEGORY_LABELS.get(c, c) for c in cat_raw]
    cat_label = c2.selectbox(
        '品类', cat_labels, index=0, key=f'feed_cat_{scope_key}',
        help='仔猪=哺乳+保育，母猪=后备+其他母猪；分项不可相加',
    )
    cat = cat_raw[cat_labels.index(cat_label)]

    all_years = sorted({int(m[:4]) for m in scope['months']}, reverse=True)
    default_years = [y for y in (2026, 2025, 2024) if y in all_years] or all_years[:3]
    years = c3.multiselect(
        '对比年份', all_years, default=default_years, key=f'feed_years_{scope_key}',
        max_selections=6,
        help='默认近三年；2026 蓝色高亮，2025 橙色，2024 青色，更早年份灰色',
    )
    if not years:
        years = [all_years[0]]
    years = sorted(years)

# ---------- KPI ----------
latest_m = scope['latest']
cur = fd.value_at(scope, latest_m, cat)
mom = fd.metric_value(scope, latest_m, cat, 'mom')
yoy = fd.metric_value(scope, latest_m, cat, 'yoy')
prev_m = fd.shift_month(latest_m, -1)
prev_y = fd.shift_month(latest_m, -12)
prev_lv = fd.value_at(scope, prev_m, cat)
prev_y_lv = fd.value_at(scope, prev_y, cat)

y, latest_mo = map(int, latest_m.split('-'))
cur_vals = [fd.value_at(scope, f'{y}-{m:02d}', cat) for m in range(1, latest_mo + 1)]
prev_vals = [fd.value_at(scope, f'{y - 1}-{m:02d}', cat) for m in range(1, latest_mo + 1)]
if all(v is not None for v in cur_vals) and all(v is not None for v in prev_vals):
    ytd_cur, ytd_prev = sum(cur_vals), sum(prev_vals)
    ytd_yoy = fd.pct_change(ytd_cur, ytd_prev)
else:
    ytd_cur = ytd_prev = ytd_yoy = None

cards = [
    _card(f'最新月产量 · {latest_m}', _fmt_level(cur), '吨',
          _delta_line(mom, '环比', theme), theme),
    _card('最新月环比', _fmt_pct(mom), '',
          f'上月 {prev_m}：{_fmt_level(prev_lv)} 吨', theme,
          value_color=_sign_color(mom, theme), line_is_note=True),
    _card('最新月同比', _fmt_pct(yoy), '',
          f'上年同月 {prev_y}：{_fmt_level(prev_y_lv)} 吨', theme,
          value_color=_sign_color(yoy, theme), line_is_note=True),
    _card(f'{y}年1-{latest_mo}月累计', _fmt_level(ytd_cur), '吨',
          _delta_line(ytd_yoy, '累计同比', theme), theme),
]
st.markdown('<div class="kpi-row">' + ''.join(cards) + '</div>', unsafe_allow_html=True)

# ---------- 图 1：同月对比 ----------
st.markdown(f'#### 同月对比 · {metric_label}')
_focus_year = max(years)
_latest_year = int(scope['latest'][:4])
_focus_month = int(scope['latest'][5:]) if _focus_year == _latest_year else 12
if _focus_year == _latest_year:
    st.caption(f'标签为 {_focus_year} 年 {_focus_month} 月的同月值；'
               f'{_focus_year} 年折线只画到已公布月份。悬停可看所有年份同月数值。')
else:
    st.caption(f'标签为 {_focus_year} 年 {_focus_month} 月的同月值；'
               f'悬停可看所有年份同月数值。')
st.plotly_chart(fig_seasonal(scope, cat, metric, years, theme),
                width='stretch',
                config={'displayModeBar': False, 'responsive': True})

# ---------- 图 2：最新年份月度 ----------
st.markdown(f'#### {y} 年月度{metric_label.replace("月度", "")}')
st.caption('红色 = 环比/同比上升，蓝色 = 下降；柱顶标签带正负号，颜色不是唯一线索。'
           if metric != 'level' else '柱顶只标最新月份，其余月份看悬停或下方明细表。')
st.plotly_chart(fig_monthly_bars(scope, cat, metric, y, theme),
                width='stretch',
                config={'displayModeBar': False, 'responsive': True})

# ---------- 明细表 ----------
st.markdown(f'#### {y} 年月度明细')
rows = []
for m in range(1, latest_mo + 1):
    month = f'{y}-{m:02d}'
    lv = fd.value_at(scope, month, cat)
    if lv is None:
        continue
    py_lv = fd.value_at(scope, f'{y - 1}-{m:02d}', cat)
    rows.append({
        '月份': month,
        '产量（吨）': lv,
        '环比（%）': fd.metric_value(scope, month, cat, 'mom'),
        '同比（%）': fd.metric_value(scope, month, cat, 'yoy'),
        '上年同月（吨）': py_lv,
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


styled = (df.style
          .format({'产量（吨）': '{:,.2f}', '环比（%）': '{:+.2f}%',
                   '同比（%）': '{:+.2f}%', '上年同月（吨）': '{:,.2f}',
                   '同比增减（吨）': '{:+,.2f}'}, na_rep='—')
          .map(_color_num, subset=['环比（%）', '同比（%）', '同比增减（吨）']))
st.dataframe(styled, width='stretch', hide_index=True)

# 全历史 CSV
all_rows = []
for month in scope['months']:
    lv = fd.value_at(scope, month, cat)
    if lv is None:
        continue
    all_rows.append({
        '月份': month,
        '产量（吨）': lv,
        '环比（%）': fd.metric_value(scope, month, cat, 'mom'),
        '同比（%）': fd.metric_value(scope, month, cat, 'yoy'),
    })
csv = pd.DataFrame(all_rows).to_csv(index=False).encode('utf-8-sig')
safe_name = re.sub(r'[\\/:*?"<>|]', '', f'{scope_label}_{cat_label}')
st.download_button(
    '⬇️ 下载当前口径 CSV（全部月份）', csv,
    file_name=f'饲料_{safe_name}.csv', mime='text/csv',
)

if meta.get('warnings'):
    st.caption('数据注记：' + '；'.join(meta['warnings']) + '。各口径按源表原值展示。')

with st.expander('口径说明与更新方式'):
    st.markdown(fd.NOTE_MD)
    st.markdown('**更新流程**')
    st.code('python update_feed_monthly.py\n'
            'python build_feed_dashboard_data.py\n'
            '同步数据到GitHub.bat', language='bat')
    st.caption(f'快照生成时间：{meta.get("generated_at", "—")}'
               f'｜文件：{os.path.basename(fd.JSON_PATH)}')
