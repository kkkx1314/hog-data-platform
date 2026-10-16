# -*- coding: utf-8 -*-
"""
社会供应压力看板（Streamlit 页）

看板本体是单文件 HTML（20 张 SVG 图 + 深/浅色 + 口径板块），这里只当外壳：
上面一排控件决定初始状态，下面把它嵌进 iframe。控件改一次就重建一次 HTML。

运行：streamlit run streamlit_app.py    然后在左侧选“社会供应压力看板”
"""
import datetime as dt
import os

import streamlit as st

import social_dashboard as sd

st.set_page_config(page_title="社会供应压力看板", page_icon="🐖", layout="wide")

st.markdown("""
<style>
  :root{--accent:#2563eb;--muted:#64748b;}
  .block-container{padding-top:1.4rem;padding-bottom:2rem;max-width:1500px;}
  #MainMenu,footer{visibility:hidden;}
  .app-head{border-bottom:2px solid var(--accent);padding-bottom:10px;margin-bottom:6px;}
  .app-head h1{color:#1f2937;font-size:26px;font-weight:800;margin:0;letter-spacing:.5px;}
  .app-head .sub{color:var(--muted);font-size:13px;margin-top:4px;}
  .stElementContainer iframe{display:block;}
</style>
""", unsafe_allow_html=True)

RANGE_LABELS = {"近 30": 30, "近 90": 90, "近一年": 365, "全部": 0}
THEMES = {"跟随系统": None, "浅色": "light", "深色": "dark"}


@st.cache_data(ttl=600, show_spinner="读取过磅单数据…")
def _load(path, mtime):
    # mtime 进签名：桌面重新导出过磅单后 _wb_data.json 一变，缓存立刻失效
    return sd.load(path)


@st.cache_data(ttl=600, show_spinner="生成看板…")
def _build(path, mtime, cal, gran, end, rng, detail, theme):
    return sd.build(sd.load(path), cal=cal, gran=gran, end_date=end,
                    range_days=rng, detail=detail, theme=theme)


try:
    path = sd.find_source()
    data = _load(path, os.path.getmtime(path))
except FileNotFoundError as e:
    st.error(f'{e}\n\n桌面跑一次「更新看板数据.bat」再同步到 GitHub 即可。')
    st.stop()
meta = data['meta']
dmin, dmax = dt.date.fromisoformat(meta['date_min']), dt.date.fromisoformat(meta['date_max'])

st.markdown(f"""
<div class="app-head">
  <h1>🐖 社会供应压力看板</h1>
  <div class="sub">数据截止 {dmax}（{dmin} ~ {dmax}，{meta['cars']:,} 车 / {meta['heads']:,} 头）
    ｜ 来源：{os.path.basename(path)}</div>
</div>
""", unsafe_allow_html=True)

# 历法 × 粒度 × 截止日 × 区间 × 明细 × 深浅色 —— 就是模板里那一排控件的替身。
# 分两行摆：挤在一行时段选按钮会折行，反而看不清。
with st.container(border=True):
    c1, c2, c3, c4 = st.columns([0.85, 0.85, 1.4, 1.6])
    cal = c1.segmented_control('历法', ['公历', '农历'], default='公历', key='wb_cal') or '公历'
    gran = c2.segmented_control('粒度', ['日', '周', '月'], default='日', key='wb_gran') or '日'
    end = c3.date_input('截止日期', value=dmax, min_value=dmin, max_value=dmax,
                        format='YYYY-MM-DD', key='wb_end',
                        help='默认最新；周/月粒度会贴到该期期末，末期未完会标注')
    rl = c4.segmented_control('区间', list(RANGE_LABELS), default='全部', key='wb_rng') or '全部'

    c5, c6, c7 = st.columns([1.0, 1.3, 2.3])
    det = c5.toggle('显示明细', value=False, key='wb_det',
                    help='体重段明细 / 审计存疑 / 口径诊断三块')
    th = c6.segmented_control('主题', list(THEMES), default='跟随系统', key='wb_theme') or '跟随系统'
    html = _build(path, os.path.getmtime(path), cal, gran, end.isoformat(),
                  RANGE_LABELS[rl], det, THEMES[th])
    with c7:
        st.download_button('⬇️ 下载看板（单文件 HTML）', html.encode('utf-8'),
                           file_name='社会供应压力看板.html', mime='text/html',
                           width='stretch')

st.caption('看板内滚轮/悬停均可交互；深浅色由上方“主题”接管。'
           '数据来自车辆过磅单全量并集，明细开关只影响显示、不改口径。')

# 高度先给个估值，看板加载完会自己把真实高度回传给 iframe 覆盖掉
st.components.v1.html(html, height=sd.H_DETAIL if det else sd.H_BASE, scrolling=True)
