# -*- coding: utf-8 -*-
"""
社会供应压力看板 —— Streamlit 版的数据/外壳层。

看板本体仍是 `_wb_template.html`（单文件、20 张 SVG 图、深/浅色、口径板块一体），
这里只做三件事：
  1. 找到最新的看板数据：本地优先读根目录 `_wb_data.json`（.bat 刚跑完就是最新的），
     云端没有桌面，读仓库里的 `data/_wb_data.json`；
  2. 把外壳选的「历法 × 粒度 × 截止日 × 区间 × 明细」写成初始状态注入模板
     （模板里 data-embed 会把内嵌的那排筛选控件藏掉，避免和 Streamlit 控件打架）；
  3. 交给页面用 st.components.v1.html 嵌进去。

这样站点上和桌面上是同一份 HTML，_wb_check.py 的 38,517 项独立复算结论照样成立。

数据刷新：桌面跑「更新看板数据.bat」（_wb_watch.py）→ 重建看板 + 重写 data/ 副本，
再跑「同步数据到GitHub.bat」推上去，Streamlit Cloud 就跟着更新。
"""
import os

import _wb_build as wb

HERE = os.path.dirname(os.path.abspath(__file__))

# 本地优先，云端只有 data/ 那份（根目录那份在 .gitignore 里，推不上去）
SOURCES = (
    os.path.join(HERE, '_wb_data.json'),
    os.path.join(HERE, 'data', '_wb_data.json'),
)

# 明细（体重段明细 / 审计存疑 / 口径诊断）开与关差 2300px 左右；这是 iframe 的初始高度，
# 看板加载完会自己把真实高度回传给 window.frameElement 覆盖掉这两个值
H_BASE, H_DETAIL = 3900, 6250


def find_source():
    for p in SOURCES:
        if os.path.exists(p):
            return p
    raise FileNotFoundError('找不到 _wb_data.json —— 先跑一次 _wb_build.py')


def load(path=None):
    return wb.load_data(path or find_source())


def build(data, cal='公历', gran='日', end_date=None, range_days=0, detail=False,
          theme=None):
    """拼初始状态。end_date 传 None 表示「最新」；range_days=0 表示全部区间；
    theme 传 'light'/'dark' 就锁死深浅色，传 None 就跟随系统。"""
    return wb.build_html(data, state={
        'cal': cal,
        'gran': gran,
        'endDate': end_date,
        'range': int(range_days or 0),
        'detail': bool(detail),
        'embed': True,
        'theme': theme,
    })
