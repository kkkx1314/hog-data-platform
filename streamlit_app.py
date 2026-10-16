# -*- coding: utf-8 -*-
"""
生猪日报平台 —— 入口（多页导航）

  · 生猪价格日报：云南企业报价 + 散户价格三图、9 省价差同比季节性
  · 社会供应压力看板：车辆过磅单口径的压力看板（历法 × 粒度 × 截止日 × 明细）

运行：streamlit run streamlit_app.py
页面配置统一写在这里；各页脚本只管内容。
"""
import streamlit as st

st.set_page_config(page_title="生猪日报平台", page_icon="🐖", layout="wide")

st.navigation([
    st.Page("pages/0_生猪价格日报.py", title="生猪价格日报", icon="🔵",
            url_path="pig-price", default=True),
    st.Page("pages/1_社会供应压力看板.py", title="社会供应压力看板", icon="🟣",
            url_path="supply-pressure"),
], position="top").run()
