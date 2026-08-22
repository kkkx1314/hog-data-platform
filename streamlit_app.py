# -*- coding: utf-8 -*-
"""
生猪价格日报平台（Streamlit，蓝色主题）
- 云南价格日报：企业报价（表格）+ 散户价格三图，可下载自包含 HTML
- 价差日报：9 省价差同比季节性图（日度/周度/月度 + 公历/农历切换），可下载自包含 HTML
自动读取桌面最新日期的涌益日度数据。
运行：streamlit run streamlit_app.py
"""
import streamlit as st

import price_reports as pr

st.set_page_config(page_title="生猪价格日报平台", page_icon="🐖", layout="wide")

st.markdown("""
<style>
  :root{--accent:#2563eb;--muted:#64748b;--border:#e5e9f0;}
  .block-container{padding-top:1.6rem;padding-bottom:3rem;max-width:1200px;}
  #MainMenu,footer{visibility:hidden;}
  .app-head{border-bottom:2px solid var(--accent);padding-bottom:12px;margin-bottom:6px;}
  .app-head h1{color:#1f2937;font-size:26px;font-weight:800;margin:0;letter-spacing:.5px;}
  .app-head .sub{color:var(--muted);font-size:13px;margin-top:4px;}
</style>
""", unsafe_allow_html=True)


@st.cache_data(ttl=600, show_spinner="读取最新数据…")
def _load():
    path, latest_date = pr.latest_file()
    province_data, p_latest = pr.load_provinces(path)
    yunnan = pr.load_yunnan(path)
    yn_latest = yunnan["dates"][-1] if yunnan["dates"] else latest_date
    return path, province_data, p_latest, yunnan, yn_latest


path, province_data, p_latest, yunnan, yn_latest = _load()

price_html = pr.build_price_report_html(province_data, p_latest)
yunnan_html = pr.build_yunnan_report_html(yunnan, yn_latest)

st.markdown(f"""
<div class="app-head">
  <h1>🐖 生猪价格日报平台</h1>
  <div class="sub">数据截止 {yn_latest} ｜ 数据文件：{path}</div>
</div>
""", unsafe_allow_html=True)

tab_yn, tab_price = st.tabs(["🔵 云南价格日报", "🟢 价差日报"])

with tab_yn:
    c1, c2 = st.columns([1, 1])
    with c1:
        st.download_button(
            "⬇️ 下载 云南价格日报（自包含 HTML）",
            yunnan_html.encode("utf-8"),
            file_name="云南价格日报.html",
            mime="text/html",
            width="stretch",
        )
    with c2:
        st.caption("企业报价（表格）在上，散户价格三图在下；HTML 内可切换数据日期、编辑企业报价。")
    st.components.v1.html(yunnan_html, height=3000, scrolling=True)

with tab_price:
    c1, c2 = st.columns([1, 1])
    with c1:
        st.download_button(
            "⬇️ 下载 价差日报（自包含 HTML）",
            price_html.encode("utf-8"),
            file_name="价差日报.html",
            mime="text/html",
            width="stretch",
        )
    with c2:
        st.caption("9 省价差同比季节性图，支持日度/周度/月度与公历/农历切换。")
    st.components.v1.html(price_html, height=3200, scrolling=True)
