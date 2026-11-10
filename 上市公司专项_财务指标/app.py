# -*- coding: utf-8 -*-
"""上市猪企财务质量交互面板（Streamlit）。"""
from __future__ import annotations

import math
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
XLSX = ROOT / "output" / "上市公司财务指标_底表.xlsx"
TPL = ROOT / "output" / "上市公司财务指标_趋势与资产_模板.xlsx"
HTML = ROOT / "output" / "上市公司财务指标_趋势与资产_看板.html"

st.set_page_config(page_title="上市猪企财务质量面板", page_icon="🐖", layout="wide")

COMPANY_COLOR = {
    "牧原": "#2E5EAA", "温氏": "#D9822B", "新希望": "#2A9D8F", "正邦": "#8E44AD",
    "天邦": "#C0392B", "新五丰": "#16A085", "唐人神": "#E67E22", "巨星": "#2980B9",
    "天康": "#27AE60", "大北农": "#F39C12", "神农": "#D35400", "华统": "#7F8C8D",
    "立华": "#3498DB", "傲农": "#9B59B6", "京基智农": "#1ABC9C", "东瑞": "#E74C3C",
    "金新农": "#34495E", "罗牛山": "#B7950B", "克明": "#A04000", "益生": "#5DADE2",
    "正虹": "#AF7AC5",
}
FOCUS_DEFAULT = ["牧原", "温氏", "新希望", "正邦", "天邦", "新五丰"]

INDICATORS = [
    # key, 中文, 分类, 单位类型, 方向(+1越高越好/-1越低越好), 说明
    ("毛利率", "盈利能力", "pct", 1, "(营业收入−营业成本)/营业收入"),
    ("期间费用率", "盈利能力", "pct", -1, "(销售+管理+研发+财务费用)/营业收入"),
    ("核心利润率", "盈利能力", "pct", 1, "核心利润/营业收入"),
    ("归母扣非ROE", "盈利能力", "pct", 1, "扣非归母净利润/平均归母净资产（年化）"),
    ("有息负债率", "财务健康", "pct", -1, "有息负债/总资产"),
    ("现金利息保障倍数", "财务健康", "mult", 1, "(经营现金流净额+利息费用+所得税)/利息费用"),
    ("现金债务覆盖率", "财务健康", "mult", 1, "经营现金流净额(年化)/有息负债"),
    ("现金流量比率", "财务健康", "mult", 1, "经营现金流净额(年化)/流动负债"),
    ("重资产占比", "运营效率", "pct", -1, "(投资性房地产+固定资产+在建工程+生产性生物资产)/总资产"),
    ("总资产周转率", "运营效率", "mult", 1, "营业收入(年化)/平均总资产"),
    ("人均创收", "运营效率", "percap", 1, "营业收入(年化)/员工总数"),
    ("人均核心利润", "运营效率", "percap", 1, "核心利润(年化)/员工总数"),
    ("营收现金比率", "现金韧性", "pct", 1, "销售商品收到的现金/营业收入"),
    ("自由现金流利润率", "现金韧性", "pct", 1, "(经营现金流净额−资本开支)/营业收入"),
    ("资本开支覆盖倍数", "现金韧性", "mult", 1, "经营现金流净额/购建长期资产支付的现金"),
    ("经营现金流入占比", "现金韧性", "pct", 1, "经营现金流入/(经营+投资+筹资现金流入)"),
    ("资产复合增长率", "发展质量", "pct", 1, "近3年总资产CAGR"),
    ("营收复合增长率", "发展质量", "pct", 1, "近3年TTM营收CAGR"),
    ("累计资产现金回报率", "发展质量", "pct", 1, "2020年以来累计经营现金流/区间平均总资产"),
    ("高学历人才密度", "发展质量", "pct", 1, "本科及以上员工/员工总数"),
]
IND_META = {x[0]: {"cat": x[1], "type": x[2], "dir": x[3], "formula": x[4]} for x in INDICATORS}
IND_NAMES = [x[0] for x in INDICATORS]
CATS = ["盈利能力", "财务健康", "运营效率", "现金韧性", "发展质量"]


@st.cache_data
def load_data():
    q = pd.read_csv(DATA / "indicators_single.csv")
    c = pd.read_csv(DATA / "indicators_cum.csv")
    a = pd.read_csv(DATA / "assets_five.csv")
    kf = pd.read_csv(DATA / "kf_nonrecurring.csv")
    emp = pd.read_csv(DATA / "employees.csv")
    return q, c, a, kf, emp


def period_key(p):
    y, q = int(p[:4]), p[4:]
    return y * 4 + {"Q1": 0, "H1": 1, "Q3": 2, "FY": 3}[q]


def fmt(v, typ):
    if v is None or (isinstance(v, float) and (np.isnan(v) or np.isinf(v))):
        return "—"
    if typ == "pct":
        return f"{v*100:,.2f}%"
    if typ == "percap":
        return f"{v:,.1f}"
    return f"{v:,.2f}"


def fmt_num(v, typ):
    if typ == "pct":
        return f"{v:,.1%}"
    if typ == "percap":
        return f"{v:,.1f}"
    return f"{v:,.2f}"


q_all, c_all, assets_all, kf_all, emp_all = load_data()
all_periods = sorted(q_all["期间"].unique(), key=period_key)
all_companies = list(dict.fromkeys(q_all["公司"].tolist()))
focus_default = [c for c in FOCUS_DEFAULT if c in all_companies]

# ---------------------------------------------------------------- 侧栏
st.sidebar.title("🐖 上市猪企财务质量")
st.sidebar.caption(f"{all_periods[0]}–{all_periods[-1]} · {len(all_companies)} 家上市猪企")

companies = st.sidebar.multiselect("公司", all_companies, default=focus_default)
caliber = st.sidebar.radio("口径", ["单季", "累计"], index=0,
                           help="单季=单季度差分值；累计=年初至今累计值（比率类已年化）")
default_start = "2023Q1" if "2023Q1" in all_periods else all_periods[0]
pr = st.sidebar.select_slider("报告期", options=all_periods,
                              value=(default_start, all_periods[-1]))
st.sidebar.markdown("---")
st.sidebar.caption(
    "数据源：21 家公司定期报告（三大报表 xlsx + 556 份公告 PDF）。\n\n"
    "扣非净利润、员工学历由年报/季报原文解析；"
    "人均指标员工数取最近一期年报（Q1–Q3 用上年末；2021 年上市公司退回最早一期年报）。"
)

if not companies:
    st.warning("请在左侧至少选择一家公司。")
    st.stop()

df = (q_all if caliber == "单季" else c_all).copy()
df = df[df["公司"].isin(companies)]
mask = df["期间"].map(lambda x: period_key(pr[0]) <= period_key(x) <= period_key(pr[1]))
df = df[mask].sort_values(["公司", "期间"], key=lambda s: s.map(
    lambda x: all_companies.index(x) if x in all_companies else 0) if s.name == "公司" else s.map(period_key))
periods = sorted(df["期间"].unique(), key=period_key)

# ---------------------------------------------------------------- 标题
st.title("上市猪企财务质量面板")
st.caption(
    f"口径：**{caliber}** ｜ 报告期：**{pr[0]} – {pr[1]}** ｜ 公司：{'、'.join(companies)} ｜ "
    "盈利能力 · 财务健康 · 运营效率 · 现金韧性 · 发展质量"
)

tab1, tab2, tab3, tab4, tab5 = st.tabs(
    ["📊 总览", "📈 指标趋势", "🏭 五项资产", "📋 明细与下载", "📖 口径说明"]
)

# ---------------------------------------------------------------- 总览
with tab1:
    sel_period = st.selectbox("报告期", periods[::-1], index=0, key="overview_period")
    snap = df[df["期间"] == sel_period].set_index("公司").reindex(companies)

    focus = st.multiselect(
        "重点公司（卡片，最多 8 家）", companies,
        default=companies[:6], max_selections=8, key="overview_focus")
    if focus:
        cols = st.columns(min(4, len(focus)))
        for i, comp in enumerate(focus):
            with cols[i % len(cols)]:
                if comp in snap.index and pd.notna(snap.loc[comp, "核心利润率"]):
                    st.metric(comp, fmt(snap.loc[comp, "核心利润率"], "pct"),
                              help="核心利润率 = 核心利润 / 营业收入（最全面的盈利口径）")
                    st.caption(f"毛利率 {fmt(snap.loc[comp, '毛利率'], 'pct')} ｜ "
                               f"归母扣非ROE {fmt(snap.loc[comp, '归母扣非ROE'], 'pct')}")
                else:
                    st.metric(comp, "—")

    st.caption("口径提示：毛利率只扣营业成本；核心利润率扣营业总成本（含税金及附加和四项期间费用）。"
               "(营业总收入−营业总成本)/营业总收入 对应的是核心利润率，不是毛利率。")

    st.subheader(f"{sel_period} · 全部公司横向排名")
    rank_ind = ["核心利润率", "毛利率", "归母扣非ROE", "有息负债率",
                "营收现金比率", "自由现金流利润率", "总资产周转率", "高学历人才密度"]
    rank = q_all[q_all["期间"] == sel_period].set_index("公司").reindex(all_companies)
    rank = rank[[c for c in rank_ind if c in rank.columns]].copy()
    pct_cols_rank = [c for c in rank.columns if IND_META.get(c, {}).get("type") == "pct"]
    for c in pct_cols_rank:
        rank[c] = rank[c] * 100
    rank = rank.dropna(how="all")
    if not rank.empty:
        rank = rank.sort_values("核心利润率", ascending=False)
        st.dataframe(
            rank, width="stretch",
            column_config={c: st.column_config.NumberColumn(c, format="%.2f%%") for c in pct_cols_rank}
            | {c: st.column_config.NumberColumn(c, format="%.2f") for c in rank.columns if c not in pct_cols_rank})
        st.caption("按核心利润率降序；单位：比率类为 %，总资产周转率为倍。")
    else:
        st.info("该报告期无数据。")

    st.subheader("指标全景（原始值）")
    units = {"pct": "%", "mult": "倍", "percap": "万元/人"}
    rows = []
    for ind in IND_NAMES:
        row = {"分类": IND_META[ind]["cat"], "指标": ind, "单位": units[IND_META[ind]["type"]]}
        for comp in companies:
            row[comp] = fmt(snap.loc[comp, ind], IND_META[ind]["type"]) if comp in snap.index else "—"
        rows.append(row)
    st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True, height=690)

    st.subheader("毛利率 vs 核心利润率（同一报告期口径对照）")
    cmp_rows = []
    for comp in companies:
        qr = q_all[(q_all["公司"] == comp) & (q_all["期间"] == sel_period)]
        cr = c_all[(c_all["公司"] == comp) & (c_all["期间"] == sel_period)]

        def gv(d, k):
            return d.iloc[0][k] if len(d) else np.nan

        gm_q, core_q = gv(qr, "毛利率"), gv(qr, "核心利润率")
        gm_c, core_c = gv(cr, "毛利率"), gv(cr, "核心利润率")
        cmp_rows.append({
            "公司": comp,
            "毛利率(单季)": fmt(gm_q, "pct"),
            "核心利润率(单季)": fmt(core_q, "pct"),
            "毛利率(累计)": fmt(gm_c, "pct"),
            "核心利润率(累计)": fmt(core_c, "pct"),
            "累计差额(pp)": (f"{(core_c - gm_c) * 100:+.2f}"
                              if pd.notna(core_c) and pd.notna(gm_c) else "—"),
        })
    st.dataframe(pd.DataFrame(cmp_rows), width="stretch", hide_index=True)
    st.caption("累计差额 = 核心利润率(累计) − 毛利率(累计) = (税金及附加 + 销售 + 管理 + 研发 + 财务费用) / 营业收入。"
               "若你用 (营业总收入 − 营业总成本) / 营业总收入 计算，结果应与“核心利润率”一致，而不是“毛利率”。")

# ---------------------------------------------------------------- 趋势
with tab2:
    st.subheader("指标趋势（全部公司）")
    st.caption("趋势图与数据表默认覆盖全部公司，不受左侧“公司”筛选影响；如需聚焦可在此临时减少。")
    tcompanies = st.multiselect("趋势公司", all_companies, default=all_companies,
                                key="trend_companies")
    c1, c2 = st.columns([1, 2])
    with c1:
        cat_pick = st.multiselect("指标分类", CATS, default=CATS)
    available = [i for i in IND_NAMES if IND_META[i]["cat"] in cat_pick]
    with c2:
        pick = st.multiselect("指标（最多 6 个）", available,
                              default=[i for i in ["毛利率", "核心利润率", "归母扣非ROE", "有息负债率"] if i in available][:4],
                              max_selections=6)
    if not tcompanies:
        st.info("请选择至少一家公司。")
    elif not pick:
        st.info("请选择至少一个指标。")
    else:
        tdf = (q_all if caliber == "单季" else c_all).copy()
        tdf = tdf[tdf["公司"].isin(tcompanies)
                  & tdf["期间"].map(lambda x: period_key(pr[0]) <= period_key(x) <= period_key(pr[1]))]
        n = len(pick)
        cols = 2 if n > 1 else 1
        rows = int(np.ceil(n / cols))
        fig = make_subplots(rows=rows, cols=cols, subplot_titles=pick,
                            vertical_spacing=0.18, horizontal_spacing=0.10)
        for k, ind in enumerate(pick):
            r, c = divmod(k, cols)
            for comp in tcompanies:
                sub = tdf[tdf["公司"] == comp].sort_values("期间", key=lambda s: s.map(period_key))
                fig.add_trace(
                    go.Scatter(x=sub["期间"], y=sub[ind], name=comp, mode="lines+markers",
                               line=dict(color=COMPANY_COLOR.get(comp), width=2),
                               marker=dict(size=5), showlegend=(k == 0),
                               hovertemplate=f"{comp}<br>%{{x}}<br>%{{y:.4f}}<extra></extra>"),
                    row=r + 1, col=c + 1)
            typ = IND_META[ind]["type"]
            # y 轴顶部留白，避免折线顶到子图标题（如“核心利润率”）
            vals = pd.to_numeric(tdf[ind], errors="coerce").dropna()
            if not vals.empty:
                lo, hi = float(vals.min()), float(vals.max())
                pad = (hi - lo) * 0.18 if hi > lo else max(abs(hi) * 0.1, 0.05)
                yrange = [lo - pad, hi + pad]
            else:
                yrange = None
            fig.update_yaxes(tickformat=".1%" if typ == "pct" else ",.2f",
                             range=yrange, row=r + 1, col=c + 1)
            fig.update_xaxes(type="category", tickangle=-45, row=r + 1, col=c + 1)
        fig.update_layout(height=300 * rows, template="plotly_white",
                          legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, font=dict(size=10)),
                          margin=dict(l=10, r=10, t=80, b=10))
        st.plotly_chart(fig, width="stretch")

        st.markdown("**数据表（全部公司）**")
        tbl = tdf[["公司", "期间"] + pick].sort_values(
            ["公司", "期间"], key=lambda s: s.map(lambda x: all_companies.index(x)) if s.name == "公司" else s.map(period_key))
        st.dataframe(tbl, width="stretch", hide_index=True, height=420)

# ---------------------------------------------------------------- 资产
with tab3:
    st.subheader("五项资产（全部公司）")
    st.caption("首图为全部公司在同一报告期的资产结构横向对比，不受左侧“公司”筛选影响；下方可切换单家查看趋势与明细。")
    items = ["投资性房地产", "固定资产", "在建工程", "生产性生物资产", "存货"]
    colors = {"固定资产": "#2E5EAA", "在建工程": "#7FA8D9", "生产性生物资产": "#2A9D8F",
              "存货": "#E9C46A", "投资性房地产": "#C0392B"}
    ap = [p for p in sorted(assets_all["期间"].unique(), key=period_key)
          if period_key(pr[0]) <= period_key(p) <= period_key(pr[1])]
    if not ap:
        st.info("该区间无资产数据。")
    else:
        sel_ap = st.selectbox("对比报告期", ap[::-1], index=0, key="asset_period_all")
        snapA = assets_all[(assets_all["期间"] == sel_ap)
                           & assets_all["公司"].isin(all_companies)].copy()
        snapA = snapA.dropna(subset=["五项合计"]).sort_values("五项合计")
        if snapA.empty:
            st.info("该报告期无数据。")
        else:
            figA = go.Figure()
            for it in items:
                figA.add_trace(go.Bar(
                    y=snapA["公司"], x=snapA[it] / 1e8, name=it, orientation="h",
                    marker_color=colors[it],
                    hovertemplate=f"{it}<br>%{{y}}<br>%{{x:,.1f}} 亿元<extra></extra>"))
            figA.update_layout(barmode="stack",
                               height=max(460, 26 * len(snapA) + 140),
                               template="plotly_white", xaxis_title="亿元",
                               legend=dict(orientation="h", y=1.06),
                               margin=dict(l=10, r=10, t=40, b=10))
            st.plotly_chart(figA, width="stretch")
            st.caption(f"{sel_ap} 期末；按五项合计从大到小排列（图中自上而下）。")

        st.markdown("---")
        st.subheader("单项资产 · 历年")
        st.caption("可拖动报告期区间选择分析窗口（独立于左侧筛选）；单位：亿元。"
                   "合计按所选区间剔除未披露公司并在下方标注；报表未列示该项也按未披露处理。"
                   "右轴为相邻报告期环比（约一个季度）。")
        item_order = ["固定资产", "存货", "在建工程", "生产性生物资产", "投资性房地产"]
        ic1, ic2, ic3 = st.columns([1.0, 1.0, 1.4])
        with ic1:
            item_pick = st.selectbox("资产项目", item_order, index=0, key="asset_item_pick")
        with ic2:
            scope = st.selectbox("公司 / 合计", ["全部公司合计"] + all_companies, key="asset_scope")
        with ic3:
            basis = "整段剔除"
            if scope == "全部公司合计":
                basis = st.radio("合计口径", ["整段剔除", "逐期剔除"], horizontal=True, key="asset_basis")
        period_range = st.select_slider(
            "报告期区间（拖动两端）", options=all_periods,
            value=(all_periods[0], all_periods[-1]), key="asset_period_range")
        if isinstance(period_range, tuple):
            win_start, win_end = period_range[0], period_range[-1]
        else:
            win_start = win_end = period_range
        window = all_periods[all_periods.index(win_start):all_periods.index(win_end) + 1]

        piv = assets_all.pivot(index="期间", columns="公司", values=item_pick).reindex(all_periods)
        for c in all_companies:
            if c not in piv.columns:
                piv[c] = np.nan
        piv = piv[all_companies]
        pivw = piv.loc[window]

        if scope != "全部公司合计":
            vals = pivw[scope] / 1e8
            ns = vals.notna().astype(int)
            sample_note = f"{scope} 单家公司"
            custom = None
            hover = "%{x}<br>%{y:,.2f} 亿元<extra></extra>"
        elif basis == "整段剔除":
            eligible = [c for c in all_companies if pivw[c].notna().all()]
            excluded = [c for c in all_companies if c not in eligible]
            vals = (pivw[eligible].sum(axis=1, min_count=1) if eligible
                    else pd.Series(np.nan, index=window)) / 1e8
            ns = pd.Series(len(eligible), index=window)
            sample_note = (f"所选区间整段剔除：纳入 {len(eligible)} 家，剔除 {len(excluded)} 家"
                           "（区间内任一期未披露或报表未列示即整家剔除）")
            custom = [int(len(eligible)) for _ in window]
            hover = "%{x}<br>%{y:,.2f} 亿元<br>纳入 %{customdata} 家<extra></extra>"
        else:
            vals = pivw.sum(axis=1, min_count=1) / 1e8
            ns = pivw.notna().sum(axis=1)
            sample_note = "所选区间逐期剔除：每期只加总当期已披露公司，样本家数随期变化"
            custom = [(int(ns[p]), "、".join([c for c in all_companies if pd.isna(pivw.at[p, c])]))
                      for p in window]
            hover = ("%{x}<br>%{y:,.2f} 亿元<br>纳入 %{customdata[0]} 家"
                     "<br>剔除：%{customdata[1]}<extra></extra>")

        q = vals.pct_change().replace([np.inf, -np.inf], np.nan)
        qvals = q.dropna()
        has_q = len(qvals) > 0
        q_tickformat = ".1%"
        q_ticks = None
        q_range = None
        if has_q:
            max_abs = float(qvals.abs().max()) or 0.01
            rough = max_abs / 3
            exp = math.floor(math.log10(rough))
            f = rough / (10 ** exp)
            nf = 1 if f <= 1 else 2 if f <= 2 else 2.5 if f <= 2.5 else 5 if f <= 5 else 10
            step = nf * (10 ** exp)
            top = math.ceil(max_abs / step) * step
            q_ticks = [round(-top + i * step, 10) for i in range(int(2 * top / step) + 1)]
            q_range = [-top, top]
            q_tickformat = ".1%" if max_abs < 0.1 else ".0%"

        figI = make_subplots(specs=[[{"secondary_y": True}]])
        figI.add_trace(go.Bar(
            x=window, y=vals.values, marker_color=colors[item_pick],
            customdata=custom, hovertemplate=hover, name=item_pick), secondary_y=False)
        if has_q:
            figI.add_trace(go.Scatter(
                x=window, y=q.values, mode="lines", name="环比（右轴）",
                line=dict(color="#52514E", width=2, dash="dash"),
                hovertemplate="%{x}<br>环比 %{y:.2%}<extra></extra>"), secondary_y=True)
        tickvals = [p for p in window if p.endswith("Q1")]
        figI.update_layout(
            height=380, template="plotly_white",
            margin=dict(l=10, r=55, t=36, b=10),
            legend=dict(orientation="h", y=1.12, x=0, font=dict(size=11)),
            hovermode="x unified")
        figI.update_xaxes(tickmode="array", tickvals=tickvals,
                          ticktext=[p[:4] for p in tickvals])
        figI.update_yaxes(title_text="亿元", secondary_y=False)
        if has_q:
            figI.update_yaxes(title_text="环比", tickformat=q_tickformat,
                              range=q_range, tickvals=q_ticks, showgrid=False,
                              zeroline=True, zerolinecolor="#D9D8D2", secondary_y=True)
        else:
            figI.update_yaxes(visible=False, secondary_y=True)
        st.plotly_chart(figI, width="stretch")
        last_idx = vals.last_valid_index()
        if last_idx is None:
            st.info("所选范围暂无该项资产数据。")
        else:
            extra = f" · 纳入 {int(ns[last_idx])} 家" if scope == "全部公司合计" else ""
            if pd.notna(q.loc[last_idx]):
                extra += f" · 环比 {q.loc[last_idx] * 100:+.2f}%"
            st.caption(f"{sample_note} ｜ 最新：{last_idx} · {vals[last_idx]:,.2f} 亿元{extra}")

        if scope == "全部公司合计" and basis == "整段剔除":
            rows = []
            for c in excluded:
                miss = [p for p in window if pd.isna(pivw.at[p, c])]
                if len(miss) == len(window):
                    txt = "全部报告期"
                elif len(miss) > 12:
                    have = [p for p in window if p not in miss]
                    txt = f"共 {len(miss)} 期缺失（有值：{'、'.join(have)}）"
                else:
                    txt = "、".join(miss)
                rows.append({"公司": c, "未披露期数": len(miss), "未披露报告期": txt})
            if rows:
                st.markdown("**剔除公司（所选区间整段剔除）**")
                st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)
            else:
                st.caption("所选区间内全部公司均满足平衡样本条件。")
        elif scope == "全部公司合计":
            rows = []
            for p in window:
                miss = [c for c in all_companies if pd.isna(pivw.at[p, c])]
                if miss:
                    rows.append({
                        "报告期": p,
                        "合计(亿元)": round(float(vals[p]), 2) if pd.notna(vals[p]) else None,
                        "环比(%)": round(float(q[p]) * 100, 2) if pd.notna(q[p]) else None,
                        "纳入家数": int(ns[p]),
                        "剔除公司": "、".join(miss),
                    })
            if rows:
                st.markdown("**逐期剔除明细（所选区间）**")
                st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)
            else:
                st.caption("所选区间各期均为全部公司披露，无剔除。")
        else:
            miss = [p for p in window if pd.isna(vals[p])]
            if miss:
                st.caption(f"{scope} 在所选区间未披露报告期：{'、'.join(miss)}")
            else:
                st.caption(f"{scope} 在所选区间各期均有披露。")

        with st.expander("查看逐期数据表"):
            q_pct = [None if pd.isna(v) else round(float(v) * 100, 2) for v in q]
            if scope == "全部公司合计":
                tbl = pd.DataFrame({"报告期": window, "金额(亿元)": vals.round(2).values,
                                    "环比(%)": q_pct,
                                    "纳入家数": ns.values,
                                    "剔除公司": (["固定剔除 " + str(len(excluded)) + " 家"] * len(window)
                                               if basis == "整段剔除"
                                               else ["、".join([c for c in all_companies if pd.isna(pivw.at[p, c])]) or "—"
                                                     for p in window])})
            else:
                tbl = pd.DataFrame({"报告期": window, "金额(亿元)": vals.round(2).values,
                                    "环比(%)": q_pct,
                                    "状态": ["已披露" if pd.notna(v) else "未披露" for v in vals]})
            st.dataframe(tbl, width="stretch", hide_index=True)

        st.markdown("---")
        comp_pick = st.selectbox("单家公司明细", all_companies, index=0, key="asset_company")
        a = assets_all[(assets_all["公司"] == comp_pick)
                       & assets_all["期间"].map(lambda x: period_key(pr[0]) <= period_key(x) <= period_key(pr[1]))].copy()
        a = a.sort_values("期间", key=lambda s: s.map(period_key))
        if a.empty:
            st.info("该区间无数据。")
        else:
            fig = go.Figure()
            for it in items:
                fig.add_trace(go.Bar(x=a["期间"], y=a[it] / 1e8, name=it,
                                     marker_color=colors[it],
                                     hovertemplate=f"{it}<br>%{{x}}<br>%{{y:,.1f}} 亿元<extra></extra>"))
            fig.update_layout(barmode="stack", height=420, template="plotly_white",
                              yaxis_title="亿元", legend=dict(orientation="h", y=1.06),
                              margin=dict(l=10, r=10, t=40, b=10))
            st.plotly_chart(fig, width="stretch")

            fig2 = go.Figure()
            for col, color, name in [
                ("五项占总资产", "#2E5EAA", "五项合计/总资产"),
                ("固定资产占总资产", "#7FA8D9", "固定资产/总资产"),
                ("存货占总资产", "#E9C46A", "存货/总资产"),
            ]:
                fig2.add_trace(go.Scatter(x=a["期间"], y=a[col], name=name, mode="lines+markers",
                                          line=dict(color=color, width=2), marker=dict(size=5)))
            fig2.update_layout(height=360, template="plotly_white", yaxis_tickformat=".0%",
                               yaxis_title="占总资产", legend=dict(orientation="h", y=1.08),
                               margin=dict(l=10, r=10, t=30, b=10))
            st.plotly_chart(fig2, width="stretch")

            latest = a.iloc[-1]
            st.markdown(f"**{comp_pick} · {latest['期间']} 五项资产**")
            cards = st.columns(5)
            for col, it in zip(cards, items):
                col.metric(it, f"{latest[it]/1e8:,.1f} 亿",
                           f"占总资产 {latest[it]/latest['资产总计']*100:.1f}%"
                           if pd.notna(latest["资产总计"]) and latest["资产总计"] else None)
            st.dataframe(a[["公司", "期间"] + items + ["五项合计", "资产总计", "五项占总资产",
                                                      "固定资产占总资产", "存货占总资产",
                                                      "在建工程/固定资产", "五项合计同比"]]
                         .assign(**{c: a[c] / 1e8 for c in items + ["五项合计", "资产总计"]}),
                         width="stretch", hide_index=True)

# ---------------------------------------------------------------- 明细
with tab4:
    st.subheader("指标明细")
    show = df[["公司", "期间"] + IND_NAMES].sort_values(
        ["公司", "期间"], key=lambda s: s.map(lambda x: all_companies.index(x)) if s.name == "公司" else s.map(period_key))
    st.dataframe(show, width="stretch", hide_index=True, height=520)
    c1, c2, c3, c4 = st.columns(4)
    c1.download_button("下载筛选数据（CSV）",
                       show.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"上市猪企指标_{caliber}.csv", mime="text/csv")
    if XLSX.exists():
        c2.download_button("下载完整 Excel 底表", XLSX.read_bytes(),
                           file_name="上市公司财务指标_底表.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if TPL.exists():
        c3.download_button("下载线下趋势模板", TPL.read_bytes(),
                           file_name="上市公司财务指标_趋势与资产_模板.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if HTML.exists():
        c4.download_button("下载离线 HTML 看板", HTML.read_bytes(),
                           file_name="上市公司财务指标_趋势与资产_看板.html", mime="text/html")
    st.caption("线下趋势模板含 01_指标趋势（全部公司）/ 02_五项资产 / 03_指标快照 三个可下拉联动的页签；离线 HTML 看板为单文件，21 家小多图 + 五项资产（含单项资产历年）+ 指标快照，可直接分发。")

# ---------------------------------------------------------------- 口径
with tab5:
    st.subheader("指标计算公式与口径")
    fdf = pd.DataFrame([{
        "分类": IND_META[i]["cat"], "指标": i, "方向": "越高越好" if IND_META[i]["dir"] > 0 else "越低越好",
        "计算公式": IND_META[i]["formula"],
    } for i in IND_NAMES])
    st.dataframe(fdf, width="stretch", hide_index=True)
    st.markdown("""
**通用约定**

- **单季 vs 累计**：原始利润表、现金流量表为年初至今累计数；单季=本期累计−上期累计（Q1 即本期）。
- **年化**：单季×4；累计×4/n（n=1..4）。仅用于流量/存量型比率（ROE、周转率、现金债务覆盖等），纯利润率类不年化。
- **平均存量**：单季取（上季末+本季末）/2；累计取（上年末+本期末）/2；2020 年缺 2019 年报时以 2020Q1 替代期初。
- **毛利率只扣营业成本**；**核心利润率**=(营业收入−营业成本−税金及附加−销售费用−管理费用−研发费用−财务费用)/营业收入，即 (营业总收入−营业总成本)/营业总收入。两者差额=税金及附加+四项期间费用占营业收入比重，不可混用。
- **单位**：比率类为 %，保障/覆盖类为倍，人均创收与人均核心利润为万元/人。
- **有息负债**=短期借款+一年内到期的非流动负债+长期借款+应付债券+租赁负债。
- **员工数**：Q1–Q3 用上年末年报、Q4 用本年末年报；2020Q1–Q3 及 2021 年新上市公司缺上市前年报时，暂用最早一期年报。
- **扣非净利润**：取自各期报告“主要会计数据”表“归属于上市公司股东的扣除非经常性损益的净利润”，与报表累计营收逐期对齐校验。
- **复合增长率**：滚动 3 年；营收用 TTM（最近 4 个单季合计）剔除季节性；不足 3 年留空。
- **累计资产现金回报率**：首个报告期至本期末累计经营现金流净额 ÷ 区间各季末总资产算术平均（2020 年上市的公司从 2020Q1 起算，神农/东瑞从 2021H1 起算）。
- **高学历人才密度**：本科及以上（含博士、硕士/研究生、本科，含“大学”）÷ 员工总数；年报“教育程度”表口径。若“本科及大专”合并披露无法拆分，只计研究生及以上，为下限（立华部分年度、巨星）。
""")
