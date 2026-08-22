# -*- coding: utf-8 -*-
"""
价差日报 & 云南价格日报（蓝色主题，自包含 HTML）
数据源：自动读取桌面（D:\\CC\\Desktop）最新日期的「*涌益咨询日度数据.xlsx」
  - 价差日报：9 省价差同比季节性图，日度/周度/月度 + 公历/农历切换，去除 2021 年
  - 云南价格日报：企业报价（表格：企业/报价/毛猪体重段）+ 散户价格三图
Plotly 自包含 HTML 导出，JS 驱动交互。
"""
import datetime
import os
import re
import glob
import json

import numpy as np
import pandas as pd
import openpyxl
import plotly.graph_objects as go
from lunardate import LunarDate

DESKTOP = r"D:\CC\Desktop"

TARGET9 = ["河南", "江苏", "陕西", "辽宁", "广西", "广东", "湖南", "四川", "云南"]

# 蓝色主题
BLUE = dict(
    bg="#f4f6f9", card="#ffffff", border="#e5e9f0",
    txt="#1f2937", muted="#64748b", accent="#2563eb",
    accent2="#1d4ed8", accent3="#1e40af", hover_bg="#eff6ff",
    red="#dc2626", blue="#2563eb", amber="#d97706",
)

# 同比图年份配色（任务指定，价差日报去除 2021）
YEAR_COLORS = {
    2022: "#7c3aed",  # 紫
    2023: "#2563eb",  # 蓝
    2024: "#000000",  # 黑
    2025: "#0e9f6e",  # 绿
    2026: "#dc2626",  # 红
}
YEAR_ORDER = [2022, 2023, 2024, 2025, 2026]
YN_YEARS = [2023, 2024, 2025, 2026]


# ============ 数据发现 ============
def discover_files():
    """扫描桌面，返回按日期升序排列的 [(date, path), ...]。"""
    pattern = os.path.join(DESKTOP, "*涌益咨询日度数据.xlsx")
    dated = []
    for f in glob.glob(pattern):
        base = os.path.basename(f)
        if base.startswith("~$"):
            continue
        m = re.search(r"(\d{4})年(\d{1,2})月(\d{1,2})日", base)
        if m:
            d = datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            dated.append((d, f))
    dated.sort(key=lambda x: x[0])
    return dated


def latest_file():
    """返回最新文件路径与日期字符串。"""
    dated = discover_files()
    if not dated:
        raise FileNotFoundError(f"桌面上未找到涌益日度数据文件: {DESKTOP}")
    d, path = dated[-1]
    return path, d.strftime("%Y-%m-%d")


# ============ 数据读取 ============
def _recent_slope(s, n=10):
    v = s.dropna().tail(n)
    if len(v) < 3:
        return np.nan
    x = np.arange(len(v), dtype=float)
    return float(np.polyfit(x, v.values, 1)[0])


def load_yunnan(path):
    """读取云南散户标肥价差，返回对齐的时间序列 dict。"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["散户标肥价差"]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    raw_dates = rows[0]
    sub = rows[1]
    prov_rows = {r[0]: r for r in rows[2:] if r and r[0]}
    yn_row = prov_rows["云南"]

    fdates = []
    last = None
    for x in raw_dates:
        if isinstance(x, datetime.datetime):
            last = x
        fdates.append(last)

    by_date = {}
    for i in range(1, len(sub)):
        d = fdates[i]
        metric = sub[i]
        val = yn_row[i]
        if d is None or val is None:
            continue
        key = d.strftime("%Y-%m-%d")
        if key not in by_date:
            by_date[key] = {"biao": None, "s150": None, "s175": None}
        if metric == "市场散户标重猪":
            by_date[key]["biao"] = float(val)
        elif metric == "150公斤左右较标猪":
            by_date[key]["s150"] = float(val)
        elif metric == "175公斤左右较标猪":
            by_date[key]["s175"] = float(val)

    dates = sorted(by_date.keys())
    return {
        "dates": dates,
        "biao": [by_date[d]["biao"] for d in dates],
        "s150": [by_date[d]["s150"] for d in dates],
        "s175": [by_date[d]["s175"] for d in dates],
    }


def load_provinces(path):
    """读取 9 省价差（去除 2021），返回 (province_data, latest_date_str)。"""
    df = pd.read_excel(path, sheet_name="各省份均价")
    df["日期"] = pd.to_datetime(df["日期"])
    df = df.sort_values("日期").reset_index(drop=True)

    province_cols = [c for c in df.columns if c not in ("日期", "全国均价")]
    missing = [p for p in TARGET9 if p not in province_cols]
    if missing:
        raise ValueError(f"缺少目标省份列: {missing}")

    df = df[df["日期"] >= pd.Timestamp("2022-01-01")].reset_index(drop=True)
    national = df[province_cols].mean(axis=1, skipna=True)

    province_data = []
    for p in TARGET9:
        spread = df[p] - national
        clean = spread.dropna()
        latest = float(clean.iloc[-1])
        latest_date = df.loc[clean.index[-1], "日期"]
        hist_mean = float(clean.mean())

        slope = _recent_slope(spread)
        if pd.isna(slope):
            trend = "数据不足"
        elif slope > 0.02:
            trend = "走扩"
        elif slope < -0.02:
            trend = "收窄"
        else:
            trend = "震荡"

        pts = []
        for idx, v in spread.dropna().items():
            d = df.loc[idx, "日期"]
            ld = LunarDate.fromSolarDate(d.year, d.month, d.day)
            week = d.isocalendar()[1]
            pts.append([d.year, d.month, d.day, ld.month, ld.day, week,
                        round(float(v), 4), d.strftime("%Y-%m-%d")])

        province_data.append({
            "name": p,
            "latest_date": latest_date.strftime("%Y-%m-%d"),
            "latest": round(latest, 4),
            "hist_mean": round(hist_mean, 4),
            "trend": trend,
            "points": pts,
        })

    return province_data, latest_date.strftime("%Y-%m-%d")


# ============ HTML 通用 ============
_PLOTLY_JS = None


def _plotly_js():
    global _PLOTLY_JS
    if _PLOTLY_JS is None:
        import plotly.offline as pyo
        _PLOTLY_JS = pyo.get_plotlyjs()
    return _PLOTLY_JS


_CSS = """
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:"Microsoft YaHei","PingFang SC",-apple-system,sans-serif;background:#f4f6f9;color:#1f2937;line-height:1.55;}
.page{max-width:1120px;margin:0 auto;padding:28px 30px 52px;}
.title{font-size:24px;font-weight:800;color:#1f2937;letter-spacing:.5px;padding:0 18px;}
.sub{color:#64748b;font-size:13px;margin-top:6px;padding:0 18px;}
.controls{display:flex;align-items:center;gap:14px;flex-wrap:wrap;margin-bottom:18px;padding:0 18px;}
.ctrl{display:flex;align-items:center;gap:8px;}
.ctrl label{font-size:13px;color:#64748b;}
.ctrl select{background:#fff;color:#1f2937;border:1px solid #cbd5e1;border-radius:8px;padding:8px 12px;font-size:14px;font-weight:600;cursor:pointer;min-width:168px;}
.seg{display:inline-flex;border:1px solid #cbd5e1;border-radius:9px;overflow:hidden;background:#fff;}
.seg-btn{border:none;background:#fff;color:#475569;padding:8px 16px;font-size:13px;font-weight:600;cursor:pointer;transition:all .15s;}
.seg-btn.active{background:#2563eb;color:#fff;}
.section{margin-top:26px;}
.section-title{font-size:17px;font-weight:700;display:flex;align-items:center;gap:9px;margin-bottom:14px;color:#111827;padding:0 18px;}
.section-title::before{content:"";width:5px;height:20px;background:#2563eb;border-radius:3px;}
.card{background:#fff;border:1px solid #e5e9f0;border-radius:14px;padding:18px;box-shadow:0 1px 4px rgba(15,23,42,.05);}
.kpis{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:18px;}
.kpi{background:#fff;border:1px solid #e5e9f0;border-radius:12px;padding:12px 18px;min-width:118px;box-shadow:0 1px 3px rgba(15,23,42,.04);}
.kpi span{display:block;color:#64748b;font-size:12px;margin-bottom:5px;}
.kpi b{font-size:20px;color:#111827;font-weight:700;}
.kpi b.small{font-size:15px;}
table.quote{width:100%;border-collapse:collapse;font-size:15px;}
table.quote th,table.quote td{border:1px solid #e5e9f0;padding:12px 16px;text-align:left;}
table.quote thead th{background:#2563eb;color:#fff;font-weight:600;letter-spacing:.5px;}
table.quote tbody tr:nth-child(even){background:#f8fafc;}
table.quote tbody tr:hover{background:#eff6ff;}
table.quote td:first-child{font-weight:600;color:#111827;}
.btn{background:#2563eb;color:#fff;border:none;padding:10px 22px;border-radius:8px;font-weight:700;cursor:pointer;font-size:14px;margin-top:10px;transition:background .15s;}
.btn:hover{background:#1d4ed8;}
.btn.ghost{background:#fff;color:#2563eb;border:1px solid #2563eb;}
.btn.ghost:hover{background:#eff6ff;}
.edit-wrap{margin-top:14px;}
.edit-wrap textarea{width:100%;min-height:150px;background:#f8fafc;color:#1f2937;border:1px solid #e2e8f0;border-radius:10px;padding:12px;font-family:Consolas,"Microsoft YaHei",monospace;font-size:14px;line-height:1.9;resize:vertical;}
.chart{margin-top:14px;}
.grid3{display:grid;grid-template-columns:repeat(3,1fr);gap:16px;}
.cell h3{font-size:14px;margin-bottom:6px;}
.cell .info{font-size:12px;color:#64748b;margin-bottom:6px;}
.cell .info b{color:#111827;}
.legend{margin-bottom:12px;font-size:12px;color:#64748b;padding:0 18px;}
.legend i{display:inline-block;width:14px;height:3px;vertical-align:middle;margin-right:5px;border-radius:2px;}
.wide{color:#dc2626;} .narrow{color:#2563eb;} .flat{color:#64748b;}
@media(max-width:1000px){.grid3{grid-template-columns:repeat(2,1fr);}}
@media(max-width:640px){.grid3{grid-template-columns:1fr;}}
"""


# ============ 价差日报（任务二，JS 驱动） ============
_PRICE_JS = r"""
const YEAR_COLORS = {2022:"#7c3aed",2023:"#2563eb",2024:"#000000",2025:"#0e9f6e",2026:"#dc2626"};
const YEAR_ORDER = [2022,2023,2024,2025,2026];
const PRICE = __PRICE_DATA__;
let freq='daily', cal='solar';

const SOLAR_TICKS=[1,32,63,94,125,156,187,218,249,280,311,342];
const SOLAR_LABELS=['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'];
const LUNAR_LABELS=['正月','二月','三月','四月','五月','六月','七月','八月','九月','十月','冬月','腊月'];

function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
function selectedDate(){return document.getElementById('dateSel').value;}

function aggregate(points, freq, cal, cutoff){
  const groups={};
  points.forEach(function(p){
    const y=p[0], sm=p[1], sd=p[2], lm=p[3], ld=p[4], w=p[5], s=p[6], d=p[7];
    if(s==null) return;
    if(cutoff && d>cutoff) return;
    let key;
    if(freq==='daily') key=y+'-'+d;
    else if(freq==='weekly') key=y+'-W'+w;
    else key=y+'-'+sm;
    if(!groups[key]) groups[key]={y:y,sum:0,n:0,sm:sm,sd:sd,lm:lm,ld:ld,d:d};
    const g=groups[key];
    g.sum+=s; g.n+=1;
    if(d<g.d){g.sm=sm;g.sd=sd;g.lm=lm;g.ld=ld;g.d=d;}
  });
  const byYear={};
  Object.keys(groups).forEach(function(k){
    const g=groups[k];
    const val=g.sum/g.n;
    let x;
    if(freq==='monthly') x=(cal==='solar'?(g.sm-1)*31+1:(g.lm-1)*31+1);
    else x=(cal==='solar'?(g.sm-1)*31+g.sd:(g.lm-1)*31+g.ld);
    if(!byYear[g.y]) byYear[g.y]=[];
    byYear[g.y].push({x:x,y:+val.toFixed(4),d:g.d});
  });
  Object.keys(byYear).forEach(function(y){byYear[y].sort(function(a,b){return a.x-b.x;});});
  return byYear;
}

function tracesFor(prov, cutoff){
  const byYear=aggregate(prov.points, freq, cal, cutoff);
  const traces=[];
  YEAR_ORDER.forEach(function(y){
    if(byYear[y]){
      traces.push({
        x:byYear[y].map(function(p){return p.x;}),
        y:byYear[y].map(function(p){return p.y;}),
        mode:'lines', name:y+'年',
        line:{color:YEAR_COLORS[y], width:1.8},
        customdata:byYear[y].map(function(p){return p.d;}),
        hovertemplate:y+'年<br>公历 %{customdata}<br>价差 %{y:.2f} 元/kg<extra></extra>'
      });
    }
  });
  return traces;
}

function chartLayout(h){
  return {
    template:'plotly_white', paper_bgcolor:'#ffffff', plot_bgcolor:'#ffffff',
    font:{family:'Microsoft YaHei, PingFang SC, sans-serif', color:'#1f2937', size:11},
    margin:{l:42,r:10,t:8,b:34}, height:h,
    hoverlabel:{font:{family:'Microsoft YaHei'}},
    showlegend:false,
    xaxis:{tickvals:cal==='solar'?SOLAR_TICKS:SOLAR_TICKS,
           ticktext:cal==='solar'?SOLAR_LABELS:LUNAR_LABELS,
           range:[1,372], showgrid:false, zeroline:false},
    yaxis:{gridcolor:'#eef2f7', zeroline:false},
    shapes:[{type:'line',x0:1,x1:372,y0:0,y1:0,line:{color:'#cbd5e1',width:1,dash:'dash'}}]
  };
}

function latestOf(prov, cutoff){
  let v=null, d='';
  prov.points.forEach(function(p){
    if(p[6]==null) return;
    if(cutoff && p[7]>cutoff) return;
    v=p[6]; d=p[7];
  });
  return {v:v,d:d};
}

function render(){
  const cutoff=selectedDate();
  PRICE.provinces.forEach(function(prov, i){
    const lv=latestOf(prov, cutoff);
    document.getElementById('latest_'+i).textContent=(lv.v==null)?'—':lv.v.toFixed(4);
    Plotly.react('chart_'+i, tracesFor(prov, cutoff), chartLayout(260), {responsive:true});
  });
}

function setFreq(f){
  freq=f;
  document.querySelectorAll('#freqSeg .seg-btn').forEach(function(b){b.classList.toggle('active', b.dataset.freq===f);});
  render();
}
function setCal(c){
  cal=c;
  document.querySelectorAll('#calSeg .seg-btn').forEach(function(b){b.classList.toggle('active', b.dataset.cal===c);});
  render();
}

function init(){
  const sel=document.getElementById('dateSel');
  const dates=PRICE.dates;
  for(let i=dates.length-1;i>=0;i--){
    const o=document.createElement('option');
    o.value=dates[i]; o.textContent=dates[i];
    sel.appendChild(o);
  }
  sel.value=dates[dates.length-1];
  document.querySelectorAll('#freqSeg .seg-btn').forEach(function(b){
    b.addEventListener('click', function(){setFreq(b.dataset.freq);});
  });
  document.querySelectorAll('#calSeg .seg-btn').forEach(function(b){
    b.addEventListener('click', function(){setCal(b.dataset.cal);});
  });
  render();
}
window.addEventListener('resize', render);
document.addEventListener('DOMContentLoaded', init);
"""


def build_price_report_html(province_data, latest_date):
    # 日期下拉：用所有省份点中出现的日期（全国均价口径的交易日）
    all_dates = set()
    for p in province_data:
        for pt in p["points"]:
            all_dates.add(pt[7])
    dates = sorted(all_dates)

    payload = {
        "dates": dates,
        "provinces": [
            {
                "name": p["name"],
                "hist_mean": p["hist_mean"],
                "trend": p["trend"],
                "points": p["points"],
            }
            for p in province_data
        ],
    }
    js = _PRICE_JS.replace("__PRICE_DATA__", json.dumps(payload, ensure_ascii=False))

    legend = "".join(
        f'<span style="margin-right:12px;">'
        f'<i style="background:{YEAR_COLORS[y]};"></i>{y}年</span>'
        for y in YEAR_ORDER
    )

    cards = []
    for i, p in enumerate(province_data):
        t = p["trend"]
        tcls = "wide" if t == "走扩" else ("narrow" if t == "收窄" else "flat")
        cards.append(f"""
<div class="cell card">
  <h3>{p['name']}</h3>
  <div class="info">当日价差 <b id="latest_{i}">—</b> 元/kg ｜ 历史均值 <b>{p['hist_mean']}</b><br>
  近期走势 <b class="{tcls}">{t}</b></div>
  <div class="chart" id="chart_{i}"></div>
</div>""")

    body = f"""
<div class="controls">
  <div class="ctrl"><label>数据日期</label><select id="dateSel" onchange="render()"></select></div>
  <div class="ctrl"><label>频率</label><div class="seg" id="freqSeg">
    <button class="seg-btn active" data-freq="daily">日度</button>
    <button class="seg-btn" data-freq="weekly">周度</button>
    <button class="seg-btn" data-freq="monthly">月度</button>
  </div></div>
  <div class="ctrl"><label>日历</label><div class="seg" id="calSeg">
    <button class="seg-btn active" data-cal="solar">公历</button>
    <button class="seg-btn" data-cal="lunar">农历</button>
  </div></div>
</div>
<h1 class="title">价差日报</h1>
<div class="sub">数据截止 {latest_date} ｜ 价差 = 省份均价 − 全国均价（各省份简单算术平均）｜ 横轴为日期，每年一条线</div>
<div class="section">
  <div class="legend">{legend}</div>
  <div class="grid3">{''.join(cards)}</div>
</div>"""
    return f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>价差日报</title>
<script>{_plotly_js()}</script>
<style>{_CSS}</style>
</head><body>
<div class="page">{body}</div>
<script>{js}</script>
</body></html>"""


# ============ 云南价格日报（任务一，JS 驱动） ============
_YUNNAN_JS = r"""
const YEAR_COLORS = {2022:"#7c3aed",2023:"#2563eb",2024:"#000000",2025:"#0e9f6e",2026:"#dc2626"};
const YN_YEARS = [2023,2024,2025,2026];
const YN = __YN_DATA__;

const SOLAR_TICKS=[1,32,63,94,125,156,187,218,249,280,311,342];
const SOLAR_LABELS=['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'];
const LUNAR_LABELS=['正月','二月','三月','四月','五月','六月','七月','八月','九月','十月','冬月','腊月'];

function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
function lastVal(arr){for(let i=arr.length-1;i>=0;i--){if(arr[i]!=null)return arr[i];}return null;}
function mdSolar(d){const p=d.split('-').map(Number);return (p[1]-1)*31+p[2];}
function fmt(v){return (v==null)?'—':v.toFixed(2);}

function buildSeries(selIdx){
  const dates=YN.dates.slice(0,selIdx+1);
  const biao=YN.biao.slice(0,selIdx+1);
  const s150=YN.s150.slice(0,selIdx+1);
  const s175=YN.s175.slice(0,selIdx+1);
  const p150=biao.map((v,i)=>(v==null||s150[i]==null)?null:+(v+s150[i]).toFixed(4));
  const p175=biao.map((v,i)=>(v==null||s175[i]==null)?null:+(v+s175[i]).toFixed(4));
  return {dates,biao,s150,s175,p150,p175};
}

function seasonalTraces(dates,series,years){
  const byYear={};
  dates.forEach((d,i)=>{
    const v=series[i];
    if(v==null)return;
    const y=d.slice(0,4);
    if(!byYear[y])byYear[y]={x:[],y:[],d:[]};
    byYear[y].x.push(mdSolar(d));
    byYear[y].y.push(v);
    byYear[y].d.push(d);
  });
  const traces=[];
  years.forEach(y=>{
    if(byYear[y]){
      traces.push({x:byYear[y].x,y:byYear[y].y,mode:'lines',name:y+'年',
        line:{color:YEAR_COLORS[y],width:2.2},
        customdata:byYear[y].d,
        hovertemplate:y+'年<br>公历 %{customdata}<br>价差 %{y:.2f} 元/kg<extra></extra>'});
    }
  });
  return traces;
}

const LAYOUT_BASE={template:'plotly_white',paper_bgcolor:'#ffffff',plot_bgcolor:'#ffffff',
  font:{family:'Microsoft YaHei, PingFang SC, sans-serif',color:'#1f2937',size:12},
  margin:{l:50,r:16,t:14,b:36},hoverlabel:{font:{family:'Microsoft YaHei'}}};

function absLayout(t){return Object.assign({},LAYOUT_BASE,{title:{text:t,font:{size:14}},height:370,legend:{orientation:'h',y:1.08,x:0},
  xaxis:{showgrid:false,tickformat:'%Y年%m月'},yaxis:{gridcolor:'#eef2f7',title:'元/kg'}});}
function seasonLayout(t){return Object.assign({},LAYOUT_BASE,{title:{text:t,font:{size:14}},height:320,legend:{orientation:'h',y:1.08,x:0},
  xaxis:{tickvals:SOLAR_TICKS,ticktext:SOLAR_LABELS,range:[1,372],showgrid:false},yaxis:{gridcolor:'#eef2f7',title:'价差(元/kg)'}});}

function render(){
  const sel=document.getElementById('dateSel');
  const selIdx=YN.dates.indexOf(sel.value);
  const s=buildSeries(selIdx);
  const dts=s.dates.map(d=>new Date(d));
  const lv={biao:lastVal(s.biao),p150:lastVal(s.p150),p175:lastVal(s.p175),s150:lastVal(s.s150),s175:lastVal(s.s175)};

  document.getElementById('kpiBiao').textContent=fmt(lv.biao);
  document.getElementById('kpi150').textContent=fmt(lv.p150);
  document.getElementById('kpi175').textContent=fmt(lv.p175);
  document.getElementById('kpiS150').textContent=fmt(lv.s150);
  document.getElementById('kpiS175').textContent=fmt(lv.s175);
  document.getElementById('kpiDate').textContent=sel.value;

  const tr1=[
    {x:dts,y:s.biao,mode:'lines',name:'散户标猪价格',line:{color:'#2563eb',width:2.4},hovertemplate:'散户标猪<br>%{x|%Y年%m月%d日}<br>%{y:.2f} 元/kg<extra></extra>'},
    {x:dts,y:s.p150,mode:'lines',name:'150kg猪价格',line:{color:'#d97706',width:2.4},hovertemplate:'150kg猪<br>%{x|%Y年%m月%d日}<br>%{y:.2f} 元/kg<extra></extra>'},
    {x:dts,y:s.p175,mode:'lines',name:'175kg猪价格',line:{color:'#dc2626',width:2.4},hovertemplate:'175kg猪<br>%{x|%Y年%m月%d日}<br>%{y:.2f} 元/kg<extra></extra>'}
  ];
  Plotly.react('chart1',tr1,absLayout('云南散户标猪 / 150kg / 175kg 价格走势'),{responsive:true});

  Plotly.react('chart2',seasonalTraces(s.dates,s.s150,YN_YEARS),seasonLayout('150kg 猪与标猪价差 · 季节性（同比）'),{responsive:true});
  Plotly.react('chart3',seasonalTraces(s.dates,s.s175,YN_YEARS),seasonLayout('175kg 猪与标猪价差 · 季节性（同比）'),{responsive:true});
}

function parseQuotes(txt){
  const rows=[];
  txt.split('\n').forEach(function(line){
    const t=line.trim();
    if(!t)return;
    const m=t.match(/^(\S+)\s+(\S+)\s+(.+)$/);
    if(m)rows.push({name:m[1],quote:m[2],weight:m[3]});
    else{
      const m2=t.match(/^(\S+)\s+(.+)$/);
      if(m2)rows.push({name:m2[1],quote:m2[2],weight:''});
      else rows.push({name:t,quote:'',weight:''});
    }
  });
  return rows;
}
function updateQuotes(){
  const rows=parseQuotes(document.getElementById('quoteTxt').value);
  const tb=document.getElementById('quoteTableBody');
  tb.innerHTML='';
  rows.forEach(function(r){
    const tr=document.createElement('tr');
    tr.innerHTML='<td>'+esc(r.name)+'</td><td>'+esc(r.weight)+'</td><td>'+esc(r.quote)+'</td>';
    tb.appendChild(tr);
  });
}
function toggleEdit(){
  const w=document.getElementById('editWrap');
  w.style.display=(w.style.display==='none')?'block':'none';
}

function init(){
  const sel=document.getElementById('dateSel');
  for(let i=YN.dates.length-1;i>=0;i--){
    const o=document.createElement('option');
    o.value=YN.dates[i];o.textContent=YN.dates[i];
    sel.appendChild(o);
  }
  sel.value=YN.dates[YN.dates.length-1];
  updateQuotes();
  render();
}
window.addEventListener('resize',render);
document.addEventListener('DOMContentLoaded',init);
"""

_DEFAULT_QUOTES = (
    "神农      11.3    125-130kg\n"
    "双胞胎    11.2    120-135kg\n"
    "温氏      11.3    110-130kg\n"
    "德康      11.1    120-135kg\n"
    "正大      11.2-11.3\n"
    "力源      11.3    120-130kg\n"
    "主流偏强 0.1"
)


def build_yunnan_report_html(yunnan, latest_date):
    data_json = json.dumps(yunnan, ensure_ascii=False)
    js = _YUNNAN_JS.replace("__YN_DATA__", data_json)

    body = f"""
<div class="controls">
  <div class="ctrl"><label>数据日期</label><select id="dateSel" onchange="render()"></select></div>
</div>
<h1 class="title">云南价格日报</h1>
<div class="sub">数据截止 {latest_date} ｜ 企业报价 + 云南散户标猪 / 150kg / 175kg 价格与肥标价差</div>

<div class="section">
  <div class="section-title">企业报价</div>
  <div class="card">
    <table class="quote">
      <thead><tr><th style="width:34%;">企业</th><th style="width:32%;">毛猪体重段</th><th>报价（元/kg）</th></tr></thead>
      <tbody id="quoteTableBody"></tbody>
    </table>
    <button class="btn ghost" onclick="toggleEdit()">编辑报价</button>
    <div class="edit-wrap" id="editWrap" style="display:none;">
      <textarea id="quoteTxt">{_DEFAULT_QUOTES}</textarea>
      <button class="btn" onclick="updateQuotes()">更新</button>
    </div>
  </div>
</div>

<div class="section">
  <div class="section-title">散户价格</div>
  <div class="kpis">
    <div class="kpi"><span>散户标猪价(元/kg)</span><b id="kpiBiao">—</b></div>
    <div class="kpi"><span>150kg价(元/kg)</span><b id="kpi150">—</b></div>
    <div class="kpi"><span>175kg价(元/kg)</span><b id="kpi175">—</b></div>
    <div class="kpi"><span>150kg价差(元/kg)</span><b id="kpiS150">—</b></div>
    <div class="kpi"><span>175kg价差(元/kg)</span><b id="kpiS175">—</b></div>
    <div class="kpi"><span>最新日期</span><b class="small" id="kpiDate">—</b></div>
  </div>
  <div class="card"><div class="chart" id="chart1"></div></div>
  <div class="card chart"><div class="chart" id="chart2"></div></div>
  <div class="card chart"><div class="chart" id="chart3"></div></div>
</div>"""

    return f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>云南价格日报</title>
<script>{_plotly_js()}</script>
<style>{_CSS}</style>
</head><body>
<div class="page">{body}</div>
<script>{js}</script>
</body></html>"""


# ============ 主流程 ============
if __name__ == "__main__":
    OUT = r"C:\Users\CC\test-claude"
    path, latest_date = latest_file()
    print(f"数据文件: {path}")

    province_data, p_latest = load_provinces(path)
    yunnan = load_yunnan(path)
    yn_latest = yunnan["dates"][-1] if yunnan["dates"] else latest_date

    price_html = build_price_report_html(province_data, p_latest)
    with open(os.path.join(OUT, "价差日报.html"), "w", encoding="utf-8") as f:
        f.write(price_html)

    yunnan_html = build_yunnan_report_html(yunnan, yn_latest)
    with open(os.path.join(OUT, "云南日报.html"), "w", encoding="utf-8") as f:
        f.write(yunnan_html)

    print("=== 价差日报：各省价差 ===")
    for p in province_data:
        print(f"{p['name']:4s} 当日价差={p['latest']:>7} 历史均值={p['hist_mean']:>7} 走势={p['trend']}")
    print("\n=== 云南价格日报 ===")
    print("最新日期:", yn_latest, "| 数据点:", len(yunnan["dates"]))
    print("已输出: 价差日报.html / 云南日报.html")
