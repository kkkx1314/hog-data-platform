# -*- coding: utf-8 -*-
"""生成公开静态页 docs/feed/index.html（不依赖 Streamlit / Plotly）。

数据来自 data/feed_dashboard.json。页面是一个自包含的单文件：
筛选（口径 × 月度/季度 × 产量/环比/同比 × 品类 × 年份）、KPI、
同期折线、当期柱图、明细表和 CSV 下载全部由原生 JS + SVG 完成，
没有外部 CDN 依赖，断网/弱网环境也能打开。

用法：python build_feed_public_site.py
发布：同步数据到GitHub.bat 会把 docs/feed/index.html 一起推到 gh-pages。
"""
import json
import os

import feed_dashboard as fd

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'docs', 'feed')
OUT_PATH = os.path.join(OUT_DIR, 'index.html')

TEMPLATE = r'''<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="robots" content="noindex,nofollow,noarchive,nosnippet,noimageindex">
<title>饲料产量看板 · 云南神农肉业</title>
<style>
:root{
  --plane:#f9f9f7; --surface:#fcfcfb;
  --ink-1:#0b0b0b; --ink-2:#52514e; --ink-3:#898781;
  --grid:#e1e0d9; --axis:#c3c2b7; --hair:rgba(11,11,11,.10);
  --accent:#0f8f45;
  --y26:#d2352f; --y25:#0f8f45; --y24:#0b0b0b;
  --y23:#d63384; --y22:#1565c0; --y21:#b8860b; --y20:#6a1b9a;
  --bar-level:#d66b66; --bar-level-latest:#8f1712;
  --red:#d2352f; --blue:#2a78d6; --zero:#898781;
  --shadow:0 1px 2px rgba(11,11,11,.04),0 4px 16px rgba(11,11,11,.05);
}
@media (prefers-color-scheme:dark){
  :root:where(:not([data-theme="light"])){
    --plane:#0d0d0d; --surface:#1a1a19;
    --ink-1:#fff; --ink-2:#c3c2b7; --ink-3:#898781;
    --grid:#2c2c2a; --axis:#383835; --hair:rgba(255,255,255,.10);
    --accent:#35c46e;
    --y26:#ef5b52; --y25:#35c46e; --y24:#f0f0f0;
    --y23:#ff79b0; --y22:#64b5f6; --y21:#ffd54f; --y20:#c084fc;
    --bar-level:#b8534e; --bar-level-latest:#ff8a80;
    --red:#ef5b52; --blue:#3987e5; --zero:#898781;
    --shadow:0 1px 2px rgba(0,0,0,.3),0 4px 16px rgba(0,0,0,.35);
  }
}
:root[data-theme="dark"]{
  --plane:#0d0d0d; --surface:#1a1a19;
  --ink-1:#fff; --ink-2:#c3c2b7; --ink-3:#898781;
  --grid:#2c2c2a; --axis:#383835; --hair:rgba(255,255,255,.10);
  --accent:#35c46e;
  --y26:#ef5b52; --y25:#35c46e; --y24:#f0f0f0;
  --y23:#ff79b0; --y22:#64b5f6; --y21:#ffd54f; --y20:#c084fc;
  --bar-level:#b8534e; --bar-level-latest:#ff8a80;
  --red:#ef5b52; --blue:#3987e5; --zero:#898781;
  --shadow:0 1px 2px rgba(0,0,0,.3),0 4px 16px rgba(0,0,0,.35);
}
*{box-sizing:border-box}
html,body{margin:0;padding:0}
body{
  background:var(--plane); color:var(--ink-1);
  font-family:system-ui,-apple-system,"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif;
  font-size:14px; line-height:1.5; -webkit-font-smoothing:antialiased;
}
.wrap{max-width:1440px;margin:0 auto;padding:24px 20px 70px}
header.top{display:flex;flex-wrap:wrap;align-items:flex-end;gap:16px;margin-bottom:18px}
h1{font-size:22px;font-weight:700;margin:0;letter-spacing:-.01em}
.sub{color:var(--ink-3);font-size:13px;margin-top:3px;font-variant-numeric:tabular-nums}
.spacer{flex:1}
.themebtn{
  background:var(--surface);border:1px solid var(--hair);color:var(--ink-2);
  border-radius:8px;padding:6px 12px;font:inherit;font-size:13px;cursor:pointer;
}
.themebtn:hover{border-color:var(--axis)}

.filters{
  display:flex;flex-wrap:wrap;gap:18px;align-items:center;
  background:var(--surface);border:1px solid var(--hair);border-radius:12px;
  padding:12px 16px;margin-bottom:18px;box-shadow:var(--shadow);
  position:sticky;top:0;z-index:20;
}
.fgroup{display:flex;align-items:center;gap:8px}
.flabel{color:var(--ink-3);font-size:12px;font-weight:600;letter-spacing:.02em}
.seg{display:flex;background:var(--plane);border:1px solid var(--hair);border-radius:8px;padding:2px;flex-wrap:wrap}
.seg button{
  background:transparent;border:0;color:var(--ink-2);font:inherit;font-size:13px;
  padding:4px 11px;border-radius:6px;cursor:pointer;white-space:nowrap;
}
.seg button[aria-pressed="true"]{background:var(--surface);color:var(--ink-1);font-weight:600;box-shadow:0 1px 2px rgba(0,0,0,.08)}
.seg button:hover:not([aria-pressed="true"]){color:var(--ink-1)}
.seg.years button{padding:4px 8px;font-variant-numeric:tabular-nums}
.sel{
  background:var(--plane);border:1px solid var(--hair);color:var(--ink-1);
  border-radius:8px;padding:5px 8px;font:inherit;font-size:13px;max-width:220px;
}
.hint{color:var(--ink-3);font-size:12px;font-variant-numeric:tabular-nums}

.grid{display:grid;gap:14px}
.kpis{grid-template-columns:repeat(auto-fit,minmax(200px,1fr));margin-bottom:18px}
.card{
  background:var(--surface);border:1px solid var(--hair);border-radius:12px;
  padding:16px 18px;box-shadow:var(--shadow);margin-bottom:16px;
}
.card h2{font-size:15px;font-weight:650;margin:0 0 4px}
.chead{display:flex;align-items:flex-start;gap:10px;margin-bottom:4px}
.chead h2{flex:1}
.tile{display:flex;flex-direction:column;gap:2px;min-height:92px}
.tile .k{color:var(--ink-3);font-size:12px;font-weight:600}
.tile .v{font-size:27px;font-weight:700;letter-spacing:-.02em;line-height:1.2}
.tile .v .u{font-size:13px;font-weight:500;color:var(--ink-3);margin-left:3px;letter-spacing:0}
.tile .d{font-size:12px;font-weight:600;font-variant-numeric:tabular-nums}
.tile .n{font-size:12px;color:var(--ink-3);font-variant-numeric:tabular-nums}
.up{color:var(--red)} .dn{color:var(--blue)} .flat{color:var(--zero)}
.legend{display:flex;flex-wrap:wrap;gap:16px;margin:8px 0 2px}
.legend .li{display:flex;align-items:center;gap:6px;font-size:12px;color:var(--ink-2)}
.legend .sw{width:16px;height:0;border-top:3px solid;flex:none}
svg.chart{display:block;width:100%;overflow:visible}
.tt{
  position:fixed;pointer-events:none;z-index:60;opacity:0;transition:opacity .08s;
  background:var(--surface);border:1px solid var(--hair);border-radius:8px;
  padding:8px 10px;font-size:12px;color:var(--ink-1);box-shadow:var(--shadow);
  font-variant-numeric:tabular-nums;max-width:280px;
}
.tt .tth{font-weight:650;margin-bottom:5px;font-size:12px}
.tt .ttr{display:flex;align-items:center;gap:7px;white-space:nowrap}
.tt .ttr .sw{width:10px;height:0;border-top:3px solid;flex:none}
.tt .ttr .nm{color:var(--ink-2);flex:1}
.tt .ttr .vl{font-weight:600;margin-left:10px}
table{width:100%;border-collapse:collapse;font-size:12.5px;font-variant-numeric:tabular-nums}
th,td{padding:7px 9px;text-align:right;border-bottom:1px solid var(--grid);white-space:nowrap}
th{color:var(--ink-3);font-weight:600;font-size:11.5px;position:sticky;top:0;background:var(--surface)}
th.l,td.l{text-align:left}
tbody tr:hover{background:color-mix(in srgb,var(--ink-1) 4%,transparent)}
.tblwrap{max-height:420px;overflow:auto;margin-top:8px;border:1px solid var(--hair);border-radius:8px}
.minibtn{
  background:var(--plane);border:1px solid var(--hair);color:var(--ink-2);
  border-radius:8px;padding:4px 10px;font:inherit;font-size:12px;cursor:pointer;white-space:nowrap;
}
.minibtn:hover{color:var(--ink-1);border-color:var(--axis)}
.note-card{font-size:13px;color:var(--ink-2);line-height:1.75}
.note-card h2{margin-bottom:8px}
.note-card b{color:var(--ink-1)}
.src{color:var(--ink-3);font-size:12px;margin-top:10px;font-variant-numeric:tabular-nums}
@media(max-width:900px){
  .filters{position:static}
  .wrap{padding:16px 12px 56px}
  .kpis{grid-template-columns:repeat(2,minmax(0,1fr))}
}
@media(max-width:560px){
  .filters{gap:12px;padding:10px 12px}
  .fgroup{width:100%}
  .seg{width:100%}
  .seg button{flex:1}
  .sel{max-width:none;width:100%}
  .tile .v{font-size:23px}
}
/* 访问密码遮罩 */
.lock{position:fixed;inset:0;z-index:9999;display:flex;align-items:center;justify-content:center;background:var(--plane);padding:20px}
.lock[hidden]{display:none}
.lockcard{background:var(--surface);border:1px solid var(--grid);border-radius:16px;padding:28px 26px;width:min(360px,100%);box-shadow:var(--shadow);display:flex;flex-direction:column;gap:10px}
.lockcard h2{font-size:18px;margin:0}
.lockcard p{font-size:13px;color:var(--ink-2);margin:0 0 6px;line-height:1.5}
.lockcard input{font:inherit;padding:10px 12px;border:1px solid var(--axis);border-radius:10px;background:var(--plane);color:var(--ink-1)}
.lockcard input:focus{outline:2px solid var(--accent);outline-offset:1px}
.lockcard button{font:inherit;font-weight:700;padding:10px 12px;border:0;border-radius:10px;background:var(--accent);color:#fff;cursor:pointer}
.lockerr{font-size:12px;color:var(--red);min-height:16px}
</style>
</head>
<body>
<div class="wrap">
<header class="top">
  <div>
    <h1>🌾 云南饲料产量看板</h1>
    <div class="sub" id="subtitle"></div>
  </div>
  <div class="spacer"></div>
  <button class="themebtn" id="themeBtn" type="button">切换深/浅色</button>
</header>

<div class="filters">
  <div class="fgroup"><span class="flabel">饲料口径</span><div class="seg" id="segScope"></div></div>
  <div class="fgroup"><span class="flabel">频率</span><div class="seg" id="segFreq"></div></div>
  <div class="fgroup"><span class="flabel">指标</span><div class="seg" id="segMetric"></div></div>
  <div class="fgroup"><span class="flabel">品类</span><select class="sel" id="selCat"></select></div>
  <div class="fgroup"><span class="flabel">年份</span><div class="seg years" id="segYears"></div></div>
</div>

<div class="grid kpis" id="kpis"></div>

<div class="card">
  <div class="chead"><h2 id="lineTitle"></h2></div>
  <div class="legend" id="legend"></div>
  <div class="hint" id="lineHint"></div>
  <div id="lineWrap"></div>
</div>

<div class="card">
  <div class="chead"><h2 id="barTitle"></h2></div>
  <div class="hint" id="barHint"></div>
  <div id="barWrap"></div>
</div>

<div class="card">
  <div class="chead"><h2 id="tableTitle"></h2><button class="minibtn" id="csvBtn" type="button">下载 CSV</button></div>
  <div class="tblwrap"><table id="tbl"></table></div>
</div>

<div class="card note-card">
  <h2>口径说明</h2>
  <div>
    <div>· <b>仔猪饲料</b> = 哺乳仔猪饲料 + 保育饲料；<b>母猪饲料</b> = 后备母猪饲料 + 其他母猪饲料。</div>
    <div>· <b>猪料（合计）</b> = 仔猪饲料 + 母猪饲料 + 生长肥育猪饲料 + 其他猪饲料；分项不可相加。</div>
    <div>· <b>季度产量</b> = 该季度 3 个月产量之和；<b>季度环比</b> = 本季 ÷ 上季 − 1；<b>季度同比</b> = 本季 ÷ 上年同季 − 1。只有 3 个月齐全的季度才计算。</div>
    <div>· 四个口径互为校验：全部饲料 = 配合饲料 + 浓缩饲料 + 添加剂预混合饲料（2021-12 为已知历史修订差异）。</div>
  </div>
  <div class="src" id="srcNote"></div>
</div>

<div class="tt" id="tt"></div>
</div>

<div class="lock" id="lock" hidden>
  <form class="lockcard" id="lockForm">
    <h2>🔒 需要访问密码</h2>
    <p>该看板包含内部数据，请输入访问密码后查看。</p>
    <input id="lockPw" type="password" autocomplete="current-password" placeholder="请输入密码" required>
    <button type="submit">进入看板</button>
    <div class="lockerr" id="lockErr"></div>
  </form>
</div>

<script>
/* 数据在构建时用 AES-GCM 加密，密码只存在于本地 .dashboard_password。
   浏览器输入密码后经 PBKDF2 派生密钥解密；同一标签页会话内免重复输入。 */
const ENC = __ENC__;
const PW_KEY = 'feed-dash-pw';

function b64u8(s){
  return Uint8Array.from(atob(s), c => c.charCodeAt(0));
}
async function decPayload(enc, pw){
  const km = await crypto.subtle.importKey(
    'raw', new TextEncoder().encode(pw), 'PBKDF2', false, ['deriveKey']);
  const key = await crypto.subtle.deriveKey(
    {name:'PBKDF2', salt:b64u8(enc.salt), iterations:enc.iter, hash:'SHA-256'},
    km, {name:'AES-GCM', length:256}, false, ['decrypt']);
  const plain = await crypto.subtle.decrypt(
    {name:'AES-GCM', iv:b64u8(enc.iv)}, key, b64u8(enc.data));
  return JSON.parse(new TextDecoder().decode(plain));
}
function askPassword(){
  return new Promise(resolve => {
    const lock=document.getElementById('lock');
    const form=document.getElementById('lockForm');
    const inp=document.getElementById('lockPw');
    const err=document.getElementById('lockErr');
    lock.hidden=false;
    inp.value='';
    err.textContent='';
    setTimeout(()=>inp.focus(), 50);
    form.onsubmit=async e=>{
      e.preventDefault();
      err.textContent='正在验证…';
      try{
        const data=await decPayload(ENC, inp.value);
        try{ sessionStorage.setItem(PW_KEY, inp.value); }catch(_){}
        lock.hidden=true;
        resolve(data);
      }catch(_){
        err.textContent='密码不正确，请重试';
        inp.select();
      }
    };
  });
}
async function unlockFeed(){
  let cached=null;
  try{ cached=sessionStorage.getItem(PW_KEY); }catch(_){}
  if(cached){
    try{ return await decPayload(ENC, cached); }
    catch(_){ try{ sessionStorage.removeItem(PW_KEY); }catch(__){} }
  }
  return askPassword();
}

(async function bootFeed(){
const DATA = await unlockFeed();
const META = DATA.meta, SCOPES = DATA.scopes, ORDER = DATA.scope_order;
const SCOPE_LABELS = {all:'全部饲料', hp:'配合饲料', ns:'浓缩饲料', tjj:'添加剂预混料'};
const CAT_LABELS = {
  '猪料':'猪料（合计）',
  '仔猪饲料':'仔猪饲料（=哺乳+保育）',
  '哺乳仔猪饲料':'哺乳仔猪饲料', '保育饲料':'保育饲料',
  '母猪饲料':'母猪饲料（=后备+其他母猪）',
  '后备母猪饲料':'后备母猪饲料', '其他母猪饲料':'其他母猪饲料',
  '生长肥育猪饲料':'生长肥育猪饲料', '其他猪饲料':'其他猪饲料'
};
const FREQ_LABELS = {M:'月度', Q:'季度'};
const METRIC_LABELS = {level:'产量', mom:'环比', yoy:'同比'};
const YEARS = Array.from(new Set(ORDER.flatMap(k => SCOPES[k].months.map(m => +m.slice(0,4)))))
  .sort((a,b) => a-b);
const state = {
  scope: ORDER[0], freq: 'M', metric: 'mom', cat: '猪料',
  years: YEARS.slice(-3)
};

const $ = id => document.getElementById(id);
const fmt2 = v => (v==null || !isFinite(v)) ? '—'
  : v.toLocaleString('zh-CN',{minimumFractionDigits:2,maximumFractionDigits:2});
const fmtPct = v => (v==null || !isFinite(v)) ? '—' : (v>0?'+':'') + v.toFixed(2) + '%';
const round2 = x => Math.round((x + Number.EPSILON) * 100) / 100;

function scopeData(){ return SCOPES[state.scope]; }
function catLabel(c){ return CAT_LABELS[c] || c; }
function qMonths(q){
  const p=q.split('-Q'), y=+p[0], n=+p[1];
  return [1,2,3].map(i => `${y}-${String((n-1)*3+i).padStart(2,'0')}`);
}
function qShift(q,d){
  const p=q.split('-Q'), idx=(+p[0])*4 + (+p[1]-1) + d;
  return `${Math.floor(idx/4)}-Q${idx%4+1}`;
}
function pval(scope, p, cat){
  if(p.includes('-Q')){
    const vs=qMonths(p).map(m => scope.levels[m] ? scope.levels[m][cat] : null);
    if(vs.some(v => v==null)) return null;
    return round2(vs.reduce((a,b)=>a+b,0));
  }
  const lv=scope.levels[p];
  return lv && lv[cat]!=null ? lv[cat] : null;
}
function pct(cur, prev){
  if(cur==null || prev==null || prev===0) return null;
  return round2((cur/prev - 1) * 100);
}
function mval(scope, p, cat, metric){
  if(p.includes('-Q')){
    const cur=pval(scope,p,cat);
    if(metric==='level') return cur;
    return pct(cur, pval(scope, qShift(p, metric==='mom'?-1:-4), cat));
  }
  const cur=pval(scope,p,cat);
  if(metric==='level') return cur;
  if(metric==='mom') return pct(cur, pval(scope, shiftMonth(p,-1), cat));
  return pct(cur, pval(scope, shiftMonth(p,-12), cat));
}
function shiftMonth(m,d){
  const p=m.split('-'), y=+p[0], mo=+p[1];
  const idx=y*12 + (mo-1) + d;
  return `${Math.floor(idx/12)}-${String(idx%12+1).padStart(2,'0')}`;
}
function latestQuarterIdx(scope, year){
  const p=scope.latest.split('-'), ly=+p[0], lm=+p[1];
  if(year<ly) return 3;
  if(year>ly) return -1;
  return Math.floor(lm/3) - 1;
}
function latestPeriod(scope, freq){
  if(freq==='M') return scope.latest;
  const y=+scope.latest.slice(0,4), i=latestQuarterIdx(scope,y);
  return i>=0 ? `${y}-Q${i+1}` : null;
}
function shiftPeriod(p, freq, d){
  return freq==='Q' ? qShift(p,d) : shiftMonth(p,d);
}
function periodNo(p, freq){
  return freq==='Q' ? +p.split('-Q')[1] : +p.slice(5);
}
function periodKey(freq, year, i){
  return freq==='Q' ? `${year}-Q${i+1}` : `${year}-${String(i+1).padStart(2,'0')}`;
}
function periodsOfYear(scope, freq, year){
  if(freq==='M'){
    const ly=+scope.latest.slice(0,4), lm=+scope.latest.slice(5);
    const n = year===ly ? lm : 12;
    return Array.from({length:n}, (_,i) => ({key:periodKey('M',year,i), label:`${i+1}月`, no:i+1}));
  }
  let n=4;
  if(year > +scope.latest.slice(0,4)) n=0;
  else if(year === +scope.latest.slice(0,4)) n=latestQuarterIdx(scope,year)+1;
  return Array.from({length:Math.max(n,0)}, (_,i) => ({key:periodKey('Q',year,i), label:`Q${i+1}`, no:i+1}));
}
function yearColor(y){
  if(y===2026) return 'var(--y26)';
  if(y===2025) return 'var(--y25)';
  if(y===2024) return 'var(--y24)';
  if(y===2023) return 'var(--y23)';
  if(y===2022) return 'var(--y22)';
  if(y===2021) return 'var(--y21)';
  if(y===2020) return 'var(--y20)';
  return 'var(--ink-3)';
}
function yearDash(y){
  // 2020-2023 统一实线（用户指定）；2025/2024 保留虚线/点线，
  // 作为红绿之外的第二重编码。
  if(y===2025) return '7 4';
  if(y===2024) return '2 3';
  return '';
}
function yearFocus(y){ return y===2026 || y===2025 || y===2024; }
function fmtAxis(v, metric){
  if(metric==='level'){
    if(Math.abs(v)>=10000) return (v/10000).toFixed(v>=100000?0:1)+'万';
    return Math.round(v).toString();
  }
  return (v>0?'+':'') + Math.round(v) + '%';
}
function esc(s){
  return String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}
function niceNum(range, round){
  const exp=Math.floor(Math.log10(range));
  const f=range/Math.pow(10,exp);
  let nf;
  if(round) nf = f<1.5?1:(f<3?2:(f<7?5:10));
  else nf = f<=1?1:(f<=2?2:(f<=5?5:10));
  return nf*Math.pow(10,exp);
}
function niceScale(lo, hi, count){
  if(!isFinite(lo)||!isFinite(hi)||lo===hi){hi=lo+1;}
  const step=niceNum(niceNum(hi-lo,false)/(count-1),true);
  const nlo=Math.floor(lo/step)*step;
  const nhi=Math.ceil(hi/step)*step;
  const vals=[];
  for(let v=nlo; v<=nhi+step*0.5; v+=step) vals.push(v);
  return {lo:nlo, hi:nhi, vals};
}

/* ---------- 控件 ---------- */
function buildControls(){
  $('segScope').innerHTML = ORDER.map(k =>
    `<button type="button" data-scope="${k}" aria-pressed="${k===state.scope}">${SCOPE_LABELS[k]||SCOPES[k].label}</button>`).join('');
  $('segFreq').innerHTML = Object.keys(FREQ_LABELS).map(k =>
    `<button type="button" data-freq="${k}" aria-pressed="${k===state.freq}">${FREQ_LABELS[k]}</button>`).join('');
  $('segMetric').innerHTML = Object.keys(METRIC_LABELS).map(k =>
    `<button type="button" data-metric="${k}" aria-pressed="${k===state.metric}">${METRIC_LABELS[k]}</button>`).join('');
  $('segYears').innerHTML = YEARS.map(y =>
    `<button type="button" data-year="${y}" aria-pressed="${state.years.includes(y)}">${y}</button>`).join('');
  const sel=$('selCat');
  sel.innerHTML = scopeData().categories.map(c =>
    `<option value="${esc(c)}" ${c===state.cat?'selected':''}>${esc(catLabel(c))}</option>`).join('');
  sel.onchange = () => { state.cat=sel.value; renderAll(); };
}

function bindControls(){
  $('segScope').addEventListener('click', e => {
    const b=e.target.closest('button[data-scope]'); if(!b) return;
    state.scope=b.dataset.scope;
    const cats=scopeData().categories;
    if(!cats.includes(state.cat)) state.cat = cats.includes('猪料') ? '猪料' : cats[0];
    buildControls(); renderAll();
  });
  $('segFreq').addEventListener('click', e => {
    const b=e.target.closest('button[data-freq]'); if(!b) return;
    state.freq=b.dataset.freq; buildControls(); renderAll();
  });
  $('segMetric').addEventListener('click', e => {
    const b=e.target.closest('button[data-metric]'); if(!b) return;
    state.metric=b.dataset.metric; buildControls(); renderAll();
  });
  $('segYears').addEventListener('click', e => {
    const b=e.target.closest('button[data-year]'); if(!b) return;
    const y=+b.dataset.year;
    if(state.years.includes(y)){
      if(state.years.length<=1) return;
      state.years=state.years.filter(v=>v!==y);
    }else{
      state.years=[...state.years,y].sort((a,b)=>a-b);
    }
    buildControls(); renderAll();
  });
}

/* ---------- KPI ---------- */
function renderKpis(){
  const s=scopeData(), lp=latestPeriod(s, state.freq), freq=state.freq;
  const cur=pval(s,lp,state.cat);
  const mom=mval(s,lp,state.cat,'mom'), yoy=mval(s,lp,state.cat,'yoy');
  const prevP=shiftPeriod(lp,freq,-1), prevY=shiftPeriod(lp,freq,freq==='M'?-12:-4);
  const prevLv=pval(s,prevP,state.cat), prevYLv=pval(s,prevY,state.cat);
  const y=+lp.slice(0,4), n=periodNo(lp,freq);
  const curVals=Array.from({length:n},(_,i)=>pval(s,periodKey(freq,y,i),state.cat));
  const prevVals=Array.from({length:n},(_,i)=>pval(s,periodKey(freq,y-1,i),state.cat));
  let ytd=null, ytdYoy=null;
  if(curVals.every(v=>v!=null) && prevVals.every(v=>v!=null)){
    ytd=round2(curVals.reduce((a,b)=>a+b,0));
    ytdYoy=pct(ytd, prevVals.reduce((a,b)=>a+b,0));
  }
  const unitWord = freq==='M' ? '月度' : '季度';
  const ytdLabel = freq==='M' ? `${y}年1-${n}月累计` : `${y}年Q1-Q${n}累计`;

  function delta(v,prefix){
    if(v==null) return `<span class="flat">—</span>`;
    const cls=v>0?'up':(v<0?'dn':'flat');
    const arrow=v>0?'▲':(v<0?'▼':'■');
    return `<span class="${cls}">${arrow} ${prefix} ${fmtPct(v)}</span>`;
  }
  function tile(k, v, unit, d, cls, note){
    return `<div class="tile card"><div class="k">${k}</div>`
      + `<div class="v ${cls||''}">${v}<span class="u">${unit}</span></div>`
      + `<div class="${note?'n':'d'}">${d}</div></div>`;
  }
  const kpiCls = v => v==null?'':(v>0?'up':(v<0?'dn':'flat'));
  $('kpis').innerHTML =
    tile(`最新${unitWord}产量 · ${lp}`, fmt2(cur), '吨', delta(mom,'环比'), '', false)
    + tile(`最新${unitWord}环比`, fmtPct(mom), '', `上期 ${prevP}：${fmt2(prevLv)} 吨`, kpiCls(mom), true)
    + tile(`最新${unitWord}同比`, fmtPct(yoy), '', `上年同期 ${prevY}：${fmt2(prevYLv)} 吨`, kpiCls(yoy), true)
    + tile(ytdLabel, fmt2(ytd), '吨', delta(ytdYoy,'累计同比'), '', false);
}

/* ---------- 折线图 ---------- */
function drawLine(){
  const s=scopeData(), freq=state.freq, metric=state.metric, cat=state.cat;
  const wrap=$('lineWrap');
  const W=Math.max(wrap.clientWidth, 320), H=W<640?300:370;
  const ml=54, mr=64, mt=18, mb=34;
  const iw=W-ml-mr, ih=H-mt-mb;
  const n=freq==='M'?12:4;
  const x=i => ml + (n<=1?0:i*iw/(n-1));
  const labels=freq==='M'?Array.from({length:12},(_,i)=>`${i+1}月`):['Q1','Q2','Q3','Q4'];
  const series=state.years.map(y => ({
    year:y,
    vals:Array.from({length:n},(_,i)=>mval(s,periodKey(freq,y,i),cat,metric))
  }));
  const all=series.flatMap(x=>x.vals).filter(v=>v!=null);
  let lo=Math.min(0,...all), hi=Math.max(0,...all);
  if(!isFinite(lo)||!isFinite(hi)){lo=0;hi=1;}
  if(lo===hi){lo-=1;hi+=1;}
  const pad=(hi-lo)*0.12; lo-=pad; hi+=pad;
  if(metric==='level') lo=Math.max(0, lo);
  const sc=niceScale(lo,hi,5);
  lo=sc.lo; hi=sc.hi;
  const yv=v => mt + ih - (v-lo)/(hi-lo)*ih;

  let svg=`<svg class="chart" viewBox="0 0 ${W} ${H}" role="img" aria-label="同期对比折线图">`;
  // grid + y labels
  sc.vals.forEach(v=>{
    const yy=yv(v);
    svg+=`<line x1="${ml}" y1="${yy.toFixed(1)}" x2="${(ml+iw).toFixed(1)}" y2="${yy.toFixed(1)}" style="stroke:var(--grid);stroke-width:1"/>`;
    svg+=`<text x="${ml-8}" y="${(yy+4).toFixed(1)}" text-anchor="end" style="fill:var(--ink-3);font-size:11px">${esc(fmtAxis(v,metric))}</text>`;
  });
  if(metric!=='level' && lo<0 && hi>0){
    svg+=`<line x1="${ml}" y1="${yv(0).toFixed(1)}" x2="${(ml+iw).toFixed(1)}" y2="${yv(0).toFixed(1)}" style="stroke:var(--axis);stroke-width:1"/>`;
  }
  // x labels
  labels.forEach((lab,i)=>{
    if(n===12 && W<520 && i%2===1) return;
    svg+=`<text x="${x(i).toFixed(1)}" y="${H-10}" text-anchor="middle" style="fill:var(--ink-3);font-size:11px">${lab}</text>`;
  });

  // lines + markers
  const focusYear=Math.max(...state.years);
  const labelIdx = freq==='M'
    ? (focusYear===+s.latest.slice(0,4) ? +s.latest.slice(5)-1 : 11)
    : (latestQuarterIdx(s,focusYear)>=0 ? latestQuarterIdx(s,focusYear) : 3);
  let labelPts=[];
  series.forEach(se=>{
    let d='', started=false;
    se.vals.forEach((v,i)=>{
      if(v==null){started=false;return;}
      d += `${started?'L':'M'}${x(i).toFixed(1)},${yv(v).toFixed(1)}`;
      started=true;
    });
    const w = se.year===2026?3:(se.year===2025?2.4:(se.year===2024?2.2:1.8));
    const dash=yearDash(se.year);
    svg+=`<path d="${d}" fill="none" style="stroke:${yearColor(se.year)};stroke-width:${w};${dash?`stroke-dasharray:${dash};`:''}stroke-linejoin:round;stroke-linecap:round"/>`;
    if(yearFocus(se.year)){
      se.vals.forEach((v,i)=>{ if(v!=null){
        svg+=`<circle cx="${x(i).toFixed(1)}" cy="${yv(v).toFixed(1)}" r="3.6" style="fill:${yearColor(se.year)};stroke:var(--surface);stroke-width:2"/>`;
      }});
      const v=se.vals[labelIdx];
      if(v!=null) labelPts.push({year:se.year, v, py:yv(v), px:x(labelIdx)});
    }
  });
  // direct labels with simple collision adjustment
  if(W<560) labelPts=labelPts.filter(p=>p.year===focusYear);
  labelPts.sort((a,b)=>a.py-b.py);
  for(let i=1;i<labelPts.length;i++){
    if(labelPts[i].py-labelPts[i-1].py<20) labelPts[i].py=labelPts[i-1].py+20;
  }
  labelPts.forEach(p=>{
    const txt=metric==='level'?fmtAxis(p.v,'level'):fmtPct(p.v);
    svg+=`<text x="${(p.px+7).toFixed(1)}" y="${(p.py+4).toFixed(1)}" style="fill:var(--ink-1);font-size:11px;font-weight:600">${esc(txt)}</text>`;
  });

  const cross=`<line id="crossLine" x1="0" y1="${mt}" x2="0" y2="${mt+ih}" style="stroke:var(--axis);stroke-width:1;opacity:0"/>`;
  svg+=cross;
  svg+=`<rect id="hitArea" x="${ml}" y="${mt}" width="${iw}" height="${ih}" fill="transparent" style="cursor:crosshair"/>`;
  svg+='</svg>';
  wrap.innerHTML=svg;

  const svgEl=wrap.querySelector('svg');
  const over=wrap.querySelector('#hitArea');
  const crossLine=wrap.querySelector('#crossLine');
  over.addEventListener('mousemove', e=>{
    const r=svgEl.getBoundingClientRect();
    const scale=W/r.width;
    const px=(e.clientX-r.left)*scale;
    let i=Math.round((px-ml)/(iw/Math.max(n-1,1)));
    i=Math.max(0,Math.min(n-1,i));
    crossLine.setAttribute('x1',x(i)); crossLine.setAttribute('x2',x(i));
    crossLine.style.opacity=1;
    let html=`<div class="tth">${esc(labels[i])}</div>`;
    state.years.slice().sort((a,b)=>b-a).forEach(y=>{
      const v=mval(s,periodKey(freq,y,i),cat,metric);
      html+=`<div class="ttr"><span class="sw" style="border-color:${yearColor(y)}"></span>`
        + `<span class="nm">${y}年</span><span class="vl">${metric==='level'?fmt2(v)+' 吨':fmtPct(v)}</span></div>`;
    });
    showTip(e,html);
  });
  over.addEventListener('mouseleave', ()=>{crossLine.style.opacity=0; hideTip();});

  // legend
  $('legend').innerHTML=state.years.slice().sort((a,b)=>b-a).map(y=>{
    const style=`border-color:${yearColor(y)};border-top-style:${y===2025?'dashed':(y===2024?'dotted':'solid')}`;
    return `<span class="li"><span class="sw" style="${style}"></span>${y}年</span>`;
  }).join('');
  $('lineHint').textContent = `标签为 ${focusYear} 年 ${freq==='M'?(labelIdx+1)+'月':'Q'+(labelIdx+1)} 的同期值；悬停查看所有年份同一期数值。`;
}

/* ---------- 柱图 ---------- */
function drawBars(){
  const s=scopeData(), freq=state.freq, metric=state.metric, cat=state.cat;
  const year=+latestPeriod(s,freq).slice(0,4);
  const periods=periodsOfYear(s,freq,year);
  const vals=periods.map(p=>mval(s,p.key,cat,metric));
  const wrap=$('barWrap');
  const W=Math.max(wrap.clientWidth,320), H=310;
  const ml=54, mr=16, mt=20, mb=34;
  const iw=W-ml-mr, ih=H-mt-mb;
  const n=Math.max(periods.length,1);
  const all=vals.filter(v=>v!=null);
  let lo=Math.min(0,...all), hi=Math.max(0,...all);
  if(!isFinite(lo)||!isFinite(hi)){lo=0;hi=1;}
  if(lo===hi){lo-=1;hi+=1;}
  if(metric==='level'){
    lo=0;
    hi += (hi-lo)*0.12;
  }else{
    hi += (hi-lo)*0.22;
    const p=(hi-lo)*0.15; lo-=p; hi+=p;
  }
  const sc=niceScale(lo,hi,5);
  lo=sc.lo; hi=sc.hi;
  const yv=v => mt + ih - (v-lo)/(hi-lo)*ih;
  const band=iw/n, bw=Math.min(56,band*0.56);

  let svg=`<svg class="chart" viewBox="0 0 ${W} ${H}" role="img" aria-label="当期柱状图">`;
  sc.vals.forEach(v=>{
    const yy=yv(v);
    svg+=`<line x1="${ml}" y1="${yy.toFixed(1)}" x2="${(ml+iw).toFixed(1)}" y2="${yy.toFixed(1)}" style="stroke:var(--grid);stroke-width:1"/>`;
    svg+=`<text x="${ml-8}" y="${(yy+4).toFixed(1)}" text-anchor="end" style="fill:var(--ink-3);font-size:11px">${esc(fmtAxis(v,metric))}</text>`;
  });
  const zy=yv(0);
  if(lo<0 && hi>0){
    svg+=`<line x1="${ml}" y1="${zy.toFixed(1)}" x2="${(ml+iw).toFixed(1)}" y2="${zy.toFixed(1)}" style="stroke:var(--axis);stroke-width:1"/>`;
  }
  const validIdx=vals.map((v,i)=>v!=null?i:-1).filter(i=>i>=0);
  const lastIdx=validIdx.length?validIdx[validIdx.length-1]:-1;
  periods.forEach((p,i)=>{
    const v=vals[i];
    const cx=ml+band*i+band/2;
    svg+=`<text x="${cx.toFixed(1)}" y="${H-10}" text-anchor="middle" style="fill:var(--ink-3);font-size:11px">${p.label}</text>`;
    if(v==null) return;
    const y0=Math.min(yv(v),zy), y1=Math.max(yv(v),zy);
    let fill;
    if(metric==='level') fill = (i===lastIdx)?'var(--bar-level-latest)':'var(--bar-level)';
    else if(v>0) fill='var(--red)'; else if(v<0) fill='var(--blue)'; else fill='var(--zero)';
    svg+=`<rect x="${(cx-bw/2).toFixed(1)}" y="${y0.toFixed(1)}" width="${bw.toFixed(1)}" height="${Math.max(y1-y0,0.8).toFixed(1)}" rx="2" style="fill:${fill}" data-i="${i}"/>`;
    if(metric==='level'){
      if(i===lastIdx){
        svg+=`<text x="${cx.toFixed(1)}" y="${(y0-6).toFixed(1)}" text-anchor="middle" style="fill:var(--ink-1);font-size:11px;font-weight:600">${fmtAxis(v,'level')}</text>`;
      }
    }else{
      const tx=cx;
      const ty=v>=0?y0-6:y1+14;
      svg+=`<text x="${tx.toFixed(1)}" y="${ty.toFixed(1)}" text-anchor="middle" style="fill:var(--ink-1);font-size:10px">${fmtPct(v)}</text>`;
    }
  });
  svg+='</svg>';
  wrap.innerHTML=svg;
  wrap.querySelectorAll('rect[data-i]').forEach(rect=>{
    rect.addEventListener('mousemove', e=>{
      const i=+rect.dataset.i, p=periods[i], v=vals[i];
      showTip(e, `<div class="tth">${year} 年 ${p.label}</div>`
        + `<div class="ttr"><span class="nm">${METRIC_LABELS[metric]}</span>`
        + `<span class="vl">${metric==='level'?fmt2(v)+' 吨':fmtPct(v)}</span></div>`);
    });
    rect.addEventListener('mouseleave', hideTip);
  });
  $('barTitle').textContent=`${year} 年${FREQ_LABELS[freq]}${METRIC_LABELS[metric]}`;
  $('barHint').textContent = metric==='level'
    ? '柱顶只标最新一期，其余看悬停或下方明细表。'
    : '红色 = 环比/同比上升，蓝色 = 下降；柱顶标签带正负号，颜色不是唯一线索。';
}

/* ---------- 明细表 ---------- */
function renderTable(){
  const s=scopeData(), freq=state.freq, cat=state.cat;
  const year=+latestPeriod(s,freq).slice(0,4);
  const periods=periodsOfYear(s,freq,year);
  const pcol=freq==='M'?'月份':'季度';
  const prevCol=freq==='M'?'上年同月（吨）':'上年同季（吨）';
  let html=`<thead><tr><th class="l">${pcol}</th><th>产量（吨）</th><th>环比（%）</th><th>同比（%）</th><th>${prevCol}</th><th>同比增减（吨）</th></tr></thead><tbody>`;
  periods.forEach(p=>{
    const lv=pval(s,p.key,cat);
    if(lv==null) return;
    const py=pval(s,periodKey(freq,year-1,p.no-1),cat);
    const mom=mval(s,p.key,cat,'mom'), yoy=mval(s,p.key,cat,'yoy');
    const diff=py!=null?round2(lv-py):null;
    const cls=v=>v==null?'flat':(v>0?'up':(v<0?'dn':'flat'));
    html+=`<tr><td class="l">${p.key}</td><td>${fmt2(lv)}</td>`
      + `<td class="${cls(mom)}">${fmtPct(mom)}</td><td class="${cls(yoy)}">${fmtPct(yoy)}</td>`
      + `<td>${fmt2(py)}</td><td class="${cls(diff)}">${diff==null?'—':(diff>0?'+':'')+fmt2(diff)}</td></tr>`;
  });
  html+='</tbody>';
  $('tbl').innerHTML=html;
  $('tableTitle').textContent=`${year} 年${FREQ_LABELS[freq]}明细`;
}

function downloadCsv(){
  const s=scopeData(), freq=state.freq, cat=state.cat;
  const periods = freq==='M'
    ? s.months.slice()
    : YEARS.flatMap(y=>Array.from({length:4},(_,i)=>`${y}-Q${i+1}`));
  const pcol=freq==='M'?'月份':'季度';
  const rows=[[pcol,'产量（吨）','环比（%）','同比（%）']];
  periods.forEach(p=>{
    const lv=pval(s,p,cat);
    if(lv==null) return;
    const mom=mval(s,p,cat,'mom'), yoy=mval(s,p,cat,'yoy');
    rows.push([p, lv, mom==null?'':mom, yoy==null?'':yoy]);
  });
  const text='﻿'+rows.map(r=>r.join(',')).join('\n');
  const blob=new Blob([text],{type:'text/csv;charset=utf-8'});
  const a=document.createElement('a');
  a.href=URL.createObjectURL(blob);
  const safe=(SCOPE_LABELS[state.scope]||'')+'_'+catLabel(cat);
  a.download=`饲料_${safe}_${FREQ_LABELS[freq]}.csv`;
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(a.href);
}

/* ---------- tooltip ---------- */
function showTip(e, html){
  const tt=$('tt');
  tt.innerHTML=html;
  tt.style.opacity=1;
  const pad=14, w=tt.offsetWidth, h=tt.offsetHeight;
  let left=e.clientX+pad, top=e.clientY+pad;
  if(left+w>window.innerWidth-8) left=e.clientX-w-pad;
  if(top+h>window.innerHeight-8) top=e.clientY-h-pad;
  tt.style.left=left+'px'; tt.style.top=top+'px';
}
function hideTip(){ $('tt').style.opacity=0; }

/* ---------- 总渲染 ---------- */
function renderAll(){
  const s=scopeData(), lp=latestPeriod(s,state.freq);
  $('subtitle').textContent=`月度 / 季度产量 · 环比 · 同比 ｜ 数据截至 ${META.latest} ｜ 来源：${META.source}`;
  $('lineTitle').textContent=`同期对比 · ${FREQ_LABELS[state.freq]}${METRIC_LABELS[state.metric]}`;
  renderKpis();
  drawLine();
  drawBars();
  renderTable();
}

function initNote(){
  const warn=(META.warnings&&META.warnings.length)
    ? ' 数据注记：'+META.warnings.join('；')+'。' : '';
  $('srcNote').textContent=`快照生成：${META.generated_at||'—'} ｜ 文件：data/feed_dashboard.json。${warn}`;
}

/* ---------- 主题 ---------- */
function initTheme(){
  const root=document.documentElement;
  const saved=localStorage.getItem('feed-theme');
  if(saved) root.dataset.theme=saved;
  $('themeBtn').addEventListener('click', ()=>{
    const dark = root.dataset.theme
      ? root.dataset.theme==='dark'
      : matchMedia('(prefers-color-scheme: dark)').matches;
    const next = dark?'light':'dark';
    root.dataset.theme=next;
    localStorage.setItem('feed-theme',next);
    renderAll();
  });
}

let resizeTimer=null;
window.addEventListener('resize', ()=>{
  clearTimeout(resizeTimer);
  resizeTimer=setTimeout(()=>{ drawLine(); drawBars(); },150);
});

bindControls();
buildControls();
initTheme();
initNote();
renderAll();
$('csvBtn').addEventListener('click', downloadCsv);
})();
</script>
</body>
</html>
'''


def main():
    import dashboard_crypto as dc
    data = fd.load_json()
    enc = dc.enc_js(data)
    html = TEMPLATE.replace('__ENC__', enc)
    assert '__ENC__' not in html
    os.makedirs(OUT_DIR, exist_ok=True)
    with open(OUT_PATH, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'写入 {OUT_PATH} ({os.path.getsize(OUT_PATH):,} bytes，已加密)')
    print(f"数据截至 {data['meta']['latest']}；口径 {len(data['scope_order'])} 个")
    print(f'访问密码：见 {dc.PW_FILE}')


if __name__ == '__main__':
    main()
