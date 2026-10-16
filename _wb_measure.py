# -*- coding: utf-8 -*-
"""量每张图的 viewBox 是否等于渲染尺寸 —— 不等就是被压扁了。"""
from playwright.sync_api import sync_playwright

JS_ALL = """() => [...document.querySelectorAll('svg.chart')].map((x,i)=>{
  const bb=x.getBoundingClientRect(), vb=x.getAttribute('viewBox').split(' ');
  const t=x.closest('.card')?.querySelector('.ctitle')?.innerText.slice(0,14)||('#'+i);
  return {i, t, vw:+vb[2], vh:+vb[3], rw:Math.round(bb.width), rh:Math.round(bb.height)};
})"""

JS_YOY = """() => [...document.querySelectorAll('#yoy svg.chart')].map(x=>{
  const bb=x.getBoundingClientRect(), vb=x.getAttribute('viewBox').split(' ');
  const p=x.parentElement;
  return {vw:+vb[2], rw:Math.round(bb.width), rh:Math.round(bb.height),
          plotW:p.clientWidth, hostW:document.getElementById('yoy').clientWidth,
          cards:document.querySelectorAll('#yoy .card').length};
})"""

with sync_playwright() as p:
    b = p.chromium.launch()
    pg = b.new_page(viewport={'width': 1500, 'height': 1000})
    pg.goto('file:///D:/CC/Desktop/社会供应压力看板.html', wait_until='load')
    pg.wait_for_timeout(2600)
    def report(tag):
        allc = pg.evaluate(JS_ALL)
        bad = [r for r in allc if abs(r['vw'] - r['rw']) > 2]
        print(f'  {tag}: 不匹配 {len(bad)}/{len(allc)}  ' +
              ' '.join(f"{r['t']}(vb{r['vw']}≠{r['rw']})" for r in bad[:6]))
        return bad

    print('--- 首次加载 ---')
    report('load')
    for r in pg.evaluate(JS_YOY):
        print('  ', r)
    print('--- 点了各切换之后 ---')
    total_bad = []
    for sel in ['#segCal button[data-c="农历"]',
                '#segYoySmooth button[data-s="1"]',
                '#segGran button[data-g="周"]',
                '#segGran button[data-g="日"]',
                '#segCal button[data-c="公历"]',
                '#segRange button[data-r="30"]',
                '#themeBtn']:
        pg.click(sel)
        pg.wait_for_timeout(1200)
        total_bad += report(sel)
    print(f'\n合计不匹配 {len(total_bad)} 张')
    b.close()
