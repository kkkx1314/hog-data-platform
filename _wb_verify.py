# -*- coding: utf-8 -*-
"""看板自检: 交互 + 悬停 + 深浅色 + 各粒度截图"""
import sys
from playwright.sync_api import sync_playwright

URL = 'file:///D:/CC/Desktop/社会供应压力看板.html'
TT = "()=>document.getElementById('tt').innerText.split('\\n').join(' | ')"

with sync_playwright() as p:
    b = p.chromium.launch()
    pg = b.new_page(viewport={'width': 1500, 'height': 1150})
    errs = []
    pg.on('pageerror', lambda e: errs.append('PE:' + str(e)))
    pg.on('console', lambda m: errs.append('ERR:' + m.text) if m.type == 'error' else None)
    pg.goto(URL, wait_until='load')
    pg.wait_for_timeout(2500)

    print('SVG:', pg.eval_on_selector_all('svg.chart', 'e=>e.length'))
    print('KPI:', pg.eval_on_selector_all('.tile', 'e=>e.length'))
    print('Tables:', pg.eval_on_selector_all('table', 'e=>e.length'))
    print('Legends:', pg.eval_on_selector_all('.legend', 'e=>e.length'))
    print('审计存疑表行:', pg.eval_on_selector_all('#qcTbl tr', 'e=>e.length'))
    for gone in ['lunar', 'cust']:
        print(f'  已删板块 #{gone} 残留数:', pg.eval_on_selector_all(f'#{gone}', 'e=>e.length'))

    def hover(sel, fx, fy, tag):
        """先滚到视口内再悬停, 否则鼠标坐标落不到图上"""
        pg.evaluate(f"()=>document.querySelector('{sel}').scrollIntoView({{block:'center'}})")
        pg.wait_for_timeout(350)
        bb = pg.query_selector(sel).bounding_box()
        pg.mouse.move(bb['x'] + bb['width'] * fx, bb['y'] + bb['height'] * fy)
        pg.wait_for_timeout(400)
        print(tag, pg.evaluate(TT)[:190])
        pg.mouse.move(3, 3); pg.wait_for_timeout(150)

    # ---- 季节同比 ----
    print('同比图数:', pg.eval_on_selector_all('#yoy svg.chart', 'e=>e.length'))
    print('同比图例:', pg.eval_on_selector_all('#yoy .legend .li', 'e=>e.map(x=>x.innerText).join(" | ")'))
    print('同比线色:', pg.eval_on_selector_all(
        '#yoy svg.chart path[stroke]', 'e=>[...new Set(e.map(x=>x.getAttribute("stroke")))].join(" , ")'))
    print('同比虚线:', pg.eval_on_selector_all(
        '#yoy svg.chart path[stroke-dasharray]', 'e=>e.length'))
    print('横轴刻度(公历):', pg.eval_on_selector_all(
        '#yoy svg.chart:first-of-type text', 'e=>e.map(x=>x.textContent).filter(t=>/月$/.test(t)).join(" ")'))
    hover('#yoy svg.chart', .45, .5, 'TOOLTIP公历:')
    # 农历对齐（对齐方式跟着顶部「历法」走，这里切的也是它）
    pg.click('#segCal button[data-c="农历"]'); pg.wait_for_timeout(1500)
    print('横轴刻度(农历):', pg.eval_on_selector_all(
        '#yoy svg.chart:first-of-type text', 'e=>e.map(x=>x.textContent).filter(t=>/月$/.test(t)).join(" ")'))
    hover('#yoy svg.chart', .45, .5, 'TOOLTIP农历:')
    pg.click('#segYoySmooth button[data-s="1"]'); pg.wait_for_timeout(1500)
    print('原始日值(农历)图数:', pg.eval_on_selector_all('#yoy svg.chart', 'e=>e.length'))
    y = pg.evaluate("()=>document.getElementById('yoy').getBoundingClientRect().top+scrollY")
    pg.evaluate(f'window.scrollTo(0,{y-120})'); pg.wait_for_timeout(500)
    pg.screenshot(path='_wb_yoy_lunar.png')
    pg.click('#segYoySmooth button[data-s="7"]'); pg.wait_for_timeout(1200)

    hover('#trends svg.chart', .6, .5, 'TOOLTIP:')
    hover('#structure svg.chart', .5, .6, 'STACK  :')
    hover('#diag svg.chart', .5, .3, 'DIAGBAR:')

    # 表格视图切换
    btns = pg.query_selector_all('.tbtn')
    btns[0].click(); pg.wait_for_timeout(300)
    print('表视图可见:', pg.eval_on_selector_all('.tblwrap:not(.hide)', 'e=>e.length'))
    btns[0].click(); pg.wait_for_timeout(300)

    # 深色 + 口径板块截图
    pg.click('#themeBtn'); pg.wait_for_timeout(900)
    y = pg.evaluate("()=>document.getElementById('diag').getBoundingClientRect().top+scrollY")
    pg.evaluate(f'window.scrollTo(0,{y-40})'); pg.wait_for_timeout(400)
    pg.screenshot(path='_wb_diag_dark.png')

    # 各粒度截图(深色)
    pg.evaluate('window.scrollTo(0,0)'); pg.wait_for_timeout(300)
    pg.click('#themeBtn'); pg.wait_for_timeout(800)
    for c, g, tag in [('公历', '月', '月'), ('农历', '月', '农历月')]:
        pg.click(f'#segCal button[data-c="{c}"]')
        pg.click(f'#segGran button[data-g="{g}"]'); pg.wait_for_timeout(900)
        y = pg.evaluate("()=>document.getElementById('structure').getBoundingClientRect().top+scrollY")
        pg.evaluate(f'window.scrollTo(0,{y-90})'); pg.wait_for_timeout(400)
        pg.screenshot(path=f'_wb_gran_{tag}.png')
    pg.click('#segCal button[data-c="公历"]')
    pg.click('#segGran button[data-g="日"]'); pg.wait_for_timeout(800)
    y = pg.evaluate("()=>document.getElementById('structure').getBoundingClientRect().top+scrollY")
    pg.evaluate(f'window.scrollTo(0,{y-90})'); pg.wait_for_timeout(400)
    pg.screenshot(path='_wb_gran_日.png')

    print('ERRORS:', errs[:8] if errs else 'none')
    b.close()
