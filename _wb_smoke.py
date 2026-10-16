# -*- coding: utf-8 -*-
"""看板冒烟: 加载 -> 无 JS 错误 -> 六种历法×粒度组合都能切换 -> 截止日生效 -> 明细开关生效"""
import sys
from playwright.sync_api import sync_playwright

URL = 'file:///D:/CC/Desktop/社会供应压力看板.html'

with sync_playwright() as p:
    b = p.chromium.launch()
    pg = b.new_page(viewport={'width': 1500, 'height': 1100})
    errs = []
    pg.on('pageerror', lambda e: errs.append('PE:' + str(e)))
    pg.on('console', lambda m: errs.append('ERR:' + m.text) if m.type == 'error' else None)
    pg.goto(URL, wait_until='load')
    pg.wait_for_timeout(2600)

    print('SVG:', pg.eval_on_selector_all('svg.chart', 'e=>e.length'))
    print('KPI:', pg.eval_on_selector_all('.tile', 'e=>e.length'))
    print('scopeHint:', pg.eval_on_selector('#scopeHint', 'e=>e.textContent'))
    print('yoyHint  :', pg.eval_on_selector('#yoyHint', 'e=>e.textContent')[:110])
    print('截止框默认:', pg.eval_on_selector('#dtEnd', 'e=>e.value'))
    print('明细板块可见:', pg.eval_on_selector_all('.det:not(.hide)', 'e=>e.map(x=>x.id)'))

    # 六种组合
    for cal, gran in [('公历', '日'), ('公历', '周'), ('公历', '月'),
                      ('农历', '日'), ('农历', '周'), ('农历', '月')]:
        pg.click(f'#segCal button[data-c="{cal}"]')
        pg.click(f'#segGran button[data-g="{gran}"]')
        pg.wait_for_timeout(1000)
        hint = pg.eval_on_selector('#scopeHint', 'e=>e.textContent')
        n = pg.eval_on_selector_all('svg.chart', 'e=>e.length')
        first = pg.eval_on_selector('#trends svg.chart text', 'e=>e.textContent')
        print(f'  {cal}{gran}: SVG={n:2d} | {hint[:88]}')

    # 明细开关
    pg.click('#segDetail button[data-d="0"]'); pg.wait_for_timeout(400)
    print('明细隐藏后可见板块:', pg.eval_on_selector_all('.det:not(.hide)', 'e=>e.map(x=>x.id)'))
    pg.click('#segDetail button[data-d="1"]'); pg.wait_for_timeout(400)
    print('明细恢复后可见板块:', pg.eval_on_selector_all('.det:not(.hide)', 'e=>e.map(x=>x.id)'))

    # 截止日
    pg.click('#segCal button[data-c="公历"]')
    pg.click('#segGran button[data-g="日"]'); pg.wait_for_timeout(600)
    pg.fill('#dtEnd', '2026-06-30'); pg.dispatch_event('#dtEnd', 'change')
    pg.wait_for_timeout(1200)
    print('截止 2026-06-30:', pg.eval_on_selector('#scopeHint', 'e=>e.textContent'))
    print('  末期指标卡:', pg.eval_on_selector('#rulesCard', 'e=>e.textContent')[:60].strip())
    pg.click('#btnLatest'); pg.wait_for_timeout(1200)
    print('点最新后:', pg.eval_on_selector('#scopeHint', 'e=>e.textContent'))
    print('  截止框:', pg.eval_on_selector('#dtEnd', 'e=>e.value'))

    print('ERRORS:', errs[:8] if errs else 'none')
    b.close()
