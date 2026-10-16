# -*- coding: utf-8 -*-
"""
把 _wb_data.json 注入模板，产出单文件看板。

两个出口：
  · 桌面 `社会供应压力看板.html`（state=null → 模板默认：最新日期 / 公历 / 日 / 显示明细）
  · 仓库 `data/_wb_data.json`（Streamlit Cloud 上读它；本地则直接读根目录那份）

Streamlit 页复用这里的 `build_html(data, state)` 生成带初始状态的 HTML 再嵌进 iframe。
"""
import json, math, os

HERE = os.path.dirname(os.path.abspath(__file__))
TPL = os.path.join(HERE, '_wb_template.html')
SRC = os.path.join(HERE, '_wb_data.json')
OUT = r'D:\CC\Desktop\社会供应压力看板.html'
CLOUD_SRC = os.path.join(HERE, 'data', '_wb_data.json')

# 构建器自检要看的关键字段
NEED_KEYS = ['periods', 'audit', 'bin_labels', 'ext_detail', 'qc_rows', 'bin_group', 'yoy',
             'lunar_of_day']
# 六个「历法 × 粒度」组合对应的 periods 数组键（模板端用 granKey() 映射）
GRAN_KEYS = ['日', '周', '月', '农历周', '农历月']


def scrub(o):
    """NaN/Inf -> None，保证 JS 端 null 判断可用（NaN 会漏过 == null）。"""
    if isinstance(o, float):
        return None if (math.isnan(o) or math.isinf(o)) else o
    if isinstance(o, dict):
        return {k: scrub(v) for k, v in o.items()}
    if isinstance(o, list):
        return [scrub(v) for v in o]
    return o


def load_data(path=None):
    with open(path or SRC, encoding='utf-8') as f:
        return json.load(f)


def build_html(data=None, state=None, tpl_path=None):
    """state: {cal,gran,endDate,range,detail,embed,theme}；None/{} = 全部走模板默认。"""
    if data is None:
        data = load_data()
    tpl = open(tpl_path or TPL, encoding='utf-8').read()
    for ph in ('__WB_STATE__', '__WB_DATA__', '__WB_HTMLATTR__', '__WB_BODYATTR__'):
        assert ph in tpl, f'模板缺少 {ph} 占位符'
    blob = json.dumps(scrub(data), ensure_ascii=False, separators=(',', ':'))
    assert 'NaN' not in blob, '仍有 NaN 残留'
    st = json.dumps(state, ensure_ascii=False) if state else 'null'
    # 外壳指定的 {embed, theme} 直接写进 <html>/<body> 标签，避免首屏闪一下再变
    sm = state or {}
    html_attr = f' data-theme="{sm["theme"]}"' if sm.get('theme') in ('dark', 'light') else ''
    body_attr = ' class="embed"' if sm.get('embed') else ''
    # 先换 STATE 再换 DATA：state 里不会有 __WB_DATA__，但反过来的顺序没必要冒险
    return (tpl.replace('__WB_HTMLATTR__', html_attr)
               .replace('__WB_BODYATTR__', body_attr)
               .replace('__WB_STATE__', st)
               .replace('__WB_DATA__', blob))


def main():
    data = load_data()
    html = build_html(data)
    with open(OUT, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'看板 -> {OUT}  ({os.path.getsize(OUT)/1024/1024:.2f} MB)')

    # 自检
    for k in NEED_KEYS:
        assert k in data, f'缺字段 {k}'
    for k in GRAN_KEYS:
        assert k in data['periods'], f'periods 缺粒度 {k}'
        assert data['periods'][k], f'periods.{k} 是空的'
    miss = [k for k in ('d0', 'd1') if k not in data['periods']['日'][0]]
    assert not miss, f'periods 缺 {miss}（看板按截止日筛选要用）'
    assert data['yoy']['modes']['solar'].get('dates'), '同比缺 dates 槽位日期'
    assert data['yoy']['years'] == ['2025', '2026'], '同比年份表不对'
    print('自检通过：无 NaN，关键字段齐全')

    # 云端副本：Streamlit Cloud 上没有 D:\CC\Desktop，只能读仓库里的 data/
    os.makedirs(os.path.dirname(CLOUD_SRC), exist_ok=True)
    with open(CLOUD_SRC, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, separators=(',', ':'))
    print(f'云端副本 -> {CLOUD_SRC}  ({os.path.getsize(CLOUD_SRC)/1024/1024:.2f} MB)')


if __name__ == '__main__':
    main()
