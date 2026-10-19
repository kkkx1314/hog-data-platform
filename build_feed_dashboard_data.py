# -*- coding: utf-8 -*-
"""从桌面四份饲料 Excel 生成线上看板快照 data/feed_dashboard.json。

用法：python build_feed_dashboard_data.py
跑完后再执行「同步数据到GitHub.bat」，Streamlit Cloud 就能读到新快照。

校验（不通过就不写文件）：
  1. 每个口径都有完整月份，且最新月一致；
  2. 仔猪 = 哺乳 + 保育；母猪 = 后备 + 其他母猪；
  3. 9 品类口径下：猪料 = 仔猪 + 母猪 + 生长肥育 + 其他猪；
  4. 全部饲料 = 配合 + 浓缩 + 预混（逐月、逐品类）。
"""
import json
import os
import sys

import feed_dashboard as fd


def _close(a, b, tol=2.0):
    # 2022 年源表按月取整，分项相加与合计最多差 1.1 吨；2.0 足够容纳舍入，
    # 又能拦住真正的解析错位（那种通常是成百上千吨）。
    return a is not None and b is not None and abs(a - b) <= tol


# 2021-12 的四口径历史修订不一致（全部饲料与分类型文件差约 1 万吨），
# 属于源表历史版本差异，不是解析错误；保留各口径原值并在页面注明。
KNOWN_CROSS_EXCEPTIONS = {'2021-12'}


def validate(data):
    issues = []
    warnings = []
    scopes = data['scopes']

    for key in data['scope_order']:
        s = scopes[key]
        if not s['months']:
            issues.append(f'{s["label"]}: 没有月度数据')
            continue
        if len(s['months']) < 80:
            issues.append(f'{s["label"]}: 只有 {len(s["months"])} 个月，预期 >= 80')
        if s['latest'] != data['meta']['latest']:
            issues.append(f'{s["label"]}: 最新月 {s["latest"]} != 全局 {data["meta"]["latest"]}')

        for m in s['months']:
            lv = s['levels'][m]
            a, b, c = lv.get('仔猪饲料'), lv.get('哺乳仔猪饲料'), lv.get('保育饲料')
            if a is not None and b is not None and c is not None and not _close(a, b + c):
                issues.append(f'{s["label"]} {m}: 仔猪 {a} != 哺乳 {b} + 保育 {c}')
            a, b, c = lv.get('母猪饲料'), lv.get('后备母猪饲料'), lv.get('其他母猪饲料')
            if a is not None and b is not None and c is not None and not _close(a, b + c):
                issues.append(f'{s["label"]} {m}: 母猪 {a} != 后备 {b} + 其他母猪 {c}')

            if key != 'all':
                parts = [lv.get(x) for x in
                         ('仔猪饲料', '母猪饲料', '生长肥育猪饲料', '其他猪饲料')]
                total = lv.get('猪料')
                if total is not None and all(v is not None for v in parts):
                    if not _close(total, sum(parts)):
                        issues.append(f'{s["label"]} {m}: 猪料 {total} != 四个一级品类之和 {sum(parts):.2f}')
            else:
                # 全部饲料没有单列其他猪，只能验证猪料 >= 三个已知一级品类之和
                parts = [lv.get(x) for x in ('仔猪饲料', '母猪饲料', '生长肥育猪饲料')]
                total = lv.get('猪料')
                if total is not None and all(v is not None for v in parts):
                    if total + 2.0 < sum(parts):
                        issues.append(f'{s["label"]} {m}: 猪料 {total} < 已知一级品类之和 {sum(parts):.2f}')

    # 跨口径恒等式：全部 = 配合 + 浓缩 + 预混
    for m in scopes['all']['months']:
        all_lv = scopes['all']['levels'].get(m, {})
        bad = []
        for cat in scopes['all']['categories']:
            parts = [scopes[k]['levels'].get(m, {}).get(cat) for k in ('hp', 'ns', 'tjj')]
            if all_lv.get(cat) is not None and all(v is not None for v in parts):
                if not _close(all_lv[cat], sum(parts)):
                    bad.append(f'{cat} 差 {all_lv[cat] - sum(parts):+.2f}')
        if bad:
            msg = f'{m}: 全部 != 配合+浓缩+预混（' + '；'.join(bad) + '）'
            if m in KNOWN_CROSS_EXCEPTIONS:
                warnings.append(msg)
            else:
                issues.append(msg)

    return issues, warnings


def main():
    print('读取四份 Excel ...')
    for spec in fd.SCOPES:
        print(f'  {spec["short"]}: {fd.EXCEL_PATHS[spec["key"]]}')
    data = fd.extract_all()

    print('\n校验 ...')
    issues, warnings = validate(data)
    if issues:
        print(f'  [FAIL] {len(issues)} 项不通过：')
        for x in issues[:30]:
            print('   -', x)
        if len(issues) > 30:
            print(f'   ... 另有 {len(issues) - 30} 项')
        raise SystemExit(1)
    print('  [OK] 嵌套恒等式 + 四口径互校验全部通过')
    if warnings:
        print(f'  [WARN] {len(warnings)} 项已知历史差异（保留原值）：')
        for x in warnings:
            print('   -', x)

    data['meta']['warnings'] = warnings

    os.makedirs(os.path.dirname(fd.JSON_PATH), exist_ok=True)
    with open(fd.JSON_PATH, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, separators=(',', ':'))
    size = os.path.getsize(fd.JSON_PATH)
    print(f'\n写入 {fd.JSON_PATH} ({size:,} bytes)')

    latest = data['meta']['latest']
    print(f'\n数据截至 {latest}：')
    for key in data['scope_order']:
        s = data['scopes'][key]
        v = fd.value_at(s, latest, '猪料')
        print(f'  {s["label"]:<10} {len(s["months"]):>3} 个月  '
              f'猪料 {v:,.2f} 吨  ({len(s["categories"])} 个品类)')
    print('\n下一步：同步数据到GitHub.bat（或 python sync_data.py）')


if __name__ == '__main__':
    main()
