# -*- coding: utf-8 -*-
"""生成离线版饲料产量看板：单个 HTML，双击即开，不需要密码、不联网。

数据来自 data/feed_dashboard.json（由 build_feed_dashboard_data.py 生成），
页面复用 build_feed_public_site.TEMPLATE，但去掉加密与密码遮罩，直接把
JSON 内嵌进 <script>，因此 file:// 打开、断网、发给同事都能用。

用法：python build_feed_offline.py [输出路径]
默认输出：D:\\CC\\Desktop\\云南饲料产量看板_离线版.html
"""
import json
import os
import re
import sys

import build_feed_public_site as pub
import feed_dashboard as fd

DEFAULT_OUT = r'D:\CC\Desktop\云南饲料产量看板_离线版.html'

OLD_COMMENT = (
    '/* 数据在构建时用 AES-GCM 加密，密码只存在于本地 .dashboard_password。\n'
    '   浏览器输入密码后经 PBKDF2 派生密钥解密；同一标签页会话内免重复输入。 */'
)
NEW_COMMENT = '/* 离线版：数据在构建时直接内嵌，无需密码、无需联网。 */'


def build(data):
    html = pub.TEMPLATE

    # 1) 去掉加密常量 + 解密/密码函数（保留 bootFeed 的开头），改成直接内嵌 JSON
    pat = re.compile(r'const ENC = __ENC__;.*?return askPassword\(\);\n\}', re.S)
    html, n = pat.subn('', html, count=1)
    if n != 1:
        raise RuntimeError('模板结构与预期不符：找不到加密/解密函数块')
    payload = json.dumps(data, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
    html, n = re.subn(r'const DATA = await unlockFeed\(\);',
                      lambda m: 'const DATA = ' + payload + ';', html, count=1)
    if n != 1:
        raise RuntimeError('模板结构与预期不符：找不到 unlockFeed 调用')

    html = html.replace(OLD_COMMENT, NEW_COMMENT, 1)

    # 2) 去掉密码遮罩（离线版没有密码）
    html = re.sub(r'\n<div class="lock" id="lock" hidden>.*?\n</div>\n', '\n', html,
                  count=1, flags=re.S)

    # 3) 标识为离线版
    html = html.replace('<title>饲料产量看板 · 云南神农肉业</title>',
                        '<title>饲料产量看板（离线版）· 云南神农肉业</title>', 1)
    html = html.replace('文件：data/feed_dashboard.json', '离线单文件，数据内嵌', 1)

    for token in ('__ENC__', '__DATA__', 'unlockFeed', 'askPassword'):
        if token in html:
            raise RuntimeError(f'模板里仍残留 {token}')
    return html


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_OUT
    data = fd.load_json()
    html = build(data)
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    with open(out, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'写入 {out} ({os.path.getsize(out):,} bytes，未加密)')
    print(f"数据截至 {data['meta']['latest']}；口径 {len(data['scope_order'])} 个")


if __name__ == '__main__':
    main()
