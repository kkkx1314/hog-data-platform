# -*- coding: utf-8 -*-
"""公开看板的客户端加密。

页面里嵌的是一段 AES-GCM 密文；浏览器输入访问密码后，用 PBKDF2-SHA256
派生密钥解密，再渲染看板。密码保存在仓库根目录的 `.dashboard_password`
（已 gitignore），不存在时自动生成一个随机密码。

注意：这是访问门槛，不是账号体系。密码泄露或文件被拿走，数据就能解开。
"""
import base64
import hashlib
import json
import os
import secrets

from cryptography.hazmat.primitives.ciphers.aead import AESGCM

HERE = os.path.dirname(os.path.abspath(__file__))
PW_FILE = os.path.join(HERE, '.dashboard_password')
ITER = 150_000


def get_password(path=PW_FILE):
    """读取本地密码文件；不存在则生成 12 位随机密码并落盘。"""
    if os.path.exists(path):
        with open(path, encoding='utf-8') as f:
            pw = f.read().strip()
        if pw:
            return pw
    pw = secrets.token_urlsafe(9)
    with open(path, 'w', encoding='utf-8', newline='\n') as f:
        f.write(pw + '\n')
    return pw


def encrypt_json(obj, password=None):
    """把可 JSON 序列化的对象加密成前端可解的字典。"""
    pw = password or get_password()
    blob = json.dumps(obj, ensure_ascii=False, separators=(',', ':')).encode('utf-8')
    salt = secrets.token_bytes(16)
    iv = secrets.token_bytes(12)
    key = hashlib.pbkdf2_hmac('sha256', pw.encode('utf-8'), salt, ITER, dklen=32)
    ct = AESGCM(key).encrypt(iv, blob, None)
    b64 = lambda b: base64.b64encode(b).decode('ascii')
    return {'v': 1, 'kdf': 'PBKDF2-SHA256', 'iter': ITER,
            'salt': b64(salt), 'iv': b64(iv), 'data': b64(ct)}


def enc_js(obj, password=None):
    """加密并序列化为可直接嵌入 JS 的 JSON 字面量。"""
    return json.dumps(encrypt_json(obj, password), ensure_ascii=False, separators=(',', ':'))
