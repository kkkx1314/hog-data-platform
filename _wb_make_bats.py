# -*- coding: utf-8 -*-
"""生成桌面上的两个 .bat（写成 GBK —— 中文 Windows 控制台默认 936 码页，
   存成 UTF-8 的话 echo 的中文会变乱码）。

【为什么必须写死 python 的绝对路径】双击 bat 时 cmd 继承的是 Explorer 的环境，
而注册表里（系统 + 用户）的 PATH **没有** miniconda —— 我的终端里有，是因为
Claude Code 是从 conda 已激活的 shell 起的，那只是进程内的 PATH。
所以裸写 `python` 在双击场景下会落到
    C:\\Users\\CC\\AppData\\Local\\Microsoft\\WindowsApps\\python.exe
—— 那是微软商店的占位程序（AppInstallerPythonRedirector），它不跑任何脚本，
直接退出并返回 9009：`if errorlevel 1` 判定失败 → 报「更新失败」，
_wb_watch.py 一行都没执行、看板数据永远不更新。（2026-09-22 实测就是这个。）
生成时探测一个 import 得动 pandas/openpyxl/numpy/lunardate 的解释器，把绝对
路径写进 bat；运行期只做「文件还在不在」的兜底，找不到才退回裸 python。
"""
import os
import subprocess
import sys

HERE = r'C:\Users\CC\test-claude'
DESK = r'D:\CC\Desktop'
REQUIRED = 'import pandas, openpyxl, numpy, lunardate'

# 按优先级探测；sys.executable 排第一（它至少能跑起这个生成器）
CANDIDATES = [
    sys.executable,
    r'D:\miniconda\python.exe',
    r'D:\miniconda\envs\pytorch_env\python.exe',
    os.path.expandvars(r'%LOCALAPPDATA%\miniconda3\python.exe'),
    os.path.expandvars(r'%USERPROFILE%\miniconda3\python.exe'),
    os.path.expandvars(r'%USERPROFILE%\anaconda3\python.exe'),
]


def find_python():
    """第一个能 import 齐依赖的解释器。都要能过 —— pytorch_env 就缺 lunardate。"""
    tried, seen = [], set()
    for p in CANDIDATES:
        if not p or p.lower() in seen:
            continue
        seen.add(p.lower())
        if os.sep in p and not os.path.isfile(p):
            tried.append(f'{p} (不存在)')
            continue
        try:
            r = subprocess.run([p, '-c', REQUIRED], capture_output=True, timeout=120)
        except Exception as e:
            tried.append(f'{p} ({e})')
            continue
        if r.returncode == 0:
            return p, tried
        tried.append(f'{p} ({r.stderr.decode("utf-8", "replace").strip().splitlines()[-1] if r.stderr else "失败"})')
    return None, tried


ONE_SHOT = '''@echo off
title 更新社会供应压力看板
cd /d "{HERE}"
rem 见 _wb_make_bats.py 的注释：双击时 PATH 里没有 conda，裸 python 会落到商店占位程序
set "PY={PY}"
if not exist "%PY%" set "PY=python"
echo 正在读最新的过磅单并重建看板, 大约需要 20 秒 ...
echo.
"%PY%" _wb_watch.py --once --open
echo.
if errorlevel 1 (
  echo *** 更新失败 ***
  if exist _wb_last_error.log (
    echo 报错已存到 _wb_last_error.log, 正在用记事本打开 —— 把这个文件的内容发给 Claude。
    start "" notepad _wb_last_error.log
  ) else (
    echo 脚本根本没启动, 多半是找不到 Python 解释器。
    echo 本次用的解释器: %PY%
    echo 请把上面几行原样发给 Claude。
  )
) else (
  echo *** 看板已更新完成 ***
)
pause
'''

WATCH = '''@echo off
title 过磅单监听 - 社会供应压力看板
cd /d "{HERE}"
rem 见 _wb_make_bats.py 的注释：双击时 PATH 里没有 conda，裸 python 会落到商店占位程序
set "PY={PY}"
if not exist "%PY%" set "PY=python"
echo.
echo   正在监听 {SRC}
echo   你往里放过磅单 / 改过磅单, 看板会自动重建。
echo   看板文件: {OUT}
echo.
echo   这个黑窗口要一直开着; 关掉它就停止监听了。
echo.
echo   本次用的解释器: %PY%
echo.
"%PY%" _wb_watch.py
pause
'''


def main():
    py, tried = find_python()
    if py is None:
        print('!! 没找到能 import 齐 pandas/openpyxl/numpy/lunardate 的 Python:')
        for t in tried:
            print('   ', t)
        sys.exit('先生成一个能跑的解释器再回来（例如 D:\\miniconda\\python.exe）')
    print(f'用解释器: {py}')
    for t in tried:
        print(f'   (跳过 {t})')

    src = r'D:\CC\Desktop\车辆过磅单'
    out = r'D:\CC\Desktop\社会供应压力看板.html'
    tmpl = {
        '更新看板.bat': ONE_SHOT.format(HERE=HERE, PY=py),
        '启动过磅单监听.bat': WATCH.format(HERE=HERE, PY=py, SRC=src, OUT=out),
    }
    for name, txt in tmpl.items():
        # 中文 Windows 控制台按 GBK 解码 echo 的内容，写不成 GBK 的字符要先发现
        txt.encode('gbk')
        p = os.path.join(DESK, name)
        with open(p, 'w', encoding='gbk', newline='\r\n') as f:
            f.write(txt)
        print(f'-> {p}')


if __name__ == '__main__':
    main()
