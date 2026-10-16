# -*- coding: utf-8 -*-
"""
社会供应压力看板 —— 过磅单自动更新

盯着源目录，一有新文件或文件被改动，就自动重跑管道并重建看板。
不需要装任何第三方库（用轮询 mtime/size，不用 watchdog）。

用法:
    python _wb_watch.py                 常驻监听（启动时先构建一次）
    python _wb_watch.py --once          只构建一次就退出（给 .bat 用）
    python _wb_watch.py --open          构建完自动用默认浏览器打开看板
    python _wb_watch.py --no-initial    启动时不构建，只等变化
    python _wb_watch.py --settle 16     文件静默多少秒后才认为写完了（默认 8）

Ctrl+C 退出。
"""
import argparse
import glob
import os
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
SRC_DIR = r'D:\CC\Desktop\车辆过磅单'
OUT_HTML = r'D:\CC\Desktop\社会供应压力看板.html'
PIPELINE = os.path.join(HERE, '_wb_pipeline.py')
BUILD = os.path.join(HERE, '_wb_build.py')
ERRLOG = os.path.join(HERE, '_wb_last_error.log')

POLL = 4           # 轮询间隔（秒）
RETRY = 3          # 读取源文件失败时的重试次数（Excel 可能正占着文件）
RETRY_WAIT = 10    # 重试间隔（秒）


def log(msg):
    print(f'[{time.strftime("%Y-%m-%d %H:%M:%S")}] {msg}', flush=True)


def snapshot():
    """源目录指纹：文件名 + 修改时间 + 大小。任何一项变了就说明数据变了。

    盯的是目录里所有 .xlsx，不只是 车辆过磅单-*.xlsx —— 名字起歪了的导出
    (车辆过磅单(1).xlsx 之类)也得能触发一次重建，好让管道把"这份没被读进来"
    的警告打出来。~$ 开头的是 Excel 的锁文件，一开一关就变，必须排除掉，
    否则每次打开源文件都会白跑一次重建。
    """
    out = []
    for p in sorted(glob.glob(os.path.join(SRC_DIR, '*.xlsx'))):
        name = os.path.basename(p)
        if name.startswith('~$'):
            continue
        try:
            st = os.stat(p)
        except OSError:
            continue
        out.append((name, int(st.st_mtime_ns), st.st_size))
    return tuple(out)


def run(script):
    """跑一个子脚本，实时回显它的输出。返回 (是否成功, 输出文本)。"""
    env = dict(os.environ, PYTHONIOENCODING='utf-8')
    r = subprocess.run([sys.executable, script], capture_output=True, text=True,
                       encoding='utf-8', errors='replace', env=env, cwd=HERE)
    txt = (r.stdout or '') + (r.stderr or '')
    return r.returncode == 0, txt


def write_errlog(what, detail):
    """失败时把完整现场落到 _wb_last_error.log。

    以前只往控制台打最后几行，用户一关窗口就再也查不到了（2026-09-22 就是这样，
    只能靠复现猜）。现在不管谁问，日志里都带着解释器路径和全部输出。
    """
    try:
        with open(ERRLOG, 'w', encoding='utf-8') as f:
            f.write(f'时间      : {time.strftime("%Y-%m-%d %H:%M:%S")}\n')
            f.write(f'解释器    : {sys.executable}\n')
            f.write(f'Python    : {sys.version.splitlines()[0]}\n')
            f.write(f'工作目录  : {HERE}\n')
            f.write(f'问题      : {what}\n')
            f.write('-' * 70 + '\n')
            f.write((detail or '(子进程没有任何输出)').rstrip() + '\n')
        log(f'  完整报错已写入 {ERRLOG}')
    except Exception as e:                      # 写日志本身失败也不能盖掉原本的错
        log(f'  !! 写 {ERRLOG} 失败: {e}')


def clear_errlog():
    """这次成功了 —— 把上次的报错清掉，免得下次误当成新问题。"""
    try:
        if os.path.exists(ERRLOG):
            os.remove(ERRLOG)
    except OSError:
        pass


def rebuild(open_after=False):
    log('检测到过磅单变动，开始重建看板 ...')
    log(f'  解释器 {sys.executable}')
    t0 = time.time()
    tried = []
    for attempt in range(1, RETRY + 1):
        ok, out = run(PIPELINE)
        if ok:
            break
        tried.append(f'--- 第 {attempt} 次尝试 ---\n' + out)
        log(f'  管道第 {attempt} 次失败({"可能是 Excel 正占着文件" if attempt < RETRY else "不再重试"})')
        print('    ' + '\n    '.join(out.strip().splitlines()[-4:]), flush=True)
        if attempt == RETRY:
            log('  放弃本次更新，继续监听。')
            write_errlog(f'管道连续 {RETRY} 次失败', '\n'.join(tried))
            return False
        time.sleep(RETRY_WAIT)

    # 回显管道实际读了哪几个源文件、一共多少行；名字不对的文件必须让用户看见
    for line in out.splitlines():
        if '合并' in line or '跳过(读不出)' in line or line.strip().startswith('!!'):
            log('  ' + line.strip())

    ok, bout = run(BUILD)
    if not ok:
        log('  注入看板失败：')
        print('    ' + '\n    '.join(bout.strip().splitlines()[-6:]), flush=True)
        write_errlog('注入看板失败（_wb_build.py）', bout)
        return False
    log(f'  完成，用时 {time.time() - t0:.1f}s  ->  {OUT_HTML}')
    clear_errlog()
    if open_after:
        try:
            os.startfile(OUT_HTML)      # 只在 Windows 上有
        except Exception as e:
            log(f'  自动打开看板失败: {e}')
    return True


def main():
    ap = argparse.ArgumentParser(description='过磅单 -> 看板 自动更新')
    ap.add_argument('--once', action='store_true', help='只构建一次就退出')
    ap.add_argument('--open', action='store_true', help='构建完自动打开看板')
    ap.add_argument('--no-initial', action='store_true', help='启动时不先构建')
    ap.add_argument('--settle', type=int, default=8,
                    help='文件静默多少秒后才动手（默认 8，防止读到写了一半的文件）')
    a = ap.parse_args()

    if a.once:
        sys.exit(0 if rebuild(a.open) else 1)

    if not os.path.isdir(SRC_DIR):
        log(f'源目录不存在: {SRC_DIR}')
        sys.exit(1)

    log(f'监听 {SRC_DIR}（每 {POLL}s 查一次，文件静默 {a.settle}s 后才重建）')
    log(f'产出 {OUT_HTML}')
    if not a.no_initial:
        rebuild(a.open)

    last = snapshot()
    changed_at = None          # 指纹最后一次变化的时间
    while True:
        try:
            time.sleep(POLL)
            cur = snapshot()
            if cur != last:
                last = cur
                changed_at = time.time()
                log(f'源文件有变动（{len(cur)} 个文件），等文件写完 ...')
                continue
            if changed_at and time.time() - changed_at >= a.settle:
                changed_at = None
                rebuild(a.open)
        except KeyboardInterrupt:
            log('已停止监听。')
            break


def _entry():
    """任何预料之外的异常也要落盘 —— 不然窗口一关又变成"查不到"。"""
    try:
        main()
    except SystemExit:
        raise                      # --once 的退出码要原样带出去给 bat 判断
    except KeyboardInterrupt:
        log('已停止监听。')
    except Exception:
        import traceback
        tb = traceback.format_exc()
        write_errlog('脚本自身异常', tb)
        print(tb, flush=True)
        sys.exit(1)


if __name__ == '__main__':
    _entry()
