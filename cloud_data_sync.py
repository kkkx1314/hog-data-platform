# -*- coding: utf-8 -*-
"""私有数据仓库同步。

公开仓库（kkkx1314/hog-data-platform）只放代码，不放任何 Excel 原始数据；
数据统一放在私有仓库 kkkx1314/swine-dashboard：

  · 本机：sync_data.py 把 data/*.xlsx 强制推成单提交（push_to_private_repo）
  · Streamlit Cloud：没有 D:\\CC\\Desktop\\平台数据，启动时用 GITHUB_TOKEN
    （Streamlit secret 或环境变量）把 xlsx 拉到 data/（sync_from_private_repo）

Token 只需要私有仓库的 Contents: Read 权限。
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import tempfile
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

GITHUB_OWNER = "kkkx1314"
GITHUB_REPO = "swine-dashboard"
GITHUB_BRANCH = "main"
# 本机 22 端口被封，统一走 443
PRIVATE_REMOTE = "ssh://git@ssh.github.com:443/kkkx1314/swine-dashboard.git"

SYNC_STATE = ".data_sync_state.json"
SYNC_TTL_SECONDS = 6 * 3600


def _token_from_mapping(obj) -> str:
    """在任意层级的 secrets 结构里找 PAT 值（键名不限）。"""
    if isinstance(obj, str):
        value = obj.strip()
        return value if value.startswith(("github_pat_", "ghp_")) else ""
    if isinstance(obj, dict):
        for value in obj.values():
            found = _token_from_mapping(value)
            if found:
                return found
    return ""


def get_github_token() -> str:
    """优先环境变量（本机测试/兼容旧部署），其次 Streamlit secrets。

    Secrets 的键名不强制叫 GITHUB_TOKEN：token 名叫 pig-dashboard 时，
    直接写 pig-dashboard = "github_pat_..." 也能读到。
    """
    token = os.environ.get("GITHUB_TOKEN", "").strip()
    if token:
        return token
    try:
        import streamlit as st

        secrets = st.secrets
    except Exception:
        return ""
    for key in ("GITHUB_TOKEN", "github_token", "pig-dashboard", "pig_dashboard"):
        try:
            value = secrets.get(key, "")
        except Exception:
            value = ""
        if isinstance(value, str) and value.strip():
            return value.strip()
    try:
        return _token_from_mapping(dict(secrets))
    except Exception:
        return ""


def _request(url: str, token: str, raw: bool = False) -> bytes:
    req = urllib.request.Request(url)
    if token:
        req.add_header("Authorization", f"Bearer {token}")
    req.add_header("Accept", "application/vnd.github.raw" if raw else "application/vnd.github+json")
    req.add_header("User-Agent", "hog-data-platform")
    try:
        with urllib.request.urlopen(req, timeout=180) as resp:
            return resp.read()
    except urllib.error.HTTPError as exc:
        if exc.code in {401, 403}:
            raise RuntimeError(
                "GITHUB_TOKEN 无效或没有私有仓库读权限（需要 Contents: Read）"
            ) from exc
        if exc.code == 404:
            raise RuntimeError(
                f"私有仓库 {GITHUB_OWNER}/{GITHUB_REPO} 或文件不存在：{url}"
            ) from exc
        raise RuntimeError(f"GitHub 请求失败（HTTP {exc.code}）：{url}") from exc


def list_remote_files(token: str) -> list[dict]:
    """列出私有仓库根目录下的 xlsx（name / sha / size）。"""
    url = (
        f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/contents"
        f"?ref={urllib.parse.quote(GITHUB_BRANCH)}"
    )
    items = json.loads(_request(url, token).decode("utf-8"))
    files = []
    for item in items:
        name = item.get("name", "")
        if item.get("type") != "file" or name.startswith("~$"):
            continue
        if not name.lower().endswith((".xlsx", ".xls")):
            continue
        files.append({"name": name, "sha": item.get("sha", ""), "size": int(item.get("size") or 0)})
    return files


def download_remote_file(name: str, token: str) -> bytes:
    path = urllib.parse.quote(name, safe="")
    url = (
        f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/contents/{path}"
        f"?ref={urllib.parse.quote(GITHUB_BRANCH)}"
    )
    return _request(url, token, raw=True)


def _load_state(dest: Path) -> dict:
    try:
        return json.loads((dest / SYNC_STATE).read_text(encoding="utf-8"))
    except Exception:
        return {}


def sync_from_private_repo(
    dest_dir: str | Path,
    token: str,
    force: bool = False,
    ttl: int = SYNC_TTL_SECONDS,
) -> dict:
    """把私有仓库的 xlsx 同步到 dest_dir（镜像：删除仓库里已不存在的 xlsx）。

    返回 {"downloaded": n, "skipped": n, "cached": bool, "checked_at": ts}
    """
    dest = Path(dest_dir)
    dest.mkdir(parents=True, exist_ok=True)
    state = _load_state(dest)
    now = time.time()
    checked_at = float(state.get("checked_at") or 0)
    if not force and checked_at and now - checked_at < ttl and any(dest.glob("*.xlsx")):
        return {
            "downloaded": 0,
            "skipped": len(state.get("files") or {}),
            "cached": True,
            "checked_at": checked_at,
        }

    remote = list_remote_files(token)
    if not remote:
        raise RuntimeError("私有数据仓库里没有 xlsx 文件")

    prev = state.get("files") or {}
    downloaded = skipped = 0
    for meta in remote:
        name, sha = meta["name"], meta["sha"]
        target = dest / name
        same = target.exists() and prev.get(name) == sha
        if same and (not meta["size"] or target.stat().st_size == meta["size"]):
            skipped += 1
            continue
        content = download_remote_file(name, token)
        if meta["size"] and len(content) != meta["size"]:
            raise RuntimeError(f"下载不完整：{name}（{len(content)} / {meta['size']} bytes）")
        tmp = target.with_name(target.name + ".part")
        tmp.write_bytes(content)
        os.replace(tmp, target)
        downloaded += 1

    # 仓库里已删除的文件，本地也清掉，避免 _find_latest_file 选到旧文件
    keep = {m["name"] for m in remote}
    for old in dest.glob("*.xlsx"):
        if old.name not in keep and not old.name.startswith("~$"):
            try:
                old.unlink()
            except OSError:
                pass

    state = {"checked_at": now, "files": {m["name"]: m["sha"] for m in remote}}
    (dest / SYNC_STATE).write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    return {"downloaded": downloaded, "skipped": skipped, "cached": False, "checked_at": now}


def _git(args: list[str], cwd: Path | None = None) -> str:
    result = subprocess.run(
        ["git", *args],
        cwd=str(cwd) if cwd else None,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if result.returncode:
        detail = (result.stderr or result.stdout).strip()[:400]
        raise RuntimeError(f"git {' '.join(args[:3])} 失败：{detail}")
    return result.stdout


def push_to_private_repo(src_dir: str | Path, message: str = "data: 同步平台 Excel 数据") -> str:
    """把 src_dir 下的 xlsx 推成私有仓库的一次性提交（单提交，避免历史膨胀）。"""
    src = Path(src_dir)
    files = sorted(f for f in src.glob("*.xlsx") if not f.name.startswith("~$"))
    if not files:
        return "跳过（data/ 里没有 xlsx）"

    with tempfile.TemporaryDirectory() as td:
        clone = Path(td) / "repo"
        _git(["clone", "--depth", "1", PRIVATE_REMOTE, str(clone)])
        for old in clone.glob("*.xlsx"):
            old.unlink()
        for f in files:
            shutil.copy2(f, clone / f.name)
        # 孤立分支：远端只保留这一个快照提交，旧的数据 blob 交给 GitHub 回收
        _git(["checkout", "--orphan", "data-snapshot"], cwd=clone)
        _git(["add", "-A"], cwd=clone)
        _git(
            ["-c", "user.name=data-sync", "-c", "user.email=data-sync@local",
             "commit", "-q", "-m", message],
            cwd=clone,
        )
        _git(["push", "-f", "origin", "HEAD:main"], cwd=clone)

    total_mb = sum(f.stat().st_size for f in files) / 1024 / 1024
    return f"已推送 {len(files)} 个文件（{total_mb:.1f} MB）到 {GITHUB_OWNER}/{GITHUB_REPO}"
