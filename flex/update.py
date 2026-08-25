#!/usr/bin/env python3
"""
内置更新模块 - 检测GitHub Release新版本并下载
============================================
- 检测: GitHub API releases/latest (匿名可访问)
- 版本比较: tag v6.1.7.1 → 数字比较
- 下载: releases/download/{tag}/{asset} 固定URL
- 替换: 下载到临时文件 → updater.bat 延迟替换并重启(运行时不能覆盖自身)
"""
import os
import re
import sys
import json
import time
import urllib.request

REPO = "Akie-tu/table-match-tool"
# 当前版本 (打包时由构建注入, 或在此维护)
CURRENT_VERSION = "v6.2.1"


def parse_version(tag):
    """'v6.1.7.1' -> (6,1,7,1)"""
    m = re.search(r"v?(\d+(?:\.\d+)*)", tag or "")
    if not m:
        return (0,)
    return tuple(int(x) for x in m.group(1).split("."))


def version_gt(a, b):
    """a > b?"""
    pa, pb = parse_version(a), parse_version(b)
    la = max(len(pa), len(pb))
    pa += (0,) * (la - len(pa))
    pb += (0,) * (la - len(pb))
    return pa > pb


def get_latest_release(timeout=15):
    """查询最新Release, 返回 (tag, body, assets) 或 None(失败/无更新)
    双通道: 先API, 失败fallback网页版(国内api.github.com常被墙但网页可达)"""
    # 通道1: API
    try:
        url = f"https://api.github.com/repos/{REPO}/releases/latest"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=timeout) as r:
            d = json.loads(r.read().decode())
        assets = {a["name"]: a["browser_download_url"] for a in d.get("assets", [])}
        return d.get("tag_name"), d.get("body") or "", assets
    except Exception:
        pass
    # 通道2: 网页版 (releases/latest 302重定向 → /releases/tag/vX.Y.Z)
    try:
        url = f"https://github.com/{REPO}/releases/latest"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"}, method="HEAD")
        with urllib.request.urlopen(req, timeout=timeout) as r:
            final = r.geturl()
        m = re.search(r"/releases/tag/([^/?#]+)", final)
        if not m:
            return None
        tag = m.group(1)
        # 资产直链: releases/download/{tag}/{asset}
        # asset名即exe文件名, 由check_update传入
        return tag, "", {}
    except Exception:
        return None


def check_update(current=None, asset_name=None, timeout=15):
    """
    检查更新. 返回 dict 或 None:
      {latest_tag, current_tag, has_update, body, download_url, asset_name}
    """
    cur = current or CURRENT_VERSION
    try:
        result = get_latest_release(timeout=timeout)
        if result is None:
            return None
        tag, body, assets = result
    except Exception:
        return None  # 网络/API失败
    if not tag:
        return None
    has = version_gt(tag, cur)
    url = assets.get(asset_name) if asset_name else (list(assets.values())[0] if assets else None)
    if not url and asset_name and tag:
        # 网页通道: 用直链下载
        url = f"https://github.com/{REPO}/releases/download/{tag}/{asset_name}"
    return {
        "latest_tag": tag,
        "current_tag": cur,
        "has_update": has,
        "body": body,
        "download_url": url,
        "asset_name": asset_name,
    }


def download_file(url, dest, progress_cb=None, timeout=120):
    """下载文件(支持进度回调), 返回 (ok, size, err)"""
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            total = int(resp.headers.get("Content-Length") or 0)
            got = 0
            chunk = 64 * 1024
            with open(dest, "wb") as f:
                while True:
                    buf = resp.read(chunk)
                    if not buf:
                        break
                    f.write(buf)
                    got += len(buf)
                    if progress_cb and total:
                        progress_cb(got, total)
            return True, got, None
    except Exception as e:
        try:
            if os.path.exists(dest):
                os.remove(dest)
        except Exception:
            pass
        return False, 0, str(e)


def self_exe_name():
    """当前运行的exe文件名 (打包时= sys.executable 文件名)"""
    if getattr(sys, "frozen", False):
        return os.path.basename(sys.executable)
    return "table-match-gui.exe"


def make_updater_bat(new_exe, old_exe, temp_dir):
    """
    生成 updater.bat (纯ASCII): 等待主进程退出→替换exe→重启
    Windows CMD 的 GBK 编码, bat 必须纯 ASCII (windows-bat-encoding skill)
    """
    bat = f"""@echo off
setlocal
rem updater for table-match-gui (generated)
rem wait for main process to exit
:loop
tasklist /FI "IMAGENAME eq {old_exe}" 2>nul | find /I "{old_exe}" >nul
if %errorlevel%==0 (
  timeout /t 1 /nobreak >nul
  goto loop
)
rem replace
copy /Y "{new_exe}" "{old_exe}" >nul 2>&1
if %errorlevel%==0 goto restart
echo FAILED_TO_REPLACE > {temp_dir}\\updater_result.txt
exit /b 1
:restart
echo OK > {temp_dir}\\updater_result.txt
del "{new_exe}" >nul 2>&1
start "" "{old_exe}"
exit /b 0
"""
    return bat


def run_updater(bat_path, new_exe, old_exe, temp_dir):
    """启动 updater.bat 并退出当前程序"""
    try:
        os.startfile(bat_path)  # Windows
    except Exception:
        import subprocess
        subprocess.Popen(["cmd", "/c", bat_path],
                         creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0)

if __name__ == "__main__":
    info = check_update()
    print(info)