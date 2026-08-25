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
CURRENT_VERSION = "v6.3.5"


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


def get_latest_release(asset_name=None, timeout=15):
    """
    查询含指定资产的最新Release, 返回 (tag, body, assets) 或 None
    按资产归口: 主版找含 table-match-gui.exe 的最新release, flex 找 flex 资产
    双通道: 先API(翻列表找资产), 失败fallback网页版
    """
    # 通道1: API — releases?per_page=15 翻找含目标资产的最新release
    try:
        url = f"https://api.github.com/repos/{REPO}/releases?per_page=15"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=timeout) as r:
            releases = json.loads(r.read().decode())
        for rel in releases:
            assets = {a["name"]: a["browser_download_url"] for a in rel.get("assets", [])}
            # 无资产名要求 → 返回最新; 有资产名 → 找含该资产的
            if asset_name is None:
                return rel.get("tag_name"), rel.get("body") or "", assets
            if asset_name in assets:
                return rel.get("tag_name"), rel.get("body") or "", assets
        return None  # 15个release里都没有目标资产
    except Exception:
        pass
    # 通道2: 网页版 — releases?per_page=5 列表页解析tag+资产链接
    try:
        url = f"https://github.com/{REPO}/releases?per_page=5"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=timeout) as r:
            html = r.read().decode("utf-8", errors="ignore")
        # 解析每个release的tag和资产 (页面结构: /releases/tag/xxx 和 href资产链接)
        if asset_name is None:
            m = re.search(r"/releases/tag/([^/?#\"]+)", html)
            return (m.group(1), "", {}) if m else None
        # 找含目标资产的release: 按出现顺序, 资产链接 /releases/download/{tag}/{asset}
        m = re.search(rf"/releases/download/([^/?#\"]+)/{re.escape(asset_name)}", html)
        if m:
            return m.group(1), "", {m.group(1): f"https://github.com/{REPO}/releases/download/{m.group(1)}/{asset_name}"}
        return None
    except Exception:
        return None


def check_update(current=None, asset_name=None, timeout=15):
    """
    检查更新. 返回 dict 或 None:
      {latest_tag, current_tag, has_update, body, download_url, asset_name}
    按资产名归口: 主版(asset=table-match-gui.exe)忽略-flex tag, flex忽略无-flex tag
    """
    cur = current or CURRENT_VERSION
    is_flex = bool(asset_name and "flex" in asset_name)
    try:
        result = get_latest_release(asset_name=asset_name, timeout=timeout)
        if result is None:
            return None
        tag, body, assets = result
    except Exception:
        return None  # 网络/API失败
    if not tag:
        return None
    tag_l = tag.lower()
    # 归口双保险(asset匹配已保证, 这里再校验tag后缀一致性)
    if is_flex and not tag_l.endswith("-flex"):
        return None
    if not is_flex and tag_l.endswith("-flex"):
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


def _dl_single(url, dest, timeout=120, progress_cb=None):
    """单通道下载(支持断点续传+大小校验), 返回 (ok, size, err)"""
    try:
        # 已下载部分(续传)
        resume_from = 0
        if os.path.exists(dest):
            resume_from = os.path.getsize(dest)
        headers = {"User-Agent": "Mozilla/5.0"}
        if resume_from > 0:
            headers["Range"] = f"bytes={resume_from}-"
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            total = int(resp.headers.get("Content-Length") or 0)
            if resp.status == 206:  # 续传响应
                total = resume_from + total
            elif resume_from > 0:
                # 服务器不支持Range, 重头下
                resume_from = 0
            got = resume_from
            chunk = 128 * 1024
            mode = "ab" if resume_from > 0 else "wb"
            with open(dest, mode) as f:
                while True:
                    buf = resp.read(chunk)
                    if not buf:
                        break
                    f.write(buf)
                    got += len(buf)
                    if progress_cb and total:
                        progress_cb(got, total)
            # 大小校验: Content-Length 已知且不匹配 → 失败重下
            if total and got != total:
                return False, got, f"大小不符({got}/{total})"
            return True, got, None
    except Exception as e:
        return False, 0, str(e)


def _mirror_urls(url, asset_name):
    """生成多通道URL列表: 直连 + 镜像"""
    mirrors = [
        url,                                                  # 1. GitHub直连
        f"https://ghfast.top/{url}",                          # 2. ghfast镜像
        f"https://gh-proxy.com/{url}",                        # 3. gh-proxy镜像
        f"https://ghproxy.net/{url}",                         # 4. ghproxy镜像
    ]
    return mirrors


def _asset_ready(url, timeout=20):
    """HEAD预检资产是否已上传(CI构建中会404)"""
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"}, method="HEAD")
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.status == 200
    except Exception:
        return False


def download_file(url, dest, progress_cb=None, timeout=120):
    """多通道下载(直连+镜像自动切换, 带断点续传+大小校验), 返回 (ok, size, err)"""
    # 预检: 资产未上传(404) → 直接返回友好错误
    if not _asset_ready(url):
        return False, 0, "新版本正在构建中, 请稍候几分钟再试(资产尚未上传)"
    # 预检: 已存在且非0字节 → 尝试续传优先
    for idx, u in enumerate(_mirror_urls(url, "")):
        if os.path.exists(dest) and os.path.getsize(dest) > 0 and idx != 0:
            # 续传优先走直连(第0个), 镜像不续传(避免错乱)
            continue
        ok, size, err = _dl_single(u, dest, timeout=timeout, progress_cb=progress_cb)
        if ok:
            return True, size, None
        # 失败清理半成品? 保留以便下次续传, 但超过2次失败则重置
        if idx >= 2 and os.path.exists(dest):
            try:
                os.remove(dest)
            except Exception:
                pass
    return False, 0, "所有下载通道失败"


def self_exe_name():
    """当前运行的exe文件名 (打包时= sys.executable 文件名)"""
    if getattr(sys, "frozen", False):
        return os.path.basename(sys.executable)
    return "table-match-gui.exe"


def self_asset_name():
    """
    GitHub Release 资产标准名 (不依赖当前exe文件名!)
    用户可能把 exe 改名(如 new_table-match-gui.exe)导致按文件名找资产404
    按是否 flex 版判定:
      - 主版: table-match-gui.exe
      - flex版: table-match-gui-flex.exe
    """
    exe = self_exe_name().lower()
    if "flex" in exe:
        return "table-match-gui-flex.exe"
    return "table-match-gui.exe"


def make_updater_bat(new_exe, old_exe, temp_dir, target=None):
    """
    生成 updater.bat (纯ASCII): 等待主进程退出→替换exe→重启
    全部用绝对路径 + cd /d 到exe目录(修复Security validation/path失败)
    target: 替换目标路径(标准名)。若当前运行exe名≠标准名(new_残留), 删除旧的并启动标准名
    Windows CMD 的 GBK 编码, bat 必须纯 ASCII (windows-bat-encoding skill)
    """
    # 绝对路径(带引号处理) — Windows盘符在Linux isabs不识别, 用splitdrive兼容
    def _abs(p, base):
        if os.path.isabs(p) or (len(p) > 2 and p[1] == ":"):
            return p
        return os.path.join(base, p).replace("/", "\\")

    new_exe = _abs(new_exe, temp_dir)
    old_exe = _abs(old_exe, temp_dir)
    if target:
        target = _abs(target, temp_dir)
    else:
        target = new_exe.replace("\\new_", "\\")  # new_X → X 默认
    # tasklist IMAGENAME 只认文件名; 手动split兼容Linux测试
    old_name = old_exe.replace("/", "\\").split("\\")[-1]
    new_q = f'"{new_exe}"'
    old_q = f'"{old_exe}"'
    tgt_q = f'"{target}"'
    dir_q = f'"{temp_dir}"'
    bat = f"""@echo off
setlocal
rem updater for table-match-gui (generated)
cd /d {dir_q}
rem wait for main process to exit (max 120s)
set /a tries=0
:loop
tasklist /FI "IMAGENAME eq {old_name}" 2>nul | find /I "{old_name}" >nul
if %errorlevel%==0 (
  set /a tries+=1
  if %tries% GEQ 120 goto force
  timeout /t 1 /nobreak >nul
  goto loop
)
:force
rem 若当前exe名不是标准名(new_残留), 先删掉它
if /I not {old_q}=={tgt_q} (
  del "{old_exe}" >nul 2>&1
)
rem 移到标准名(move = 原子替换)
move /y {new_q} {tgt_q} >nul 2>&1
if %errorlevel%==0 goto restart
echo FAILED_TO_REPLACE > {dir_q}\\updater_result.txt
exit /b 1
:restart
echo OK > {dir_q}\\updater_result.txt
start "" {tgt_q}
exit /b 0
"""
    return bat


def run_updater(bat_path, new_exe, old_exe, temp_dir):
    """启动 updater 并退出当前程序 — 无窗口运行"""
    import subprocess
    try:
        if os.name == "nt":
            # CREATE_NO_WINDOW=0x08000000 彻底隐藏cmd窗口
            subprocess.Popen(["cmd", "/c", bat_path],
                             creationflags=0x08000000,
                             stdin=subprocess.DEVNULL,
                             stdout=subprocess.DEVNULL,
                             stderr=subprocess.DEVNULL,
                             close_fds=True)
        else:
            subprocess.Popen(["bash", bat_path])
    except Exception:
        try:
            os.startfile(bat_path)
        except Exception:
            pass

if __name__ == "__main__":
    info = check_update()
    print(info)