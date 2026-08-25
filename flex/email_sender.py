#!/usr/bin/env python3
"""
邮箱发送模块 - 本地配置 + DPAPI 加密 + SMTP 发送
==============================================
- 配置首次弹窗填写, 存 exe 同目录 smtp.conf (或 %APPDATA%\\table-match\\smtp.conf)
- Windows 用 DPAPI (CryptProtectData) 加密, 非 Windows 用 base64 混淆兜底
- From 头必须纯 ASCII 邮箱地址 (QQ SMTP 550 坑)
- 支持自定义 SMTP 服务器/端口/SSL/STARTTLS
"""
import os
import sys
import json
import base64
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.header import Header
from email import encoders

CONF_NAME = "smtp.conf"
APPDATA_DIR = "table-match"


# ---------- DPAPI 加解密 (ctypes 调 crypt32, 零依赖) ----------
def _dpapi_protect(data: bytes) -> bytes:
    """Windows DPAPI 加密 (仅当前用户可解)"""
    import ctypes
    import ctypes.wintypes

    class DATA_BLOB(ctypes.Structure):
        _fields_ = [("cbData", ctypes.wintypes.DWORD),
                    ("pbData", ctypes.POINTER(ctypes.c_char))]

    buf = ctypes.create_string_buffer(data, len(data))
    blob_in = DATA_BLOB(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))
    blob_out = DATA_BLOB()
    if ctypes.windll.crypt32.CryptProtectData(
            ctypes.byref(blob_in), None, None, None, None, 0, ctypes.byref(blob_out)):
        out = ctypes.string_at(blob_out.pbData, blob_out.cbData)
        ctypes.windll.kernel32.LocalFree(blob_out.pbData)
        return out
    raise RuntimeError("DPAPI protect failed")


def _dpapi_unprotect(data: bytes) -> bytes:
    import ctypes
    import ctypes.wintypes

    class DATA_BLOB(ctypes.Structure):
        _fields_ = [("cbData", ctypes.wintypes.DWORD),
                    ("pbData", ctypes.POINTER(ctypes.c_char))]

    buf = ctypes.create_string_buffer(data, len(data))
    blob_in = DATA_BLOB(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))
    blob_out = DATA_BLOB()
    if ctypes.windll.crypt32.CryptUnprotectData(
            ctypes.byref(blob_in), None, None, None, None, 0, ctypes.byref(blob_out)):
        out = ctypes.string_at(blob_out.pbData, blob_out.cbData)
        ctypes.windll.kernel32.LocalFree(blob_out.pbData)
        return out
    raise RuntimeError("DPAPI unprotect failed")


_IS_WIN = os.name == "nt"


def encrypt_text(text: str) -> str:
    """加密配置文本 → base64字符串存文件"""
    raw = text.encode("utf-8")
    try:
        if _IS_WIN:
            enc = _dpapi_protect(raw)
            return base64.b64encode(enc).decode("ascii")
    except Exception:
        pass
    # 非Windows或DPAPI失败: 单层base64(非加密, 仅防直读)
    return base64.b64encode(raw).decode("ascii")


def decrypt_text(data: str) -> str:
    """解密密文"""
    try:
        enc = base64.b64decode(data.encode("ascii"))
        if _IS_WIN:
            try:
                return _dpapi_unprotect(enc).decode("utf-8")
            except Exception:
                return enc.decode("utf-8", errors="ignore")  # 可能非DPAPI
        return enc.decode("utf-8", errors="ignore")
    except Exception:
        return ""


# ---------- 配置路径 ----------
def config_path():
    """exe 同目录优先, 无写权限回退 %APPDATA%\\table-match"""
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(os.path.abspath(sys.executable))
    else:
        exe_dir = os.path.dirname(os.path.abspath(__file__))
    # 若可写, 用 exe 同目录
    try:
        test = os.path.join(exe_dir, ".wtest")
        with open(test, "w") as f:
            f.write("")
        os.remove(test)
        return os.path.join(exe_dir, CONF_NAME)
    except Exception:
        ap = os.environ.get("APPDATA") or os.path.expanduser("~")
        d = os.path.join(ap, APPDATA_DIR)
        os.makedirs(d, exist_ok=True)
        return os.path.join(d, CONF_NAME)


# ---------- 配置读写 ----------
def load_config(path=None):
    """读取配置, 无则返回 None"""
    p = path or config_path()
    if not os.path.isfile(p):
        return None
    try:
        with open(p, "r", encoding="utf-8") as f:
            data = f.read().strip()
        if not data:
            return None
        d = json.loads(data)
        # 敏感字段解密
        for k in ("auth_code", "sender_email"):
            if d.get(k):
                d[k] = decrypt_text(str(d[k]))
        return d
    except Exception:
        return None


def save_config(cfg, path=None):
    """保存配置(敏感字段加密)"""
    p = path or config_path()
    d = dict(cfg)
    for k in ("auth_code", "sender_email"):
        if d.get(k):
            d[k] = encrypt_text(str(d[k]))
    with open(p, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=2)
    return p


def reset_config(path=None):
    """删除配置"""
    p = path or config_path()
    try:
        os.remove(p)
        return True
    except Exception:
        return False


# ---------- SMTP 发送 ----------
def send_email(cfg, to_addr, subject, body, attachment=None,
               progress_cb=None, timeout=60):
    """
    发送邮件
    cfg: {smtp_server, smtp_port, use_ssl, sender_email, auth_code}
    返回 (ok, msg)
    """
    server = (cfg.get("smtp_server") or "").strip()
    port = int(cfg.get("smtp_port") or 465)
    sender = (cfg.get("sender_email") or "").strip()
    auth = (cfg.get("auth_code") or "").strip()
    use_ssl = bool(cfg.get("use_ssl", True))
    if not server or not sender or not auth:
        return False, "SMTP 配置不完整"
    if not to_addr:
        return False, "收件人为空"

    msg = MIMEMultipart()
    msg["From"] = sender  # ⚠️ 必须纯 ASCII(QQ 550 坑), 中文显示名会拒
    msg["To"] = to_addr
    msg["Subject"] = Header(subject or "", "utf-8")
    msg.attach(MIMEText(body or "", "plain", "utf-8"))

    # 附件(可选)
    if attachment and os.path.isfile(attachment):
        fname = os.path.basename(attachment)
        with open(attachment, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", "attachment", filename=("utf-8", "", fname))
            msg.attach(part)

    try:
        if use_ssl:
            s = smtplib.SMTP_SSL(server, port, timeout=timeout)
        else:
            s = smtplib.SMTP(server, port, timeout=timeout)
            s.starttls()
        s.login(sender, auth)
        s.sendmail(sender, [to_addr], msg.as_string())
        s.quit()
        return True, "发送成功"
    except Exception as e:
        return False, str(e)


# ---------- 常用SMTP预设 ----------
PRESETS = {
    "QQ邮箱": {"smtp_server": "smtp.qq.com", "smtp_port": 465, "use_ssl": True},
    "163邮箱": {"smtp_server": "smtp.163.com", "smtp_port": 465, "use_ssl": True},
}


if __name__ == "__main__":
    # 自测: 加密解密
    t = "sk-test-abc123"
    e = encrypt_text(t)
    d = decrypt_text(e)
    print("加密解密:", "OK" if d == t else f"FAIL {d}")
    print("配置路径:", config_path())