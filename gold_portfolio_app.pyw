# -*- coding: utf-8 -*-
"""
محفظة الذهب — Gold Portfolio
تطبيق سطح مكتب (Tkinter) لتتبّع محفظة الذهب: شراء / بيع / ربح وخسارة /
حاسبة زكاة / سجل عمليات / مخطط سعر الغرام. يستخدم Requests لطلبات أسعار السوق
(التصدير إلى Excel اختياري إن توفّر openpyxl).
"""
import os
import json
import time
import hashlib
import hmac
import shutil
import threading
import queue
try:
    import requests
except ImportError:
    requests = None
import math
import copy
from live_chart import clean_points, render_chart, hover_chart
import re
import sys
import glob
import base64
import subprocess
import tempfile
from datetime import datetime, timezone
import tkinter as tk
from tkinter import ttk, messagebox, colorchooser, filedialog
import urllib.request
import urllib.error
import urllib.parse
import socket
import smtplib
import imaplib
import ssl
import webbrowser
from email.utils import parsedate_to_datetime
from email.mime.text import MIMEText
from email import message_from_bytes

# ============================ ثوابت عامة ============================
OZ_TO_G = 31.1034768                      # غرام في الأونصة الواحدة
APP_VERSION = "PriceFix-18-Sidebar"
HTTP_BACKEND = "Requests" if requests is not None else "urllib (بديل تلقائي)"
LIVE_MS = 5000                            # تحديث السعر المباشر كل 5 ثوانٍ
FLASH_HOLD = 350                          # مدة بقاء وميض تغيّر السعر (مللي ثانية)
HOME_DIR = os.path.expanduser("~")
# ملف واحد للبيانات القديمة (قبل نظام الحسابات) — يُتبنّى عند إنشاء أول حساب
DATA_FILE = os.path.join(HOME_DIR, "gold_portfolio_data.json")
# ملف الحسابات (أسماء المستخدمين + بصمات كلمات المرور)
USERS_FILE = os.path.join(HOME_DIR, "gold_portfolio_users.json")
# تفضيلات بسيطة (اسم المستخدم المتذكّر فقط — لا تُحفظ كلمة المرور أبداً)
PREFS_FILE = os.path.join(HOME_DIR, "gold_portfolio_prefs.json")


def load_prefs():
    try:
        with open(PREFS_FILE, "r", encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def save_prefs(prefs):
    try:
        tmp = PREFS_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(prefs, f, ensure_ascii=False, indent=2)
        os.replace(tmp, PREFS_FILE)
        return True
    except Exception:
        return False


def user_data_file(username):
    """ملف بيانات مستقل لكل مستخدم (لا يرى أحد بيانات غيره)."""
    safe = "".join(c if (c.isalnum() or c in "-_.") else "_" for c in (username or ""))
    return os.path.join(HOME_DIR, "gold_portfolio_data_{}.json".format(safe.lower() or "user"))


def load_users():
    try:
        with open(USERS_FILE, "r", encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def save_users(users):
    try:
        tmp = USERS_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(users, f, ensure_ascii=False, indent=2)
        os.replace(tmp, USERS_FILE)
        return True
    except Exception:
        return False


def hash_password(password, salt_hex=None):
    """بصمة كلمة المرور بـ PBKDF2-SHA256 مع ملح عشوائي (لا تُخزَّن كلمة المرور أبداً)."""
    salt = bytes.fromhex(salt_hex) if salt_hex else os.urandom(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 200000)
    return salt.hex(), dk.hex()


def verify_password(password, salt_hex, hash_hex):
    try:
        _, calc = hash_password(password, salt_hex)
        return hmac.compare_digest(calc, hash_hex or "")
    except Exception:
        return False

# ============================ المزامنة السحابية (Firebase) ============================
# تُحفظ إعدادات المزامنة في ملف مستقل حتى لا تتأثر ببيانات المحفظة
CLOUD_FILE = os.path.join(HOME_DIR, "gold_portfolio_cloud.json")
FB_SIGNUP = "https://identitytoolkit.googleapis.com/v1/accounts:signUp?key={key}"
FB_SIGNIN = ("https://identitytoolkit.googleapis.com/v1/accounts:"
             "signInWithPassword?key={key}")
FB_REFRESH = "https://securetoken.googleapis.com/v1/token?key={key}"
FB_RESET = ("https://identitytoolkit.googleapis.com/v1/accounts:"
            "sendOobCode?key={key}")

# ==================== إعداد الخادم (مرة واحدة فقط) ====================
# الصق هنا رابط قاعدة بيانات مشروعك في Firebase ومفتاح الويب.
# بعد ذلك لا يرى المستخدم أي رابط: يسجّل بالبريد وكلمة المرور فقط،
# وتُسترجع بياناته تلقائياً على أي جهاز.
FB_URL = ""
FB_KEY = ""

# ==================== المزامنة عبر البريد (Gmail) ====================
# لا قاعدة بيانات ولا رابط: تُحفظ نسخة مشفّرة في رسالة داخل بريدك نفسه،
# ويستعيدها أي جهاز يسجّل بالبريد نفسه + كلمة مرور التطبيق.
GMAIL_SMTP = "smtp.gmail.com"
GMAIL_IMAP = "imap.gmail.com"
GMAIL_PORT_SMTP = 465
GMAIL_PORT_IMAP = 993
SYNC_SUBJECT = "GOLDPORTFOLIO-SYNC"
SYNC_MARK = "===== GOLDPORTFOLIO SYNC ====="

FB_ERRORS = {
    "EMAIL_EXISTS": "هذا البريد مسجَّل سابقاً — استخدم زر «دخول» بدل «إنشاء حساب».",
    "EMAIL_NOT_FOUND": "لا يوجد حساب سحابي بهذا البريد.",
    "INVALID_PASSWORD": "كلمة المرور غير صحيحة.",
    "INVALID_LOGIN_CREDENTIALS": "البريد أو كلمة المرور غير صحيحة.",
    "WEAK_PASSWORD": "كلمة المرور قصيرة — يجب أن تكون ٦ أحرف على الأقل.",
    "OPERATION_NOT_ALLOWED": "تسجيل الدخول بالبريد غير مُفعّل — فعّل Email/Password "
                             "في مشروع Firebase.",
    "CONFIGURATION_NOT_FOUND": "مفتاح المشروع (API key) غير صحيح.",
    "API_KEY_INVALID": "مفتاح الويب (API key) غير صحيح.",
    "INVALID_EMAIL": "صيغة البريد الإلكتروني غير صحيحة.",
    "MISSING_PASSWORD": "أدخل كلمة المرور.",
    "PERMISSION_DENIED": "لا توجد صلاحية للوصول إلى قاعدة البيانات — راجع قواعد "
                         "الأمان في Firebase.",
}

# المفاتيح التي تُزامن بين الأجهزة (أسعار السوق والسجل اللحظي تبقى محلية)
SYNC_KEYS = ("holdings", "sells", "history", "alerts", "cur", "karat",
             "chart_range", "adjust", "buy_margin", "theme", "titles",
             "sound_on", "source_mode", "dark_mode", "lock_minutes", "chart_font_size")


class CloudError(Exception):
    """خطأ في المزامنة السحابية برسالة عربية جاهزة للعرض."""

    def __init__(self, msg, status=0):
        super().__init__(msg)
        self.status = status


def load_cloud():
    """يقرأ إعدادات المزامنة (روابط/رموز الدخول) من ملفها المستقل."""
    try:
        with open(CLOUD_FILE, "r", encoding="utf-8") as f:
            d = json.load(f)
        if not isinstance(d, dict):
            raise ValueError
    except Exception:
        d = {}
    d.setdefault("url", "")
    d.setdefault("key", "")
    d.setdefault("auto", True)
    if not isinstance(d.get("accounts"), dict):
        d["accounts"] = {}
    return d


def save_cloud(conf):
    """حفظ إعدادات المزامنة كتابةً ذرّية مع تقييد صلاحيات الملف."""
    try:
        tmp = CLOUD_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(conf, f, ensure_ascii=False, indent=2)
        os.replace(tmp, CLOUD_FILE)
        try:
            os.chmod(CLOUD_FILE, 0o600)
        except Exception:
            pass
        return True
    except Exception:
        return False


def cloud_http(method, url, payload=None, headers=None, timeout=15):
    """طلب JSON إلى خدمات Firebase (يعيد البيانات و ETag)."""
    body = None
    h = {"User-Agent": BROWSER_UA, "Accept": "application/json"}
    if payload is not None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        h["Content-Type"] = "application/json"
    h.update(headers or {})
    req = urllib.request.Request(url, data=body, headers=h, method=method)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            raw = r.read().decode("utf-8", "ignore")
            etag = r.headers.get("ETag") or r.headers.get("Etag") or ""
            return (json.loads(raw) if raw.strip() else None), etag
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", "ignore")
        code = ""
        try:
            err = json.loads(raw).get("error", {})
            if isinstance(err, dict):
                code = str(err.get("message", "") or "").split(" ")[0]
                if not code:
                    msg = err.get("errors", [{}])[0].get("message", "")
                    code = str(msg)
            else:
                code = str(err)
        except Exception:
            code = ""
        if e.code == 412:
            raise CloudError("تغيّرت البيانات في السحابة من جهاز آخر.", 412)
        if e.code == 401:
            raise CloudError("انتهت صلاحية الجلسة — أعد ربط الحساب السحابي.", 401)
        if e.code == 403:
            raise CloudError(FB_ERRORS["PERMISSION_DENIED"], 403)
        raise CloudError(FB_ERRORS.get(code) or
                         "خطأ في الاتصال بالسحابة ({}).".format(e.code), e.code)
    except Exception as e:
        raise CloudError("تعذّر الاتصال بالإنترنت: {}".format(e))


# ---------------- تعمية نسخة البريد (مكتبة قياسية فقط) ----------------
def _sync_keys(secret, email):
    """يشتق مفتاحَي التعمية والتحقق من كلمة مرور التطبيق."""
    pw = "".join((secret or "").split())
    raw = hashlib.pbkdf2_hmac("sha256", pw.encode("utf-8"),
                              ("goldport|" + (email or "").lower()).encode("utf-8"),
                              120000, dklen=64)
    return raw[:32], raw[32:]


def _keystream(key, nonce, n):
    out = bytearray()
    i = 0
    while len(out) < n:
        out += hmac.new(key, nonce + i.to_bytes(8, "big"), hashlib.sha256).digest()
        i += 1
    return bytes(out[:n])


def seal_sync(obj, secret, email):
    """يحوّل البيانات إلى نص مشفّر (AES-less: HMAC-SHA256 CTR + MAC)."""
    enc, mac = _sync_keys(secret, email)
    data = json.dumps(obj, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    nonce = os.urandom(16)
    ct = bytes(a ^ b for a, b in zip(data, _keystream(enc, nonce, len(data))))
    tag = hmac.new(mac, b"GP1" + nonce + ct, hashlib.sha256).digest()
    return base64.b64encode(b"GP1" + nonce + ct + tag).decode("ascii")


def open_sync(blob, secret, email):
    """يفكّ نصاً مشفّراً أنتجته seal_sync."""
    enc, mac = _sync_keys(secret, email)
    txt = "".join(str(blob or "").split())
    if not txt:
        raise CloudError("لا توجد نسخة في بريدك بعد.")
    try:
        raw = base64.b64decode(txt)
    except Exception:
        raise CloudError("نسخة البريد غير صالحة.")
    if len(raw) < 3 + 16 + 32 or not raw.startswith(b"GP1"):
        raise CloudError("نسخة البريد غير صالحة.")
    nonce, ct, tag = raw[3:19], raw[19:-32], raw[-32:]
    calc = hmac.new(mac, b"GP1" + nonce + ct, hashlib.sha256).digest()
    if not hmac.compare_digest(tag, calc):
        raise CloudError("تعذّر فتح النسخة — كلمة مرور التطبيق غير مطابقة.")
    data = bytes(a ^ b for a, b in zip(ct, _keystream(enc, nonce, len(ct))))
    try:
        return json.loads(data.decode("utf-8"))
    except Exception:
        raise CloudError("تعذّر قراءة بيانات النسخة.")


def gmail_check(email, pw):
    """يتأكد أن البريد وكلمة مرور التطبيق يعملان (IMAP)."""
    email = (email or "").strip()
    pw = "".join((pw or "").split())
    if "@" not in email or "." not in email:
        raise CloudError("أدخل بريد Gmail صحيحاً.")
    if len(pw) < 8:
        raise CloudError("أدخل «كلمة مرور التطبيق» (16 حرفاً) لا كلمة مرور حسابك.")
    M = None
    try:
        M = imaplib.IMAP4_SSL(GMAIL_IMAP, GMAIL_PORT_IMAP, timeout=25)
        M.login(email, pw)
        return True
    except imaplib.IMAP4.error as e:
        msg = str(e).lower()
        if "auth" in msg or "credentials" in msg or "login" in msg:
            raise CloudError("البريد أو كلمة مرور التطبيق غير صحيحة.\n"
                             "استخدم «كلمة مرور التطبيق» من إعدادات أمان Google.")
        raise CloudError("تعذّر الاتصال بالبريد: {}".format(e))
    except (socket.error, ssl.SSLError, OSError) as e:
        raise CloudError("تعذّر الاتصال بالبريد — تحقق من الإنترنت: {}".format(e))
    finally:
        if M is not None:
            try:
                M.logout()
            except Exception:
                pass


def gmail_send(email, pw, blob):
    """يرسل نسخة المزامنة إلى البريد نفسه."""
    email = (email or "").strip()
    pw = "".join((pw or "").split())
    msg = MIMEText(SYNC_MARK + "\n" + blob, "plain", "utf-8")
    msg["Subject"] = SYNC_SUBJECT
    msg["From"] = email
    msg["To"] = email
    msg["X-Gold-Sync"] = "1"
    try:
        with smtplib.SMTP_SSL(GMAIL_SMTP, GMAIL_PORT_SMTP, timeout=30) as s:
            s.login(email, pw)
            s.send_message(msg)
        return True
    except smtplib.SMTPAuthenticationError:
        raise CloudError("رفض Gmail كلمة مرور التطبيق — أنشئ واحدة جديدة.")
    except (smtplib.SMTPException, socket.error, ssl.SSLError, OSError) as e:
        raise CloudError("تعذّر إرسال النسخة إلى بريدك: {}".format(e))


def mail_body(raw):
    """يستخرج نص الرسالة من بايتاتها الخام (يفكّ ترميز base64)."""
    try:
        msg = message_from_bytes(raw)
        data = msg.get_payload(decode=True)
        if data is None:
            data = b""
        return data.decode("utf-8", "ignore")
    except Exception:
        return raw.decode("utf-8", "ignore")


def gmail_blobs(email, pw, limit=6):
    """يعيد نصوص أحدث نسخ المزامنة الموجودة في البريد (الأحدث أولاً)."""
    email = (email or "").strip()
    pw = "".join((pw or "").split())
    M = None
    out = []
    try:
        M = imaplib.IMAP4_SSL(GMAIL_IMAP, GMAIL_PORT_IMAP, timeout=30)
        M.login(email, pw)
        M.select("INBOX", readonly=True)
        typ, data = M.search(None, '(SUBJECT "{}")'.format(SYNC_SUBJECT))
        ids = (data[0].split() if data and data[0] else [])
        for i in ids[-limit:][::-1]:
            typ, d = M.fetch(i, "(BODY.PEEK[])")
            raw = b""
            for part in (d or []):
                if isinstance(part, tuple) and len(part) > 1:
                    raw += part[1]
            text = mail_body(raw)
            if SYNC_MARK in text:
                out.append(text.split(SYNC_MARK, 1)[1].strip())
        return out
    except imaplib.IMAP4.error as e:
        raise CloudError("تعذّر قراءة بريدك: {}".format(e))
    except (socket.error, ssl.SSLError, OSError) as e:
        raise CloudError("تعذّر الاتصال بالبريد — تحقق من الإنترنت: {}".format(e))
    finally:
        if M is not None:
            try:
                M.close()
            except Exception:
                pass
            try:
                M.logout()
            except Exception:
                pass


# العملات المدعومة (أُزيل الدرهم الإماراتي وأُضيف الدينار الأردني)
CUR = {
    "USD": {"name": "دولار أمريكي", "sym": "$"},
    "SAR": {"name": "ريال سعودي", "sym": "ر.س "},
    "JOD": {"name": "دينار أردني", "sym": "د.أ "},
    "KWD": {"name": "دينار كويتي", "sym": "د.ك "},
    "QAR": {"name": "ريال قطري", "sym": "ر.ق "},
    "EGP": {"name": "جنيه مصري", "sym": "ج.م "},
}
# أسعار صرف تقريبية مقابل الدولار (تُحدّث تلقائياً من الإنترنت)
DEFAULT_RATES = {"USD": 1.0, "SAR": 3.75, "JOD": 0.709, "KWD": 0.306,
                 "QAR": 3.64, "EGP": 48.5}

UNIT_MAP = {"غرام": "g", "أونصة": "oz"}
UNIT_REV = {"g": "غرام", "oz": "أونصة"}

# العيارات المدعومة (ذهب خالص = 24)
KARATS = (24, 22, 21, 18)

# النطاقات الزمنية للمخطط: (المفتاح، الاسم، عدد الثواني)
CHART_RANGES = (("live", "دقيقة", 60),
                ("quarter", "١٥ دقيقة", 900),
                ("hour", "ساعة", 3600),
                ("day", "يوم", 86400),
                ("week", "أسبوع", 7 * 86400),
                ("month", "شهر", 30 * 86400))


def karat_factor(karat):
    """معامل نقاوة العيار نسبةً إلى الذهب الخالص 24."""
    try:
        k = float(karat)
    except (TypeError, ValueError):
        k = 24.0
    return k / 24.0

LEGACY_DARK_THEME = {
    "bg": "#0a0c11",
    "card": "#12161e",
    "card2": "#1a2029",
    "accent": "#e9c684",
    "accent_text": "#151008",
    "text": "#eef2f9",
    "muted": "#9ba6ba",
    "profit": "#3ddc97",
    "loss": "#ff6b6b",
    "total": "#ffffff",
    "total_bg": "#1b2230",
    "border": "#252c39",
    "entry_bg": "#151a23",
    "row_alt": "#141922",
    "hover": "#f3d8a5",
    "flash_same": "#ffd93b",
    "sold_bg": "#12301f",
}
LEGACY_LIGHT_THEME = {
    "bg": "#f4f6fa",
    "card": "#ffffff",
    "card2": "#e9eef7",
    "accent": "#a8761d",
    "accent_text": "#ffffff",
    "text": "#161b26",
    "muted": "#46536b",
    "profit": "#0f8a4f",
    "loss": "#c62828",
    "total": "#0b1220",
    "total_bg": "#eef1f7",
    "border": "#d7dde9",
    "entry_bg": "#ffffff",
    "row_alt": "#f5f7fb",
    "hover": "#c9922e",
    "flash_same": "#b8860b",
    "sold_bg": "#e6f6ec",
}
DEFAULT_THEME = {
    "bg": "#0e1522",
    "card": "#162031",
    "card2": "#1c293c",
    "accent": "#d8b775",
    "accent_text": "#19202b",
    "text": "#edf2fa",
    "muted": "#a4b1c5",
    "profit": "#40ce98",
    "loss": "#f27d85",
    "total": "#ffffff",
    "total_bg": "#24354c",
    "border": "#2c3b50",
    "entry_bg": "#121d2d",
    "row_alt": "#19263a",
    "hover": "#e7ca95",
    "flash_same": "#d8b775",
    "sold_bg": "#16312f",
}
LIGHT_THEME = {
    "bg": "#f2f5fa",
    "card": "#ffffff",
    "card2": "#edf2f8",
    "accent": "#987022",
    "accent_text": "#ffffff",
    "text": "#203047",
    "muted": "#5e6d82",
    "profit": "#137c58",
    "loss": "#bc3e4b",
    "total": "#203047",
    "total_bg": "#e8eff8",
    "border": "#d8e1ed",
    "entry_bg": "#ffffff",
    "row_alt": "#f4f7fb",
    "hover": "#ad843c",
    "flash_same": "#987022",
    "sold_bg": "#e7f4ef",
}
DEFAULT_TITLES = {
    "title": "✦ محفظة الذهب ✦",
    "subtitle": "تتبّع الربح والخسارة مع جلب السعر مباشرة من الإنترنت",
    "window": "محفظة الذهب - Gold Portfolio",
}
THEME_VER = 5

# كثافة شاشة ويندوز (96 = 100%). تُضبط عند الإقلاع لدعم شاشات 125% و150%.
SYSTEM_DPI = 96

BROWSER_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/122.0 Safari/537.36")


# ============================ أدوات مساعدة ============================
_market_http_local = threading.local()


def _market_session():
    # Sessions are private to each worker, never shared across threads.
    session = getattr(_market_http_local, "session", None)
    if session is None:
        session = requests.Session()
        session.headers.update({"User-Agent": BROWSER_UA,
                                "Accept": "application/json, text/plain, */*",
                                "Cache-Control": "no-cache, no-store, max-age=0",
                                "Pragma": "no-cache"})
        # No hidden HTTP retries: the five-second supervisor handles retrying.
        adapter = requests.adapters.HTTPAdapter(max_retries=0,
                                                pool_connections=2, pool_maxsize=2)
        session.mount("https://", adapter)
        _market_http_local.session = session
    return session


def _close_market_session():
    session = getattr(_market_http_local, "session", None)
    if session is not None:
        try:
            session.close()
        finally:
            del _market_http_local.session


def _market_get(url, timeout, as_json):
    budget = float(timeout)
    if not math.isfinite(budget) or budget <= 0:
        raise ValueError("مهلة طلب غير صحيحة")
    # Separate connect/read limits; DNS and redirects can still exceed them.
    limits = (budget / 2.0, budget / 2.0)
    if requests is None:
        # Startup never depends on installing an optional third-party package.
        req = urllib.request.Request(url, headers={
            "User-Agent": BROWSER_UA,
            "Accept": "application/json, text/plain, */*",
            "Cache-Control": "no-cache, no-store, max-age=0",
            "Pragma": "no-cache"})
        with urllib.request.urlopen(req, timeout=budget) as response:
            _remember_response_clock(response.headers.get("Date"))
            text = response.read().decode("utf-8", "replace")
            return json.loads(text) if as_json else text
    try:
        with _market_session().get(url, timeout=limits, verify=True) as response:
            response.raise_for_status()
            _remember_response_clock(response.headers.get("Date"))
            if as_json:
                return response.json()
            response.encoding = response.encoding or "utf-8"
            return response.text
    except requests.exceptions.Timeout as exc:
        raise RuntimeError("Requests: انتهت مهلة الاتصال أو قراءة المصدر") from exc
    except requests.exceptions.SSLError as exc:
        raise RuntimeError("Requests: تعذّر التحقق من شهادة HTTPS") from exc
    except requests.exceptions.HTTPError as exc:
        status = exc.response.status_code if exc.response is not None else "?"
        raise RuntimeError("Requests: رفض المصدر الطلب HTTP {}".format(status)) from exc
    except requests.exceptions.RequestException as exc:
        raise RuntimeError("Requests: تعذّر الاتصال بمصدر السعر") from exc
    except ValueError as exc:
        raise RuntimeError("Requests: استجابة المصدر ليست JSON صالحاً") from exc


def http_json(url, timeout=10):
    return _market_get(url, timeout, as_json=True)


def _http_text(url, timeout=10):
    return _market_get(url, timeout, as_json=False)


def _valid_price(value):
    """Reject non-numeric, non-finite and non-positive market quotes."""
    if isinstance(value, bool):
        raise ValueError("قيمة سعر غير صحيحة")
    price = float(value)
    if not math.isfinite(price) or price <= 0:
        raise ValueError("قيمة سعر غير صحيحة")
    return price


def _yahoo_last(symbol, timeout=2):
    data = http_json("https://query1.finance.yahoo.com/v8/finance/chart/{}"
                     "?interval=1m&range=1d".format(symbol), timeout=timeout)
    res = (data.get("chart") or {}).get("result") or []
    if not res:
        raise ValueError("لا توجد بيانات")
    price = (res[0].get("meta") or {}).get("regularMarketPrice")
    if price is None:
        q = ((res[0].get("indicators") or {}).get("quote") or [{}])[0]
        closes = [c for c in (q.get("close") or []) if c is not None]
        price = closes[-1] if closes else None
    return _valid_price(price)


QUOTE_MAX_AGE = 20.0


class MarketQuote(tuple):
    """Backward-compatible (price, label), with verified source timestamp."""
    def __new__(cls, price, label, source_id, quote_time):
        obj = super().__new__(cls, (price, label))
        obj.source_id = source_id
        obj.quote_time = quote_time
        return obj


def _remember_response_clock(value):
    _market_http_local.server_time = None
    if isinstance(value, str):
        try:
            _market_http_local.server_time = parsedate_to_datetime(value).timestamp()
        except (ValueError, TypeError, OverflowError):
            pass


def _quote_epoch(value):
    if isinstance(value, bool) or value is None:
        raise ValueError("وقت التسعير مفقود")
    if isinstance(value, str) and not value.replace(".", "", 1).isdigit():
        stamp = datetime.fromisoformat(value.replace("Z", "+00:00"))
        if stamp.tzinfo is None:
            raise ValueError("منطقة وقت التسعير مفقودة")
        value = stamp.timestamp()
    value = float(value)
    if value > 1e11:
        value /= 1000.0
    if not math.isfinite(value) or value <= 0:
        raise ValueError("وقت التسعير غير صحيح")
    return value


def _fresh_quote(value, source_id, label, timestamp):
    stamp = _quote_epoch(timestamp)
    now = getattr(_market_http_local, "server_time", None) or time.time()
    age = now - stamp
    if age > QUOTE_MAX_AGE or age < -10:
        raise ValueError("بيانات المصدر قديمة أو وقتها غير صحيح (العمر {:.0f} ثانية)".format(age))
    utc = datetime.fromtimestamp(stamp, timezone.utc).strftime("%H:%M:%S UTC")
    return MarketQuote(_valid_price(value), label + " • وقت السعر " + utc + " • " + HTTP_BACKEND,
                       source_id, stamp)


def _swissquote(data):
    if not isinstance(data, list):
        raise ValueError("استجابة Swissquote غير صحيحة")
    # Use one stable platform/profile, not whichever spread happens to move.
    rows = [r for r in data if isinstance(r, dict) and
            (r.get("topo") or {}).get("platform") == "AT"]
    if not rows:
        raise ValueError("منصة AT غير متوفرة")
    row = max(rows, key=lambda r: _quote_epoch(r.get("ts")))
    prices = [q for q in row.get("spreadProfilePrices", [])
              if q.get("spreadProfile") == "standard"]
    if not prices:
        raise ValueError("تسعير standard غير متوفر")
    bid, ask = _valid_price(prices[0].get("bid")), _valid_price(prices[0].get("ask"))
    if ask < bid:
        raise ValueError("سعر العرض أقل من سعر الطلب")
    # Broker XAU/USD quote midpoint, not a futures or a token price.
    return _fresh_quote((bid + ask) / 2, "swissquote",
                        "Swissquote XAU/USD • متوسط Bid/Ask", row.get("ts"))


def fetch_spot_usd(mode="auto", timeout=2, only=False):
    timeout = max(0.1, min(float(timeout), 6.0))
    spot_names = ("swissquote", "gold-api", "goldprice.org")
    names = list(spot_names)
    if only:
        if mode not in spot_names + ("stooq", "yahoo", "paxg"):
            raise ValueError("مصدر غير معروف")
        names = [mode]
    elif mode in spot_names:
        names.remove(mode)
        names.insert(0, mode)
    deadline = time.monotonic() + (timeout if only else min(4.0, timeout * 2))
    errors = []
    for name in names:
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            break
        limit = min(timeout, remaining)
        try:
            # Do not reuse another response's HTTP Date in freshness checks.
            _market_http_local.server_time = None
            if name == "swissquote":
                data = http_json("https://forex-data-feed.swissquote.com/public-quotes/bboquotes/instrument/XAU/USD",
                                 timeout=limit)
                quote = _swissquote(data)
            elif name == "gold-api":
                data = http_json("https://api.gold-api.com/price/XAU", timeout=limit)
                if data.get("symbol") not in (None, "XAU") or data.get("currency") not in (None, "USD"):
                    raise ValueError("المعدن أو العملة غير صحيح")
                quote = _fresh_quote(data.get("price"), name, "gold-api", data.get("updatedAt"))
            elif name == "goldprice.org":
                data = http_json("https://data-asg.goldprice.org/dbXRates/USD", timeout=limit)
                item = data["items"][0]
                if item.get("curr") not in (None, "USD"):
                    raise ValueError("عملة المصدر ليست الدولار")
                quote = _fresh_quote(item["xauPrice"], name, name, data.get("ts") or item.get("ts"))
            elif name == "yahoo":
                quote = (_yahoo_last("GC=F", timeout=limit), "yahoo — عقود آجلة للتشخيص فقط")
            elif name == "paxg":
                value = http_json("https://api.coingecko.com/api/v3/simple/price?ids=pax-gold&vs_currencies=usd",
                                  timeout=limit)["pax-gold"]["usd"]
                quote = (_valid_price(value), "paxg — رمز رقمي للتشخيص فقط")
            else:
                fields = _http_text("https://stooq.com/q/l/?s=xauusd&f=sd2t2ohlcv&h&e=csv",
                                    timeout=limit).strip().splitlines()[-1].split(",")
                if fields[0].strip().upper() != "XAUUSD":
                    raise ValueError("رمز المصدر غير صحيح")
                quote = (_valid_price(fields[6]), "stooq — بيانات قد تكون متأخرة؛ للتشخيص فقط")
            if time.monotonic() > deadline:
                raise TimeoutError("انتهت مهلة جلب السعر")
            previous = getattr(_market_http_local, "previous_quote", None)
            if (isinstance(quote, MarketQuote) and previous and
                    previous[1] is not None and quote.quote_time <= previous[1]):
                raise ValueError("المصدر يعيد وقت التسعير السابق؛ تجربة مصدر بديل")
            return quote
        except Exception as exc:
            errors.append("{}: {}".format(name, exc))
    raise RuntimeError("تعذّر جلب سعر حديث: " + ("; ".join(errors) or "انتهت المهلة"))


def market_closed():
    """سوق الذهب العالمي مغلق (السبت والأحد المبكر)."""
    wd = datetime.utcnow().weekday()   # 0=الاثنين ... 5=السبت 6=الأحد
    return wd == 5 or (wd == 6 and datetime.utcnow().hour < 22)


# ==================== أيقونة التطبيق (مضمّنة داخل الملف) ====================
# صورة ICO لعملة ذهبية، مضمّنة كنص Base64 حتى يعمل الملف وحده دون ملفات مساعدة.
APP_ICON_ICO_B64 = 'AAABAAcAEBAAAAAAIAB4AgAAdgAAABgYAAAAACAAbgQAAO4CAAAgIAAAAAAgAMcGAABcBwAAMDAAAAAAIADhDAAAIw4AAEBAAAAAACAA4BMAAAQbAACAgAAAAAAgAPY4AADkLgAAAAAAAAAAIAAyxwAA2mcAAIlQTkcNChoKAAAADUlIRFIAAAAQAAAAEAgGAAAAH/P/YQAAAj9JREFUeJxNk8FqFEEQhr/q6ZnZzSYbzSaBJCYSE2UlkJwUxShmLxE8eRH0DTx58yzEqw8gvoUv4FUERfCg4CWYg6gQs0Z2Z2emuzz0zmYaipkpur+uqfp/WV67qM3UIiIA4SkgCCIyyasqioIqquF7kBXYZmIRIC/cZHMAhJcqpcp4Kd4raWJpNWKsCOSlJy/8ZHMF0PFtgSYIgeJVERHSJMIgILUwRlCvZKMSEKw12MhgBKIoQsRg5KwyQ22JCNmoYH11gVcHj5lpNRgMCwbDHBD+DUbo2b/UAQEXGSEvHPfubLG5PEvv1hYLc9PcvXEVVWXv+mWsNdQZpjqMQJaXrC7Nsbkyy7OXb9jf7fLk0S73ezsURcHD/W1iG+FrBAMKKJEI2chxe+cC56YsvZtdOu0Gw0HGcDBkvtOmzHOKosRMug22ut6rMtWwLHVavHj9lq+Hv9j7/J3tjQ42tjx/+oCjwyOc86HbhJDNKxtalEpeOACmWynOefqnGV49jSRmtj1Dv39CkjZR9eR5TiNNSGJTVRDGN8hKetdWmDvf4sOXH6TWsHapS5Ik9Pt/sTbi3fuPjCYtEOxYMjjnabcSfh+fcjrM6a4v4L1QOiXLMhYX5zn+c4LzrjZErSoIyiqd59O3nzjvacQW5z3GmLEiQ0RRhDFn8rGVzlXBe8VGQXnOKyIG73XcL8FI0ICvJA5YVSWJI0AwEy/UnDkpVsPEUbxCGge4HY5KmqkliQ2IVL4JdjbCOFOz8xjig53/AyljAlKAIvo6AAAAAElFTkSuQmCCiVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAENUlEQVR4nH2WTWxUVRTHf+fOe49pOxWKLdLSUihoCSNYK8iX3xhJBEzUxBhJjAsTjbrUGPcsXLl3wUKX7tTExIUxLEwIpgQFIwH6EaCWCp12ptP5ePPePS7em5n3hsa7ePee8+49H/9z7jlXhraPqiq4jqE1BKS5SPESdGIogGqK5zdCRMBRhU05D2MMVhWRWLTIg2sBIwIioKAoqgoazdFSMUaw1lIoVnFcx2CM4d5KJTrcslIT34StLYakpuQytMpQfy+e6+AAWFUyxmCMoJo8I5FQ6RTR8T9pSbzFxpA5rW0tyxRNeNGER2MY1hWe5Gqam4hs0oRoLQJhaCmVa3HQkh50KlvfSZOiEmdEBL9h6d/cy9dnz3BschdrVZ+MMRGcrQSAjFlfBsQQpUe0wxhhrVLjrVMnOHl4jK4NHr9dmkZVKJUruG6GrqxH3Q/xfZ+ebm9dJ9q6RdAWNEIQWAYe3shrL+7lq3M/s3f3IEcOjFNYKfL5h6/y8rE8heUyB/LDfHzmWUKrLenrxCDNNWIoV+ocP7qHenmVs+fOc2HqBmdOH+S91w/xzumDTORHqZSr7Nq2iVPP57H2wYAnFGgKPKuK52Y4cWgnt/+5z0tH9rBULHN4/whLxQp/35inWmtgMoZ6YFlcXELVxtiko5yKgRClZLXeYM+Ofob6sswsrPLR28/gW6hWa3iOw9L9ZQb6erB+lcGBjfi1esf9WU9BXApEILCWN18YZ26hyCdf/kjWy7BSqvDZ+6/w7skJvv/1Kp9+cIrhb75gZCDHzI0ZTKf1say2ghih0Co9WY9fpm4xO/8nXZ5LNpvByeT47qcprk8vcPHKHHN3Szy2e4SLf8zgGSWTMahGV1TimKoqsn3nTs11uRRW64ncFtaqPpP5YW7O3aNS9XFchzAMKSwX2T22nSCE6elZcr3djO0YZblQIAgDjAiBhS193ZTWaskYNCup4DdCRrY+xBvH83z7wxSO41JaXWODa3jqyX08d/RpLvx+icn94xSWV+jtzXG5tEIQRlALNoZI2gqatV5QwtByYHwrt+/c59jEKL5mmJmdZ/LxR1koZ7l2/SbDQ4N0d3fxxL48C4v/cuXqtUStasfCadsPqBIq9Pa4LC6VOX/5FnvHtuB6LhP5HUz9NUOub5htg1uYu3WHocFHmF+4S6m0ShAmalU8qZKIQalGxkhcSZVGYDEihNa20ltV42yBIAxxHAe1FhBc12khEVplYHM3pXK9oxaJgEYKPDcDCo6TrofN4bluqiRE6Gir3DRxSdwD02JCs/FokyR9jbTlaWcMacKkEdeBqM+qKtYmGk+qWUWlJGpMzR+2LSze1aSiDhlRTiMICaylf2M2Yay052TjN226eZGizIkaPhpBZERoBCH1RoAjIhTLPq5jotaIPvg8+Z8nS9O/JPSK0miEGBH+A6BP3TNxckozAAAAAElFTkSuQmCCiVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAGjklEQVR4nIWXyY9cVxWHv3Pfq1dDd7t68tDGdjtu24lj43bsJoltYoKwcTASCkJKBLGEsiCwYpNl2CAL/gA2iBWTkJDYwIJJBCESJYoTDwTjQVHspoMhsd3d7uqu7qpX9d49LN5cXU6u9HTPHd490++ce65s3TGpAGGoIBSbKr1t3UxhjxS6iFBAkLhPznCdiHaTM8bqFUquibZJtCgiCAIS0xLTSDoWIzEDUBRVBQVVTT/QSE6NlOx0QxaWWoiAG4bKWL2C6xgWl9sYY3o00ligTINkqWAwkR7zxEzT/6LF0Crj9Spjw1XuLq7iIlByI+bdwCJi4x+Sn4pGX++C3olETilsSEbWwkKjxabRgcwFKDiOIQgjRbIfk0F2WC9M+k/0wUUyMopjTOKNWADpxVuvWv3U7NfWSdO3aa43xWnbc1CRsSA4jtODk96je7XvJ1TiFkkEyJn6AUoIEKpyv7HGymo7jZRPbg8ECaD0qKKI9m6KJA1CpT5Y4QcvP8tLzz9FEISIPMgSH8e8uOr2btSesAEwjqG5ssaLz3+Os18+RKPp8/aVOa7cuE2t6qXhpjkgRYEghblCi41tsrGAmHXMQQiDkPpQla9/5TiXL9/AX2vxwlefwu90MGLoBhbfD2JkG8AQWvA73X58iTNVpFxBHF0PItcxrKy1OfHEPqa2VPjV7y/z+sX3+dozR9ixbRPz9xt88dhe/vCzl3lkaoKW30FV+d5LX+CV7zxDJwjJw0UTYErUuwWH9GkKOMbw7KmDzM3+l79emGO5HfKts6c5eXw/QXsnM/t3cPjATkbqAwRhSFkdpj41jFsbIgwtlNyMQT4GhV4QFpsRodXu8Miuzczs3cRv/nyZ2/eWef3iLG9euMFzX3qMyYlhPlpYYXm5iYqJdDRCJ1Tu3p3vTUPkUA6qGQg1jeEsdsUI3cBy8okpgnaLg3sn+MUPv0HJ86iUXbZuGeHcj27hOcp3XzxNEIS4joPGmbXT9vnYuKYXhHnmInS7IePDVY7u38rNDxb4cH4VEEJVrt68w9qqz6kn99DthnRba2ybGGXxziKDtTLjwwOsLDfj2zLv0KzPhaGk4SdxWjYGmq2AM0d3snvbCOd++ga//tM/GKqVsWpp+wE/OfcCZ47v5m9vXeefV2d55duneXT3BA9PjrN92xgXL7/XR21N6VwmzNlEI4BaC5WSw8aRQV59Z47zV/7DxpFBBmtl6kM1Bmtlfvnbt7h68y6OEb7/479w/sI1Th3bxxuXbvHz373D/GoYHZSzQMo2jjjZvG2HbhkdYGm1TaejccjEWBAhDEOaaz6e6zJQ89Iiw5jIRZ3AMlArs9RYwfM8PNeh5XcxAkFoGRqokg8xq1DxSozVK9y+08hhQNff/wYIAsvR6Z18dmYXa60O3W6ItdDpKsZx8Bxh/t4Cp54+zujwEM3VVRyxjNQHOXni6LpMmL+echhIUnAEPtXcbuDEkUk2bxrhzUu32PPQZubvNym7htW1FpsmH2Lf3ilmDk/z7w9uM33gUdq+z7+u3aBaq1CMLM3gJ/l6gKTmy0kqsNYK2D+1kbv3GjSaPo/t28rBfTt5691ZhsvK+Ph2dn/6GH9/7TUenzmMEaFe38CZJz+DMYbJHdtTuBV1kvSuyDJhYnzNwIIo01PjlByDAY5NT7K01mVy6wgTI2XagYeq5eatObqdLoemD1DfsIH7Sw02jo9RKrmkyaeQ4eOsk09EaSeRyUKrDFRLzP5vifPXPqTTDfn847v4aL7J00cfBoU/vnqJPfdDvnn2Oa5cv867V64yc+QQN2fnuPHe+8wvLKZhXfBpOpIsChrNNn43zBUakc/8bkjFcxGBVruLMUKnE2AVBqoenY5PaC0oeJ6HtSHWWly3hLWWctkrsFSFStlhdEOV23caeRDGMuVBCFQ8F43r/GqlhCqUPRdBsGqp1WqZC1UR8XI0WJu3/Xo6A2EuERVwkZMmOSyr95N17bM/KUb61JeavZCyTBjX8lmW/qRKWPrQ/fJ+npaoHhBAJAtDQbFWY7MlLsjydl9RVBFRtJDAcsGWPqaSeRPXPIq1Ng1HNwHa6IYKC8ttnPitl4+I/CMlyxVSoNMko73/aOoKAUJrGR+u0vK7KIrrOoalps/wIIwOVXIKRHeBJNVTDFCI6oQocZl0LbGKoqjVlHGKF41EFKDtBywst3Adg6uxJosrPuCvK0mTg/OaFtySf4D2qT36lyOCEz/P/w+A1flbPWAL3AAAAABJRU5ErkJggolQTkcNChoKAAAADUlIRFIAAAAwAAAAMAgGAAAAVwL5hwAADKhJREFUeJy1mmtsHNd1gL9zZ3aWy6VELUVKMkmRelG2HpGsio4VuY1N23IiO7br2kag1obdoA8UNto/dYq+27hAkwYt4AYtiv5xkSJFDbQ/UjhImkfrJLDTOGrixpZiy5Is6kk9KC6X+5jdnbm3P+axM7NLSinQSyzv3bnn3nvOuee9I6MTk4awiQS9iZ8AKCB6kJro2QTBANILXuJ/4SHR3sHYxGPJnCeAQUQAwRgdb2kn9257GhFBCYnm3xzqpmtwI8B479RYeu1hEMDXGq0Fx46IDQkQwNOGW9YWWV100CZiVkCJiITUd8YiAhJ+JzEWgbBXIoiSDgwEBxswxsQfEuPgE0JG8wTwSoTyksuFKxVsS2EAW0Roez4bhgLkT19cjJHtbkL2WpNj6XX1y20VzZuVxTMSKBOKzvbJYQyG85eXcHIWdiRbgwMB8p5vsJTEV9RBMELYZLYP+5BrAVgv2CwThN570c2bsPO14cTsPDs2j3DxyhIAtjGBpBgDSgQrlME047KczY5Jc1pW0oPlOB4xSXfvF8KrUCwDnAMAld3XJC1FT85n57JtRZlZoZkO8iu29M2p1HNJDCQJmDRvP42VWakp/k/EiunYe0yCgPDsYEr3MIvZw7I2+4YnZ8bL6cgNtjBpCVVZGJPwYtIT6eV2TopWsl9u7XI6kCU00UfmNZIUk3RkEs4n1pt48UrcCq7UUgEvjNHoG4pyj71iBJLzYS+mtyBIgoAUlRJZpqS7X64JxgiLlTraQN6xKOQdtPkpxcOsQHXP0CbAzU5OpS49vg4dLDJJ4etAGqPJOw6HD08zUHR479QcPzx+lryTC2/iJglZSZ1S4ZFJbWl3AccXEa0KiYgtVEd5LUuxWHG5/2d38dIfH6E8v8Dc9Qa/8Nzf4nmB3TY3S4AkByv4mMz3hBKbVNe7ZSaNwbaFZ56c4fr8Ij/8wY+Z2jbGzEduo1pvYNkCogCFiEIphVLd2CgRlIrirRsQnAlz4u06wpG0BMsrrxKhWneZ3rOVuw/sZu78RRxb0Viq8MwTMzi2hdHBdSsRWm2Pas2l3vDCMxSRL3BbbeqNNq223/OsLNOSKKmk0Yw66emNk2NBFLTbHkceOYhp1zh95hKvv32BhWvXuevDO9i/ZzPVmottBwiOry9xcP8Uu7ePhjFecLLWmh1bx9h32yijw0VuSvdTIpQIljq5RpLz6TBCBJQS3KbH5o0jzBzcwYXZM7z1/hz//K13uXKtjK18nnp8Bs/3aXkeg/0WTz16B1/54m/z+d87gtY60CIFnu/zqUf38aW/eJrnnz1ErdHEyopZKkJIU2Gn47MwqlvWuQbASgkNt82DM3so5g3vn5jnzWOXOH95kaPHzrNr51Ue+OhuNowMMrphiCfu3cF8pYbfbKfDRGNQSlFZrKBRiOUEYpfVha4YUmJVCGm9mXCgI0qepymt7uPQwVuplSv85PQcJ85dZ6Dg8J8/+ID5+TKlosXhe/Zy5IHd3LF/B81mG9/zsG079qhCaIb7CnjNJrXFMqJkGbVLSEIi0Ql1ICEuKfnPbgCWUtTcJgdv38zG4QEW5hf48nffY+5aBd/XvHb0FEePzdKo13lwZi8nz1zk69/5H9qeptVqIpbqiGrYW5aFp32arov0Mr3ZBCvxNeUHAhqzuUB6lTYG27K4/84ttJst5stVPrx7gtt3TJB3bFqeT1+hj6vXyuzcOswr/2Z46Z++ygufuhe30cBo3cn4QjztnI3vezTdZu9sMKnZkmav3TGYIeIpj9gBFQnkrdFss3PzCLs2jTB/fYHS6iK/9OA+Cv0FnEIfOceh2WxxrVyj1XC5785b+dKrR0Fr3Lobs90KfUKzZSj05fFbbbxmO5E/ZzNC6ETJHbzstP1P3lYC+QRhnq+5Z98EljI0WvDl147T9AyWbSFKYVmKtqfZ/6FN7Nm6gV2bhti1dQPz5RqVxQqjYxNs3riON390EitnsXPbLWweG6JWd/G8NoikIuLUVYU4JUWshwgFjideFq8Vmm2P9UMD7JwcxPN9jn0wz+defo183g5jMYMooVZvMr17kpc/+wxat5n5mUneePsctVqdwtISX3jxWf7mi9+g6bo8+9hHWDtcYu7cLL7vpyOW5PVHpl0CTsZWKElAxGljOrYprCWhBFptn0PTE0xNrmOgWORrr7/LQH+e0qp+1qwqsGZ1P4MDBTauL3Hig8scfWeWdcMlDh/YQrlS59XvnMBdLLNxyOHvXnyaf3zpOUZKA5y+uMCt2zbS59i9o9hYbAIPnrScXcFcFI8YE9XYAlitoc/J0fYN//7mLOVamx+/fwkn59D2dIplRgRRin/4l9e5PLcDt15npFTkC698j4tXlzg4PcX6dUNcrbT565e/xubRQR65bx/Hz5axlfTwxhHXAxNrwj8QZHRi0mhj2DZeYvbSIp42SCrmTlAv4DbbeJ6PZSmKhTxa907GJYx/3GYbUYpC3sa2LKr1JrYV6Eq97tKXd0CEhtuiv5APYqhlyjHGgG0rtk8OcezkFSxL0qXFoF8hIjRQyOew+h0abptqzaVYzGN0xwBE3DPG4ORs+pxckP0Z8DyfgYKDhFweKq1Ch+lbzhY8T2Owep+9XFDZDRdxvnuBiEEpoVxp8PihPbzwq/dRrblhrcbg+aGpMoLWAREGodXyuHLlGk8+9jAfP3Qfi5UqWmuWlmqAsFStcfddB3n+15+lWquiVAatOHwIGRWjJj2SegluIaoEp0PrIIwYGuzn/gNbuPfOrWzftI5Gs0nesRkc6MMYjbKgWLAREaq1GmOjG5iYGGOotIb160dwXZe1QyX27d0doiF8cPYsb3z/KI7jdJvRZJBpQlfSywr1YHlqFxFFtd7inulJ/uutM/z9K9/jk4f3Mn+9yr5dE/zhb36ChUqN9UNFnntimspSlad/8XH+7I8+zfO/9ss88tAhrly9ytTWLfzWc7/C4Y/dy2c/87sMFPvZOD7Kzx08gNtwEzcgGXQk9sSR31IJ5oYPOr5aYkcQfowhl1Mc2DXOO+9f5r/fnmXP9g2sGx7E93z68jmUUtiWRd6xGB+9hcce+jgv/P6LfPoPXuT4eydZvWo1Z2ZnefWr3+S1b7/B2OgtTO/fi6AYGVmLnypppG/CpIxKiK/Exj6YiMM6AyZhXZSA2/KYGh9kaMBhZnoTTz10O6sKeT66f5KlmosymvrCElprHMemWCzSdJucO3+BVr3OhfMXaTTqPHD/PTzx6EMY32NhYRHLsvG1T7vd7orbUgibKFLtRKN2p/IQUWhCR9Yth23f5+G7pvjX/zjO14/OYivhG98/xW8cOcgLn3+VwQGHP/mdx9k2MczIqjzvnvgK5y5d4q8+96f86K13+OSTP89n/vwvWTM4SM7JUXddxsfHgpzYVhQKhWWqKwZEIWFxIeFxkLGJSeMbw9R4iTOXFvF1YlFKgQOpu22yxMkLZZotH6UEbQx37BrnrXcvsrZU5PDdH+InJy9RWapz/PQcgvDwgw+glOLEydNUqzVOnjzNw5/4GKsGBjh79jynzpxBKcXI2mHePna8hyJHWmuwrMAPHD91NSgExI5srMTs3CK+3ysjC8kWg9vyyecsknlHvdGmv5Cj7fksVV3yjo1lWfTlg+RlcWkJMORzDpZt4TgOlUogarlcjnzeAQNt36PQ19cjmOsQYFuKqYkhjp26gmWpbkcWmdEO59MA/XkbHRcGgv8DRQetA8c1snZVbO60MaBgqDQYQBvin43WDK6OHVyEcM7JxY4tjVUPesJmx086+hFXJkz8sKMnuke5MQgnAkR8P1s7At/vFmy/RwF1+TA6+6STQarkYyFTj06tT87o7GQGTjLfb9SSa7IOtEcznTnVdUj8W2xqRQLhGyF0M7DL59y9vydwyjA19sQmVNLuK8iu6sH5VCEgG36kctRl2o1FJ947dUEm+Jk1ML4mhWsQWUb+IG1O05jrRHCS5HwEm4Q3ib7X8wyy2bRWgrw94HMioTE6KhAJvtZYKlkIkE5yk3CTHYXLWqxQnxLvLUToJh1QGtGgT7/q0O1Ifa2xLQsw+FqjVPg7McB8xWXL2CCnL5Q7biu8MWOSByVxlvT3FMckRDUq06QJTaPXgTAp2PDwMEu0lOLWTcNcnq/G07Yxgm3BtUUXgE0bVqMjbkhYq+56zSDBsehZaHtFha8YRB/ViR3jO9JRyKIxOgxhkq8adK4h9bqBAJfnl5ibr5GzFcYY7CiLspVwtdzgWrkRvN+QYECIfYIjXSzs0W4EcDPmNd2ihD9APnhmxy9TADl7uXQuffBN4f//0CLskg7vfwFrebOP8vzDyAAAAABJRU5ErkJggolQTkcNChoKAAAADUlIRFIAAABAAAAAQAgGAAAAqmlx3gAAE6dJREFUeJy9m3mwHVWZwH/ndN++b8/ysr6EkIXsJiQkIYRFloSAgBujqINaKhYzFjpVjqXlUiV/UDNKzVDjLFUuYzlazIzM6IDlgooOzCACGgMY0CQmhAcJIXl5y93v7fXMH73c0337vvdCzcyp3Lzuc75z+jvf/n19WoysuFAxqyaiXwDI6O/5NgEo7e9MsGjw8RwVPZ8MDjp8Xl/eHDA7Hhs9J+jAL354vMgs6daxRhbJmWDja51owTRw+WtJGRJBqTTxUwQQAhw3QEpBsWDMAsnUbBAquRbTwrandO9QuZfdm+oEizqEgJbtESgfq2CgNEAzpqwQ4LoBIwv6GVkwgBeoBB0hRISbQIjoPrpu98nwXgiE0Mfac5J5UiT3UgiI7zWRVUqFiKroGlBKQRAk1yoIx9rw0RxU9C9eAwwJp8aqvHKmHBFBxQRQCeeXDvexcukcnjl6FscLNH5kdVahOnQuc68ioil9/PWoTZc2KxELZVEphVUw2LP1AoIg4NRYBatgopRqq4AUgmULB3jm6Fls16dghvoiRKSBSnuqEIlWhvqYyAptW5E1Yq9nd93mBhqcTmD9+QAyWcJ2fZ46dJKdm0Z49Vw1gTEBggAsyyAIwPUCCoZMGBeLSpt3EUVCAdMQTUN1v56tB8h6mexGO6Wy81lB0lUwJa4XoJSkxzJxPT9Uwfb8PKSyfRp1U0Pnw+HZqEHGyqvsRrNq12Vt3Y6qPAKhEaBjXkzluGV0XOj93RDKtvNVhemmBcwYiwj9UoYdmbVkKjiJLbzQxUwnhB5M5AQ0QtCFplF7PUawy5rnTctIZUV6opk1NAqFRKJzWuWqRw4GShfP/63WZafn+5iY+aqbCghNt2O/G/vRacNMHakZxmeDZeo6ZsTrCbu7PSGNl0wPhk2R5qQQuhroxMiwQWRvsuI7E1HyjN35srrL81JGsL2fjlwgmZc8V2mM1QdmQiybqGSt9/nmA1kEs2HzTHlC1JPhgdkxkhufZzMyOvfQsadwISnbyCrl53vbmZrQMY82K0QX190Fpy7C1yZAFLtnpS6O6TukIC9EyMGg0XSSYatgYBpGO17vwLJLS8HHfbNQj1Q+FTEiWStkaq4KpARdJfzTRlMD7b6Ma5QCtqxfhmkIDCk5dabEZLmOaRoEQUxI1ZU7qY10xDznN09EG09P6UIAlblLT9NEsAM6pLKUklq9yZ7ta/jmX99BuVJn3lAfD/30IJ+89z8YGuzTpsxCJ/I2mkv86eZFfq0zEMqZ1PFAFYHqNkAXQZVsRAiFFBAEAe96824GBvo4dWKUc2MT7L9mG2tXLabRssM0uCPQej2uM2bITHCCbFYDSidAtzhfHw/St9lhQpfZtB3WrV7KzdfvYqpUozY1SalUZvGi+bxl7zZaLRdphMRqR52zcZe53JkGZ21eXG/IwGnJUNs+iA7OzPiEBEZKSaNp885b9rBgZDFnTp6ir68Xu1GnXqnwrjdfydKFc3FdLym/JchIiWEYGIaBELHESe0nMAyBaRgYUkZSNJumUHH43pUAekUmqeIAKap144BI/ndcl0XDQ9x64y7qEyXq5RKjZ6s0ak0mxqdYt+FC9l+9lWqthZQGoWFSCCGo1VtMlWpMTlVp2S5SikRwI7SoNxxKlQaVmk3TdqIgreu+23MTG66rm4hLYkEGXqU9T+4zwgWEUBFhBdIQlMt13rZ/J6tXLOTo74/g+wFffegg779pGxeuauI6Dre//Qr+/QdP4ftB5M4VrZbNB//octZfNIJpmvz0vw/xn0/+nv7eIn6gEFLQarm855ZdbFi9mHOvvcbhkxUe/fVxei2TIC8m6IhM4z9tw5kOhXXJ1/1lspBmJ3TrHc0JFPQUC9y6fzu1ao1meYojL0/wwolz/PrIGVy7xdiZs+zetZnLd22g3mjiBwEDvQU2rpzHlTvXcOcfX8sdH7yRrRsuoGW7EYcVAkEQKDavnMsNu1bwybvezpWXbaLZtBFSTCeYgMwxgCFAu2akt6yvyFP/rC5JSaPhcOnWVWxZP8Lp02fxXJ/HfjNKwTB46tDLvDpWoV4qI4Xitpt34fsBthPwiQ9cx+bVizk7XmZisopbaeI4HkJoCqBUpAItzp6bwhUW5Uqz7U264tjmfBLPivav0w0mTNblf1ryhgQAgsDnluu2gIBmpcypM2UOHD7NYH+RsYkaTx46hfA9xs+Mse+KTSxbMo+3XruJD7/vBpwgEjzfo2CZSMMIN51ikKJQKNDT00OjXMJ3nERCuhMg8l4qimYUSfVYqSwBUjqfib2naUIIWrbLmhUL2L15ORPjJQwUT/3uVcq1FlbBoLdY4LHfvEij5VApl5g32Mtb9+9kbjFA2XXmLRjGbdl4rhsaYSMW22iDIrRLQkikYWA3mrhOi3bVdnoe5YdLGRvQDpfV9NY106QQNB2X6/esZ05fkWa1wtmJCt9//AjNlstkqU7L8Tjwwkn+6zfHKUjFZLnGTVdv5tCJcX7++HP87Inf0dtj0Ww2USp0pyEmaZSllEhD0nLsqLCZj1O7spUhQSby1LLBNGU6jWr31Mr1fYbn9HHVthXUmw4qCKjUW1yzcw1WwQQEhilpOS6mZeEGYFcrXLRimDWrlvGRex5kstrCMAS24+D7nlaDSBs4w5AYUmK3Wviel8FL24FKi0U72U0TRcsGo7csUVSb86IpmZQEMEJhSEG97nDNjlWMLBigUi7heopF8we567bdFCyLQrGIWShgmBLH9alV6xhmAVMI9u1Zz4OPPEexYOJ5Hp7t4rtu2gvFjgcwDBPDkDRrNr7rt2OVGS11lnlRwtZllx2bTv9tt0ApigWD63asxHU9XNdjoM9i7mA/0iggpQlSoqREISlaJsPz52IYkkq5xpbVi9mxaTm1poPyA1zbwbUd5s7pRSCSiE9IgSEFC+YN4nk+ru0QBIHmu/OwlslYW/K7qQDpSDAHNoJpe1QpoNny2LJmIRctm0O51mTunAF+8sRRDo9O0NtrJXQThJtQQmCZBh+5/TqkWcB3W9xw+Toe+/VxyrUmvu8xOVXimj2bmTPYx2SpRm+PxVSlwdWXrmPthfNpNGx8zyfwvUyqnm5qFrXEdDqsGUGUSNkQvYihlEhIGijF1dsvwPccDKuXct3jG98/yNnJGgVTJil/InJSUKm1WHvhYt5x4zZGXxxn+9pFXLB0Ls8eOc21u1Zx+vQY2y7Zwlfv/TD3fe1HlKt1rt2zgY++by+FgoFZDGgqv0PPO5peNRKxQsU2RWQIoGgD50WVid4TL4Pj+Cwd7mfL6gVUGy4XzJvP93/xB6aqTRbNHyDwM9EioccomCbf+/mz3LJvK0pKCgRct2MlDz52mNv2b2H5onkcfuEoey9bzxt3b2B8vIKFg0BxruoCDo7tRllEp6do70nX9zjFC6PKmLmZdFho2SCdnkALpyUCx/N502WrWb1sPsuXLcEwCzzy5FEKhoHvBfiBin4Bvh/+XM+nxzJ54dhrHHj2RS5asZTBoUFuuXw1/X1F/uLrj3FmvEqPITj58svUSxMsmj/I4pERvvOzQzz046e5+A1rWLt+DeaMxZTsQQ6RBELxvdnefL49zJ6oiMnjAz1Fk1PjNb7+w+eRUlKuO4yeKVEsFKITJt11UABf+fYveOb5UeqNJhKYN1Dk2MkpPn7fD9i3Zz2XbFnFguE51O2XePSpI3z3R0+zbf0IgZK4ts3zoxNYpjGLQqv2oifZV1iPECMrLlRKgWkINq9awHPHx8JKbiofyFAxic4EzZaD74eHFgwp6O8tajrfHTMhBI7j0XTcMOhR0NtjYpkGrq9oNMOqkZQCxwmTojlDfdiOR6NhIwxJb48VEWA6CkRZK+AHissuHuGZ37+WvB3uUhSNbX03/xqNKejvsZKTIaDw/TCK7P5GKepRCssy6ekpJLoWKHA9nyAImDPQGxlj6O8rAuD7AVbBpHe+hed5oDrP/HTsPakeh8avrTVZI6jj2wXpdn+7AKJU6GxEoLScPMm7UvDtv1H9QYWbilugoK+3j6GhAcbGxtvgmjFVSuE4Hv39vXhugOM63YOhlOBG+XpOEpfGTfcGaNTLayKsFRhSUm/YfO5Pr+eBv/0QpiEJgiCSozCu13MLKSR6nm8YJqZp0Go1uXTndv7hvr/E8wOEEBiGEamCxDDCebZt860vf4k/ueN2yuUKhmF0bj5FgDxJDK8zbjANJrSAKH6xIIRKDEi4dFip2bhmMdvXLWL+3B5uuGoj3/7hQeYN9eIHUKs3KVoWphG+jW86LoYM3WEQKEpTJYrFAqWpMq1WC8uyokNbHtVqjZ4eK5QUIRgc6KdoWdz9hfuYKpUYGOgj8LsYWz2r7WIntFxAtKU2TrFVFEsI4v8SA6ciGyCloN5yuHXvJp54ZpTDo+f40Duu4Ls/eRZfKaQU3Pe5d3H/Q0/y/NFTSMPg/W96A394ZZJfPv8qA70FPvXxj3Dlnkt5+ZVTnHz1NC3bxg98Fi9cyKc/cRcbNqxDAN/+twd5+JHHUCh27riYY8dPcOKlV7AsC5UcbBRJmS3FTm1/bZuVqQq3CacfddNbp5i4fsDi4X62rV3IT355lEd+eYy5/Qa7L15Bre4gheCGN25m0YIhXC/A9xW7Ni1n5chcSqUyd3/mz3nzTdfz91/5Bk889Sve/c63hR7Cdli2bAkKxT1fuI9v3v8An//sJ7h053ampqa4af9etmzaSMu2O1L3tFfQuarvMPyrGcF01pVdIp0fhnbBEFBrONx0+UYmSi3+MDpOrd7i8QPHec/NO3n8wCgomJiq4vsqTGoAN1C0bJcVy5ez7+qruP3DH+Pgs4ewHZtisYcPvvfd9PX28uTTBxgbG2fXJduoVmtUqzU2b1zHww8/Qq3WoJW8YMkyKY15sjclOkyFRgCZTb0jasbkSM8UKJQQFC3JrvWLmNNnce/H9mEVTRYMz2PxgkFWjsxhstIKk6amDYDteFiWiRDQ39+H53mMT0wwPDyPWq3OmdfGkEJQrde47da3cNedH+Bnj/6CqVIJ13Xb2Eoxy1xW90Bh3S0uiYVvsZJToKpNqHi70wQ0QioaLZdtaxczPFTk7q89ypce+BV/8y9P82df+B4nTo7x7psvYWyySrVSYe/l66k3bXZtXcHu7WsQssBLo68wVS1z54feS6VSpb+/jzvveC9CQLPV4uYb93Hg4HN85rOf58v/+C0Ms4CIKkXxadOZ968inVdJQhSxFgAzLoJ0DKm00GebQhAEAXt3rODg0TEOnTjHUBSwlKtN/vXhQ3z09iv4u/sf54tf+zl/9elbuWr3eqbKDcammlhFC8/3+MzdX+Teez7Hj3ftwHFdWi2byXKZOYND/NP9D3DP5z/Fd7/zzziOQ29PD47jghR4vpdOtroimop+Qppo5TItFJZsXjXMc8fHMGQoW3rSoG89bCEnFs7poVS3cb0gwxHFkuEBJsstStUmi4cHWLd6CYePnwEVIA1JpW5jt2z6+3q5eOtmxiemeGn0ZUaWLmHs3DiNRpORkSVsXLeWo8eO07JtDMNgcnKKpUuX0Gw0qVSrSEN2idvarhsEvh9w2dblPHvkNRzPD0d1AmxaNcxvj41hSDEN7zW/KsKTpaaRL46O42MWJKYRlsJs26WnWEjydMOQSCHxA59Go4lpmhSLFq7jUrAKSClxHIdm06anpxidNgHTNHEcBykl5rTJkE4A8H3Fnq3LOXj4dH4u0D1yz1k4WjR9/Dw9s1gMDyQHgaJgGqlT2hDmC4EKkFIyODiQnPi2ilY0L6BQKIR+XtulChSWZUV86JKzZ1RaEtE9UxtMH5HJBEIpt5CyE+nYPIkuMsioDFw3TsUbT6+ZPxbjmtfXxlHkDqXys6hJHShnfxp8XkCh72KWff9XTeXdaB6sQzDCi9xXYx3fe2S8RE60wOwOM+YRcBqiTrtO3vy8taNYJvu+M2r5h3D1JCJcossDsgjM1LIxRUrfpp+WmtOJQSeT0GReiwQzTaZgaWd58dKqQ7RmcUp71k2XnDwiypyhzsAsPAQd32nAHX2dhE5JQFJfbccJpLnUTdFVd+HIiquQnX3ZycltkLmfpnXApNfVDaDQ8DLbIqJtTsU5AETvyXIWjgkikjntB8ZfauSIeO5amcEsnacre6XmdBrqJKIRSTScGk8lU21dimroQmlLphcPx/WzRJAW5yyXO1uShhCvEUtHZr6Kx2KBlRpxdYnSpTTObUTKS4e/NjFl+Co6/KhIAH09BVzPJ+8IjI5wVrzbdjh68YD2yVxEwPAwVFg0lVIm3kZEZw5D2Jj47eeFc6Jfso6+vmjPS3CIiRpe247PQG8BCGhGB7AgDoQirr90psLFaxfy22PnqLfchLMi9gyRpVXxMdVcqZwpIMjo5XkECyJ7E+RJVmcQEwQw2Gdx6RuW8dtjZ9qnSiDMBSCklusFzB/qYeWSIQLVFm8lRMLFUBJizkU5gMY59L6YS5H4SBl9NIlMfTyZl0eoSF5V5gfRB5ORiAcxXDhJm0sCE6vGi6emGJtspD+cjB+ulKJgSiYrLSYrrfxPZzWxF5EIdAaes2g5G872dMiFfjj6PCNMRfjprBBk8pG4HqA9oxBVbr0o1+58Qxx+KClIG+d0IVLf0v9nPNytCaxCNsUPW+bjac095MWIiQSEiY1MWcUQIK3hOT6+a5su2syAzWbJFFx4ozRmQcjQ/wEejiarmDrEVQAAAABJRU5ErkJggolQTkcNChoKAAAADUlIRFIAAACAAAAAgAgGAAAAwz5hywAAOL1JREFUeJzlvXe0Xtd12Pk7t3zlVfSOh0IALABJgCRYxV5ESqKkkJIi25IcWY41SSZZnoxir5Hs2PFyZCejsceeJceJYy3JFpPYkiWrW4UUJYoiRVGgRJAECRAAQQIgyuvva/e75cwft51bv/sAUCuz5ix8eLecs88+e++zy2lXrJvYJLngSQQ/ABn8tNT9oPLni5Zav4pTlfrLYKbblVcnSp68sgBeSR155YvyltUxoK1CRFy5wGkxjchLF4L56u9Cwi4jalGd5DzLy1PlXZW6qsMxKtR0Dkmm/kKxxL9R9aeZdD49X4Ux6H0R0T0GM0yto0pdZc+E8rcAlnzDBeD/b2mQkFUVwqr0qyoExamyAAhACFFNO12QJBDIoE9doEp/brjnpFA5nHPfyBYsAiWlRMoybRSnSgKgaYK+7dG3baSUviCgtEkm7qqABBmWkKkSCgyZyF4OO5E3BxeZzp6tQ80q1DyBCOaSNJfGisgWvI/+L+VR8qXIc2dyEGrUDOp1A8+VGeqmk1HGNCF85i60LdatGGHHxFrGRupoQijYyKAhWgpJEWmMkBxCiOB9/Dd6L8L3IpEvLhdfR39FHty4vAzqEQFiQhOIRJ74F9aPEBG7hQ8korUvtQFJpURKgp4W3AdSLWWcT0rP7yAyeIeMeqgMYQRwUcqHnSr9F+U+qi+glOdJZhe6PPfyKV49NcfocEPpoPlJrJvYLMkRACHA8/wG3HfDFpaNNTl6co7p+R6u5zs0Qi2nimfIjAS8VB8SYXkRlQkFIRaQGKAQijFI1SWUC1XYEnlEsh5VYFSQ/r2SVyrMJ6J/fB/5bNn3sXCkYMg0Q1UBC8vG6EslX0wymWKs36blY00u2bKC6XmLz3/7ORC+Bs8XAlEsAACO6/LBt17B61Mtvvb4yziuRNe0/MiqIPIrFr6MTi67zVYmUblTkrsSMslXRV2mQK+XW9qs31RiFSpDDXNEpA9uHNfD1AXvvvtyNq1byp/8t8fRdb0QQq4AaJpgoW3x9pu3AfB33z3IsvEmoNImzXWhPK+Iecl7UQoqlUEqz5UkhcwnYzX/qLjuZKXE4V2mm6L28gyKRe1T5FvNmimaQ8dQs52dafH+t1yJYQg++/VnWTo6FGjuZEVGFgLYjseaZcOsXjbEp7+2n6VjTTwvD9vwWdEon0osJS5NdjeljIgeyUTZvLpFTk9NkSm0k3kpjeZAuc3jnGrCQhsuyHYMpQ0V+ocPQwY0iAvEV0J5kKRTaFpWLh3m7x5+jv/tl25kYs04k7NdTF1LVi/IjgRqCKy+w8WblnHs9Tn6jsx6nxESaSblMcSHmsyfl0eFk77OKSOLBpbSuIR1l6QqTMnt5eFzcq7TZas8U1PImjKapeuOaSwA25UcOTHHFdvX0LMchJY1Xdmh4KCDjg/XmVmw0DWtxIscxKwy4kuKRwclxUK1mJSjXQoFcbFww+QRqfnQk8eLn0e/xdSvlCskfppGyXweoGs6k7Ndlo0P544LSFk4DuB7lK47iPhlurOskT/vYWGh8OBcBSoEUGQKy0wkpHtofp6c+kQZvsXvhCTgoes77gX5CyaDwspTKGfKy5LbQWYhg3LB8wuRZDEhK9nk8L8BglsKS/EBBtUp1YsCulTGO4aTBylfANJmLkxlPMpFqEAIMnkvhFo+x1Sl2ip5SniVyCRkYThZHc4iUgksQZEJSDrSObzNgZppVNq5Ezl51QfnY+vz0hsBc0B1lZLqE+Tgl4FT0IbC+pJ09gOF/LokgzRAYQVVvN2qXUsjX8yKbGYV+G+UNrmQcN8o4VToJgdMowlZNBfgOx/hcLsQ6lCiSP3NQSBSAoLYbqaZXCHOT6B2DgSrogTCPJXyyoI8aog4AEiegnxDUuj3pMdilIpl0YogoboMKoXSTE/10tDLTtAh1WJBskyaErnyVRALF4p3GEIVvU9lzUGjNG8iqaFuRV9mQF3yvBRN0HYB4SCYOrGWSKLCdLA/exXf+TOEwUROyldQZkYoZHAmUsjBP4KV0y3VRz9HEz84VfTeKuBcGvlBJW0V+/0h/3IKSdBKmRBdq45LSY2ZBwWqfaAZCcurgynh45+Ti7zo8hVNwIWov4KPpgqRKChQHAVk1KsiPRLAi+WCPHMxCEGlnrIUvn7DIsSy+suMdVqAz1ElyRIzds7hoExcihKzJEXhQFAyCeUqb35Nhs7iQBsoFPyq6MI8cEHkULWzXQgzkakr7cNUQUYU+8F59Q1ifqTWz6N3yIoCkKw1GzEMtFkhkhKiWbMygpW+VtSCCMPIAue0MiEXi0ROjx8oA4uQxIQizZtIKwuRk4hI1GVhSXaXmIAifPMbIRGI3EmLFIKiCkeqZYmjCfVvqntlJLPAqaySEvlyuF0Z56KHRfDynGEZCP/5z6kUjgSqU8Dlsisz/0dAipA/L6Oer4WydeTlOccg/Lx9kDynN+3gFDmRRarl/G2bpHAkMGTSIgkWFVmMU5gDI+86N6nmZFDmgvdVApKqSWQusi8TedL5LrBgVmhbgQ8QIjvIDc9xhmRcJtd/E2HeghT5CUXZyuyfRrFbo87RF9RblKqqd6nepHu7F/s/Qnl2rj15UN8Siu0vyBd6UOeZRNzrL4THLSiJ9UXBD5IEL3IMc9LPK5JAKp57GngonBcyxQ5yoSsrShaERLgNqibsrQJkinHpNW3nT+y0jxHbTRE4fTJYT4cQqR4Z/s0j9DmYqjLUyuqouGNncfVl65JSi2giogLJQkKWDQRVxiJuTETGSI2fi+NVhRmxedE0fzZR3SjiE0ASL4KVwb+wV6QZdEG6+CLSIuobJC+571N7MgrqlQyaC8hfDZpffxQahzYuRzWHglHYqEBgVD+guEaEkCy0unhBfilFNPVtGBpDzTrSkxFYQaghFCFLByqLJrbyblD5QTDyYJ2TbxtUEu1GKk7lArAolR3s8hPpFylBKGRuXmvzwqb4WtMEb7/7Kpp1A8d1ogbXaiYzc22+/+NDCC01+KGaiRz0CtMgxl0AK7KoNFCQwsWqxZkEFVYEVcEktDI+DdKUUGxPpBmKzEI6XxqZmPHttsVN11zEn/7O+7D6FtJzEULDcVwMXafXs3jgX3ySYydmaTbNgn0Ni0gifZMDb9FClNJGKsxz6vk5FZVFXmLgiqDF+AIeyXmCMl1ZFMoVECISiHgDqus5PHjfXhzHYm6hw0LH5vjx07S6NlMz8zQbBu+8aw/9fj9Y0BJu1qza5cuS6kOUhaWpVJglLK+dP2p5YBU3MJ0Kw0C/3GIxCaso2wgC+TF5DvMjnSKjaEMIjW6vz46ta7jzpl1YtsvwcBPXtjlx5AgCj2ajhmW73H/3Naxds5R+38nptIOYVvY+LURZD7taUkPAc4UxqIZ4bCbTGlkiAFL5v3JKqO6CBi1StSV0ipAYuqBn9XngvmtZvXIcDw3DNJk+/TqubTNz+iRm3aTvwOaJVdx50y4WOt1gbXxICFXDqD24zAfJS6k2Vu0vhY6bLHmfAlDYOeO2ROY4E5EFNRaagAihEGZFQYhiwPRvESlhfmRibEEI6Dsua1aM88C919GzHBr1Jv2exezkFI1Gg9nJKTwPjJqJ67o8cM8emvXYBxDBsmwhZLxUKiEE4WjiOXh1i7GYRalKtbl5UoIcbnsvqU8xOgU4hgKZUZ1512XP0oDPLWmaRqvV5c237WHH1rVYtkO9UWP69Bl/m5Np4riS1vw8jUaDXtfimsu3cOOebbQ7Vu42aTEoTlJyVjML4bWW+lVUD5Uir7yOlepwUgYjz0VnA/ipZGeQmqqqxkX0+nMQBM+TNBom73vgFjwpMM06jmUxdfoUplkDAYZZY/rMaYR0EZpGvV7jF//RLeQvZSEzepmPZJr5qrZIaw+RCL18TRNqmwt8Kl8uDQO/Qgzmgyz2AYokXXmeOsZlAFbJ11UHQ2ToJAW9v93lxmsuZu+eHfS6NkPNIabPnqbbbqMbBoahYxgm3YUFOvPTNBp1Oj2bu2/dzc6LJ+h2Lf94m8IKi9qrlJFCeaPaVhm/L6hCUcrVUhkpB9Ewz6fOwWcRIqnFroXSxkgIEgJXoiorRkxqPv9gKv/Ekve+7QZ0TeB6Es9zOXvqDJpu4EmPuY6NofvNmpmcwjAMPE+wbHyYB95yPV2rn9kiHVSQrV744w2aJtAEaCL4q+7aDgysr11E/NwHGmXUNC2ABZrQEEKUCCKL6ySDkggxy9cIi9NJAmSRQ5hgbDocjAqnM6YK50uHJgTdbp/Ltq3nrjftZGFuDt3QmJk6y8LsPM1mk8nZNn/++SdBSgyjxsJ8B7vvYNbrdHs93vnma1i9YhyrbxMJrFq9QiAhwLZdWp0e7Y5Fu9sLfhadnqW0J+7VufpDgON4LLQ6tDt9Or0+HcvC6rv0+m58hlApLUtSlWghMyKbvC49JSyZwoGeWJ7KPMwsElKpV1E+widf4dp1QNMFXavPu95yPUvHh5meWaDRFJx9/SSOKxlv1PjpwdM88tQRHrx9F7t3rGWh22dmeoq1E5uwrB5bNq3gvtt386m/eZQVy8dwHImQCm5BqzRN4Dgea1aOsmblGJ5LeAAa4J/Bc/jYGRxXxlvuBBktEo5Krlg2xJZNW+m1uziujed69C0LKSVn5vpYfWcg5QvJOiiDgHgzqmq+4yYb0VlwedCj8wCDVoZ/ZCwExU7fAMmI1KdyK9InX/ko9Psua1cu4e13X83CQgehaXRas8ycPUutZtLtuzz201cRms7DTx/hml2b0GyX2clJVq1bi2YYuK7kfe+6hb/58g9xXS9odoy7xNc0ridZaHf5tX/8Jv73f/oWFtoWw6NDkcP0+tk53v6hT9BvW+i6pvSCZG8TwhfavZdt449+51ew+jZTZ04xMznD3NQkmy7Zwa/+27/hyLEz1EyjWBssKuV1oBx1Ep5oJoo2hmTgimifYBwf5ozmZaQyRzhCqQxClExZ4dcX3muaTqvT486bLmNi/TI63R5CSGanp+l2LYabNfYfneTga9MsGx/m6QOv8/pM23cAO10W5qao10263T5XX7mDm2/YSavTQ9NiYklA1wRdy+bSzcsYHTKxrD6W1cOy+ziOgyc9XE9i9+0ks6IwK6RF2BB/itp1XA789GecePU4tt3H6nW47KorWLV+PX3LziNaNROQyZM0oVlTrfJCRI+0Iv4LRb9X8lurWhIfu8HApB8/e1LSrBk8cM8eep02ngTXsZmdnEHTdNAEDz9xENfxqJk6Z2faPLbvKEMNE0/C7NlJ8Py1AYam8d533JzQeDHzXbatH+f991yC7bg+uTwPPA+h6ei6iaYROJGxzS9qik8+XwgM0+TMyVMszC5w8ZW7Wb56DZ1WV8Egj0aDaKjmUaXBv1Z1dDpPdEYihRogZHxCQZdjt6hRsHyzIVMXuibodCyu272Vyy/ZwNxCG/Botzq0Wz2Ghhq8dmqenxw4SaNu4nq+EHz7yYMsdPrUTJOFhQ6dVgtd01mYX+COGy7j4ovW0elaiMDDdxyPZk3jNz90F3WzhuN4gV8i8YIpZi3w3HUtjoby2508G0gIgab79YDErJt02hau6wUniipFY+VxHinQalLEYxyB9o5PVA1zln0vQMHNP/lUlAyaVHVbc6pR6JiG4vtWkgfuvQZN4PdM6TI3PY3regw1Gnx/3yvMtfsYuobnSeo1k5dfnWLfC8cZHW7gOjA3PQPSpd+3WTpa4513XxWNCWiaRrdr8bF/8Tau3b0NT6+h6RoSiSc9PM/Dc/3JJE3TESJNsvIBl+g4Wk1D0zSsTge738d2nIQm8o+lPd+BIgWeor1DHzXjHQhZMhCUMM8hlwYN8lQQggLdmWC8kGgCupbNjs2ruPHqbcy3uiAlrm2zMD1LzdSZb1v84KfHqNcMv6eGFUj4zo9eRtN0DF1nYW4Ox7YQAlrtNm+/ew9rVi3F8yQzsy0+8MBNfPj99yHMGuOjDbzgKFzpgue5uK7tx++arwWSrmPSH1CjHhl0PU3XEZrmRwa2jWV1sfv9hK8TH7hdMSWqDn2yJD7hCqj4vOHwfOOYTqUaIBG1+ZzJyajoraqTRsTHzRaJjB/6WdzzpksZaRpYlo3QNLqtNr2exchQgx+/cJxXTs1Sr+mRY+Z5kqFmjZ88f5zDx6cYGqrT6/XptFsIIWl3LbZOrOKem3dydmqeay5dxwfvv5Kp6XnqzSZmzYzst+d5SNfDc3wNIPyRnArtC+xwoHo1NISmo2kaTt/GsSzcIBSMU8C8xaj/QlTScP2/segmooA8wCn7VopUuQrMwk5hneNeCPyBmJVLh7nzhh0stNq4roOuaSzMtdB0HVdKvvPUYfIOQdE1jfm2xTd/+BL1ugkIFuZbSM8FT2JZfd51715M0+TSi1bRW5jllUMHAY0lK1ai6zqu6+J5Lp4Ex3EhGNFLLzFLNELZ9yjxzVd4Ork/Eqjh2g5Or49jWfkataozvSiLq0Yn6uMiExCeLZMoEy7MSIYb55XyNIbwx/07vT63XbuDiZXjdFtdNCmxez06rQ6jI0McfHWKnx18nWYju9zLkx6NhsmjTx1ieq5Ho1Gn22lj230AFlo9rty5mVuvu4SvP/YSM/NtcLvU6zpLli/1z9SVEuk4SE9GYSDBMG6Vpqv+XEgzoeu4noftOdiuS+ExtlUGeRZN/2xdQpQ5gUJVAIpNo+jo2EUiIyj0GTzp0WwYvPXWXVh9B9d2MXTBwuwsjudSb9T53jOv0Os76HrMkIjoEho1k5Nn53ni2WOMDtdwbYdup4sQAsdxEUjefd8ejp+a5akDZ+h2Heamp/gvDz1Cq+M7lf58g4dj95Guh6bFx8vHNRZ02WgFU/CdgsALd10X13ZxXXcxQXaizvhE+zKnnEpyUrweIC88EbIkEqiYIvMShysJhDRBu9dn98XruGTzCtpBuCY9j9Z8i6GGyampeR7bd5R63Z//97w4ZAsdHU96CF3nS4/8lF7fxqjX6HRDuyuZX1jgpqs2s3PHer706AFA8tAXf8hf/o9HGR6q40mCvQWuP4Qrg28kJDacQOJGbUo4ayhiE+DPDdg4roPruqmQWCSLVrHvmZ6oFFTRVPZLpNOA9QApyQ6It2iPVcUvMXgRwJFJT8JzPe6+fht4Ho7tYJomvZ6F7bo0GnV+sO8wU9MtRho1aoZG3dCphT9Tp1bTMHWN8aEaBw+/zrMHTzMyMozjuNi2BXhYPYvRkSEeuG8vh09M88VHnudTf/9jhoYbAeO96Oc6Hp7nIqXnDwSJvN6RTxIhYt9B0wSO62I7TuBXpOigPikhbzSCXWn4uMxxL9kXkFBP6jr6RMRehEC6UUH+9LYo4a9aCafQhRD0+g5b1y9j72UbabU7/mYPDboLPWqmSbdnc/0Vm7jqsgkMw/CJK3ziCi0O1fzhawlCY2x0hL7toWtgdTuYpokUgna3z103XspnPv8Yn/7KszSaNUxD9zWI5+G4Dp6UeK7Ec13fiUyMtKltLKCEIMDJjwJc18FxHFzH9zNyO/pAvhZlUOgrQ8qH3n9eeVm+MUQWaYJSDAskLjLVMu8V4M+19/sOd163ndGhGnOtLo16zQ+bXIlu1hDAqmVjgUrVot4VXiOCSE3z42/N8IeEpedimgZCMxC6joZHt2OxYeUIN1+9jS89st8/T98L5valRHoe0vNwPS+ICrzYdA2kg/9eaMIfQApMgOe4eLaLdLxiciVBnIPHD/HeyPAbSPl8KRWAWNMtNjZR8we2sEJ52/FYPt7klj2b6fT6IMFzHeZm23hSIDQdIQR2HzQhMUwDwzQxDAOjpqGbOrpuBGrar9dzXXAcXKePLQVSGCAlpqnjuZK+ZXP/7VfwzcdfxHX96eCwGa4rkY6Hi+OPBwSziFrGDyhIEuqmGQiAhsDFs22k6+G6sTD5BBIkD9ZMkTQHdtERsOH72JSE+dImveKawMW7fQWbL0uGDHRN89X7rg2sWT5M33IilelJfMYKf0WNbpjoZh1NN0BoSASe68frfccPsVwnsN2u648S6jUIhnmFBq5toUlJp22x86LVXLFjLZ1evGJISonnuLie56ts18VxXJp1k5GhehB6piiTpS8rlg77eOsGMrD/MhhiVp20RccCQu3ZxaVlkb8SpOKjYheHDnm9PmEvlYMj8o5H8zxJw9S545qt/pSr5/nOodAxawaGrtNomAiho2u6z0hlnJ0wRNMFWmgahEAThj8KpxsgBL1uD00X9Nod7J7tCxcab77pYp7a/2o00Ne3faZLD6Trh4K27TA+NsTEhpUcee2sPwGV6MVKe5CYhmD7xCp//YF0cV2B53o4jovnZu31YlJcerB2LjLlwfkA+YXVMDPSNgjlyxMpB69AiES0G9jPk1n4IP0vkXV6Nru3r2L7xqW0uxaapuF5oBkGmoC+bXPw2KRvwzVfDYvI/vuGXxO+X4CIwy9/4EZDSknN1Lhk62qkNDAbQ3QWJtE0jYW2zQ1XbGLLhmWcmmqhCcHMfDcQAIHnSXrdLo3hEUxD57YbdvKt7z2LPq7j9rNnAeiaoGPZXLRhBbt2rKXXs/xp5XoNz/VwIwGI2bj4BSGDPwrpZyvvzOUHROSEKP44vlQAx8fG5rYhamNOSKn2GOlx21WbEXi4roNERxcmnpSMDjf5+mMH+Pd/8TBjI01c6WWdowz4yA2O9gZ6rstn//jDXL5jI52eP75g2xYSwdhwnTv2buNTX3wKwzQ4PdWi77iEmr69MM+yVauZm+/wnvuv568+9z2OvHaWJWPD/s7kkDqahhOYi1/7xdsZHRni1MlTGLVGsDLJDcYs4kO0q3pYi08BXwo0uqDMB0ibNwnqEqr4JHGZlZe0sZdheUHCcZF+T+7bLhtXjXLF9pW0u5Y/Deu66Lof5vVtj+/86GUazTqmqVMzTWo10/8b/mrJX71e83+1GqZpMtRs4Hjw9UefwzAEjt2nMdSkH6zwabe73HjlRsZHG2iaxomzC0zNtSHYX9Budej3fTOwdLTBJz/+IdasGufM1Cytdo9Or0+rYzE1s4Bju3z0n93PvbfvYWaujUDSGB5CBkyPvu0baFP//0WaXUHOQJD6XgZ9NG2WE1yt+vHolIwWhCYiGi6OHZOEfAuUsYBAkDSB5Ti8afcEo02DhbaF0MBs+N8pHGrUef7oGZ4/fJpm3cT1FFhq+9LoJtsZrB+o863H9vP+d+xl6WgDR6+jmya2ZWF5kvXLh7l210Yeefook7Md9r1wnHtv3I7lSoQwOP3acTbv2Mb8fJurd07wrYc+xmc+9whP/ewo8wtdDB22blzBg/fu5erLL+L0mWn6vT5LVq6Ozi/wnXKZ6iyqnc6SOz/lRVx5eUINHY+NRGswZEkYWDzaJwEtWqufrjA8TTx6EnihMjOY5KPnuB7LRhvcePl6rL4/OuZ5LrV6Hc91MXWNh588RN92qdcM3CSI4pRyVaSUmIbOidOzfPOxF/jgu29kdr7L0MgIZ2ZnMXQDR5PcsXcLj/30GLbj8o0fvsRtV28BTcesN1iYneP4kaNMbN9Gp2ezatkIv/WvHqTdtel2+wg86gZ0uz1eP3GSbsdizcZ1OI6D0/NXAEnXDeiiBsY54UOFVMz6IKSMaB59UJeYR/4AUel6gAFBaDFiIrmlO36WhKkFod/ubatYu3QoWCItaDaH0TSdes3gtVMzPPaTl/1Zv9zVyzmhUIGQSClp1Gv8w2PP0+q6gMQwNZpDQ7iuS7vTZ9u6JWxdO4YmBIePz/H1Hx5idLjmO4HNERbmFjj64kFac7NYVo/5dh/bcRFCYvUcTpyc4tWjx3D7Fjsu28bRE1NMz7Wp1eqBA2gHq4oqjiUUpWiFVv4AT8joLFHUbANmA0ut0qDFHwWFVc3nSUnd1Lll9wb6Vh9N0xgaG6MxPIrrejQbJo8+fYTJ2Q6GoedPRBVqKvW9nzwpaTZqvHTkND/a9zJNU6Pf6zM6voTRJUtxXYlw+9yyewLHlQwP1fns137Kw08eYslIA8dxqTdGcD3JiVePceSlg7x2+BCvH3uFMydPMT87i24YrN4wwa7dl/Hci6/ymx9/yB+M8mxWrV3Nui0X4fRdvGgu4FxTOJAkU80XJPYChFkkKLo/os2AoeAIZGC6FYmT8SBDOHS6GD9GC5ZhX33JGi7fthrP9RgdGkIYJkLTMOoNupbLIz8+RL1mKhtHhNqijC3NNCCFU2h2vvDNH3PznglMzUB6HvWazvLlo3RaLW68YiPffOoVTk62EAj+r79+nLmWzQN37aZmCmw09OY4pmlgGBq1ep3m0CgjI2MMDzdASL72rSf5X3/rU7gShhoGOh5Lli6j3hiiOzfD5OTMuYyw5TRQbaOIO27iUIgUbUTMrwFhoFqJqm5CCQpi++hRuJAkdjTCr1YkofmKR0rJzi0rafWhb0vmrHaw9k5nbGyYJ/e/yisnZxlq1sj57vHglJPH9TyGGjWe3v8qj+97hY2rxmi3O3iuE4z5S0zN4NJNKzhyYoax4TquJvjzv3uKp55/lXfcvos9OzexYmSYoeEmjUaDesOPNByvz779J/jM577L33zlCRzXY+n4EPufP8rqFeN0vCad1hzzfcm8JfPPLVIfDWpj4r0Wm9yws+YoBz/6iDuxWDexKZHH/3J4n3feuoPXTs/xzMEzDDVMhaGqC5eHefGgUPaRpKbHu4zVwRBN88NDL9Hz0/WduxGVEgxdYOiaP+oYLJkOoxbpSWxHRj1KaP7sIVIysXYJ2yZWsWHtMsZGhwGYnGnx4pGTHHj5JK2OxfioPwTseRJDgGEa/vCvlMEoJVh99zxaoCbF5geRhiYErW6fG67cwLaNy/n0l/axZLQRR1H4jmjJQZFKRoJ+nvbwg3zxQINHsdiqIYgfmkgp6Nnx7qLgaVRCC5dfZah0/mTzdxtLf3l2nmwJoh28EpCeZLhZA+Dk2QVeOTkbMTQMbU1Tp9mosXR8JFj374OzPej37JAMRFvez39pVQpvlRdZjZ2mmhzkA0QZC5+HUheHGzHD0mGfSDwO32hCQyiTR+FS6vjBefb8kjA5nKsPsymoZosEvRmgZhrUa0ZqeRjRqiTXTdqrcAYxvRMu9uJzZgDPQcZl0dkEBeZYiMKBoPQMsuJV5lWKxD+VQqVcjgcm8xgbfdwsxAoUOZaJYec3IIUMXwTBw2Vni+FSlD3qACqFFcbkPMpNic6ddohFRMccLJTbsoEgtS6FscU4FXSbSkkU3oVOJaBEAumceVohdnTysVYINCiiDSaVpPB9g/NKMtAGwWBasqOkvbZyOAF2DOBK6k4mbhd3TmCid8v0wwIss8wRyi/j4AVEUWfzQltZbjJF6m+2puIyRe8EQujYtk2r3cbq9QchMTAJ4R9z57quYkLUXdaLgZ+XVxX8uM68JKn62bgED2XOw5DIMllm0AcngtHB8JDnsLMKBK7r+Q6aB2iCRt1MzKBF++hUtZcrBOn7rBosOsBaCA3HdVizejVrVq1kZm6W146fRNOMNFEqJ9u22bJpI5NTM7Q7XcwLejYAUYjv0yjpBObhW7IbMa1a0vPPKaJKWbhYNh980hmJEBKCnmWzdf0yPvUH7+ehP/4n/PoH76Tb66OJ+LBHX1IC36OIGbl0TZuN0HnNOpuaBq1Wiwfe8Va++ZW/5V/9L79Kt9MLjqjPqauEj5qm0W53+JX3vZevff6v+OQff5zRkaGUJggAVZaHbLSWuAoEu1C+5MAzghJuYE4ViooVWqmvlvCyC1oopB//9yybt956MXdcu5Wrd23il//R9WxYPUbPshVipSmeU7nqW0VObMokJJ4XgPD8/YGO6wahQ1qLlIKIs0nJdXv3sHL5UvZefQWrV63Atu3FaX21ThXLkC6qb5CQ7ayKE6L0s3F5opOgZhJc5NkqCIRPQrWUQygpY9yF8LeAb1g9zp3XbeP461N0bdiycRlvu/1y/u9Pf5ehhr8ZJFm5QA8GlPztf2rPJsDD35whpZeaVIoBaZqR1EZa+NUNERAya1+BaHmamqIxgqidklqtxh9/8r/S7Vk8/uRTHDp8lEajnnUsqwhERmHlaOgw4ijqmbLiOEC2oE9cKYK55gBa3OEKYpmUpg3zymB9vKZrLHR6PHDX5Vy0YTnfeuIlHtt3lD/4yDt46+2X89BXnsbqu2hRwOHjIQQstDr+QQ/NOnVT3S7ub83q2w69vr/FrFELmx2cPq757Wh32v56xAD6ULPhMzeElFRjaJqOlB7tdjvY7+8bSl3XaTYamKYZbDX3BaDZbPDscy/wi7/8YYxajVWrVl4g+1+SSlyjEidQpDJmkQxEIFsk5XjI6JkMtJKCTUpgXc9jZMjkruu3YXuSR58+yt8/vJ9/+p6buOyiNdx+w6X83TeeZnysieuGqkPgScl9t+5izaolPL3/GAdePhlMIPlrFmzbY9PacS6eWMbZmTbPvjwZrBX01yP2bRvHcbhq9y52X7GT8bFx5ubm2P/cAR559AfYjq0McchIG8zPL1Cvm1y950ouuXgbK5YvxXZcXn3tBE/+6GlOn5lkdGwU6XmB929x2aU7uHr35bx++gw/eOKpwK8pSyoHFbMX9q3Qgc14sAr9QVGzJOheogFEBmZZyoqIIB4aTkUHOYCFELQ6FjdesZ6LNy3llRMzPPnsa8y3+nzniUP8861redutl/Klb+/DS8mdbTv8+q++lWvv2MXHfvNT/GT/UZqNGq4n/V1Fls3Vl6zlX//CtXz3x0d45uAZTEP3ZyR7PVatXMHvfvQj3PKm61m+bBxdNwBBu9XiP//lZ3C90KyEpk9gO33uuuNmfu2Dv8T1117tMxqJhkan0+bI0WP84Sf+H77+re8yOjqCENCe63D1niv40z/5j3z3O9/lez94At3QcztYlj4BQzNWOR2BpdcBJKmtDreX+wApG5vwvTJqPEAkgcsgrzwZQvrwPe7cu5UhU+MrT7/M8dNzjI0O8Q8/eJEH77mCvTvXsPvS9ex74US8Nj+odn6+hTvbpu+48WmgQYcR+OP+na6D5XjR0m/bdhgbHeWTf/Rxrtu7h9Nnp/iz//xFnjvwEiPDI9x1+018+Fd/mddPn2VhoR0cOe8PX3e7Pe68/WbefO9dfPELX+bxH/2E6elZhpp13vrmO7nu2qv4xB/+Lmcmp3nmZ/sZHvYnhxzHpT03S6fbQ0MLRjpTdE+TTuSMPBZEOKJos4jaaUT8AY2SJWEyn5EF0iVFeFxq2jiEV2oIF5b0OaQJDavvsH1iOddcto7JmRbf+P4BkNCsm/zsxRM8ue9l7r9jF/ffdhk/fvZVNKHhoWywlC66JjAMM+V2+BjquhYtJAXfwZuZneVf/rMPcf21V/Hy4SN85KO/z2M//BG6piGlx1/+1UP8zkf/Db/8S+/BcZyALv4K5nqjzn/6i8/w46ef4evfephup+sfLOFJvvjlb/AXn/wEd9x+C+942708+dTTjIyMIKTviOq6ga7pIeEYyNiSyCmmZUxP1bz6TnixE5g1QEFlWYeuPImwcCTQaYSV0a7IlAU1aL73f8c1F7F8fJj9hyd54egkzbqJxJ+W/YfHX6Ld6XHH9ZewdWIlvX4yJPQ8D7yC41cDLaAbehTDO67LqhUruOvWm5Cux1//ty/wyKM/YNXKFSxZMs6yZcsYag7xR3/657xw4CD1ej3YN0iwvtBgamqa//H5v6fbtTBNX/DqpsnJk6f4xje/i6FrbFi3mlqtruAlghNKw1tVE+YkZSV2cZLRL169rWpB5SxjJQnyNEAUFycfDbRSEW8r5Sb8ioXAZ/7K8SbX7dxA1/L49pOH6VgOIwL6todpaDz61GGePzzJ7ss2cd/NO/nkQ9+jPjYUzdBJKcFxSz8OpWsCPdhBZPdttm/dwsb165mamub7jz/JkiVjOLYdRRC1Wo2pqRkef+IpbrzuGn97uNKbOp0uOy/ZwT133s5ll15Mo17DdR2E0Fi9aiXz8/MMDw9jmgbRTmApiTqDyKFtWh7K+l+uIKWYXwJq8JdDo7vkPH0GDxkE+iKstThvQr6EP/DTave59cp1rBirMTnT4p7rtnHHNZvRNH/xhKHroOs0anUWFtrce9M2PvulJ3EcNzq1w98f6AVbuIl6vX8pEbqGrmtour97SEqJaRqYpkm726XT6ao0C9rlRxHz861gx5GIwrxWq81999zGx3/3Y6xds5pTp8+w0OoA/s6fet2MRvlC11HdA5gfVy0iRXzOkxCRuAx0A7FVDzSvVAQgiuiVs+oiMGlHRcUjOoywWgMkykkbwp8/bxiCN12xgb5lYWgG2zcuxTA0NB0Mw8A0DOqNOp7U6HR6bF0/zk1XbeHr3z/A0iX+4gvPk3iug398bzy3EE4ojQ7VMGq+7Q17crvTpWdZjAw1GRsd5vjrXuQgAZG5WLNqhb+1KyBK37ZZu2Y1//b/+AjLlo7x+//hj/jMZ/8WgrMAp6dn+IV3v4M/+cTv+wtDIvr4oWu0PyavW1WNvBL5QmlQPXSZA0smwggJGCJaIJh0+tR+rJrV8ODo9EjB4pIPRReCTs/hyotWsH3DUtpdm7/86k84M9vFMILpUiGCHcEava7Nh9+9l7fcfAlvv2Mn33nyEEL43xGanZvDtnts37TCPwNI+KODtutiGoLrrtyCJwW6aeB5/qjca6+d4OixY1y1exf33XsnT+17lpHhIYTwt6EvLLTYuGE9t916E51OB4I9iVavx5YrdrJ69UqOHjvOpx/6HLbjUKuZ6LqO47o0Gg2E0GMHLBW6+o9CZ62i2SygZTzUq/4VSo582EIETmDgJwRPs65fRqnL1IGD54y8j8TtV21huGHy3NGzPPKTYxw6PsOBVyZ58dgUL74yyQtHzvDcodPse/Ekn/v2s8y1e1y9axO7L9tIt2djOy779r9Cr9Pltusv5l1v3cv0XIu5+Q6WZfEvP3AH1121nZ7tYeiG7/zoGrNzc3zxK98ATeOX3vsgH/jFdzE7N8/s3BzT0zMsXTLOx//dR5nYuIG+YwdskghNo9/v4zouw8NDbFy3lrmFBVrtDidPnWLLpgne9wvvoWdZ0VkCipuWZJEIYyZ1wGdxBCyICIFy7pQsCUvLTODdC6GELTlgBwlyQjtJepbLuuVD7Ny6kr4DTzx3AtPUGQq+9Jm2KvVlIzzz4ileODrN3ismuPdNO/jRT48wOtzk7x/ez7vuu4ZLtq3jDz7yDh5885VMzVlcunUlV+3axIuHXme0bqIZ/qSV47qMjY/y3z//JW64/lredt9d/Pvf/Q3uf8vdHDr8Cs1Gg1tuupbNmyfY//xLXHrxdiT+qZtDzSYHXnqZZ559jltu2sv/+Qe/zd9+4atMTk6zdu1qPvSB97Js2TLfrxBRgwOie8HPJ0jSuipeW9U+JYujvCztU4AFGPn1lDhxQrH76XyJ25xWqHULge043LZnglXjDQ6emGX/kUl/z73rKeFknHRd0Oo4fPuJg+y5dC137N3KQxtXcPLsAnMtj3/9h1/gt/75fey5dC1vuf1yas1heu0ef/bX3+bAoRN84jcfwDBP43keuqahaxqO6/Hr/+a3eemlQ7znwfu5+47beNt9NTzP4+Sp0/zGR38PNI0/+5P/iGEYCPwFKl3L4jc+9nt8/Pc+yo3X7eWG668LvkMg+erXv83Dj3yG//Dx30bdro70D7qo1erBaON5Mj8MvguEQCZyxgtsQh9HAGLdxOYEqeNl4ds5fnqefcGy8AxwxV+IkVGxF2RPClHVnL9DdmL1KDVDMNfqc2aui66VfOYscFcaNZ1Na5ZgGDrHT88xM9+lVjPoWTaagCsuXsfGdcuRCA4fm+SZ519h1fJRtk+sYK7V45WTM1E4J4TA8zwWFlpctGUTF2/fxujoCJ1Ol+dfeomXXz7Kls0TbN40wdT0NK8cOx5FHpblf4pu95W72Lh+Ha7r8uqrx/nZ/hcYGRnh0ku3s7DQ4sgrxxAIHMdh1coVbNywnrm5eY4cPZb/DaPKKZjMUoRHRm8EC12LG6/cyLYNy/n0l59h6VgTV5mlFIKyfQEpAVAdTKoIAMq9TOWBcJ6gb/sncRm6f7RbFeH3pMSyHKSEek0Pto1JNOFPDHV6fRzHj7VrpsFws47tuPQsG13XqNeSlk8g0HSNnmVh9Sw8zz8UstFo+N8f7FlYloVpGtRr9chAapqGJz06nR6u6wQ9XGd4aAhPenS7/uhgvV4npLhjO/T7/QCP+jmZ/TglIzaprLDSBLQ6FjdcuZFtG5fx6S8/w5KxZmLFsqDskCglYhAB0dNCkFsg931kAEmGJ5J6TY/GGaKeX1qPHz2Ea/SjvfYhjsDIUD1SuVJKXM9D1wQjw3WQZDaZSilxXZeaadKo1QhX1HrBMbGmYVALZhelF+MfTvWODA/5cZHwYYUCNDIyEsEJ228aRjRTmd3sWq39iYzRUnBlnlWENkFLgsmprtJA0KDHxe9jhVSEhZRVzg9LQQ0YW4SPT/CMG1u8ojeST5k884f4uUyfS5Coz8tEXz6O2eNjJDKxO6cwVRECJU9u9gpjMwMmo3OC1/NKxUxLj16VpioR07lHpvn1DYJ3jlFcNVgFwBPxfx6I8JTnMKU6hShbFVxYLF3NIiidmFMugq4pz8+DixeKGW8UvHOuPM+fCu/TPBNkV+pK5a7ogAgRnkhdEO4FxTPn1EWxZhp5lZkDBgoElVRXXN8FzPfzTJVjd8jvDGkA2fdSzZpXn8w9ICKILSvoYRkMQ8bCIlJCkJZcmYtqMq9U4JRWPziVye9i4FRNi4E1CKdEZ83xS5Akl9angas9Pa8CP0euBqhmzlIbIKO6Va2Q/hUhpCRl7/rAdCHtfFmqytgLhU+Z9YuUrRKmZUbNQg+gCF48EFT8vYBMPF+UYmmq9D3AwiRzL6tkf+PyLEIY3/AUcDGDS3jodNYkDJbb3N3BMiic/vhgvsMX7UiXcdlk/sU4iufp+GVSFRtQlu/nwfk8+uQ5fSUpHKIpsPNlUUB5GCjBH0xIHjOmIpo0PTKVR/0rKmgIpd5B7R5Il8XUNyCv5DxxCpVtnoMcjtGn3kc4FeCVg7NUnwX4lhlfIXMFIGQW/uxf6k3RXRzjDzjzMpeQIvl+kBNSSbVL4pO0ilJAsCqRxyBZqmJKSgupozoShOf/BsJTP2DtLz5Rj48rD+gLhoJjX10pHmnJGHjh51MFinQqDcsQOa0pLmD4V6nzC4V4JXXnBTVp1T1Q6otMYXxsjn8rBgujQtIKlaNyNPFU5A4ExTY8sWBVpgXDP6c/b/+gzKwZKGpQDmEuZFg2MFUZm8hLqoqq6mmW0UC5riTcgzqSeitzX4X15QhA3LC0YxF17IG2tUh1qYBiJCpmPIc0iEl5z0OrOKANEfwLgV8RLmWCVsD06HXoAKgaO4mvyB8ISlYXrSRNVKcCEqlfBVtamWYD1LJQb/LyCMJZseopp84MrQu8wgq9N+vX5TByUD0IkmcM5SVVcOJd2glcyB0KTjsmRUw4B11dlflVQEeo5REidGTVyvK++h0Kh0YSjszPKtMPF9n7VZcjASOP+SVAqmRT/DO1RGIbuxDKJ2MK1bJPyAjNIudOpstlkS7OkxHN7LOCnpBPCtWjLkApV2gSCOTgVJby/CEFdKG8pF4k8pR8pzgDI+ZLeBhGzHhFb6eE2EiEHtHzHGyDOmJRCCVKDTLyelCaSYKkfR2kyorSoHJpVzlHeBLH2qVh55UZhEtaTahRUpH6K1OLSVrn1hyhmnYgsw6uDy05xmsIoZ71my6fjOnDxaAiL0wRsXMqExtBs8hHUBO2Tq2jioAUw0+Uy4RMIbLqp27CfGpmlaHpugoYFjIh8p7TflJenrx6yMEn3XYf56iGVLaIxtKvM3aXYk0lwnMChUzLhcBxPWqmljF/pUuQI77KALk85g8KdWQx09K9JZEv3dtlCuH4vYjyKOACrzlco5/V5sWrmrLOoFqn2hNFNk/6IMxMp/dSDwLmIXPpm8DKk9RNg74dfKlcZq2U5u8TV3qc9BeGnp7usHrpEK7nKV+/ltG+8lw6JIhQoCkGpMHGIO41SWFUiRSqdlUQgj9lGjXilfqxC0FyL0SJulY21cTr8lJCoJyvFH5EI65LK+lhOXUnZD5ZTgiB63msXTHMibPzweGUSRiSHC/Dk5JGzeClV6dZs2yYpSN1HFfdFeuLUSwIVWxUXDbcXOH/gp6miqUgRQSZ8wsJnMCcxJFxGfOS84mbSIXGeZLHIuTBGiTU/t6f+BDpZFmBF7QxFrA4WxH8LA1k6rkkXpeoCYHteCwba7BxzTg/O3jK/+JKzlrE3IEgXYOFXp+fHT7Dg7dfwqe++izNRg1D95dBq50qUkQh4xLvoibGd5HPEfeEtAD5xwGKEpMbMiXUAjKbP+MQJU2ADC6SS1JF4jphKnLwLE9JWCoK6TZH9Iv8glRKtz8jyDFgTWjYrkur0+MjH7iZpw+cZK7dZ2y4nhEAAdl9AWHShKDd6/P2m7ezYskQn3/kADMLlr8pIte7Vhwr5VV0XkDg/JSbj7THUZw3vx/KDL/SYU9RB1ZrVs1yrk+XlxKaLudVWE8Et1BaMyTIdYdybsIl8EtHG/zKO69meq7HX3/1GUZymB8WLRQA8O1Ip9fnup1rufKilZycbHFmpkvfiU2CGkn513HTkv5ZSu4T++Yz2aKCkQlRCRfeKr0leh45V3FZIeL3efeE90r1kUBED6Tin8qE2pbB+2izTFrOlDP8/H4Q5lHP/hERnESSSh0K4Lj6OH/N1FmzfJiJNeP85MXXeeSpI4wMNQqPogt2Bm0u7hYEBzh0bUabBhdPLGPl0mH/wIaAWuE5+OlJ4OizrUG+NOPCPfwRk0VoRcLPmqn5FAEI4Spbu+K/obMqcn9k7pN4CgVegv7BO/WDj1IqbAnMYvg+Pl8pyBPtpg5MoUwKjSoIkRDEbolSTwg0Dt2lIhS263Lq7DzPHjpVqPaTPAKxftNmRZjyM2tC4HqSXt+JdsNUTVlJDbidM9kiEwVyoSSfyjyDITOXsuRdBsdSMy/y7yI7UZKvXI9nHg82Jqk3QafRdY1m3UTXRSnz/TKp9QDqzlE1ecFRKUMNs9pW5PzqMs5POixJvs2rSKXyIKOci8K5J1V+q+RJm/ZSVHP8gIGVFDyVPr8GMT/Ma6jVDzq2NP5SRpxEwlCGcAtsTlirkl2tM/qapgRZOB3rYxvCqvQF7Tc8nYMwZtIb3Y48XSkxQgace/UiK7wFajwvAjg30v3PwvgwFUQv/1OlPNzKdgdXBZtQCYtXz2qu7Jbzc0+hM/eGH8b8/wnmFyUZHBJ1QVIajvoVkCKTkKM9yPdDCkrnagI1RLxwgqAGiGoqmyMogROcrJov9MVtqwZf5vzNz/v/Ahj/+oLZ9nq4AAAAAElFTkSuQmCCiVBORw0KGgoAAAANSUhEUgAAAQAAAAEACAYAAABccqhmAADG+UlEQVR4nOz9efwtx1Ufin5Xde/5N51Rw9E8z/Ms2ZI8ygZjY8A2cwbIDSGQkHtxIPBIciEhublhupAECJA84BHAgAdG21jGtiRLsi1ZskZLPpqOdMbfuOfurrp/1NBV1dW9e+8jIPe9V+fzO3vv7qpVq6rWVFWrVtHpZ50t8D9bIgAWVgIAmYfMeorAdwoACTdRWLnrIVUOazGYcyYCIHS7fHz85yHM/qYS5R+BaiWWdfCxAeTtqleevE8AJAAxbz/Y/akpUZxEb/ptqoJECI9nNV3XSwJEkeGmkwY3fyojVFKI5ANocgrm5gvCsJ9TRb5FGHV2D4VhLtqznmATGjqD27ZZ/XCSaCyazDi6eNbrd52fIW9vnWTldegFqv/mHXWbliRs4TyfN9nMX4XPrHcnWb1KTg/NB/avi5p86aa12qKaTdTIMw9ei+RZdMB87VOm4e02+v1laZm/NtOkDB/yftv4LCIqZbtmN0O1/2+svfMmn54Xoe1FhBkA0iwvy7LynDMhLV70da1yFh6vl6Cq0+F/3X0SYqhQvVz9AX/zat/HQ1h//u9if9XrwbptCuWbd4x8PPXvRfv19WB+u+y8LMyt7wLxnKVfxzRPB/rzwLL5P+A2cLHk1jZ7LeFvLvmMX4XP62X5+FBFLR2c4xB6Vmc9ZdHVFD3+gTWARRJBrRvUWwOanYprGuUw83rn6/e6if52BMBiQ/vXQ9Ch5OLmMr8t/8l58reZAjjUQmsR3OfJX8c6q8oT0r51U2jKsQD9iBoCd+5urJrO+a+kQMtzvr609rciAP562eWvlyGLeuVvm/mBIA610Jof9/KZezWssMCsIyAqmLbWUOtpR2j3aJ5UUW5hEiiBWQovZD2cXDqJNYD/WdMcnfN6WXP/P5/qLcvVzRtOgc6uvcmhtf//Uwes7rRv/vT/hQJgji56PXZy/v/pbyjNI2T+9tPry6be4i+9frtar68A+OsWsDXh/89ACKWo/nUponlh/j9VGValk9kceJ3T3EsCtaGSWpc4uUboFa3Xdw3g9Vhw9T4BS+FSeD/A/x565sCxasu7MV9lLSp4bwGJXHhu3vDv2S8q8vtjHVqQrgPX332qmkML7+usbXWhM4jFLKSygQTCfRCqvlCmpCCbDS+YFLhaS5eBqktpMTQuARTlSgYBJCAEcDIeiaSESVzdu3MuqCkwRARiklu5EBAB98tQjY4AENYYLoDObJop46qSrKERLUXI40wh8vKGUXTxAMPM2nXi7ksiKvZxiMG59byMw/xm1lo8J9Tafp0ldOwdPBtcnY0CjmJbIcLt4V57ATUuFnJ668/AFFb/+dLQbRQJa4i85IxcqC8sXCMigJGUV0KAW0whmb/W4FSmWILyal4wEZOtmUxTJGmCiIBWI0IcM9mIgrjTKaQyfM6vrDnP5n3NpWTeZaEyhobU7xyUu4DkCCWtYSoFkzD9S04lFbylTB1HwgfNIhc38xjFUTU5Sdk5QuIjaSqMialCfyhBKFSbc694f49avhfMaqOFb1C2Kf7L8VTjZ7ShypB3vtvKQr/IwqTyi7wBFnNbwtfA0OMpQl2SN6Yw3rb0qlJvNoic+PJ9fgKRAOcCaSYwGiVIM6DVbKLTagAQ4BBSyZagOE+Kq42I+tqfMYbhaAqRJTj3wCquu/QCXHjmPuxa7aLRiMGIDKNpUz7nIzVQBBCYHFhSDANSn4pQiazvEhizAJJ+T+TkyWEwkKI0WY+bn0x+OWCMJBXLTw8OAMHkwGnciCwYqq2k3C8FCIy0NUAOrhq2bLJ8zkjixyQCZtAlOEU+uv0aDgNATPW3KqthG/ykhebA1YKBZLuIaXyZ6VdbQZL1P4RwzijJKaoWMEpEcCkouBBgSpkKIUBCKIEiAQhlMeZlYD3X3zmEAITgSlrk+bVwgpCMIvOrPFzkFqlVjykPAXChG5CLfC7ABddN9fCROYUQUhxa7dCslVtn9nuei0KBHG+tU7jAJElxYquPp776Kj73pefw9AsnwKIGlrptcHAYBBZMAgCdfta5lhypD9CMNwFCEEbDIS45exfe86YrcN5Ze3HsxABfe/k4Xj22je3hBJwrrUAAM4LcarGBKj9J41MwCsj6IRyGgZKM5p1GEEJqPqGlLUDannNMc7teS/8TwARTQimndIOaNsEp10mmHQ7OVmu1cDPfhWFUIfLnbslcwAEwMwsp/3KBA5tZdQN0WbJx1/UY6Snxp9wm0P1reoZyTeVoc22x2PhDMaQlKAzu5r/8h/nJbWtLWHkJpNnaWF+GC60pv2UdWEwuPyyVJ9yxMnkQSrkAKryxnudWukSIjHDM+14ja7S+LUChnhOh245x5qm7cfl5p+Gs03fhuZeO4jc/+hAeffYIektLYBDI+El4vhIDnX7WObbhNVd5RoSMA8Qn+MDbLsebb7kEX3jyZfzxXz2Jg69tYZIKMIqUBgyZTbA6xP2pGSG3AWXnuWtNthmoU97JBAEI5hKhRZR67kxWySBuBrQLx22PkWoeNnkZcp6pJ8KCY4bBU7UOYrptufmticbpHIugizAsnL1B8VoIY6YKABRySC3CKPQAwZjiIVxsQWMzpS1sZ9WSl9EgJO24C7vuWNhrhLoGV0Bbmcr60EZK2D3nqQmTT+R5yVN/JiuZOT/nGRoMuPDM3fjAO67H3bdeio99+iv45d+/D6AGIgYlBOYbFV0PHTjrHMFNB9RPRAQugAYS/LPveAPOP2sv/q/f+Swe/MqriOIG2u2GMU1rYeQbAbbiqNGYIDxHjdUE5CkF2EKogOA8ya/YFlQW0QlbKDjk4UIzbckbVcDIRx1lXaC5gLwMSjDbz0MAQt0zK/njMgNLjYbhn5KqXD3iIpY/s1cPciRsZZJXVSAI5225cA0gOS8hC81nAqNJgmQyxhuuPRs/8f3vwnMvHsWP/cLHMOUxGNnWSf0KiBjowFnnilyT1ENKWoCELBnjh7/jDpx1+l781K98HK+cGKPXbYNzbuY986BUNajzCQC3xEICZFapxYCimmLmrWxOJBbGebH011XdbLi2NfQ3nBZtdKGcLenkOg9jEXaGYxzY3cDP/ej78eqxLXzwZz4CipoLVUnEwOYmRRJgRBgOhnjvXZfigrP34yd/9eN4bTPBUq+NjGeO5K1EzJM7xjgjgu2jNH/jqPiL/AezWu6Lbu9xGVKlYCkvpCfu3uNwYTsYhg6OocuRW54889UHW4YzwcOjIm/dtDD/Wcj4OFk5aldOgbFTaS7a11sUJTjNgVwgjz+dBPLgJfIFFwJplmKl18KR7Qw/+NO/h/PP3Iu//403YqffB2OzDf5CnQDYvNqIEcN4muL8Ayu4547L8Mt/8ABePjpCp91AmmYeuBqwyf6iiN3MpcqoaNYouLWbabUuK6y6aiBozESycStxoqwUDApvgRxOqfFlC0GL8fVvQcVymo785/NMe5xnoiLDbBAuAlVOp/q9H+mIQfjtNnnqjF+1Og6+KTRZj9PraE4UPPkqpJz3KskydFoNvHJsjJ/8z3+Md7/pKlxx3j4MRlNngbgILlSn6d05hAAR0mSCd999GZ46eBgPPP4Slnotyfw53PKGzTVfpEDH6wrMXMR7l7t5ycVHBujPYCitKnUp//TKupHKZDMmeeXKmmKvslSpYhuP4gqvO11jFQKkRqrT/4WMJZaReU2SrIyQ9S2Y0Jh5iNjGktDl/HBweUbhjbuTz35Up58KTbboQQvdoJlVJ/l059NlvfBnaZphZbmFz3zpJTz8+Ev47m+4CVkyza3KQBJOo/I2zTQcfPSnSYqzTlnBpeeeij/5zJNgrJlrV4f5K4C4X1BssPD+QgBsQeAzorYkVM6AheXWFUQwr0MvzAGl3qWzCYKD9NFUEgDV3b7RgkC2tSjJTyKVgqla1bJTzqlkPSM7vyEOPxqQW97kJ+QastZ8mqwa54kdOGciW+noNE99Vj7y243Ab/3Me663MTnQaDbxoY8/gqsvORMXnLUbk0laaQWE6pprDYAYIZkmuPaS07C5M8KzL62j1WoqpwbUpMuQBnaMdRQJxU5+qKsQRxY7rhy1Kpu9aAIXD2JV2vyBJ4opane83zf5c72i7Ur3BZMDvo6F4uPn/7SmTSFONgQTmseQeV/oppn9VlsDzZWEsUTKLMeZEPL6ZwYntd8XrTACwAVHt93Ck187ivWtIW69+lxMk6TGWoCL05xrANJN8aKz9uKlVzcwTgQYk50ycxvRee0zPXIJW9ssDX23bcfyUq4QqcLb11b6t465V1W+2JCZTas7FHomoqcoQdO6Ppz5igUY1oxhmVYLTRVCuGgzm4yWLAhcQzqBvqeS8XfqWkwAkFPU9sGoU9AuO48FpxrrWM25QI0ihuGY46svHseVF54BEvM7Bc11HFgIoBkx7F3t4dCxHeSMX8NWcxoR0PBGHsxr3toUHIBnP3DWE3goU0UKWSWvozkO1JcQptqQEApyVvidwJzWSChZ8CvHrqqvhPda4xTqa3v6ZE8vPNdY2xIpWDYLam8RmopVtLmg5wIWrh6DEosztxjsPLq/CYIYXjq8idP2r6LVZNLjtnaaIyagNDuAKGZoNWLsDMcgFllv50m6EYtQnq2R3Xpn2iH+vLuG3PqbTRUIERBeidaExHI1ZcBUCYZZz2GUT7Ha17PjbKKnkudV5fwgoH7y5+yLKJhQvfqnr6XLsxZfqjYH+9euu2QcSSBiDJvbA3RaTTTiCJwLMCqrujhutS0AvRVGqhRPF1UbNhKLEFG5uVt+sMmXvGV5aqa5mj7v1GheMHo+YGvhEs1ZB3GfvkvrrJtO1sqatywQXhfxuaLuAl4VLrTYcoBf3knayiC44d1DeMnCWZZJR6GZC4DWe9UsVwCE2uoRqGGyqrBEtcewBOFK4V8CuKztNiwf7jxzQ5e+5kjFTj/ZWYNrHXtADUH6U5wagsg2zPT3um01aGiLZEFmDvVRqRldnkjorUjAbN/p87MhH4q6OJ3k2BVgFdprz/P1bwS/62Uz4QmSeVB0BUBosM2aypwcYBNS6YvqRYuCFQv9IABYVOEl3PWpytlHhQleKDcnNRR87a3vc4ByF6R0Wb9ffO6dIWyDlSyQzGKuEgI19EVpvbYgMu8IxbahWEhP9+zyRqiJGe3zkPX50XRz0aSuTFZ2sptRCaZCiGvjT50dkW9ZcCJVSCr/HFMAXc52vAji4w5Y6fymPJnxmjlI9bWMkQ8zZVjxRTm5aQ0zs3qDRamVNi/D2XWe7HTc1/aLwKPA1wW10tzWRylOIhdI8xUsPvHwqbXdXgbWR+ck2ilQLcyqQBPmvhdAz6xsE89/ixrEWaGCtdKoNWahTCdjcoe7y2lXSKDVpQaB4uW+wdrmsLD090XSnAqsLjxHOZpYA3PAmAO/6q4sq3vORguPhJwZxGJTHD8UQTFThcVG+Q/r60JDOXdQUGEsO03NAXauxKTS/q7ZilAHcRSskrnn6r4KDJhcwXJzVIF8JqAfLTytdKquyc0Fk/r1SKoVwRmHb3/XABX6XpU9ZOZb5aVVakX4WSQtUizU5FlwBKoViqPfCHpDcqFQ4WKhsOB6oHPSDWr+YHImUAWwYRh1e55Q2NM2M4RZ9mSochGeYZS2saJNXhXB7WqTKiS/bktwblSzn2xGmdtCKsOpDoAK/Hw86s/sijAKz3XoL7/BNSvQTFyavdySLVS1UDeRBY48GMJ5Z6In1U20gAUgWalkMIOPBWrtv5ZqJd9mrtAkQWmre76KaRiCIz0XnwVehLTZolqXNHZh62Rua96bzy4kAyr7ohilJpysKdSifVRFGkHzeoHWltL2PLwwK5U3wPxvIsuqqTiEdVK5hkeuV9VcFkCInmdXZ49qhSaulLA2jHmNFl2nL5J9fOYZsRltqUyL38WSd09x92RuTCqtECxmUjqprnvynNOCk85XU+vbc7STq3COVAbP70f9W/UxAZXnJyqqqkeNIv8w4ZkpFKct1ICQw0WJtiw8zolI2Plmol0l5aNARXPYaOYoMAX+6qRFGcsSfsGYBHPC9dCtPRWpBdgXtqEknI/Keuc2b04yVQ4n1ZtN1q3H/qwNR04DnZnTghYHq1XSzPlZzopmKSAPOV1c8Z0h0QqYWz3iLIQoM1E7cpiAeGWCQBNe2aEd336ZMVEjv4jfxgqrxofjt22m/LArtqcmun0nO7fwUCqYdvMKLG9CP6td1nthazbK+1gsgkahOnFSXZSXrTmVrXqnx8/rqmrIHq0SN1ZaPkkQM2ipmAoCYFY/G+NZVeTMUEj5CIgQUQaYlfzoLg5QR+uTLSz0ol5Q9OUtcNfKLMomN19J8QC+3mfhpYd4CDUjxFyUypPwvmsCgCGAhWcjIRznz2CSY6XVapsN3z7Xr8tadzGE9EQN0Gb992Qi+rxe1oeJGS6BCq9NBWc1WcgGYP1Zr2xmFJhj9kb6ODBZj+qWtctphPSmhK2dy7Dxa8rdjHMXAzsmfACUsUwqUITtF+4LJK1FK9CCgNldsEdMI+aALWq1AjQ72IVtcJQOmk8Vuq9Jxcb34MybauMxO9nWYX2z3VYANlqWsPddd+viqJnBadtiDXQ2mGqBCGWicjiLjqGmp7Ir5mZgGPux2KuSy29kbj8pHWtdoAw+5UQj4O8VqIs9gqZ7VbLyk/dczB8BzRHLfkPrLJ+bgZHMa24rEnVHybVYbBQcuXMyJnJJ+do8rHNS/nVhgrZlqOmiunapC8d56pgndZKr4AqRdmpbNuqrvjjGAltliLqF7amevuQGrimOQFtnNFWK3oUJR8ez10JAeOMkAMFBInd/dQLhWszvIi8FRnHMfW2tM5SfJ9BXQuXa21YFNYMnCNciyUW4L8o1TJ9QbBNBbdwIu8yM5FkXRbER6qs6SVj/l+mrunC0ZvPnbvXxcH6aMasLK2B5aaUoIM+xz51yYNL5Lcdpsf6mnGTmapvBwsAhv39Un8kcdXeZTO76WLgXqM6QiDqPLZaEYoGQf7bmExK5YCtD3HBkwHy3wje76xT19ZnbBhQXyOrabnY/WbgU2lZ3CAhu2wC3i+dKUv7nl4RhMUBqsY7qLScHcHD7x6WbOQS1Ael2JgU7vC4coWCQYwHUhmZLa4cW50bGq5VcWArX+eSSLB3PIwBI5O6UYlaFBRNXaWA9IyAZXky/NlrRnteGAZcPqDG1HaxDLUFuWlW1IfTpM7/+6prq8pni1JCH4iwUq/CZ9axWssZHoUd2O2fipDJV1R+C4zwj94F9JZqw3/t4l1QaGv8y/EwAlRIaKKz3LJDKSAVzChFntVC43yiP1FHqoFeB33whwQDoiyOtaPmBXFpy+trRXrgSEBz5bwD5pZN1EFlAbTrFFuac8mpC1ZYS4Otb/fzw9dhYFoAtq2emGpnKZK/zxWd4Hz8vneSNuDkesyy3GmkhU35OK2AmbEVoIg9TVsaZxaJiHgHgS56qSvRkiaBXcHObId+XF9Zn8a7zAHzT2QuY8wWwJ0lIswZezcu8SgN4nDwqwVSre7xG1JqaWHOrmZXYefwyfvnQtEoUXy047IWkT7UF380q6+GyME4BC6rK8i19rCxrxRe+IVqV5vIE1IMyK/Y4OYW0rVDkGJPPXJVszjahlMBqGgkmc6EjfDxeB+6rnguVZ5hL4/4tpDK8RA0BEJwu6TJlW8S24xZK8rxOaS46CqTXZexEsXwQ3qx1EG3+k/1odt2gmgLANtv8mzECEkxNABzSrx5GPZ9Uixv2Cpe9BzxvZwe32shDquaUw58T2iZ+JQiS8017r/v1pOkFzdBCef8ZAs9NCvWrh4uR+d7gkV8ghARZsMjNIgJgfNwEwn+hqmalEJxC82sCdOBoXrLXP8oKVfS3Uq+6n4STbzZe8wcEUStFQtia3cY1NPHNEQrfL6/CGBnbJV8gFEF4NZMqWjgjNafpRqRxz38beHVg2JdjCJiFG5E3WqX520kWfyyUqubpwbYJFC9nKYEXKu9YBiWEPZvLK9pLcHYClJINOqfOFN7ue4cSnbWMmjRqN9cvXxm5KIyo44JPuSN1LXyUsl3saJrQaIvAWoCoqDu0Tindf4SVJ9f6dam7xJ5zDt5rvOa9DwD5dFH9GbC1wQjzQUoISD4guJYFBYvlydYYMHACwZnqoxXSkDOT14e1NKyn4cs472S0tcHNLevIHDvVtSitvi5aXLkbe/kith4zi66NFaCBVUX/9ZCw28IhF9MLFpQ91SpHqdoCCAke8yywDehIWVvKzxrFXIo60tpXFE5+G6ZGyi8oSnCaLzk4+S98FApJ5JYReYLOz2+cLKrwVP1E5I7AvNMjv4yvKWsXzH8S4OzkFVOIJry2kvV2YbPGgqUZd5H+seCYr+SjJaxvZdacbdFIe1Q4hFk3dkIRH2H5vSzSXdVhwcn/Ss7YzbbO64hwy2LQRxxtb7tSuGT92a+Em8WAn0OdeGPnSP6yVDl+vlWjJbRvl84IaaYRESVTmpNJZPfTnMm3jCrHLXRJqPXVtrYqDIWZyR/7ymbVqMS2AgovZsBwZJ21Fb4ITdrCVr0wVqCdOXgBabG62WHBq5CplPZzJF+QGP7wGcT/bv9WnekXIbvnaiLrV1cia2oliynIPAho/9p41S07wwSsrmTxIuQ/CGWusB7199eLrojguinagENXxlfBqqjLEX5Fqyb/FPnnPNNRX8+ZRxIx+b8WCWX8Igpwai8CCmujXsbEmFdaVLwmAMQhQKCqq5ZMsl/WcRV9vSTVAsla3JH/2yafLfm0eKgK61ShMRwh6hP5/JdG1kr2jGXO7i3NflLDZCsDn2MDWiZYdl7t7CmemXlDdS6WhPdNWwG5iK1qs3xeX0UoxsyljofFovM1Iw1zMDk9WavNBaFWYkLOrmx2gTptmkdL+eaZ+fQ5R4BCzFprJqVhsUAzXy/uLMlXln/W+0WTglfsEj0o3rZr4L2Y93xAgI+C8tYvUzV2J9kvDi/qKF1UlxVlOVbIXKe0bYVWDfIMWJUWlT7vXjDhdQu9mHPzSuA6oZcDcs78mrlOEYAXxMf+KzHd5jWJhd8vc0rnRYV5GZzXBZ7VHuE8qUaAgNAWW3nZCgurLEugiJkJFcbOGu+T7Bdj8JvKfClVVYF8Fxfi9c40vSkMuswUrMAj/NieI+npQJm5Vg6lmNXLN0/nOzAsIUTA7GVzrwMoN/PlEXGLS4xgswVCWS9ZHpkCyOM6kFd2Bn5Vgtuvvgqc/7wMRp3kD6+BUY+opYYSKCDhd01l5Z5QDrS7SraWg/diac5qTpXyMKSo3OoL7apqqAQc1x0bM69QRFYwNRaS9nZn27C4RdB57YUgmDPoQVg96PmGzG196fJlMGtCUYcGJXFJn4CSPnDUR4kdaXeP89DOPwNJ4fZHJXFWmmwl/bqIlhP5h2QyT/CWArWFqXRUyw+kWQRbeaCoxArzletJJC04ZtJh4KXiPgjTHgI4qdgH+aSg7uHg+ouADmKiSPxlSnkmHp5pR2Xegq5E14RRJcgdKBaFzzV+jiKd76BFsXK1wKctB/IJOjD3KPwWSsGRKwAF/C8lydNqnoFVLdQqSJYsZj0Z98TAkBWC59WxAsjDlARyN/a6xFt8nQslv845LS2y4NS2knREKTIXggICxODxTH3zyxMAsxqi3nEY91hTssBY1bAkwhbBm/37MqSVBtDWXZn1UKwoL34S2qhOVcWkNbiHg31tdbn+dYrYeXxPRJdpZzXUpWZRKF+VKmijoKUXTJVl6w5irh0FAEaQEX0cIRCqOPxOW21FxWTjNSMJ56NotNVJvvxSWt/eYJ4LJ8x1mV7OgLYC0V+K7FiNgFOWhPJr1rD0j0AJXmoQu/xmWX4eYhXJy+TDm9t0KBYQQQeNYr48cpKfjRnIBhL5mfQC6exGLyoXw5AqxqZWcb+kr1VqQA8s7uoANOXrNRomd/Fw6HzucBth9MIo1ksO/hEAlkfLFwCEPlavH8za/p0rHoBX1NHCearfNkngwT1vQ5W+aeYShPCZLEQfc1Gjh70PrwYs4XwriOt8LcE5neKqB83UBblu1Ie9A8KgzxS4hzsDfWO9smmJQvntJsyRwmKvRtLj7Qsze2oj5jzHQRzmTGr1dbxefZQPS53GVPVTqPzCEpdZfwo2CQji1k5ASCmUpVnHgQsKiKPghhrIVzuJoj97UavNIqnXQX/NC6IifxjTouGXawFLUOgFq4KQEwAFiN/qI+0+ZOAKwDHwArK0+MLDfmFunieJyp9OmoGL8MvbwsP6q3V/njfNqkzBfvJs/tct2Ta2MAJL8ky1tgqhUi0ACqal/2yOFOwI60iEZ4u6F7mohjp77zqzdT/aojg6zDhnqiQUrdIr4ArAbKmYNluDbI556huYVOAULSxJyJVuBnWEWtQSoBQa2zrMuLBAmEdFVhSv0LbVwheGdgqCYpE0E5852lXL0rKvpHPrMOeBStc3QhjJSmvuAoTN7JOxrnNwyq+gAEybM9ydDZgfPrFS/nyRwXXUp2392fvyOWoFlplppJDLmLoPST8nQAh9TNtNAcs8n0BoK8rFr4hYGXP7nWXlf13N1zkkR8j0nsckL0XB66OTSeR9zkhB1P2HRUOxsoC5SAfzDouGQ3UFgBX6g2w+8QSD35iZnTMjk73NZwjA9qMP4KG/O04x3usyAiv81LCKo1SL8U1Oq1JTdx7+XADWjWE2vmHgDGSiKAtlDQgbvMgtq5nHTguUUzbNqNFMu0yZXKmbSoZvbhg+DovgpXBx4lbOSevk4WNQ8cijGgmljJy+oXwDzeTJZiCX/57TD0BSGFdfyW+R9bX2VWy6PfYcQMcIVP4G1bD8Blq9SdZvV23WSHkmImEYbHGazLeRCMJEOmLMupvFtk0dU0BWbiwFCIDrM+XCbL17RozEXai6BODfveQ0s4oQ6whzja4wdFlj7AJJeN99/OZJlH/UF9gBMGS1CXCPK88D09OXZe/CydLz5lJQKwSJ2Q7wV/7DgDX6c4QEy4mwct947kFXUisg8Z2QYw4xCIVE6MCHTgwzQ/eUEpXqVuXcYPBQJrpbro5IsNuoBk9xx2g0dHIJYfkaGitcgIghYoqBGUOz2QLn3Al/EEwGvZJTgXWYrKaWFB6che/jtPs4t1YXA2MJpLmS1RfGuno9rRGd6go4i3ZyurT5XYAK0bXKk65qrpiAArBi9L2OuwEKaTKEaiX7BlzzzH5QYnsSdykwFL2klPk1XM+s0fALcQr0Z9klE6odhHy7RhAYZfjJH/5WHNi/hsk0AREH54De7pJyjkMQAzEGEgK9bhufeuBJ/MbvfwadTheiIPHDLRLet0Jz9atFNK1dfqaVNENg+uUXFSKwGHbR9tjfLStroWTDq7bKK8rbFgB5KJI9BLOTyjiXACD7W0gcnsRgaZykJBNWh+fLXMXcWjXotwFmCDoUlVG6xfyeZiQoeaJdeE1dNozAgqFBl0P6bBIoYhgMRrj7lovwgXfdhuGwb8Q5F5k02x03Z4K+xrHZiHDGaXvw0U98AZv9FHGkbe/yzhfW/05/WW007XNKzTmgM83+kBXika22suZNC1kvmn78sXTMUMcgPCka90l24XZanMgAQTqofqmYLyYFwqjbepJNETdZ+JfQ+3zJBqLmtoaoySNgv1zFY/sgCAAUHGUqCnsO8qJgiZSVJwQdLPUsQk8o+ATvuedmTCcj7AyGGIxGGIwnGI6n6I+mGE4yDMcp+qMpBsMEg9EYo9EY6+tb2LfWxlvuuBw7gwEiFqG4SKnOhxnbt2SQPFlQfDkj2TMsDaeyWNkYau6yhMO8zOHnF95nZSHLscoJ1ebBrcVZgSS876ICsVqWALxxy294DNnFVclQap3+tvG2F0ZOVvN7NcBMMRwbrophPQIvaDI7EcrhhayKAsAATBu0T1TWSxIgRhiORrjkvAN4w40XYzQeoBUzxBFDHEWIowaazTZ4xkFEaMQNRDEhjhiiiNBoRMiyFO952/VY6TTBs8zZlclv6SUjqN02eJ+FbliAwmvTQNkClcgFPoX2uxdIVcMcbKPdH7pgwE3mZKYTNk6V91vOAVMotqd5R07S+kKuwFQlwU46WQxNZDSvMIcFpA+0O7rkFg3CtLIXxHpJ/qAasYSS8cYRxkEH4O5UyRsZIsJ4PMZ77rkZu9Y64BnAIgbGAMYkzGa7jRe/+gzWjx1Do9ly0GeMYTie4sqLz8Qbbr4Eo/EEUaQXf7RasNSDkdJ20lOVuW1mF24dbWXylBRwLsiwBe8C9FW7iMajJECpk6cGmNcjLQpLsYLwpuSzwcm+nl8AmEFlr2v7i0kZN/qMc8F0CtReqYm0ian/3HrKf/vvHHs+3PEkHK1sy6lpkuLUfSt451uuw3g8QRTHICIwJg92NBoNjPp9bG9sYuPYUUDIrUJGACMCYwRiESLG8C1ffytYTOBE5jCIEQTBCWvInqXyLMHkWw4lBXQXiVA+z8QmoJwRdarB3ZVjX/VO/9kCoaycCBuJiyQb30Vg6VNAgpvFYE0BddPcAkC2naqr8RlxXknhh3AyJlNJj9XR+j4+foSQ0sStMtq8zgHa1pwLpkhEcRRhMBzg7Xdeg/PO3IdpkiGKIzAWgTE5FI1GjCOHXkGz3cVkPMHm+jHEzSi3jhkhjhsYjRPcesMluOKiMzAeTdQWoYWojbvBwzfBrWlKbXOZrLHwzhoEzG5jUQfTDG3vwKMKOCVJGxegYhT2SpwsQVrAZV6B6eMkZtRfB4b8EBY8zZOkcawJf24BIBVN4RxeIJP3ey74irjyMDwVMAhFxwS1huALCRH4PjMxp8PzK5hV7Xa0ZLsu9UX66hOICFwAvXYD3/Ku25FlKeKIgUURKJLTmihuYDIe4vjhV9FuNcGIcOSVFxEp6U7EwBCBGEPKBZa6TXzDPTdjMp0ALL9/2Z2C5N+C/v9KqEnjynBMMZ9JJRaT1beFm4oKfW1jWTEQVTOxyiSMhs49LFEc/0omKbnmwzdU5p52kAOjPilaAlfkv4W/ADzLkPLSnAJAIu/f3OXg6DwreKnPrsFRsMLVugVTsQy61s6+KiqrtOpdTuAUYK3gXezKgtF7/gQgihgGgyFuu+ESXH/V+RiNJojihmRqYiBGaLZaOPbaq0gnUxAEGo0YO+ub6O9sIW5EEMQAJgUAiyOMRiPcc+e1OG3/GsbjqRQ0BgmrbyyTlZT1Q7bQJH2U2NbsAXUe7DSyOyP/WdqnswRKjapnJnI/PN6rqnoe8PMXUExr4VFspm+qBhaTnQEF8jMk83dYLQEg7G9aIVr71ORnNg/Kb0urqszQk7o7jfQLoSuYocpDFc4lhcqeKZViryOUak2XFQGCEBne/+43oNVgkumIwIikEIga4GmGo6+8inanAyEEIhaBsRhHX30VjUZLMj5jIAZEjCGZpjhj/xLuecNVGAwGiIg5OOh+I1sQ2JMY41Ja1gn+n7/4qoWjBX9hzWilhebDARgCrneiqEmL3lTBCU69EH7CzJqEiglRdnFPDtvPQMHnpNZ6nHMKM7HJUy0BkFctcivDqtExj2wicErPQsV7bjTuAlR1UoMVKF8wOMjLXl0JY4TReILLLjgDb779KgyGIzTihmJmKQharTZOHD2C8WiMZqsDwQU4F2h3utg4sYFkmiCOm0rDMik8IobpdIL3vvMmLHWbSDN3ji/gm/2htlo2u3sG20qzxpAMqHDZkDE9Y2BCU4kqvCpIaW4a8PQLzUl+oYyFvqk2XmfClLpQ9q0IZylNRqHCEwCzyudzRZEHpbShzjL7HN/8kEbRPzXTlxhKxhIor8opEsJtVmPL4JSlCkIjxjAajfBNX3c79u7qIVV7/MQYGEVyZR/A4ZdfQbvTgx3ymzUaABhOHD2KZqup4MndgChiGI+nuObSs3DHTZdiMBwiiiLI2AGUhw0Ptc0yj0lFGNJGTnWH+I22chR43Zsb1AEdmFkUdJazbcisT5YjQYC7GzNH8gWhXd3CydiBM6Syb93a1pv+1DEfNJ7k7EbNg6YjAGYXzBEq0EnplNxiYs/Jg5z3NpwyjrM7wzND/epmzRQqGyucS4lcpeVZJBrXEgIhAqbTFKfuW8N73nkrRqMx4rgBMDKapdFsYWv9BPrb22i223JLh5GyDoB2t4cTRw8jSzNEUQQGyu++Iekg9P533Q6mBQfNbp5jvQkOCOEtKC2Q/D39SkK38gnrt2DFZ847BCSNrRwsu12wBS1AcsuV0jYV8zjPfdpUgnCuU0WemW0+tSkujF/OPNt/GsfCFKCaBHLGrU0qthYI8iqhuD2lM1Y5arg41U61TMIqDtKDqLOGtFteCYsi9Pt9fP1bb8J555yCyTSV234USSYngYgRXnvpRURRJMtoE1/9xXETk9EEmyeOotFsQAevZMQQxRGGownuuvUKXHnJ2RgNJ2A+g1R2kY4qq600zOjyKntY90XI/Aq98xGsMLuqAnoS5HqGXtw0NnvJDHcRoRCCYTNykClmWE2zSDtkDZD9adUhyNk0q9vGYg8Zj7sqhJj7syo5gtpDmkKZ/Cplg/VMpxa7V6F/0qacj3uZeSEzZBnHUqeBb3/v3eBJIlf+mdr7J4a40cRgewsbx46g0WgBnMuTf0QQyjOQC4E4buL4qy8DPDEOQWDSgYgLYPfqEt7/7jciScZgEUNhC3Se5s3cq64jjH3BGNpZoMAzeO8AswoXsrAs2ihiMWOaeDLppOmoCKeWBe6Z+lQoXBsYgKCIrAqaKKtUcWiC+OVWujYH83L5PA16KlOePL7SUq0WLfs8WTDjayY/r5b6juT2hJqVWMTQ7w9xxy2X49qrzsdoOEEcRYiIydV/SF//wy+/hCxJJdMrUHJ9gAEkvQDjRgOD7W3sbJxAI45B1r8oamA4GuNdb78FZ52xH9NJAlY5DbA71moPOPID5tU0UARpjbeexhWCFQhDQUBue4Txsv6MSR9CJ4fg0pR1QMZSnKVpFl1U0VSdFFDm/ru5DBNFi8IGqg7ozCOXCgLALAoHUt7BFeG81YAI9dW+/iP3WoIy25yCuobifaAoFJidlNWpD0qUKppZMOyvfvlKK0AAgiCyFN/6nrsQMTXHFvnWX9RoYDoZ4tjhVxG31NxfD6kCpxlZzvdjrB85rK67l7sAxCJEcYRpkuKs03bhXW+7Ff3BCIxFfhOCDXM1CFMLh3U6yhbqcIdHWxEmK5nHxmRXD+xes4WDU0V5Q4q4WmOij28LYrObVIMuZiqtmYVDtF+v7kJGLW+1gDQVzIdgbUcgAtRtvUD5JqbMSWDIwxPkR2llKW++T3YUW0UiGnSwLTV7y0jV12PCZyl7ZxqRw3bPBEgGH43HuPSC0/Gm267EeGcARsp6EhxCCMRxAyeOHcd0nCKKGgD0gg6h1W4iiqPc8hEccaOF7a1tjIcDRHFk6mEgRCzCZDzCN33drVhd7iDLMtn6qs4UxR/C1tzBfDoF1m3M9FRpEdIHoxTsEEPbvgjlJOV+WgJXmIFxUdXh5PSGkvCP+c6btLDyjdo502xDNBC2rVBao2CCwTvv56H42gJAa3T53do2KuDmy3TkA1F4Y0GxQmOHW1CnWcLDIzdJ5x8r25y1flfxk1U3ixgGwyHe+45bsbbawmQ6gRAcnAtwIc/qZ+kUxw4dAotbEjQxcJGh02nivi+/gONbYzTjSApJYiAWI8uAzeNHEDe0H4GcVkURw2g0wZWXnIU33nIZdgZDZQWUz5tkl1udo+93C53FcEsWO0IAQgkF4ZhwmsF5wFnFs4P1NVc2+GBeFylbBuQyQtYt25fPAWrTQYmZPrfpXwarFEaw8R4S+VdhrCvbqquf5nIF1rHxyMTksxIVvkBe9Ckf+afmgsNpLz4FLX5/bmDdkgIhv5O1VaTm2S6gspHw1ZNFUbr8TLNNmKqn0wSn71vDN77jJowGA3kWgHOl/TlYg2HzxDHsbG0jjmPAmv8nGcfv/MkX8NATL6PdbEAIDj2VYXEHm+tbyLIMUdyUPgSqrYIBEeN4/3vuAkEJGv3PbI/BCET5503RTORgv4F2X1mSWu+CKDObzG+LASEJVTKq289yyPM1Jbu73br9MczHkey1Bw3XHNXWdTI3y6zkzywsuizRfPVglcIpgef8DLM2+YvrwTyhdwvFAyCHDIrt10yfM5AwCIRbLWCTSv6wvkFDVjZuWRPCwtV2GNFJChC9mJa/974rWHmEyXI0ACBiEQaDIb7hbTfjgrNPwWA4ARccQqQQIkPGBTjnOP7qq+Bca2BZZ6fdwmNfPYyDhzZx/6MHMRyniKyFUxbFmEwz7GxuIoobksmV4IyiCIN+H3fddhkuv+gMDIYja+DVMSal9bniIdksv1E5g5sFO8uphrEIkf4jhoiR+mPSQYnJAKZRpP6YJUCt7rSNE4L+rewIh7j8oHAuDRGT26FRpPGSuySMESICWAT5SfrkpQ0mwBgBzV8+7DP0bUBmlgMLBLn1LU5dXh8B11vlVXP/UDBD1b9xWcS9IBxT0Avd5HGuqU9YdVvXB5sbpJXKk9+9Xi8E3vQQEVZeAaX5tdCxe8tH0BJCZsFSEbvCR5ayNh4DEjZoqCkjJOMCvU4T3/L1t2M8mUgmNHH+BFgcY9TfwubGBkCRCjNIEFyA4ibufeg5sKiJZw8ew1eeP4obLjkVo4nc/gMjcDSwcfwY1vadYlyJ9ZpLmmVYW23gW77hDfiRn/xv6PW64GkGYsqKMG1U454PaqAxyiSn3HgWQmB7ewDjrWgYRFgDmw+SgDy3sNTrmHwO39lDZA6A2kLR6nkDXxeQU6P+YIA0SRExeeyXgYETqSAtucIia8w7rYbcZVHPHCE409KbI1WQcTEF7m1w4GhaVNaSeubKFKZEqLaohBH8PhsQgFjY3Doj6apJVBsOQg8UydDcRJCRbSFcPteMYWIyzjBljJrwRsifszrcWa45tCAj5zUHBDOEauITet1jFl+sd0RS+2/vDPDWO67A1Zedhf5gG4zF4JyrbT4OJgibR49gOk1AUQx9M0CrGeO143188amX0e22MBiO8amHn8UNlx+AEAlABEYAKEa/38dwZwu9lTUFW5uCDJPhCO995y34pV//GDb7UzTj2IkPIiz87Y2kQq/rSbvQx5kFWjHwtrddi3YrhgxErdxvSQt1vU0sAMERRzEGkwT33v8kkiy/56HAF2RsRyv5nW43QCqOLJ3g7W+4HAfOOAXT8Rg840gmU6RZIqdcGUeWphhPpzJ8GhdodBp47PkNbO6METF5TLswxjNI0ckY0rB1UqFOW4CWZbKiOZH+sCxYb4FegNSUPWwExOFKwtg6qFW6jZIkGghnRd9oVP936Satx7xOFgrnC1KyemEPFDkfDlyXOUR11xjhmcNm4PjAu98IRlx62SoTTQgBwYDpZIj1Y0fAOSGKAQiSVkOrifseeRqb/QQrSw202y08/JWXcehYH6fu7iDjpISAUIuBR7G0uiaZAQJCyNBi42mGM8/ch3e//Vb8p//2Z+js3ZUfFNJa0eDtjEqwbUxZV5wDnaUYP/YD34R9u5eQcS5N/ThGFMdoNmNkmQDPODhPwYVAs9nCi4eO4zMPPoFpItQSTW5l5Qf2LYHrjKlRO/YQqU9Clqb4prdehW/+pneAWIz+xgmMRn2MB30MhxOMRmNMhiOMhwMMdvogZLj+7jvxbf/kl3H0xIvodtpQg+STQc1U5KxSEvRToc4KPlR95DC3ZS3lhpH0hrQtAoljKHy8QCyF9gwi94sZryzPbBK6EfKdKL40ZdyBLqvcG3Tb7C9LjnFgqRvTWbkAqnZVF16VtqrP8ZASVplbEcNwNMYVF5+BN9x8Gfr9odT6Qt1vKAQixrCztYH+zgik9uoFCcQxYWci8Nkvfk1e+iEEmo0IG1tD3P/Yi/iOd16Hrf4QsfIOhGDY3trEZDRAo7sELlIwIbUxiwjZdIr3v/dO/NYf3IuMq9V5b3zylI+NUFaZ1ihEhCST5jUEB4kIW1vbiEkyOIsYoqiBKGJodVvIMg5kco0j40AjnmBra0fSnzUOhvmhcbJDc8t1DT0l8y0t0/ccYBFhc30L933q0zj1rLPR6zbR39rCaDTCeDTEaDhGMplisN1HyjNcd/O1aHV6mEyn6hBUvvxYOhUO6ZoiSQezz5UKsHyh50E2xrt9PVgRYNUUn4lSDRculDN/IIWQdPaVhfWXIff1D8Hxfhey2bBKkm9r+ot4NeSOyWdWlfPVbCOUVRlGDJPxGN/w1puw0ouRpCkE1Mo/zwACuODY2thAlunw4FL7dLsdPPrsYbxweAvtVkO2TAi02x187pHnMZgKRBEzPUZEmE5SbG8cR8zUUiYjQMUKGA0muPaKc3HnHVdj0B8hipgx0QsqivQ80o3fQIyQpCkO7O2h3YzABcC51DCMSYEXRZESAgyNOEYjjhA3YkQN9V0FPPX7Pj+FrCIZi3xRS0C4a1r2+FvCFyQ9TuJGBBIZnnviSWweX0fcakHfqhwxQjoZo9Vu4ta778DeU0/BZDKR/W7RfqWRp3H18SnjgypgqoODbgm2wPOtHtsi0uceQIXpOGnYThnyPvO6Kpz6Q60T+Xyj8FrLUa51h3ocgD2rgyy8Z9x4XD8Jy+owgz5bVpP53zLzteLSJ9dItncySXDqniXcc/dV6Pd3ABAUf0vzXwCT0RDbmztK++hhlub7vQ8+I5es1RshgFaniedfXsejzxzCUrcNM5knAY4Im+snkCr/fzkVkDhlkEFEv/Ub75bMSlrDuYSvvxtyU0wpYxhMcequBr7nvTeC88xoTM45BBfyfkIdlBJC+iQY92XVc874hTRagbPNK3NIyUxXfFjKmuQCjATiiOHlF17GxrF1GRxFCCTTKbpLPVx/243oLfcwGU0AThBc1if1U7UicfrG58m6rGOXQZgt3Nqo5HcOR4ADpG/P1laUr9CL/Wu/YqUvg7jrM+NlrGMhSqE8tgldUadPJ4skjYrV/qLJX2OkADhzL7IWXCzxwBhhOJABP886bTeGwxF0tFbN/ADHcGsDk0kCiiLj+99qNfDsy8fx6NOvoN1sqqgxivyFQMYJn3jgKXCKVX1yDZ8xwmg4wWBrExHFWtIAACIWY7AzxF23XYnLLzlThg8nK5KzPSXyaI0YkHEgRorvf/+t6DU4JpNEhS2Xdcg2cfMdgAluql2dNdfkW8DFPs5PD2vBluPnjleA8wSQR0GMAJIxEo6++ho21tfBOcfKrj244sYb0Wh3MB0nEABSkUEQz0FVJV8hzzA6K+H4vwtwyhbX7f6zaFG5lmtrxxylLiCeP9I7B3LcRZUFEMZfWylFi9x7YiwWP07djGSZW1XnEurA0YWNgLe1Xik+rsR0ZgyCLJdSsvbGCRkHeu0Y77nnRkwmIwBAxjO1ICbAIc8FbG/2ITKogz6SgBtxA/c+9FUMhgkYg9lFEQA45+i0m/jSEy/h4Csn5PRADQIRIcsIWyeOg/PU6jN5qCdNE6wutfDNX38HJuOJ8hpUbVF4O32sLIGIxRiPRvjeb74FF5++gvXNoQw6KvQVVFDejMKsb0iUcxfw/JSeil6ghEFwOO2x8aYJTioIBKXtjNOXBBBHDFsb2+BC4NyLLgBRhOk0hRAEnmVAxgCuYwX4GtPHTeT4nYwlaoGphFVFmsYicgFYpG4YlGDJdkWnubrK4bH6HGbfPqYBk/0jXKwwmbMh5gLFPu6ppxi2cgqgE6jL+tP4FWCIIpE5ZqUyyglmLqUJK99q02XlG8YIg8EQt95wIa685AB2+iOpGXkGLjI5/xcC00mCnZ0hKIrVoBCajRjHt8f4/JdfQLvVArfCemmzM4oYtgYTfPLzT6PdaikBoS0Pec/geNQHETOhxASX8Ac7O/j6t96M009Zk0xgtHIOw9a0URRjuz/A1991Od71xitw7EQfURwb3SQNgLzTzP6/Me4i6CVp0iHPWF6BrQwkKuVWmL5oJdf4Ao40V22RdyUY5EBExiEoyxJM0ym4EMgyjiwDuEiAWRaAoUnPQgqlKoa2aJLIdI2XQX2WsqL3wqdLIOczvw7S4iAEao6LQWQbZSgwMicsPKBeR5hFnpIeMniqfIU9wyrZFLYoi9qkTFDYP4wAs9rE8z7Vp/h8eHmwTQKB473vuFle7MmhfP45BOdIswxEHFubm0iSDFEsqVVwoNvp4KGvvILDJ3bQaMauxaFR4RytVgv3PvQMjmwM0WpE+WlBxjCZCAw2ttRuqmQSoRhnPB7jnNN34W13X4ftnR3j/CL8foJcPBwMRrj47D34Z3//nZgmKbpLSyCKwaI4Z3b1KYSA4NKBSf5xgEnvAFKEx5Sl4q56aQuHXATqqEX78JA5u6Bdf3NNJzczOLIkBU85spSDc652Jzwl4NFDrilZqXVSO4mSv0LbTPXuZyFP3peF4PxauFfh47V1hgAI2WB6z9RXrcIgZRjGR7ASoeLp8L/eZI2qssuMHDLaxcKR7DK5BcQYYTSa4vILD+C26y/GzmAIYtJE5hy5/3+WYfP4BuIoUhF/pNvsOOH41IPPgEUNA94XAlwItBoNvHpsG/c/+jw67ZbiIaYW5hi2t3eQJGNZhnO1uyAhTSdDfMs7bka33VD42M2TDjWMCGkm0KQM/+aD78dKNwaLGuj0ejIUuRIuelFWQDG90v5ccHCeqWPLwmJKqyX2lomwns8cKsfuNKX0lEOvyuj4CcQIUOHWsyxFmkzB0wxpliLNMvAsgwA3vOgMtZEetkIK4WQXwkwyr0XaVdvxuaaEplXzivLX9Yx6jfBMC8DjArK0f4FM3bz5kxnYmDlAPmd0kVww2UUdVH1tBOtEVbhmbQ3mBYShZyKGZDrFN7zlBiwttZEk3DA9BEfGOeKIYbC9jdFoBMakl6EQQKfTxhPPH8XTB4+h05Z7/z7qGn8uBOKoib984GlMMrm9pQN/xlEsnV4G2xAiM+QhhACDQL/fx7WXn4k7brgIw/4QccRc4CCwuIHRoI+f/NHvwO03XoJJkqLba6HdbqLRiGAW2wQM02tS5CIz7RVkhTtHjmPRx0XADbxYlqwDXt6YGAZmUGf+9VRAWliMMTkNy1KkWYIszaQFkGkPUPsglOJ72MFuKvCalzR9dilYFdqa9E0Fe6qSd6JzQY15Z/kDaG1WQMM1Q+Y6DGQEVKF/NAPrV2p1uLbNpIjLuJ7KyK+FrZJ5TLBQWRv5vOeseqgwrrkQy6PZkJYIJDBNUhzYt4y3vvEqDAcDRXRCmZvSKy6OCevH18EEoD2yCEAjbuCzX3xeush6FobfDL0Y+PTXjuDJrx2VHmwKY+kSy7GztQPNmUJNAwQ4OAeiCPiWd90mnYKYmoKoNjXiCJubW/jub34j/v63vw3DUYKlpSVEUYRWI0aDMWRZZlbzTTOkeSLHWTFaLj5JMRWBKJILqMEBrLKvtbkRtJ1NMYJ2kc7rlhYBIUu50vqp/EvlZ+FmSyFHmIzfinDeFVCehxZDzfKa6AIOtde3V4RpqTLO1KDkaybSBwVeF7oKfIHTgBpri5G0Samlly0Yai0yaliUwyDtTx5oe91k96GmSSfenWLvWVFUnLrzvIxFGA6GuOv2y3HKviWMx5O8WiHAswxxzJCMp+hv7SCOIyOZW80IrxzfxucffwHtdls5phST3WRGwGQq8Mn7nza3BmvtyhjDYKePbKqmASr2gBAAMUK/P8Ldt12Oyy8+A+NRImMFCHld+fb2EDdcdgY++H3vxsaxdbQ6XTQ7HXR6Xays9QBkZuGRaVd0tdYg1zkEMi7dfyPoRTnLYjRn8WvZwYZQyXNasjV2rtzk1NFMB4jlUw9GyLg8C5ClKdIsRWamAKoqx1K0RVS5wKllnFaRlMvH5dZq5WMy42B8VoNKxD4Rq8W3TmweAWBhaSQJmc5w5u9muROzO6qQ9LVVIi+76GzAUkZlQKQ3lZZTYSZ08cjbnGUcvXaEd951DSbjEYSAmWMLQUizDO1OGyeOrUvHGUjTWIgM7XYT9z1yEMc3R2jE+jhQeSMJcmGx3e3ggUeex6Fj22i1Ysj+kubudJJKK4Qg6xNyOgIIJGmGXStdvOedt2A0GqkmMEwmKZbbAj/yD96GbDJAfzCEANDpLWFppYcsSyVTMXWoBJCMDylk5KJnvvAJptvCnNXouSKOa5qyyvsTT235wNnN0P4HKqoyiyAyDp5l4DwDTyXzSwFQtuI0Q73XYNZCPv8xWa9r9kuezapUL3oK5PxWmXwbrPYuQIhx3I7StwYV8s4cec/M02BOxsyyyzqmT8ggkYxpHCTKYHl/jAj94QjXXX4WLj3vFAwGIwByJVyvNkdxBCaEdE9txCYoBmMM26MU9z70LBrNJoQIzIVLlE8zjnB0Y4B7H3wWnXYn1/LKCtje2oFAqtyPpYeeUJZCvz/E1911NfbtUb7wjMCQ4fu//U6stjIcP3oC0/EIPOXob/exvbUDoga6vS5AyrgUKpYA52ahU0k+8FT6IuTyVntK5uqukhz0wFsWhu6HQvhQvQ5DOl6iWnNg0spjKs6D4BlEmkEkHEgzIEkhEgHinit3EBcU6ch+toBFGpzN2N9n9o//SPGPDvAQBFBkJr0WUFMAaI3MjZUh8pGxEPfPXsPq4aJp4nx6eFvTufmTXU5bIxSqR3g4VsAiGyeldUSKr3vTtQClyLJUOf1Iz78kSdHr9bC1sYMk46AohiDp6d7rtvDlr76G5185gVarIcfOS2WMwjlHs9nEJ+9/EjujVJ4mVK6+LGLy9Nt4bHYijLsugNF4gnNOW8M77rxaTlcE4cApa7ju0gM4fGQdo+EA2XSMI6+9imNHj2DvKaeh2euh3esilitt0PN96eEICOLypkSCikPIwFhktgHlgiAzB3sqZ1u+xefE+A9tGqptaR0wRfsfEAOY9M/gGQdPMzU+cjEw5RMISisQcXEp0OIiNOm0yxvfuvAK+ZRSs/53v1WA0j4sNat2MQghb9S2/MwPGZVxMlnlUBBSokxuzEohjV2W6szjAkqZEcN4nOCic/fj1mvOxWBnqEz0TM6FswxEAq1GE8eOHJVhvEUe74VThE/e/wwyHpXKuCBqSoO0W00888JhPPj4QXQ78h5BKPOXc2Bnu484YuBZ7oashcAkSfHN77wNS50mWBzj0NEtPP7sa1hqt9Hf3sb2xgZee+EgzjzzDLQ7bbTbbbTabRjJz5na8hNAxs1ioOAiD0QKa/WfAYLKPEEqWhzUdqKQU5BQZ77Im+kxCCb7Q4+JdATiEBkAEc3ESCMt7N/zpJCmL9P8deEZOBKY1EX2rRnzAWX1hE9Opu43F7d8mJUkNhrWFgK+qveq8KXsvNLWqYrKOaku6EAmYvJSznfceTVWui2kUxloAhkHOEeSTNDrdtHf2cFkNEYjjlV8b4Z2u4mvHdrAI0+9hG63KRfRHKzq4SIQ4c8/+zgEReaSURDQiGP0t3fkaj8JsyMhOAcRw2CU4KrLz8HN112A0WgIQTE++eBXQREw7PeRTodY272M6XQIiAztXhedXtdsX8ppDgB9IMiy1HkmFwsFs60tZ7NtRqfXzWf1g1mXkO1n2lGNSW9AuQsiBWEmuBReetsSM/rcJ9dFadH5PX8b8/xFK1afzjWh38pIPoQbAFZfXqjBN4q9qqSeMugPjWTtyl6HlE9GipjW7KVAHkZAkqY4be8y3nzLZRgOZNw9kSlPuEz+dbptHHvtCFjEpMuJany73canH3oW28MEUeTeXVcUqcWfguSKe6/bwZefeBnPv7KOTqdtphHEBNJpgn6/j1arKbWycg+GkFMIBo5vfMetQJah027ikScP4auvbODMs/ah1W4jzYBBv4/JaISVlWWMphzTaYKIIoOM/seFXmfgSLM0n+NSbgVA+wE46tRPs22EUDK+BnIhwAytXBMhFYtRIIMMksqF7D/hBdcIojMLrdr0bHwiA4VmbZaXaEKbmQJnA+omVr8VFvKlQkyZJfZP23WzEnZZmp8ocpgaH5vLcsFUA7FC9cQYhqMR7rrlUpy+fwWTJAXMNphAmibotFuYjkbY2dyU4cBk0D/EjQgb/Snue+Qg2u2W2hlwK8m/hS0lPQpRxLAzmuIvH3gK7U4HAJer/0IgiiNVt7w7gMvJujTOCdjuD3DHjRfhgnP2Yjwaoz9O8fknX0NveRUvvHQU02mCyWiMyWSIrZ0J/sN/+QgSTlDglNaXql+6AysX2yQB1LagcaglII9QU9bZJ6MZ9KUfzHFTl22V16ZlPHPcleWWa835X9kcTWep/VwpQfiCp2pztGTaHMiWW+Dl2YpIElh9BvNFIZmnfr48zpwKVeVjG4RdUbMtUWhekWB3uHCaIcqum7KTLexIzieXOzHe8cYrMZlMAJByX5AMwdMUvZUlHD96XLU8A1OmeLfbwcNPvIhDx7fRajZVRF84U5WZZKlMXSE42p0O7n3gKRzb2EGrwQDBlXCIMJ2kGI/GiFsNufcNbrRhkiRY6TC88+5rMBgMsLzUw31feB5fevxlCGLoDycY9PuIGOFH/+1/x0OPyrUGzrlzloDDcgfmHJmaAkDNS+UhIBX3QKj+LzRwceZXMwzl1CNyh0HlEyAI0vmHS58MYZ3O1OxSrN0KKz8Dt1LSITeXA4pZ0+KaW3czajNV1vOzcJGcYxcAOSeWmkTa/solpxB6eLxiFe1xrxS0V46pSENV/ULWn/7tqP+C9VSGkQxRxQjD0RjXX34WLjxrL4ajEUAELuR9R1mWodFoAEJga30DjUbTlI9IIM2ATz3wDBiLIcBz99gaONhdILuB0Gw28NKRTXz6oafR6bRkLFgmHWNYHGF7axuNRix3AjiMBiRI9+B33HU1Ttm3AgGB41sTfOmZQ+h22tja3Eav28Vvf+Tz+M0/+AzWVpeQpqlacYchXMnr0vznyufArBOQDvsFGCGg+7u0dZYGC45deZ8wa6phrkgnGObXF7LIrUvh6oVC5xfZonCrkebfmTxHHvuQUos1Jb4Tpr6IU94HdZi/VNzVS8aYzj1nvAwCxtfeOmguhD378c2UEBz1tI6qn9FuyeCWz3Rw7SIHEjLG9W01QhDAU7zt9svlVlgGsyUGQfIm4NUlbJzYAM+4cUzhXKDdbuArz72Kx545hHarKc3kkEIsa6YmOv0dAkLIKD1/fO8jGE04GCMIxgDG5P0BkwmyNEPcjMF5qvbw5dHk8TTFWaet4a1vuBLDwRDtbhcPPHYQk2mKVkR47pV1/Pyv/QWWl5fAudTshr2tnQWhzj1wkSkLQPoHCJ5rWDIm6iwd5b61XEOCtCDtC+kIZXYCKI/TwJi1M6POZXCFszsTtOs1ixhuZbbxWGd9wIeo6VC5yDsu/6WJvO+WJas+bDbSO2+ly3OBOGtzCAALY5HfYuPi5g9vXsaSBz5WxSeOdUT1ucQDa4bRwCNLlAuT0QRFFR4AByfCaDzBBWfuxXWXnYn+YCBbKAS0x10UMTQaDWysbyJq6P19yQgsbuAT9z2F0ZQj0j71xKzGhv/I+jTn3lQbMg70Ok088cyrePTpV7G01JPCVl2KAQCDnQGazQayNFWxCVR8QgDj6RTvefuN6PVaiOMYz728iWcOHkZ3uYf/41f/HIMJodHQAUhyqteaVI6rdAwSQk6PzBkEJXTJOAMBxeVYn/HycbN/lyk3ARmATjoBKXCM5ZYASDF8ZiwUzjN1Y1JZCnAmWcatx4PzJSmyZuexPwXCOAnAHKnW1g/ZBkcxBQIRzncYyEaMKOfoSklWU0wGiki6E3MVtx2U8vP85j8rp9XBFv8bWaAlKqnwW0SYTiZ40y0XY6nNkCSZWvwSAJfmf295CYPhCGmSgSKmr0VAp9nAS6+ewP1fegbNmJBMJ0CWAHwCkSVAlkDwxHw3v/WfmELwBJzn+USSIJuOwdUptz/5y0fAorZ0uiG5tMMYw3A4AJCCRQBP1TadEiSDUYJLLzyAG686F/3hAGlG+PQXX8Kv/OFDeOyrR9DtttTNRXlf6fk+lKchF6S2BYXcCRECyHQ8Ai5pk+VLTcLpf58gLQ2mpxtV5CNgbArmnEDML04VXGl+7aSltkVhjPB6KrzqPFItAADsmBJ6F6W6jO+abgkHdajH9gDQxcLyUis923wx9wLUS3oVOS9vzWNyhWrDd0qXqnK7vPfY/1YC3OQrGJmkfdjVc33giHJm19WTtzVEljidZhyn7O7hzuvOx3A4UfJPDlCm8rc7LRx+7YiMoKPkIzEgg0CnHeN//yfvQRxHkN5yOmhGZMwUecuPrFcQSWWm2gCCIXAOSB98cHklV9RA3GhgMJpI/3fk04QsSzEcDtBqRugnU3lrjupSkcnjwl//luvwVw8+g163iy8+9SoyLrC81JMnAE03kiFZoQ4BcaGxYYpJlDAlIV1TnWmetmTs0bM1GhUezUr6LIDeagTpMVOuwBGBC7nwlysTAU5csk6oLp+WhVXXQsyvk3AIzgi3UL4qXrHwysMrWLDnwofmEwAavTyUgvUw9N1JFQ2qRNzeOAmYjA78XKYbprYFFUR+sYp6GcLKhSIj5fT7A7zz1itw+v4VbO+M5Mk+ZBK/jKPTa2M8HmMymoJFyvWVyblexgVWeh3sWVuSwogxYxozIhk6CwCsM/4EqFV0yj8VcmbTgKAsjQaIGCbTRGp30vNMgSiKMRlN0VzqgBGZmAOANM/7O33ccvX5OP/MffjaoRNot1poKGb2ZTpUJF3tCSi7UBIeF3LFXUYF0sylF/9C3BMY9AUYTAjh+Bsok0P1JYBM+T4Q1ElsAZBct8jPf/gEHBBG8zJXToFFGDNhaVoOSCJV3uFA33ytxMlNcwoAsllpzrS4+PSPMOZnnm1s9ODrWHPCktpKfNhi1zcU/Id6pZuATAh02zHecsslmEwTAMqM1A4uXKDZjLGzuQN5jFnvTUsEIhA4FxiPUhCTZgFTAkAf4slXsJkqyhT/qDm9ECrYhqZzKTiQEkBT5EE4cgGmow4JApIkA4sIaSL7U6gFqTTJsLbSxD1vuBI/8xsfN9t9eb8o/wFbc0Et/qk5Doc8FZgp12Nlokh7QYQJr5DqkEeBr9RpOJafANSajdT5A731Kdc+oK5hi8AEU5GXPUPcXCRTgrf/uLJ5xV3/0nYXMvnSws3g/2Ig498/uzL9c4Hbgb0Zx0maRXUq1PMdXad9gwwK380VF5TzoFwV9qSjZ8QUkgIZkYyPf91FB3Dh2XsxGktmMw4wWYYoIvAsxXg0REQMzNayTGkjyPUpRoRIPgJTl3jIT3nDLlOHelgsL92I4hgxI/k7lp8sZmBRBBZFoIgQUQRmvPRy4agZnShCkhEYOEiH8uYyTBmIoz8Y4S13XIb9e5eRpDrwh91RyJmfyJkPC7UeoE9AymjBXC0MCqN1CbO7vDT5VrHhUelHwZgWpua0hbSaRAaRZVIxciEXADPlCkwccgLkawJNbzXxCn23NHYtSLUy2b1nwRb5QnEdtHMI8tucAkAxl41Lzb6aO2nNbfO+8yP0XXdMCZn52an4ylQv8jfEM7zltktgvOH1f8qttN1uYDqSV02xSDJnFMeI4wZY1ACLG2CNhlwbiCSzyoi52lRVjitCR3eRFoPIOESaIct0RFsBngnIIMMyv74ZhkEG4mAk1xfiOEbUbKHZ6qLZ6iJudRE3OyDICzO5DFgICGA8SXDGqWt4862XYTKeIGJ5fPnQREuYo8BKwGSQx4G1BaBOIUKtU+g9emOVzZu8sbJfMCI0YqamTdKyIsgrygkJRMbV/AQm2o8wVqEI04qcj4Ur9vUOBZ6HH8hkeEdbrNbzmamsMmX1lFN+AIIUG3FNA80CEJrPubgEcawBWaoUb+5jJvL2d12RO0Chbw5OFXgJ+73Q1gNhPJni/ANruO7iAxgOx4pRcwkaM4Zhfwcb69vmyC8DASruv55aEElrgki68cYxQ9SIlFBoIIojxA2mBEeMuKEERxQpE1fNcwXkGXelxdM0AU8TZNME6ThFkkyQJCnSTICLSPrBZwIZl34D+07ZA2Ix5G0+BMhAvsgmCd5x51X4k796TF7/ZdYAhJkO5dOafI1AcMhFNR0VmAtwAphyBOJZhkbM0IgYhkKHDKtOjgXuz/QA6EVczuUOS6/TUguuSnsTQ0TSNVlOS7jF/HIBMw+1R+Z/m8QKh4+DeMxsipus6Wjp+9LXZTyn4j2o6dY8KAlAXg46XxGZKqY93o+6Rp9rahris03/QnBI7dOttZXXAfbSgIVPJUbWS6YCfr7xxgvQ7UTY6Sfmjj0BGQ47ioAkEWi024jihrUnTWoODgBS22svNSMMGIEiOX/NBINI5C3BaZKBphmIJWpnIBcABMr5kcsTbnIeHoHiCHHcRNSS+EmHN6nl5XZlCp4mBidzpJSkh+Nl552C6y49Aw98+SUsddtyXx/k9B1BeRVmAiAOTgwZkx2XZQm44GAUyYtDhTwi3Om2sNRrY2NnDKDGMVxdnf1pfxcAiIHzDL1uG0udlrwinQEik/3MABkEVHCk6hZj46UIdTOQRWs5TyttaglAgjcroPko2xRzLFndFEvj1JmT5gXzx4oejcdV7SQQz9sUAR3VpaSMI8UWmPE5Thq2VYDcEjBJL+nLvEQc7m3A+acUFaLQj1UpSVPsX+3gjmvOxWg4kUV5ri30lpNAhHanoRbuIrV2p7enWL43rwSAbfkZt1ABZCBIHx15Zp20u7Bxc1XBNRRcvTFnruKKVF3WLoP2Y4CQaw7jQR/JaIRGq6d84mH2xmPieNvtl+O+L33NGQHTkQLS7Vl5+wEckZDCgBhDlgrwJAU1mfEWTNIMy90u9u9bw8FDJ9CiZn6EN5R8I7CIiOGkNE2xb9cylpZaSEZjkCDF2FI4p5m0lnTgUu3QZN/ulCtlLei0n0Te5gK28/KZ3TxtUJFuqCPVAg0PAUGAjksu2p2RytcAgrDIBGosNTasxwvF+Xe2c/TYiPxZsEL1XJhhQ75XJr/nfun1UsQYxuMpbrnmHJy6ZwnTqTrppl1dIcCYQCZgnYzjyjVY/ebSC03+ycUyCCEvpdFmG3EjLDRzEyPEjFmLgUxOEWJ5Ey+pG3dJLxzqAJyWuWrqy2QsPJGlyNIEjVYLKc/k5RiAWbQjkrcb3Xz1OTjnwC6MJ9PCfJ1IudZm0slH+9cLwdXR2wzTdArJZPJ5lmXotBq44NzT5E6ENZsIj783rCXDTgSk0wQXnH0qOq0mMqGvIyNApOCZgEBsXJX1GOR3GeR2DczCjnzihLvT5IgKEgylAP8IWEq6cGHOHMAdi1ZBNmsbJZUH8SMdHTkgeUpwya/LmsXetkvODIS0GBbWwgtBNsrgB7id5FKH3Z3kDSj0nN3u70rM5bH+Totw9/XnI5lOAYJZ+OKkDr9AgAsZfDLX0lJzc+U7z4VAJgQyIQ+kSH/0TLnTApyTuUUoE6k6VizN1kxAzt8FwDOBVEW0yVLJhFxp7pSrPFxHwIGJhJPPb+UiWRS1ELeWMB6PZNeoa8uEEJgmGdZ6Lbzl1sswnUzV9qTb5ZMkw2SSyGFTW2s64CnnhOl4LDWwOnOgxAyuueI8APLeANNXZWZtjSQ1qcB1l58tpzUaLiOk0wmIRdCX1AjVF5kKBMLVFKCUoUXASYjMjKle8uGastbktsCwc1jiwjIE9PfAFGMWfjFp5qglfUieZqwQVkYykdKMdcwaQm66O899vKxJvV2f2uLL5WGuDQv/WxpIFFalFfZEGI4muOnSU3HembsxGo/VxZpMopkBjHEAkZp7WQJRACwmLHc61v40zBRA3luhpwNa40tZLJh9qk0t+oGMUwuZuqSbMdOCkgjamGNEpiuzdIrBcOwIKAGg213CaHsbadxQi4lSgxABo9EEb77lYvzRJx7BOOVq4VGdZyCGyXSCndEEp6KnxperHQnletwfYG3PXmRCbjkSizAeT3DL9Rdj7+oSJmkmLRbbihQiH1bhfQZohQBMJlOcccoarr3sTIzHE6NZRZpCZCni7jKE0K7PXFlFUOc+YjC1e5JvUNpMyR2aFEWyWywZ7a+mKaTqMrXX5UPdD/l0xhYh80wEYjvzrOpd04hsgeYgZr6a8bXOWNti1GJ67YbrLtExEHiJraELWpM0yzpxW2W1TiOtpghy0CnPoQ0QkeHuGy8AU15vkXK3NMcfOIcQDWu2Ihe3ogg4sTnGxz79JBrNdm6uEQP0EVlQ7sxjukC+F2YRESrAhd085VhClu40AiAvAxKYThPsWe3h7XddoS4KVS7AQiBqxIiaTYyHQ3S6HXldtmrYeDLBmaes4pbrzsOffuYpLPc6crqg1mAnkwwnNoe45Oy9mAoVaCOSU59Gs4mdnS3kV6LL02aD0QQXn3sa7rj5Unz0E1/CrrVldV+i7PRggMzAkGsBG8cxNsd93HPnTdi3ZwmbmztSMTGG4XCEZrsHYpGyxFToMkVZnAMiSs0BMMc9RNjqIoBTINVm2UIbPbqcVVFJ3bkhXOLaXIqMzBznK+l1qs8Nf/fIobYhLDGpJahTXV5Ia265AKZO5CH33jPBRIwU0YDymrSln/vwa2rSXBWK868x0QIhP5Shj5OOJynOPW0VV154GoajicRDn3RjkewFYiAmL8zU5bKMo9Pr4qOf/iJ+9ffvw+rainXhR2GPwsEnx996Y9OGXj0yApTy/20zRkgtxzOBCBnOO/uf4IqLD2A0lnEAuJqCdJeWcOzQK2i2m9AHU7QQSJIEb7v1MnzqgWeUiS8BMxAyARw+tg0WkWJ0SXhplqLZamNnax1JMkUUxdDUKbcip/je77wHf3Hvoy6hmoEMG4F2PgG5NjNJUpx7xm68/123YzScICIGwQiT0QBZlqDZ24MkmaKhrzRX5xKMihHaQcruf7/Pq5DJB4Zy4l4szbNy70mbXG2pX9ogRIVV78499C5AbWwtWeUyukbGrHIK+VxobQd31I1VRbb4EAYmKVh5+zRzWNpa2f1Ce0NZZqTxETdWhStxc8Ghxzsf0Ol0ituvuxy9NkO/L11oTYwTIcCzFHGnZ8QQCQYuBBpxjPWdKe575Hns37cHTXXbr9PDZWPi5LS4oUrFVIxxxBg2Nrbw55/5Cq6/+kIMhlvSwYdkgIx2q4kojjEaDtHqdGSEXwVwOBzjsnP34YrzT8UjT7+GdrsJNc0HYwwvHl6XpwDBwQUDcY4MQKPZQJZx9De3sGv/fuOrwOIIO9sDvOHGi/G93/l2/NyvfhSnnLIP0+nUdInP/IWmCSCKCAIMWTLEB//h+7B/3yq21jeAiACeob+xjrV9p6pxZCCSc/4MHEwwRZLyTgBwKRxyMrb6HnPw9MkwP7RAr5+9WDXlmtATCXVg1PYE1BobkCaoCJQUllmt+N9B1AhWY7ZoeIoZLUtIz60tk8CvLbcAdC1C1y2QTyl0Xbq+XFLmwiW3VNJUYM9KCzddcSbGo7EZIw4VWFIFAGk04jycFwOEkHf3PfT4i3j1WB/NZiwXnjJ5HNX8md+Z+c6Vl58OWy0dd6y8WeCPF+GmmVo8zDjSNEWn28W9DzyJI8d20IwjsxDJIffHl9bWMB4MAaEi5nLlxsszEGW466YLZSQgs34goxC9+OoGBqNUHi7i+WWoHAKtdgcnjh2TwVC1s5CQ0ZS2tjfxoz/4Xtxz1zV47bWjiOPYxC1wxlQLdvVUevvFyDJgNOjjg9/3Hrzp9iuwszMEixtoRA1sHT+OdqeDVrcHwTNL+OfKQbuDc8YBxnOzWdh86C9ve5JJj/l81nqgPMHmgflT3kfCUqI5vnVYe46AIMKMiELYuc1Cs7jIfxqp5GpyzWpk/SrUJcp+2MmCoNptgiOAIARzJHxAIjlwAKndJpMJbrjsDJy6q4PpNIVeIVf7fZhOp2BxwwIqCYwxhmkKfPy+J6RHoO4Tj3iEbtOs8GqzkmXQCG8oSFXRajbw8uFNfOqBJ9DpNsGzRK3Oy+lKb3kJFEVS0IGUi7BcDxj2R7j+8gM487Q1TJMMkQLabDRw6Ng2Xnh1Xd09kMndDCGQJim6yyvYXN/AYHsgw3Kr8xIQ8io1yib4tZ/5AXzgPXfgxIkTGA7lEeYo0n/SjTci5c7LCNM0w/rWDnothn//o9+O73jPHdjZ2gGLIjRbbfR3dpBME+w57QA4T1Wf68i/+ShrD05SbsGWIW+MrZwk8rLCImvd3/MobmfIBAxtSDqdV5LoCVvOj/JQk6X9bT61U6Cquc8CCC01fcqGNMVtQaTLmMVDIwgqes9Y8jwXKKXiNneiMSZ9YWRyrR+EQPlnJgRaTcKdN5yPLJHmqYkmCyEPz/AMTRXRVy7SycW1breNp144iq8895qMzxe47sfII7vv5pqBuVaSEfjWVMYGxzlHHMX46J8/iOFgIncNhMjjIxBhZdcahiq6EeeZ8Z1P0gxrnSbuvOF8TKfqMlHIQ0yDUYIHHzuIiElvPJgTdwLEGuh0l/DSc88jiiJk+j3PwECYJhzNKMN//T9/AL/2sz+I6644C5PxEBsbW9jcGmCnP8LOYIyt7QHWN7awtbmNThN43ztvwG/9wg/iG95yLba3t0FMoNFsYjIYYWt9E2ecdz4EVMwEklaqubZMa0hjAQqlNfO9Im0vCnuqaVmo5tNS3vNZAWbwYdbKhD31qJsoN4y1sjP0lC8ya3ydrcEA2813HDj3YihBLX9NFjq2Gp610GP/EAZQWQk1bHpxTFg4kLUwWWM7Uob8SnDthftx4Rm7MR4NwEBmEU/EDEkyRbvVRsRis4ikvfEajRbufeirmKaEVpuCflmLWo0OhJB8s6S/LwCarQYeeepFPPT4Qdx27TnY6Y/UhRlAyjmWVlexfvQ4JuOJMemlE5HAZDLBG687D3/6uWcwTaV9xQVHI45x/yNfw7vvugLtWC2oKaWQJlMsr+3GkUOv4NALB3HgnLORTBN1sYh0BE5TIB1s431fdwu+7u6r8KWvHMQXHz+Igy8dx/rWNtJpgmYzxv7dy7j4vFNx7eVn46zT92I6nmJraweMEaJGG4OtHbz28ks456ILweII6WQMs+3KtPMPkN86DWnRQa3pCDc8uDAuiLlBrS020orGLjGXFeD5xWicFrIAFU3b6ludAyDlVCa4S+9kFbPT3AFBtAkT9qh3wz76Z5OE4f6SVtu9a/UPkbAGx/5UC4dq5VFY0iUfJGFghKq1u4hEhjdedx4iskOeqPPw6vRcu92R7sAEaU6Dod1s4LXjfXz+0efR6TSlJi2ryXevMwJVS27hvKLSvMUUekoQSFLgw3/xBdx+/QVy5PS6CudgcRNLq6vYXj+B3vKSvEhT9dt0muLM/Su48fIz8MmHvoZep4ks42g2Y7x0ZAuffeRrePcbLsZgLC0EaSlxpEmKvftPwYvPfhXtVhv7Tj8NSSIvI4VaiGTEsLW1A4DjpqvPw+03XIosI6RqGiJPGSVIkwmGwyn6O/LWYsaAZquF9aMn8NrLr+DcSy5Go9WUzlqRPmcggAz59p/IIO8PpPwCEWt8LdPU/CTzaU3XSnu+clgcuAIkYxLYa1QWNrPK5z+0BSHpX1sBJgCvPf3Rs03VLmHR0QJBQbVWtdpb/BIomzdCrglUCIEcdQ90Lh3MGf/cHipOAWzr36vOQFKCdDxNcM4pq7jmolOk4485gAMQI6RJikajgbjZNHNLJhgEz9BuNfDph5/B0fUBGnFU0jTfQLefaw9IUXhTRNh6Okt7CHkIqNfr4r4vPIuvvnAMnXYb3GgLQppOsbJrFWmWIU1SSIaR9osAIU0T3Hn9eWhE+oSv1Pitdhsf/fTjWO8niOMImdBHgWVkIEBg76mn4enHHsfhQ4fQ6nYAyMNO+uryOI5ALMJgmGJ9cweb21vY2dnGzk4fW9t9bG4Nsb0zUcFMGBqNGBEjvPTsszh+5DVccs1VaHXbECJDHEVqa5YB6rg1V27Ymt3z6VeJD0vo2UIauhyubRHPOJZmJW35SccyR2GQTVm5G3NBsKhpkb3YqSDWTQTorRR1042BVdYOY556UhZyO6cyFQSBZHiz9Vdgplye54tshDygSGluufU3nuCWq85ErxkhzSxzUU170jRBt9u1S0GQ3G7b6k/xic99BY24AZ5xK09Ju+x+MYuBBHsoS/vD/6xKim4aMcNmf4KPf/YraHc76jovGR6MpymiiLC0sozB9ja0bOZqpWs4HOOis3bjwjN3YTgcq1V7QqfVwOHjQ3z4019Bt9ORx25V0A3B5Y28UdzEqWecia899TSe+fJXIDKBZrMpzw2Y6MSEKCLEUSSDoxCBEWRgFRah0Wii1WnLLdYjR/D0I48giiJcft21YHGEiASGY47+KEEUK89LRCCK5FiYOwB8wTlj9j0P45flDVagdzoCrtClwK2pihYE+okzp7apuhIJk3MuC0DqBM9cqkpa9YYcq+eZENvLsCTFRzH0gQLIpEoXZOXRdRWsAClMpinH7uUWbrrsgJoLw3S0dCVlWF7bjUarY669Isj5cLfTxsNfeQlffekYms3Y+C3Wph/dRwUTb4Ekit8lnkC318Mn73scJzZGUlBxLqcwJE897tq9G6u7diPjcjtS7pPLbdGYZ7jtqrORpYlxP+acsLS8hD/97JN4+IlXsNxtI0mmCq4kcJEJRFETZ5x7Pob9AR554AE899RTGA52EMcRmu0m2p0Wms0W4oYMbNpoNtFqt9HudNBsdZBlAocPvYonH/0ytk6s4+KrrsAFV16GJE1khKOojX/zC7+H7cHYWF/StVoeREqzTJ2OtIVr7vRUOVJhHRPOFxqysimnUkrB9eog8LIb/PQUByZIjxPkzNbQwgWvbZA51wAkBxIgw17VonIy8x1pugNuWK8aQMj7rETRnTY4U3+rvN74YYxhPBjjDTedg1N3dzAcjBBFEQQEOMmV5k6vh6gRAxkHI3n9FhhAGcBB+Pj9TwIUy8p4wXgp8rf90khwHaGCQR4XnCNZFqFTr3nP0W7GOPjycXz2oWfwrjddhY2NDUSAvDJMRf9d27MHk+kUO5vbmIxGCgzHaDzCDZccwOn7VrDeT9BsSAcncA5BTfzi7/wVfuIf3oMz9/cwnqRgTeVnT5E8j58Ce085FWk6xdb6CRx/7RDanQ6W1tbQW15Fq92GDuGdpBkm4wmGgxGGgz44F1heWcGlV1+LldUuJskU4+EYDALd5RX8yE//Nj734NP4ke/7RkDdegSRAVmKsy84HxwxXjn4MhqNJmRgEH8huoIGta6bh/6CA2MPiyjPUoBvvRQw6xFyOYxc9MktVjzjEuLXRaIC6zl3aDHL0bQW91l5BSC3oUI+lEURVYQ9F7Ie4s65AflFCIGldoS7rjsXnKdytVrImHLtXg/NVguA3COXc0hlNKnbfp596Ti+/PQr6HbaMpIOKeR1Veo/oRFy4hUYsY18b6mGQDRTiBIVY4hC/hAkrRWKYvzxJ76At91+MaQfnzzdSFwyX8pTMGJYWV1FP2IYbG8BnCMVHLuXO7j9mrPxh/c+LT0D1f5/HBGOb03xk//lz/Hj/+DtOO/AKobjBIzaMKF4OTCdTMAaDPtOPx0EIJmMMRqPsbP9CpClEABY1ESj2UKz3cXyyjLOOPssLK2uotGIMZ2MMRmPkExTGWGo2cKP/tv/N/4/H74fZ5y2V1kEMZg6ibm6toKVlVWcf+mlSFPg6CuHZIBUnoEzvcNUg6jsbva/+30eLBwaG3s6jBK695WmKmczPyAjNRMgbHojKB8YSxkGgyzMbQGg3HHJ1z5GCHg9Jez5fx5t1sgEu4c9zTaXae0XsBFXoBkjTKcprr3oFFx41h5MRgOACHHcQKfTAcUxhIpwoxmUQZ5bEJlAu9XCXz7wDAbjDCsrpPb+lXHlbJnm7c1vSxbe0sQcLfMFRYiAyPohAME5Oq0mvvCVg/jSVw7i+ktPx85gItHh2iBUV2era8MjtoLBdh/ZdIpkmuCOa8/BvV94AZNMCRYhkKYcrVYDRzfG+Je/9Kf4X//Om3DTlWdjMJwiyzJ5D6JaT+CZQDJNEEUR2r0VLO/ag0YzliHSInlAKY7biOMmmDpzkaaJ9PHnApwDy70eBqMp/vGP/xr+6ONfxK5da5hMpnLNgrRVR2i02kg5IUs5Tj3jNEzHQ2wdO25dIDonNZHFeyWzz8pkCQmzEk+zaNoCrMx722VZR4fy3eDzcc8Pl0nvWLKmPjLNeTMQQR8ftqcX5lPANU0UEppgHX9rzQu6MzRXCqb+UOidub2vfNycd2pDRqS449pz5FwZEZqtFtodeYpPB8+Uh15yKSsEodFo4sjGCJ/90lfR7rRlSCrVcGFV7viYmYXJMoE8p5kT0kb2c7e5AATG0wwf+9SjADEZMku5HaeZdk9OkWYpkmQKAaCz1EOz08FwmuH0Pcu49uLTMBpPlCKT/gJZxtHutLAzFvipX/kE/seffRnNZhu9XhNZliJJUrM2INfkGLIsQzJNMB1NMB5PMRlL03407GM02MFk2EcynSq3aKARRdi/by+++somvu0Hfh4f+eSj2LNnl5zncxnzEEJgZ3MLacqtmAgpkskYS8tdtHtdpFxGA56rp7UFXVSg88HQoFzLvgQOK/raEaRFpfJzQ1uqY5XVbc6+kKbGMmtnrikAmQpyQhUuk5rW2WXy9+4JJdedEQDksWFhtKNjbYi8bqV/cmjO1ANOL2vrWmosnVf6po+nCc7av4xrLzoVxIDVtRXEZgMYciuQMUQkrQWK9PYSx9rqMv70T76Aw8f7WFlZkqvgVrvz6wZ95PR3QoEK5/ELr8paQt2cCyz1OvjcQ8/ipcPbOHPfEsZj6e7MOcAzDql3GTKK5H48MTSWl9Bqt8F4gjuvOwefe/QFFWwkb4bgAo1mBIgYv/bhB/Hwkwfx/nuux3WXnoF2M0KScWRZCo4YIK4iIkuvxIgDnAlAcFCWqZuWBChiaHfaaLaXcfjIBn7xv/8BfuW3/gJbwyn27FlDmqSmn5vNFjaOHkGapPJMQKOJditGlkSYtiKM4whLK0tgJJCk0kqoDE3m97XFa+Vz9tlwijLeJ3RHlVvmvzkMbgkTb7tPiHw8jN7PcS42d64pgDAtyNtALtMphJ0tOrIRrTJ4ckshd8yUce+Mi2beJKNBtRBxQyKoqYcgFSdO5Mip93IPfIwLzj4b24MEScYRxanS2NzMoaRrqfqLmPJbZzi0PsZffPYJtFpt5XBiVe+00qIaB4+TTCGtPwusUGfp+wP84ccfwXvfch22d/qQMQhTFdY7v0yDKw8/eTMwIZtO0W13cPreFbx8bBvNZsO10JTjTa/XwePPHcfjv/AxXHHhaXj7bZfjuivPwen7d6PVasgFVB0MJZL3H7AoQiOO0Gw00O400Ww0Mc4Izxx8DX/2yS/gdz78V3juxSNYWV6RAUuzTF6FLuTW4Re+9DQ6mGBpbRe20w0sbUzR6WxhNNzBznYfo/4Ao+EIPE0xSRJMp/IwUy01vijDV8DJa7brZ15mPaj6pGbIcrEYz9L8xlzVdRkvIFVMvZshAFyqMqjaZ/TtT7Kbk2tSGLq3PfrsKsgz+XVZLcFISbBgRQqMxsk72GHEidtRXHB02m3c/+WXcN8jXwOpqDn6GLJ2atHV66YQkdrHFsjA5NYft3c1Al1oXlntnMt+dMEF6bAOcZLU8r1eF3/4F1/CRz7xRQnTsRKloJJfc20kAHlPgQAoitCMrVASepgVDmkmnaMEGnjsq8fw6NOfxP49y7j0vFNxyfkHcN5Z+3D6Kbuxa3UJ3W4LzSYHEUOSZtjuj3Ho8AaeePZlfOGx5/CVZ17G5vYI3V4Xu3etqZOT6ko2IRAxhiQV+Nf/10fNJSuA9DPIskyRF1OHjCLTT0kmhaG5xdjvZLtPbdpcVAh4Sr44+lYFPh84lfomiCVGFB85VQnknp+2CaBAzBAAxdbq6ozpTV42w8wcpMxuX044PasXKcgqb3LmYb/zgSqT2sYuct8WLG0nEjwSDhAa5kZaF6TduRoOqYCaMiwXFzbzFwcm7yiv6bWSN/iGSSvyzVJqBJAQYCyCnrTojQluQFiCUBcSMHf+yRitVkRmX7EIYUau22mDCNgZJvjcIwfx6YefQxQROs0GOu0YzWYDsZpWjacphsMJBqMJ0oyj2Wyi02ljz+41Gfswzaw22sRMoKihYjDKDIIIiCJjNsu7S8grm9ulfh/5DBvOKEVjLYeeWuOulZi2gEX1dV9kLSSbwQrYFhVTnbl2AezTUiV3qniVCZg7sg2X5+KVVB6hb5h1RG4OwzWn9WdIHFsa1hz+yEOC5EwqDC7aDMxxsWBZ3/LnXOGunUm0dRK0lQK/55EAvkCqkV2E+qWIjY5PZHKTtnDkO0cJ5qafJYDCRKUth5wupdUQRQzdRlve7ChkNKHBRKA/mkAf2yYQoqiJldWWWSnnnCPNMmfo3Kod09E8y+WuNY5mClbBE4Z+St4X2nsyJkGZ0CbA8FmxxsIvyukX/rjOSHNeD25VK0pFgMnsjpPNtG78lZxFq1D26/Pz6sH1GVcUSoUskWJL3CfkfROAYzHM7mzdPsvJZ5a2rp205qiHiY1R4VkZPu6AVUAMvxQCEJkwVgdB3ocIfUmJyiQgtwvtJrlCqKzeMCo2RlI2CviWnoP1ovw8d3ItO58WSFsyXimXp9Ris8dr85DU3JeDatMknx8WXps/GUlGFF8gp1VpGBDsvXLbjV8WFbl6qjtCVrXkfKmYPpS9scq5C5x2m8pgBIbGzupNM+ZKfvVzdI8Dw4cXxCWAewiYZtxgPteW05eH6ANGpg8sPHT/i3nbZp/ld2D6DWauFq/DPXYfLSIw/PLGcrPRqFSvJpffpnn7aQEBYIurEvlkfwjrgSfmdDSUQnx0kfs3C7uqIFX5A6q/2vsCWhDpRciyHio+11GGBEXQXiRCEY5/gGduZW61abaWKwOAXAWEhqQGCMN4dvkCnCoBlxcydt6ilk2ofmGGb47kMVCwfwKcXCK4Cpf42DDDRWY9COBkhQgXfiUBG5XZ29wW+nP005weEYCeU1tk671HoC/L7EfbbBdOHie8mJ+/ADeEZ8gGtE3lkjLOSJNRIbpjZVx/gIg7Z65nYOPVU4XEHElpz9yN+HUBWftpMYc+tTl3BcXkyen6vWXb0cWCdtg4+2UpPZO0UCmwgl6ZCnk8gVSqh/S1USL/HhIClnFdySIzEptHUpM3wEGnBv//ks7SeaQpqJ+oRmpzP0jXZULAz6mlEFVIaq9jA2HOcmZVg0L5LoKRvpTDobLKfDTJFA5kqpFUeR2bYeGkZZySfeHxqsLLGUmgAp+ZRpxFV3rJSJNGmUFcfGppH1/fQIty13JzHWqKFUgru8pCqJdOXkb7DKHFbv5zXlJg8xSQg5EfY7SFh+FhNSfxrSW4qBZkr5NHWUCzDxyS3wUufIuIqNB5djnbBHSnEXY+87+wS+jgGjk/55aByCle2LkBiBjuFGI+8hAmzsEMq6asvIGDnHBC1mYtuLlNWOVcV9pKz+zXSz6Gfw1xhWEGk7DKluFkhlffwlQCUVj9VAWvNHkCxoEzz7iLQl85MBaQUWyezJqEtWcdWS+0Fepvi9SyMIwWzRmtmi0ITlCEgnq1owWpO4QJHnVq+ylQgwGjNf6MexALbVTkS1AiVksGpvZ1ySsUFpfBpCSNcdqZW3jI7VYnWiMhF2A2OkDNAYTBZ6aT4yxwVpMW0pgVfOwiYHOhtaVbwkWmn2Yi4DXQVojkPa5GNJCKea0rckyOeeTT/BaAQcOLCBSSjL5pX6ZatLovpLLOEXBv/bHvV1PldIRiS7s50IJuoGQVUIUKgRvqDZgUiD6sutRfkbR9vABOBg+/3WXj58N2zTxIYaL/7Dz2zxAnVfRDJS4zkj2kZW2xLrQg9U/YY2TiMthA57GyvPZWWlN1xq1aomm2J7LbNTtpVOL5ZYZ3C4/zyiMWgnfgQuQvSkdrxijmhgJcIeCKp3Ls80NIBG0AlPgXKqvBMLIxJQLV2mXMc7+TyvpZ9wvznnlljS93sPNVkVlj6cMtf22MJocGCQUY6nkUsRyEUF6Stvt2sO4Z+MyDfpklBkgXbtJHtvPYVhWFF8Nh7hQQxo4FRk4W6YyX00G1bwycdyacXm4cI67VcIOU5SMeqEubUZrJSKveIMDQ95ppgQEwfEkuXvrkodmSJE30gUrUnqQ8D6CdV3ReqJt1ck8sYTLYnzMxtH7bz1huypSVn6Vdy8q5boy58A7mL+l8LrC5vZmb0kKg0+2g2WxaGpZ55RcY+7nGXcJnxDAcj5AlKVrtFtqttjy9GXAKmgcHV2ksIhUCfe/PFQB5hkaNf+gYeZ5b41A+co7hjbqegL4V6NjWbt2FLijtlzLBQMUyBalYC2tTBZHtsZ2fNdQWC3ma0z/rZHcqA0OaZdje0Zdp5IJkaaknA1MK+zYkT0L7awBmwPNzD37twcYHjCb5wTwrINTP9oD60x6rmEGnYAY4oDjn6LaaeNubv0Fd/S3Dp3/hkUdx8MWXlRAIVWAhX2U6e+0MoVGeCOPREBdfeD72792Ll19+BS+8/Cq6va57irM0iTC6NoqLWAOaNshvlIsT+f1m9Ik8Si0Ab2Gy3Mp0bWvJEfNHBILFNBrpwuASNKNpoSAC78MDn/euKVeL+UteECzmJxMTrmjCemgYZs0HhwAkaYbT9nXxgb/3FogslSfTiBA3WvgfH7kfh08MpCksQnH9tIXEvf7Qv/SZiLL2lEhTmyB120yEHFu42PVYyd4HBFC8sMKu0y1LYOBZiqWl3fipf/Uv0Ou2kKQJ9uzegx/4Zz+Cp556Bu12B1mWWrD8vgkzv2f9zs38BEKaJPixD/4QPvC+94LAMZok+I+/8J/xW//jQ1jqLamTnNWaM4hmQW7VQ4qgLcyy+ny6VGNoBIYBYClc1+G9XATYoCXc+S8G0frTDwqifwlhZgjlxq9vBrony5ShLRsVsDLcZAsTBcecSNSQbAAIWvfBZO4ZUOKDCIPBDt79gdvwT7/7Tdje3kEcy1uAdu/Zi8l0iv/4K3+G1dVlZHm8DA9/N3BIWPUWSpbgF3qoJL0V+6CUuD3LxyHmOV35soxje3sLWdrBNJmCsQjT6QQg39m0hNns8+q2UVJhfJQmAbCIYXt7B/e89S58z9/9DmycOC5vSmo38c9/6Pvx8MNfxHMHX1b3JGQomOJ+CimhRTS/KRfiDrdfXMtUjqM5NVM0xMy3Sr8GL83tClwwe0J0JTQZ2iatXVA3xntvwiWTBlIsWoqRJFqCTbvCwser3jfjClVJXOwqkzTDruUW3nzrRXj11cPY3BxgfWuEza0xDr92DO+862qsrXQwmaZuy3wLT6vq0lGqoY00XMf8A/RCj+5hmeyTll41EJabbWi8ZiTLmoqiSMX3jxGxCIwiFDdQnU6wHgf6Yz4Z5CQGhiRJcNklF2M6GSFNU0RRhNFwjFariUsvvQjT6cSKEVgBzO/nRfHS+kTb7QU4wkXFMwaM8jeS0To/YcGohZ6CscBZgBxRiZnbc0S+0BfuO2ia0Yhqb2SCvcdNTmanBuuhm4Gs//0i5L8qGXDbWpAWtCwcRQyDwQh33HABzj2wislkDBbJ+X4cRZhMpjj/zF248+ZL0O8P5Em3QF154FNRomULtmU5ouov7ypt9vtk4DKcC5lJa6Ggqf36w2TFVD+RM3bCGhqbe+wbE0sEkj1WVNR0dZMAR7MR4/mDB9HudFXAkQSNOAIXAi++dAjNZtOK51DRTE+XLRzQySFdsrrA1fLODXoij6lhqILyS3EdY2JeXFDXAvA1jJJgwn6pMNTCLSdyrYms0aX8hlMT/0/DBiCE7eHnc63/p+DnwdNydB3JXdJLDs4eeYocCyGAZgx8/V1XIktSCE7gWSKJiKcQgiNLpvime65HK2bg9uqyhYdxew6io/Q2+ds7Go5dLkfYMSrMPF5D4PDn3IVzJh7IvG/9EFUBlKx2ahYnvbjmeNjZcAN9Y/22F7Xmpm3FJTzjWFpexl/e+zn8zu/9EVbWdmFldQ1xu42f/6VfwaOPfQXdjgrn5qPoa3sPx7lwCpR3u7LYr3pzItfxOYBQ3fKtHvd5bDhCrNs7I9/sx2W1Ct2InKQd4va0M5m5mMTMsGDgCiSjpS2YIi/qbpv7SDlEKDxCtXGTIcAGowluuOJs3HDFmRgMR2BRjCTjeOyxg7j+8jMgeIr+YIgbrzob115xDr705CvodVvmroB6RONSXx5CTb8WVh6riNc217GlpBojNEOwXI1Uib2lFUkI5fioL22FNwDCG8cA/iWGQaDlpfhIhaC3rBn+xb/+afzhR/8Up556Cp5+9ll85Ymn0ev1kOnbgf1F7Cr85k21yocyKT6w19r0eFhf89JaYhJqr9+IOT0B3UQ5XpXZcoQIisFFiXOtRXe+0CyTNj4cAlSIZAUwKBH9JyxnNFGUtMQYIDK8683XohFHSFOOZgy8fHgD/+pnP4SNzSEiyBj57Qbw3nuuRZakjndWeT/p0VQRka2oNYAUAq5sso1BF6jivQq7sNCrKGjkgvluBUW1NbqLZm5B2YFPQ4KP7JAg4SStQ6u+MlilRp11LJkBzVYTDzz8Rfz+hz6Mp55+FksrS8gZX/iF/4aTY3OqdkqFV+pHo8bGuTFbW3x6/CyQpYnmXAPQBrew7eVZlZgjtjrog+VAYxGRbQXILrD8+f1KzDPmSG938VHkDwMNEeZVLmHd33ne8TTBeWfswW3XnofBcAwhgLgR494Hv4rHn3oVn3n4WbRaLQgh0O+P8NY7rsQF5+7DJEnkTUPlnWPVo4SVcjJiTN50y8gWDDZu3lFRS/ATEVjE1J14DO5BF/vP6VDriLMAixjiRgQZPTnDdDrFZDpFmqYgIkTqMk/HeveYlEy73B6gwDfGmFw4ZAxcCCRpiul0immSIOOZfB/HygFLGGPC70a7lfKV/Lbc62H33t1YWloGcbs/UCxcJ52sZRCAJfk9XycRhTHyG5oHdnOOYRtSKbEsrDRnSDDtTqhFQaAhIvDdqVwYOIICwcAMFed2jrD+l2BcMpOzBoOdgWHsAKE7x76LgAo42gJOdyaLGCaDEd586/VYW4qxuTVGsxlhfXuKj3/2SfSWl/HH9z6Ob3jzVWDgmCQZ9u1dw9e96Tr8wn/7BNorTbnXbNOcb9ooVIgByWQiV6fVVdfEGFrNlhK6nuA1sNQX9TGdTuRVYJCLcowxRJEaasGMlSGt8VyAAMrDUQjs7OwgmSZYWuphbXUVnXYLBGAwGmNraxvD0RjNZhO9pS4olTH+IPSVG3p1OjRebidEUQSeZej3+0iSBO1WE8srK1jqdRFF8vKS0XiCnX4fw9EYcRyh1+1Kjcd5Po4OZPU/AQTCaDREmiTyreBottpotprqMpe6nKzo3gyB9i6xtLDz3bXUnGEvZJNCkkTomvj8d74y5Fl+TvTsXBMU3DkMPJ3YIn4Alqni953f9mBZTXyufDPMSuQiLlRHO/1imfgKSi4k7OryGnRc9UKF1g8NIxdGQJKk2L3SxJtuuQjD4QhCCPQ6HfzFA8/gqy8cwa7du/DIEy/jkScP4eYrz8BwkmI4HOKdd1+N3/6jz2GcpKYOOT2zCTYnKBYx9HcG+P7vehve9sbLsbG5g06njUOHt/AT//F3kRrG1dJCWM1XmAsgogw/8O13YFengfFkgnazgS8+cwQf++zTaLfaEGpB0O5/oRgljmMMhkMgS3HT9dfhLXe/AddcfQVOO/VUdDptAAKj4RivHTmCxx57An/+iXvxpUe+DEGEuJH7cmQaJ+5SoM0KjOTtQOub61hZXsKdt9+K2269CZdcfCFOPXUfVpaXEDca4FmG4XCMV187gsce/wru/fRn8eAXHwGLYnQ7HYeJSdFPzjBAmk7xEz/yz3D2WWdiMBxgqbeEzz3wEH71138TS0tLVkj3GcnzUTDjWGACyxWe8lOGJk5yMbtFByGTxrP8DHOQBS9fVA8WKzVBF7gb0LZX5rOCNBbqOGrZ9EEAQkfyJQEQg+DwhIAtDWwhYD0npfmtd2F83Z6yL42MiLAzHOKNt1+IM/YtYTAaIyKGVDD82V89CYoaYARMEoE/vvdx3HLNuRA8wWg4xoVn7cZt11+Ij37yUays9uRioI4J4M9shGxfkiS48LzT8Ya7r0H/6Dp6nSae/tpRNbVRC5XCxpyQL2Dq6RXHFeftw5l7utgZjNBuMhxdH8iLNDRBEjkthpAm/ebmJq696gr80A98H26+6Vo0YrmXnqYpeCYFWGvXCvbt3Y0brrkS3/r+b8Rffvpz+Lf/7j9iNBjIOPuGxsnQqNPTJEOSD4ZDNKMI3/Vt78MH3vceXHzxBWi3mkiSBJPJBCnngCBQo4GlXg8HTtuHW266Bt/9He/D5+5/CP/+//wFPPv8QayurioLyxtdJX/SNMFVV12O6665GjvbO9izew3Hjp9AkibhhV+Pugw4PWZC05bOpAUpy78by0p3gL8Lo6jRWKIhhs9Naac/DYYWYiFON1Y6GV4opioBUGBQ3WJeGFgHJx9/JwPBGDK2GYR8WmAQtsqYhbEgzBDSQllslrAK4uTDsPCDgOCEmATecvOFyKYTiCxDq9vCsweP4JEnX8JSr4ssy9DrtfG5Lz6LV47sYN9aC9MsQzYd4hvfejX+5FOP5tO6ELr6q9LC49EE6c4Ig9EUWZagPxiosQ9oGq2BLNgEwnSSYjRJMJ6m4IJhmmaWjnQrl2sOETY21vFt73sv/uWPfxDtVhNb29vop3J7s9mI0Wq1EcUxOBcYjyfY2t5BxIB33fMmXH7pRfg3/+5nMR5PsNztYqoG0DYEzQk0xjCdTnHdNVfixz74Q7jmqsswHg3R397GlhBotVpot9toxTEYgIxnmIzH2N4eQ5C83OPuO2/DtVdfiR/4p/8cD3zxUayurChXY/c+A7mAKvt0MNjBYNhHsxFjMpmqtZVw8mdoTs951i35BG8xs+Q5HeIr7xNnTUbAsiL8C2YkPQrKzy4WjI7QVpfZCSA3n0bASuW9UBCOOVIUIEe3Y8jJSWAAMWWmuMxvdVceBciO5ye8dX5n5TaXkjmaUkjpvncWLMsRhjZptC4lIgzGY1x67j5cdeFpGA7H4BlH3IjxifufwvYwRczkWfhmq4mj60N86oGn0Wo3gSzFTn+AG688E1dfeobcNiSmFUYgwKUwhz4YE2AQiBghZhEYi1yUrW4oxjeReEdxhDhiiJharIvV3FLP/fUhEhCiKMbGxiY+8C3fiJ/+yR9DkkywtbkJCIFer4u9e/ZgPJniiaefxUNfeARPPv0shsMR9uzahV6vi6PHTuCM00/FT/3vP45WsyHvDNQWhsi9DM2aCjGMR2Nce81VuO22m3Ds6DEMhmMsLfewZ9cuTMZTPP3M8/jcfQ/hL++9Dw89/CiOn9jA6toudNodAAJbW1vodVv4hZ/997jg3LMxHk9UP0lD214vIZC8eVgtWkZxhDiOSu3BQhKmaz3G91QHCXXoDBaNaqFADsmasfPzGXZ0Q8cXxbYwVqEnDaxvbvvIe6Z7Z+4pQN6OgDq1NZU3ISmY/I5kJauPrEaQcgsmHaRBc6iSIMYm8/FgkqFYKGaeLTw8s0UIeW+dIBAjZMkUb7v1YrSbhP5QoNkAjp7Ywb0PPo1WqynvzgMgOEej0cKfffrLeO9b5WJglhHaTcK733oNHnr8BRDrgkpMVf0p2wvwLIHgHMQaYFEMRoTMF+h+020lRDJARERMXZfF7LUhSAIixBFDv9/HHbfeiJ/8f/1z7GxtmVvoVldW8Mqhw/j1//Zb+NRffQ5Hjh5HlmVoNGLsXlvFDddfi7/7XR/AFZdfiq3tHaysLAFcyEs8bMkkLEQJyLIMyyvL+M3f+l289U134qbrr0Gapnjo4S/gdz/0YXzxkcdxYn0T0zSBUIuKqyvLuOO2m/CPv+97cOD0/RiPJxiOxti/fx9+4Pu+Bz/0z38C7U4bItM3B+UahhipbVyFADF3e7ZGKtBuwUQIXDQatDYLbFFWIxQhgkTYjyQXQMIwt61MAxq8UB5YKCy4krAh5O3BLqkxhJZTTNgPRE61zp6n8ABZUtSOlVfQtDYGBD+DNs0YAdMkw5mnrOGWq8/CYChvr+k0G/jcF57DwUMbaMYRBJdrBoIDnXYLX/nqa3jgi8+i3Zamcr8/xJtvuRhnn7YLo/FUWkNWpCKXgbWJxsHBwUWmtvMig5dh+lC7zDMCIyZvM1bbarq87FZuPDXTNEUzjvEj/+s/RRxHyISA4BzLy0v41Kfvwzd/69/Db/zW7+LosXVEcQOtdgcUxTi+voU/+Mgf433f8b343d/7CHbv3o3pZKru7LNQE3Zf59YVY4TxdIpf/bX/jiPHTuCDP/av8F3f+4P4g4/8GQ4fPQ4Q0Gm30et00ep0MBxP8KE/+hN859/7R3jp5dfQarYQRRG2d3Zw5x234bxzzsJoNFJ1avWq+9OnVZF397wpSE/6RT7FDRh4zvdS/gFgn9vwp8NGQVq8QBY/6P4t1Gknr91FAVCKWF4+p0FReDkjel4AvLfbSdySbrlHmS3xqnHUQsNSicFWwDGf7JVkxiKMJ1PcedP52LPSwTSRl80PpxyfeOAZMIrBRd562RcZkkzgY59+HIITIDJMpglO2dPB299wOSajSb4VV9EzQggZWIRz07+mNRXtdg71MXWrsdL+Zr5LgFDfIyYZ6E133o6rrrwEO/0+GGNYWlrGQw8/gn/4j/83bO30sXv3bjSaDcUwHASBRjPG7t17EDWa+Bf/+t/iLz75aexa24XMTLksXaTvr9PDRvLy0G63g88/9EW8673fhg99+M/Q7vSwe89utJotEAg8U1eCcQEWRdi3fz9efPkQfvE//1e02x0ADGmaYnV1BZddcjGGw1EepIXg0JC9pewz48knmwqKbteh5JjnPj6WlaaFVz7uGv9cwwunXH7ArizOiT81KAqASv6VNZGwRtN/PXci75ttq4ZGy9f+HgLmCK998MTOkq8tSDi6I/O7DlIusGe5hbtvuBCTaQoBgWarhadfWscTzx9Bt90ygUB0vVnK0e208eBjL+K5l9fRihlIMExGE7z7rddjba2rYgfIOkKLekIaANKyUILAiQFYMTbC8gIhQDG/dGMmBSK/YEm1lIB33vMWIEvBhEDEIiSZwL/7Dz+PlHO0W01Mp1OrrdrcBbIsRSOO0el08O/+w8/j2PETiCNmcDYWNzyPNtX3Qsjpwtb2NlZXliF4Jrf0IgYWybWPOIqUACNkPMPuPbvx5ceexJFjJ9BqNaXvARM45+wzwFXf6r1/WY9tgeTBUI1qCPZnSSeXkKLVqPLBqUqmOmH9zgW/HPbct5HcQsiP/rrvwt7AjjQBMPdhILJsp1ANHlnbWqkCdAFSwGwXlruuCGZUkGZqSb+jcwwESfN0NJnihsvOwFmnrGI0SSC4QKPZwl899DwmCcBivdrMrQUdoNGIsLk9wSfvfwbNZgMZzzAcp7jkvNNwxw0XYjiaIGJMCnJvcdL8Jg4hOLiQkXaE8lcPhYIKtFBumTICY5BWgLzFRJGQBMAihiRNccZpp+Haq67EZDSGEALLvSV8+tOfxeNPPovV1ZX8Nt6Ssc4yjna7jee/dhB/+JE/xlK3Y6YBur+1M1I+JdPRkqQwieNItpGAyTjBoD/A9vYONjc3sX5iAyfW17G+voHNjU0MBkO8+OLLePXVw2g0G9JK4hz79+8z2j8/DGUZ5o6arFIgjp1Z1c3u5wK8H1yEJNuC8OzsEHkbUnaVWO4O7iX/gBrqeAI6mkf2qjC1ePgX8hfqC5cJPRR+Dm3aF22GSmgi+BX5LRj2dUySmZtM4M7rz0eaZuAZQ6sV4djmCA8+/iJ63Y68UZi58EgtLnV7PXzq88/gO77hRkSRPBWYpVO8563X4S/ve0oZJ8JDSOR0qawAcC6tAGNSk9stup81ERAK20PECBERGFlTKpILa2mS4Lxzz8Ep+/djONiRjAmBT336c6rO6qTnmxnniKIYf/WZ+/F3v/0DiBlDpiwWuXgljN4Qwp40CjDGkGYptrc2sbTUwzlnnYkzD5yOvXv2oNGIwbMMGZcCJVJWQZqkWF7qSc8+1WHtdluGIjPCxsJRKMawXWxLjEP7odPbs4i3etZb7DfDyCp8m8Mzopjf/+WYL25UjUpbhKzFSsVLswVASFD5DzWW+tOI3pnQwxVamjrvG6tpQSFTIpBIGke+E2XeYdJ5Rp85GY2nuOTMXbj0nP0YjSbS/G/38Md/+Tiee+kolpeWkE0y8CxTMeXUoQyC8cH/4uPH8MnPP4f3vPkq7PSHGAyGuOnKs3HVRafj0adfQ7vdkEwmkBMAAKi73rQeEGotgGvp7zC49Sm0BUHaxpfaH8KyBCjfVQEhyzLs27cPjUYDEJLBRuMxnv/aQemea7svB5KxgjOOZqOBQ68dxnZ/gE6rCa6Z0xweIjMO2gRlUYT+zg563Q7+wd/7Lrzz7W/BBeefi+XlZcRRDECA80xqeZLbhyyKQBAYDgaYjidSEAsyC522YjCBUVSX5G7PUA5moRb5UZnhchPlJF7WJ7NJnnL7XFlIgry6ioragk7BSuyiWsiE2pjjL1/OsQ1IyD33PeiqIb5WqtkjeQOE5gcbxfwiEguVGsB0c+XA5zEHDMLIu0K2i1GENJni1qvORMw4hhN58GU4GuOM/cv44e9+ExoxM0EkhPLCk1qHgTEgimOkHNi3toTJJAUISNMMKw2Bd9x1OR788tfQ7jTlNp8R/mTmrgRSpq2AUKGqcgFYpAzjoaqku15ll+cIFFzfWiNpvjebkbkZhkUxkoRjNBojIuXbPysJ2a+MMYyHY0zGU/Q6bctesbWdFsjSMtnZ3sZVl1+Kn/7JH8cVl16M0WSC8XCMfn8bjWYTEeUXv9h2AxFHypP8+jAqbvUW75XMrbt8rcfO77JG4GUOpoL2qsmSlJXlppxRbca2e8wed+u7JcPyOLN6oiecMcjr8q2LufwAFJMoNiooI/IQmidZ5r60brS2XsiEcDrSmL76N7kXfdiKdDpNsW+1jesvPQ3D4RCCYkjPuimuvegAWs0YoAzSYpCONnKxDXK/PWJgUYy40UKWcYxHMmoQY8BwZwd333gBfnX/Kk7sTNGMY3D/lBwBxORaB4faDXDcCC0CsMx6+wAVMenco+Pgy31wr3/UtCJLMuMaDMi1gSiOpS9AQM4Hu1qtL+gFR8DWo5ZXpRJyjBiGoyEuu/gC/Pqv/ALWVno4fuIEGIuwsrKM7Z1tvHTwFRnGW6h1lowj49LTTwiBC84/B0tLclE15OLqHrdQ9Gqkf9E0nTnhMeVKxcTMZLxftThz+tYeW0uRGmYPaH3SkIrHxY28MuZ+GVYBV2BjNhfszZxQfKHkK/68WICCfEVm7xERIN33YDoriL1TmfWjaN8bgWLaZb+3EiOG0WSIN19/HvaudOXJswaQQS6ibQ8GoAEpQhfK8ytSv6UwYIyBxQyMTRBFDHEUQQuL0STFKXuW8dY3XI5f+/370dq9AmR6xR65ealWyMEFBBeGkQUHKLKkPxUFpOBAsxWh3YokPMbAKIJeBLS7ibEIg9HYzPc5z9DqNLG2uoI0y0CMCmGzbT0kH5A5Lry2topetw2epsZWhNZFAlLoKHs8iiL8+I/+b9i9toLNzQ00Gk00Gg382m/8Fn73Qx/GkWMnkHEreKoQapWfMB6N8Du/+cu46YZrMZlMfWLwEMynRq4Von94msAGUyA9kf8/h2Wr8wYPRJYlh6M95heACRnk1CP7V/myuT1S+JH3R6xNR7tCo4WVbSG/egdIYHVmqbZwjVdpsfkZhdvIwuKi7x9d0TpfoprfipjUtVAhJ5CMC/SahDdccw5SdVCEZxnabekHD1LMT/JUHjFIbcsi8xlFDFGjKbU+GKKIIU0nSKZTEBEm4wneefcV+L0//QLSVJ5xd2wpIkwT6QUoRIYkmaLXitBpNzCcThAhMgLMt5GIMWTpFLtXlrF7pQvBE+lFGEk/AGnRaztYoNFo4NXDr2E4HoMxQpqkaDUiXHrJBbj3c/fL03rInD4qaj7Z7ul0issuuQjLSz1sb20Dyu3ZCxonp1PDIa6/9mrceMN16Pf7iKIYy0tL+Llf/GX8m//j57C2azeazYZ19l9PAiQ9EqD8Gsgbywq9bHO9Q3plNIUALdd8V5Y3WEa4GYwQsqY8FrNqbS+01af72Uxv6iOkq4oBwBECmheFHkDtVKAlhzASzRWYZeftXIljcvlmjpPfR3heo8sX5RYcm9+gtSFhMJjixktPwTmnrWI8TcDA0Gw3cP/jr+DLzxxGp91CZkxAt6+ImPK+A6C88CIWYTAc4223XYw3Xn82xpMEo9EUl56zD7dffz4++cCzWF1ZcvwJMgAb2wO5+JVxTKdTrC53cOreVRw58RpaDbk4Bn+wlXBK0gxXXnwG1lY62NyYIoq0JyEzjRWQ24utZhMvvvAyXn7lFZxzxulI0gTj6QRvffNd+O+//fuqeWqBKh88R8iSmmI0Gg183TvfJmMQUL5OpBfbcuFPGI8nuPjC89FqNDBQh3+OrW/g9/7wY9h3yiloxHL1X/pByIVI229Bn9aUij1Sg2CZzp5+ydc/hCU4y1M9avMrqVWoojYy5wgKRrbwTHyh9ydcAcjdLDXxEfndgHkZf6FCX1ah86nOtu1BY+a42tw5wSTs+bemRrLqmCFyF+5kqz5NASwfQP1/HHHced25AOcAF6CIYzBJ8bsffwxfe3UHnVbD0WjaQ7EQZ0A9iYjQH0zw0qvHcdNVZyITmczHE3zj267DXz38nKrdPh9AeOHF15ClEwiRIcsInRbDG26+GA89/gJWV5eQTBPDYLI2uc0nAHSbhG+65wZwIcCiGFzIRUz7JJrsdoEojnDi+Dru+/zDuOjbvwWD0RCD4Qg33XAt3v2ut+N/fOij2L9vHybTqRkfWV4tfEKg1Wzh+Il1vPNtb8Zdb7gd/f42WBTll4BkGs/cUuGcY3VlSflQSEvk+LF1jEZjxFGcH+1VUyHjockIGedYXV3GgdNORZKkxRN9RnNRzuyesHQddcuUjc8TJcm2fu00zxQBgAyTxkAWHzjGrJNXGB828vMC+TGImon5uT0LRCV741T2qHdfgYOg/tBDb1sNxopwqiEUPPechte2b2YkoTowZwgiwjTJcOGZu3HVhadhkqQQAFqtGI888xoOHdvB3t1L6HYa6HVb6HXbWOq1sNRro9frYGmpi+Wldv7X62C510W328b+fat48mtH8KWnDqHTakIA6A+nuPma83HlxadjNJmagyqcc8RxhMefPYTNzYF0iSZCfzDEt7/3Dlx6wWnY2NxB3IgRRcz4+0eR9Fw8ceIE/sG3vQnXXnEORpPMmNBMWSdmnLSGyTjiZgN/9OE/wWA0gg5FNhqP8S8++EO44ZqrcOToMVVHpJyL9HpHhCiKcOzEMVxy0fn4iR//IJJkCpDceQDBuAC7e/+yv+XcXRJrkiTYtXsN7XYLQnBEETNbskpzAAQ04ga2Nrbwnd/6PpxzzjlI0lROw+Au6mphqvWMO1W2L7SB6QyXtHJhXL4IrTuxhCjnYX5nUOoXNHLHsIacBxj9WqM8QFoACLh2RlkpZkyQYDIvvG0IYTVSMXquxQKH5f0xWujkBlyiN5owpwppoUpnlDuuORfdZkMutgmBlBPu+/KLIJJ74lwAXAhwkYFzgSyTJirPOLj6nmUcGefgXPkJQCDJgD/9zBMgxmSACgF0WhHeefc1SJJEzs8FIDhHu9XA8y8ewyNPvoJOK0aWZZhOOXYttfBLP/U9uPrS07G+sYGNzT62+gNsbA1w/MQWiE/x4z/4Xnzfd74FOzt9ybA6JqDaBbBHQAiBlHMs9Xr40qOP4WN/9kns3rsHKZex+HrdDn75F/8jvv6eN2Nrcwvr65sYDUeYTKYYDodY39jA+sYG3vTGO/Brv/zz2Lt7BcPxGBGLDOfmTrfu2DWbTbz4yqvQU8bxZIxT9+/Be77uHhw9fASZ8v2PY3mEF0wKjFdfPYxv/ZZvxD/6X75HxiKIIge20IvHemAthSWFKawZZ2AXC8o6cDRbPtUr/OkXc2n6Klj1+D+3qa0yZDUd9VDS1nlspGDpufm8anlOXvFQIKNxr9f57RsO3EkrwrZTDgeY35wpxceYg7r71G4wAdM0xb6VJq675HSMJhMQY+g0Yxw8vIOnDh5Du9OyggwXo/KolrrPhLJ8uEC328b9jxzEC4d3cNreJSSpQL8/wJtuvgj//UOfxbGtCWJG4EKeQswQ4Tc/8nnccdOFECIFGMNgMMR5B5bxWz/3j/DJzzyGLzx+EBvbQ7RaTVxy7il40+2X48Jz92L9xDaarTZ4I8NoOAJYBOf4q9WhelrWXVrCz/zCf8EtN9+IMw+chq2tbYzGIyz1WvjFn/m3+NwDD+GTn/osnn/hRYxGI7RbbZx14ADuuut2vPnO28EFR7/fx9LyMqYTeYSXq67W/aINVSHkjcFffOTLePnQa9i3Zw3j8Rg7/W384+//+4gaMf7oo3+CY+sb0g1ZCMRxjAOnn4pv/yf/EN/9Hd+K6XSMRhxLC8AoEGUJaBq16pMXa+RBQnRen02K5JlPBaxZRW45BgvOSCJfKTPfHL03i+CVS5uaBhj0uGtRlzpyOojLwvGsNri6WRn1qoaqRdRqMVRRI8HthkWEABVRcAMkKsIkwmg0xl3XnIc9y030BxMAQKPVxINPHMJokmF5mZCZDqYiAZShoOREq9HAxvYAn3rwOfz9b7oV4+0djCcce9eWcPetF+M3PvQAdq2tyOuqucDyUhefefh5/NZHvoDv/dY34sTxTUSNGIP+CFEs8N57rsb73nUTOGuCEcDEFMPhGOvrO1hb6eHgq5s4fmwTF5+7B9tJgiiOEEVkTHF/Zttut7G+uY1/9IM/jP/6n38Op526Dxsn1jHlHNMkwR233YQ33fkGTKYp0lTCa8Yx0myK7e0+OOc4cOAAfu8PP4qzzjiAq664HNPpFGByq5GMRSDPTTTiGMePr+MXfulX8DP//l9jPBohywRAGX74f/1+fOe3vQ9Pf/U5HD+xjjiKcMqpp+Ci88/F6vISBoMdtNtdvHb4MPafsh9ZqtqknJbMQSB7/HUYcl8ReQOohYYTcs9fGLe/L6CY7J00eyePnPgWFeWN6KIcAFntJVkHgxYQesG9fGow8zBQCKe8D91ZvLPAos1sjRkI+WJiEaqwvuiTa+FOrmHkCPsvH3RzEkx1eJoJtCLgtivPxGQ8BucpIkbY6E/x8FOvoN1u56v0AkW3bQdxD0VoGuLodFr4y/ufxObWGBAcWZpiuLODe26/BMudBtIkU2wprYbllWX87G98HL/5Rw9gbdcqWo0InCeYTlOc2Ojj2IktbKxvYP3EJo6d2MZ0muC0U9bw2ok+vu9HfwVfe+kwGg15vp9DCjCuDs5oOtNDkmUZet0OnnnuIL71u74Xn7vvQezevRvtdgsQwNbWDo6dOIHhsI8kmWI4HODE+jq2NvvodNrYu3cPfvO3fx8//KP/EkmamhiCWcZVeHN9Zl12SpplWF5ewu996CP4Dz/7n7Br9x70lpbAM4Hjx9fRW+rijltvxje95134hq9/B66/+gpEkZwq7Nq9Bz/3i7+MP/zwn6DX7SJJpkiSDEmS5R3vWeVCAGnKkaYJ0jRDmmbI0kxNZXNrTjKItTrlq1H9qNxwrU5aWwtNf8yJmB6yLL3qoe1ufZhKomhrtnxZ0DlvWno6SMbqqrAZyhDRUtAHSoA6fWYJJhhXVwqV8XH05o4+agLh5xop8r8XJ1va424ymeK2a87B1RefgTiO0eu0sXvXMh559lW8dqKPRiNyLRGEuzGIj6pSCKDVauC5l07g4a+8gFP37cJyt4VG3MCNV56Ld959DXYGQzlXh5wHMALiRgc/9Yt/gg/+9O/huZfW0e22sHuljdWVNpaXO1hZ6WH3rhXs27MXUaOL3/7oF/BN/8vP44tPvIzdu1ax0mtjbXkJaysr6LY7ZhdGmT7Q8yKCDAzS63Zx6LWj+Dvf84/xIz/+kzj4wiGsrqzhlP2nYM/u3Vjq9dDr9bC6sop9+/ZhbfcufO3FQ/ihH/4J/PC/+FcAMexa24XVtRWsrC5j99qaujZNWMQvpADMOJZXVvDzv/Sr+L5/+kG88OIrWF1bw769e7DU60JAIEkm4GmCOI6wsryClw+9hu//pz+K//Bz/wlru9awtraKpeVlrK2tYGlpCXqhWY+F/srAsLa2hl1ru7CysoLVtVV0uz1wwZ2lqVyrl3M42bQVGv9KNhIBSrTezmBBp5z5Yls+OR5hGg1VIBDnM8JZW3EuGuEVUqkfdS7hPc+xCyBjVKutY7VPkyV2yIYTgOF8V5go1e14KnCgEUXoTzL82kcexnA8heBAoxHjyYNHZUhqzmcMqtVQHx/LEOKcI2408et/8AAeefIlJEkKAUKr2cCR4310Wg0zJQGUBUQcnW4HH/7kY/jL+57ATVedi+uvPAdnHtgrw28Rw8bWCM8efA2fffApPPrEC2i3O9i7ew2//ZHP43MPP43xRLocv/DaFtoducpuLl3V+Cnc0yRFu9WCaDTwm7/7R/izj9+LG6+/FldddTnOOuMMrKwsIY4iDIYjHDr0Kh5+5Mu4//MPYWNzU0bmTTP80i//OnbtWsN0OkG33cEXvvRldDod8My/El3Wv7q2hj/783tx3/0P47abb8T1112FM04/gJWVVWRZgs2tDbzyymt49LEn8OCDX8TOYIhT9p+CT3zqr3DotcMYDAZot9t46eVD6HS76p4Ad8hYFOE//cqvY++uNXWXQQsvvPASWq0WeGbvnvuDGkizlE9VEna/u/MQzVF53+RTDxP3zz7sZN5btz5bwWGL6Fn856N9+lnnijpRTCQYBsan+Hc/eA8+9ukn8fGHX8BSzw2Ooedg2qT1KxVOB4TqyEVI3iUCTutscVeaLHFtM6kAtG8DAZgok1WaVhLjdquJOGb1DKMQ8/tZSNaWJCmGo4lqmwTeasTodPTFHzbmUnAxYkiyDIPhWMbki5jCTV5FxkFoN5vodlpqt0JgOBwjSTPj4NVsNdBtu3VUpSiKkWUphkN5rTYjKIcigGfyCHAcR+j2eoijCGmaAhDY2Rkg49yMWbfblYxWEXs/YhFSnmE4GCLjHI1GhDiKIXiGaZIiSeS13ktLSzLsODjGownGo5F0VxYCrVYbnW7bjJ85aKn6cTDog2dS+Mk6Guh1u+ocRK0uOYlURiBkhIKz+GddahNkWaGUNUmPyO3+CO+5+xJ817tvwd/50d9EhgaYN7E3/jkeNCJCPM+ERjeD2z9KkrD+19mF9y1ch2WqhzqvaAdVwLOnEQQS2iUY0ItDrWaMdqvhwNCrxqWD5/fnDCFAQjojNxoxdrUa6kZavR6hz/zL+UI+FZT/Z5wjIobV5Z7yFiPog3pMGUScy7m1tryWe53cUlLChvMZ42y1gStHnqWlJeVkJKB5WHr/yfxccMX8suDa2grM1IKgtkR9s8ztqExkYETSqlG3EkHFWux09TVyattVBRvptFvodjsSH+jtWe5Oh606V5ZXYDIDEFzMPO5cO7nKPJCqCEOWt0+7Sk95TQTKec6ux6rX3SVTSlJJQO0ubE6YlpgtlQLAHy5tVZNaDSkraR0wc8sCpaiEk75swT7EqVtdZneT0QT269wrUcWoU3CE0CfuSqR0KJH3WaNBmtmzTPlIuGoqXJ0208EhuBaHOa5ZyUwq087hdZHz6tYlOOfgXnl5UImC+Ga2SR0UimFhKiD7BZlsG6kVbq59RDxYXAhACYMydWKrkcyxQIT79WSFwOtV3hi4OpaBpW4FINf2ddJTY1ao3zk6DcDsAJRUXnkcODy9VbutZdaT3iK0taTXyMrk7NXnxewgCrOkrew7mwnkQ7crRKFcOU5e1jmKhuCQ/0zbrcH5W15p5eKo/WheovRhmXELCKmQpIDwfnuw58JHqJughMHDXSB3CWumeAsdia2vgYplQkwxbxt9C9J6ka/j5xVqWaDJRGa1aF3k7JGftRMlFpFb38x4AOR/E7mZ7psmpPES3m+ToaKn9Csql+o5DLuQfk7mdX4BaI7IPHZHsU5v3EsZ1U4B9U7Wqoj3esYukJtZM4X1ubAmshjM6XffyqkDR5evsmrqJHLxEP7LeQDbBqP+vYCAJA+nYF/VRCenRyoXrmH7G3r/XwaiEfmUwTmt635WpbmvBze+1z504fJcGIGKnrKwN3ACdZQWFoCMRuvY/Xm5mb1STRnGUTIkVWsRgCgYJSH4dZOjfCslppc/8LDgBDoPLlbfipLnc8NSf+W+IPPhZdKcloitQn0fgIXQUuAc/hGFLOozjKy9QK7hyUDYi5ghC1wPnmNiq/lA3c7zOZDzBbwIPCstaK+nepxaCsN+4UmvUH21+ziQsUozvg6ack5s3D5eFJcSeTs3HB+AZ8QtlMrGvBY5kvOhy52UxeXgY9MbuTjphhfOwMgyeWxD5Xpn4r3DtcRqIDPfzUDCEgHC4m3rmaNITeYavRXC2Onoqibl4tk39cnqmPIUkGI+NdfrVRcl//s8ttk88E82LTp9sJMtbxdRRp6wLiWdOnA9q6hU5leVL8FLzCxf8rIwXl6Mf/tSUYIVx8Au7xNVfsuwUXrOlKmcSISqu7iMWFWEbCsgAPwk5qEFBisQ0WzgsjNcO9v0W0Fa2Tnsv7J6/PnavCkgXBZJtkA62fR6CiI1fmXacWZVHlmVav9KQN5LjdOsukO4hNoxU4mUaBtH8BP0wlglL4XAW99l/8i6JFgyx9xz/VtOJBqDQjyA2XhwuU2j9ydfD0IMmWoLwNUdYXwj7eXTSqAVNjDZA0uojVihLWLONoXHpdSte5FUB86sPHbfluI2I7xr2bjP3c6A2b5oXwVpchav6PoZfD9j29W99KZcncTs3Y08GjXUtmlRnMzmbDHvIqB//OdkVAiDMwPRkr8MZK2BVKuiOrMIUGQlHB2bwFZrlrtlCEBdmTJ3CpmAc4AN4bXIFGTWzGtWRjL/zQ//ZFOZXF9YgNYs+H9Xd22vuiVH/Vdr7zlnzszJJJOZwXFiMhniiEqCOkYZQjDoaCCCEUTFCBrwURT8I0TwwUcfFFHwxedoMOAlKkQlCkYnkMRLTBwJiRkTz5mc+z57lQ/d1V1VXd2r1/ftE2Md9vnWpbu6urpu3asvwSCIjUkjPD7CjMtiAMzJ2zNky/z8p8cCBMvGXhpdAzDskm961C1IVFIhoZrIvUc3VyWVW935C2qxGcZZGsVzE3VmjvUU86K8dBffRnwcvb4Ijz9Mu5ffO8vbay06ITwHCnp82UH3IyxjEMKxTtUpTw+0EFK0wJSXApBLkzsIg7Bx6bGgfZ5naIHayDrK3HWakjl527S9tn4vCoccys8MU1TXxvVcnVSGpqVcz4bjQovUm7r1ErRDGPUm9hoMIhvglPByAtexhqBnXEiVH56+swMIdQ6JV6Y9UVfQ/pFfHgGburHjuSdA4/RnkFWHEsoAYJ2WJ09PapKdefLcEmJ3LHlJgzJz11MKUNoecwb0gALngputyYtVcqG8KSTX2m9dKo2Tlc2s3FPbePWos79C55KXoKrDGMpAzACdwlM3/6SCK67XpMD3yt3p3KRNmy5T29ab5JixnK2umE4bISp/Wzwe4BY96XVZprqJqs2hxG0YqnXGXbIMaNnhMLnEyPLCdx9VPiZLny8vJqSi15pOHU/OWotsAiljMt6mSqQ+taeJeuqzqS0GdXpI3TQDfb9GM5TUfRB+ZWNC5XqmnpK5/ibFF6tG1RN08/VfiV6I4u6d/OPxka4XB/IYgk0ljqJTxG66DtX7CpkxYvMPjiKy6OsBsiNwmRBviEevK/FOQaxajn4VLlJJdLEjYPV/KoqreJbHrHRKI0zX0x3u6pdRLRDZ9wF1HUzpT/ZC1xkoc6UYg1VLlaRu4ib1jAJL35HKoaRyCaGsTZjWNFeOBF1wYWgM8Xt271tl3gZb+O5j3EKogn5s5C9edrwr1AYYq6atb7dAm7abpMpWv44uojDrUbS7YoA7+010vZWKwsv/LvQPG7QNpST/tAGgwsekbHrSbSVjpr0quVw4UrWsvzfpHrHKBkG2/54JGbm5qFZUvztCY0ho4dYee+gFo+Z9OSFzBpSYsH0cTu3eCx30+4HU/76APbRkqWRg4yPkHMpj+QM44+H1JTvFHYVF5wJYnKO8CeaH3FmyKQU2EVGv4UbFS5hG5g1pBe4i7LwQ6y9BgYTJh8iAzr85DjHGITSVLkkv/aj1dGRJ/uEEIUHg1OsCTOMLYutDd3EXtOa3eTFHF2cX1e2xbYBvhqCag5xuRqqSRY9M72ys+wP9GEBHytW71ch8n6Xa+80NtR90TCsrVxEoebucPTg7oKcpjBpa+b8RGBqU6x/l3ezmBDRs0TIbFhzikdpR28NxSb4gb38AeBLnKN8MPlaKdUCPDeg0A2+mcG8DGWjwWOokwu6LlurStCVNUGZh0gAw6pBsGnjzvRELsyTYkVpAhEcWPHhL4K/F+iXjUKximUswaPpNqRD8a/CXQfeBDMQhx0VEkbaMdpT56MJYduHZTFjJcEDRzbHGS+ebzO/izCCvdioB0pC1WgFpp2MJoIixjqzbNS2CKw2McpH5lL5FOQXUNQBzvYg46N8ZGZDfyFEYLB2mxjqodNr6rfk11cbpNvx+Mocwiesii7RYqXlyWNy7pxsneQ4o59guwnQoHl1r0OKv+3tRNm1wc9Q67KLmPLMGoqQTpxbMQBWby5UEjUL1ICYgOdeOAfBhhn8j/RyXXoyChPabDFgKw21EIaFTNnkaX0hf2fHfoi954dLbYqYhCuW2cPj3vUCm5916feGSXkdJ9dH+uimD2gj2JLJefbSjPRQPZnDozvaSohmPx4iJTS+Tz4gW9y4TVUaoFxUp7VxOU+jx+G0lo8/EehcuvfCn6twOOrIFGXcBtGnJYQpLLcokAzE7iRFp4EUsaSdMzcwjsD/GzRCXLhOu0KuRv9UPqpk05GhBmO+p9MvfwqEtPKMZkDRrjaIyGgOSjGLKkhEeQpfCT17hm6Wo+1y2ye50aB8ioGnSzQxZDlhmrgZ4GuNGJdIkVs+UwJC8ZNUt7LWZQb+466oPJEagCJdHmDeBLYuD8lcNOSKcnW3b0UyyZsYagCGDE4GpYAKbTQqRphtS/jxYTBICZkvpZHZDNWn1tjucSyV9YEgrGY0B1DMNvTVtQrsNiAKZQ4QZsEMa1Fr5bTwyzVlqPLelex+fzEiT8wNh+S/PttCoCxfNqkR7KzuApg1rZMSZP2bb77DoiHeyCad3v/pC8b9Xp6KVtQx9+KjM5LV64kIfaQtTfjXWXKaD16hhFkQ9lr5l7PCMZXYdOxWsQbgcIpo+xGj36Tcsl+ghKDWX0TKmW6XuY9ZC0PWOAX2jovc5xX7WPXhUFMGA3aZqr24x1fDVHLaihG8PuhpwJYj4PH/8bQzauGyh4ok0GpFeeJDlVeIJiD1s5GdQSGB4GFTOAeTSeLHnb72WBZnfp48X2w+EU3toZlyH5K3REMTqilCNQIMkOJO9zl/SrpFAWMDKahZPmaOLSkNwapAK2Zo1112l9y8WiBcwKGfCvS1wCqKonc4PzCkBD5NQGy/qifeH1LFXH8O7oTsueEZJYtJUbb0ey9uh0iarymqn3XBjk6ZwkcHq+ppETvTjann98DLt6a7Xol6pjgNlDoFnFgNZJWOSQwkzobINONSnO3lCKEytoDUpbZyQcqXKctOxyQcpl3vx1GrVX+MCZ93iIPwRNNH1LAR5LId24tnh7buBTsGnhYzdsx30KLpCh9UoXoeyqbpt0ccNnnZANFB+MMppK4DtpolBDJnpuxAdjhcaVpvWhKaayOAQXV/1TBsXgmsItkdUT+f8EKflhhIp5mhA2SC1DD89LVtWmwrWCpcDEIQRMvlb+kiocQWQeVUYFvXdGMC5KaY9hMRb6HREWH2HSo8+rWXI1MhVVFQXAlvRhyNhQFHGw2i6O4M4ssHn6ZkOtb0RUII/jWMEbOphPf8gTwFZyDOKgqLYakN/DK7ZSMt5ez+ySvVxsXnjsC+EiW8YVMgogsD66GFR8UoUoJSGdG4NMnCX3xHK1sY1NevkipzguSU3fCW4LYgxUa6DLKIIz07pf7AQRZ2bZOmJTSMZcKc1d+CQAGkIG+H3hZfXwxwW5CPHAW+a/J1Qxyi4CYE7+dWdqJ4EFfsMwJDaFhgo56hT7Yn0yqyePJh0YTwBlOduMaRLLax7PEdOT7PJt0zpAaZ2V7lbIEuxOq19oWT5tkJIJvl26eLSVxOEakcYYjoGtFzO8tIr/yBvmL9TbtNNmYh2CibaL5E5g48DO6lVK+RvoXpbpXD7cjEKJOfSj6jTjNR9LmXaCuF9Cz1Wn7Zwbq4IFmEgdQ2aA3xR40BU2LQTUYm8NCVRwDVEoxJ36rctjlTbTH72OMY9MMGmOfwuipvrMwT5O6WFj6NJRnsgx/06O4u2pT9SurkF8USgrrVNMYYWtHROYP6TrLa7X5HwBsNyBSr9ydiYNbRRlvaylDL68jQyCU0fTduG0g3Z7G0P8Cuc/TXQXTxFdAmQGZWjbaC7oNvF2T9bYiksDMfN3I8u+E1Xx3R1U07yqYwA7da1MMRxv8pZ9ShttjSqDk3mKsyU3hf72o3urR0wRrlF0DEATQMrYkdenGMRjNc2aXy6haTlCWYW1QFQJ8ntVbAAOL7ehbJMuglfYvc5LYYC4eFOnrHac7FbmYEBhuJxb7VhhMsvrOrAoaF+9oWG331nEIaw9tbrcndPtghPxcdlDT0FBKloaxSUyXvWfp/ztBkXTW7MvVjmOJyJNYZOe8jAODRmjNR7Z5Kdp6WL3Pxe3BMdof+DiK1PZeeNw1PZOVNf2UE5IZLdk1rp3FJEm66KHY2krUWk218iw0OaTuuJFovjtL9Ey8Ii43AJsAPAAQGetm5hW1AjRfsFrZe2g1bLjSezFMGQqfNlpGjwhed0TitqBer0XTLjAMVDRIQ3a5hVwsZQiF+rZ6BNgwthZSao/B5sUpzHH+MZx9Ha2RqLPVVX+e4rk6KgPImX7i0XgjSnnAA5VJUybq0QFpOj0bR/67mmwdsv1DY7CLhWw/A7KlPv4dfZRr5jVs2Tvlw47yEzjCSDbqY856XLQwlndGHZLrsZBtWok7oOEMfzAMpUKA3nEA/E+Zt9s+hDUyZoOaxrm1azmQ9Q/ooIMiPKdNMGemnzs01rZq9RICNWYUJpVqRZA68Ew3gkD5lXpOZaylwFp4RT4FZgkm8zJ9gWt+cP0AjkMbDZZgMTrGgp+eWVbpYmu5PQRjFd0pJI2mLJXpZtpCtu2fCHAlxy7w7KcXjMWD+lPTXrTlN1kVHLspZfgr49HTjTz6ER0KGGK0l9wwchfbGL6ipey5v7/Ji9gG2CanUfaRTjwzG+UpeqCmKAahot4V4SdD5PT1BWQ7dvGBEAawiqd5Z72Tuup/QbzCvzXJ02DHmk0kZbo2+FRtuhE+q0lFXhjqKbQRn6N4Ju9rZOhUvGaOtE2njXLchqkg3eFpo4J7HyS1idA7LynEwOG9ko/PLEd2BJoqWXKKLd/14jJh5NGsu0VuFudg0gdUeoVqyUv2b8bhlxt3EjYuryye7IeKiQ+lorqFdWdz/EFYEYmihdYHTNc7sbTAxbGpCXpe7qZEfGbkexM+8LLOqPVNV1uyT+tSjVqsatduBxml7rb6WfAs7/kaoRA8TSb18320eWEyfDs5YvMSZX5+uQPFns6P7is8d0Q0UJDUqpGIdKKysEZbUgwJl4SZQFs4S7+a9LljVg9am+X0zqckHmAvOnvmS6o6qbdNF7LU5jjbBva5OFUYYW5g1Fa0440iWwuw/57g3kBOxMPi6zet64qzi5SYdu/6YNO2A87RGgg4d8vJfIHxen2NLjNgHPTl6WMLf4kwvu1+cUoLo//8RqIr1hAlUSmnR1wVCuWPlNyVmWFWsLXwjPUUaONsSb23BU4W8Kh2FASlr77367f02ALGVuDaGKCzO3C12m0J5y6htGOWTEV8dHTOPKxXmHxpJK2ChNRNDbZ3OpIzdCqKyjqWbQD9AG6aD1qlwjpDKY05MBuOsS/w66Hoz6CZRcXQL+NtlzO5r9J3uMV0jMnIu0mjCKtftPvNTKPJwIUyov9ebjiiwxI/uQZKMOAIYbJ7DPoHZnyUV1oxttyYrC5ojBbJ8zcCnGweY4Ig/OmIHBQVa9eUNVCkVUvqbgWYzRv2aYjTeacKIreTlVVYbNSTikOaYmEcHqJpmz54a1UO2slUWFm7IN095veQyUMSEtTM3eEaOQIr8zY0Lqulg+baAEn/sy7mxeU0wU8YX05LSSR4xsoW/b+RYozcSZ50vy84zGqYi+9qDMAiHmsllBjxAC4XxlnJ0zrlw6wcqMxlOa/bfQabCxsPYe2tI2BKsoh4RTXB7PRm/mlBujNbqYSOGTIPUn5Sjl8HLMnYIM/lwu6ac6rVt3wWrKcIQy9ztr0cKzbDDUCs1QOeQ2t7tsyVVZs8MImG6YvvY4apsGFUL1yItSXkoem2zb+MVODVbZcUdXxZAzMgL6XbC1nQ6arO72QVWdwXjk4Ydwdrbi/vkKkchkR5XMd9pgSV5ICVU/LRYi3Lu/4vrXbuPJx6+C19UFJ/Kpoi5VScyXsHci2tDJ1DOgerqZ1WkmL6/Q3Y8GAlSmscONSSWB3PvFToJFGQPO16WFerRIHmr/iofkgkM3doXF4Ggng9h6pO/uWVQ4O4S8PDvxgopTNVGfDpLKX6SsM+0lo/+c8akAtexYrGTL4+REfPOO85wVd3Ah5cqmbp//kq5oNlvUeQc5IdP5PeW2a4yApKiBwVYAmEkh8Lri6acew1ev38Sde/exyDRj3TAD2k6D3fQGdWCsK+GVL/0P3vbmJ/HQibIuZJLBnPuX+T4VDW7zEckI5MGN3vpH7bq9ziYUlS5u2VSzs3rW219ATHIXS3qnF/1ERs6EGx63pSd57PStlUDBB4WKkCgSAqsJpW7UirSeZVYiBU1/2T8hMjI75EvoIqDsD4H8jZ1Ws0tPMwlNXRa/VwSupimclUZXOIk4Gy/Hf7Ov32iGiuRZg3QSzSie6HB9SQpNDS+jSCeNzzCAk4XxrW9+Ap/9z1dxP+9T2g7l9SnePdn+5PQUL//bf+GZJx/DN73hCu6dJatTylGKVQQmm7W66w9lC68t6xS9KkGOBIyOeXeUo5qh16fAY5lixId2tttWC1zM/AZxz1S9hvRtmUGQLx9KeE3BCod+LmjLKwaL5xR5peq9zULKqH5IdNkAIkeEefNXU0Z+T5o5AGyUxJ0CJ6DhMVenK2spaqPAWKzCH23lUcc20k01iAQQ1cHsglui4g77W9ByyO7/2lgMh7LoOBdPxFjLoJ60b0muuocEwtnZPTz9xit49pnH8fGX/wOnpw+JW5wEvyvwBqwr4/KlU3zm81/F9Rt38O7vfhZ3bt/FsuRjDCRyVQWwCVdXMJ+jTr+kOtLcc5wdwmWrZJNUS7sSEM/wpItOoR2zS6OSyIIWOg9a4LX3AEp/Ov8VJ9DdZkiI8tudVaG1dOhQVddWBLsWEpIuBJkJJ6psRjICLN/klW/TbUdAHREX3FubjagifXWcQmmFrPNUSCrt8tdyCaxGurRia2dR9E/h4PwXnATVrYNqJw2mLuocQHKVt3uRNaCH1k5OFty8dQcvff/zuH7zLj7xmS/gyuVTrOs+q7tjW/D0elkIt84Yf/K3/4z3vvg8nnr9Zdy7f26+hvSDTUYdb0iMkP3Jff9sTIXjXdWqxntIgmQwZK619dSMYKMRNYAyv7462qfQTe81JEac4s67/N4rizEO1ZM3jlSVUPgNqC6tTy33tk6s8zR7rWeFMY0z0aibBr/SUWldoffxq2RYmuroeLa5mb4ibgzoT8SmLs3UbFcXH9KbvEG9zSI34WeUH2jEFMmYEfJY3Nk5nnr9Zfz0+74XH/roy7hxd8Wy7A3o/clAfRdRrtZ1xaNXLuMvPvEK/vvaLfz8j72Au7fvgOikeASj8ORwuVA3hTxrFi2qDNkSCuW0Ui59WINXwLVacuPpU1mlSVj+RHDSA/aNM1Ct+L5uv6VXkbJEKKTo3ALvGXV5bBLUa1WmdElqrZTHj/qeJerUW1lrGiLjGRLcwrSz8rz0irSFi1FrbBexJYfgjUuPdndvZKmXR7/RPE5/1GvzInu2FgBhWU5w+/ZN/MrPvQfXb9zFh//y03jd1UeK95+PAaa6AG1IQwSsOMXvfOjv8X3veAt+8qXvxGtfu4HlZKnjAYW5NjIflcNYYVZLdYgev+kxFMU4BL2xnTBVoQ7o8DopEJWlb8fgrXUqITKvxc/UsfFB/dumDmze5KYeM5N/JqrbqlcnE7lfj6WERtoCV8XfqzpNuRMQjyF10gY3tBCWkwVfvXYNH3z/O/HSu96OX//dP8M9XrC43u9sKROHg3rOJgG7cvkhfO5LN/Gbf/Ax/MyPvoCfeuk7cOPGDZytK5aTpYyylkG4aarcd4lhw1avXXN26lO65noK76g1ohBukHwif+sY9wqcaxfdX+5FBsXdyTfiQZlb9WsiDJfX0HKEMVP12QhKt1lIQDMhaS9pjsesn+2SbVe+abcoSf1kuCyE+2eMa9ev44PvfwG/+LM/hF/9rT/Gp/79K3j0ymWsvB7EcXrmLc9OSqHukKSBoJOTBV+7dQfvevs345c/8AP4+Mufw+9/+B/w6rW7ePjhy3jooZNg2u1kWaP20ghl+SUDHHp/15Eyz/w7/9yFmXmFpHQceqlrsGYpZ80/Qol2dNDdryjUtuxuApZEWSp4OiaW6BQ//+4BFruVuHJUhfqqq9gkHnC/yhUga1PMykACZCdrjYGLdFj8XuIiyYTJC6zMuHd2jlu3buOJxy7hlz7wbrzvPd+FX/vtj+Ajf/2veOzqozg/Pz8oliWiPQYgQgAsy4Kbt+7ibc9cxS/8+Dvx+GNX8Kd/8y/42D+9gi9fv43z8wXLCRknHM+eY3sbsUqHo1niiYGV9AYZKkm+4jxyLUpZGyIyAhZsClLhOqwwhFhUblHMyN4kSTK5+pTYFGlTD1bs6dQpQhppGasfVSQxyiQgOdMhkS60WzNIBuGGOhNgz2IYp4/M9uimKmSPFsscsz8CKi9q71a314hWV15Z169kwWXXLczrioUYTz/xCH74xW/DT7z3e/DazXv4jd/7c3zys19Oyr/qsa19EBiAnWY912RZFty5ew8ndI4ffOGt+JEXn8ejVy7jC6++hle+eA1fuX4Ld87uF8VN+i1+dMklZ9NA2nNqeur8vyQrpPZe4IKJF9QpzSZnqXUJMEpzkGskRv0+XNJm0VfGR2I4WatQvjCUyTpakFKehaiWK48pGRb5JEqFLkefors+Tt+EhV7DP8FX+OTqCokmFlVXlZdrmspp8WWdWInrhTXfWQm5DvdqI8NgYNWRTTZquuMs5+pp+stbBq9sCCmGKpchC9m40LkCTFgh81W0t1ZmThk4VrLBJRqrJEpNxdUVv8Ta+VHhCRTJUKU/fOkSnnriKp570xvx1mcex807Z/jDj76MP/qrT+HuSnjdI1dwfn7cwbDJADz7HIefU/YiWwgrE+7cuoMrl4Bvf/ZJvOP5p/EtTz+ON1x9GCcnp0r5uBgCogWyq1XyrktR9LUIeCU4/eZr8X7qG3tRHFG0jGMRDyz4cz7KK8LqFxqqU/SJalmkvF6hiRK9C4saJrEh2Qo959PKmidM1H9I5x4qw1h+qeZfKJeV0dZ7wrLovEv5JLssJ7V8qftCyQhlXi7KIFL5lit5SNGThxKThU3tIzzQi3fKUW91lVqSLxEzLoomyl0+1WWFN19hVlGiFatortI4WQVHSOGy5ElJ1Z5/nO9ZDDXAfJ4Mu4y4Z62uE2640GO6a2ZAMZFIQq+wQ31GrvVN6VYxJLkshtCFYhDv319x7bWb+NwXXsXfffLz+MdPfxE379zH1atXsCwL1nVr4X6peTfdshDoTW95juuwm/a886BzLgvhfAXu3ruH+/fPcboQLp0uWSG0z3BAzUV9xSiNUx9Wb8loI0he1XkEAWgLbCfB+Gs9fq5eeXJLnCgCpmcq2gzmTuoGlO4F8zavzNIGnUzCABcB6Q4QQ6eNilEeLEcH2q+ZVKq8ahQzbVwDCQqkLAIxiOI2fbOP8qGten1ZdqiqFIT1ye1RopwyU7BNTI6RHh+XKLdKUGkmdk2Wr8TAn5+vuHd2jvsr4/T0FA9fvoSTZcGajdhxkGUhRQBvTfInFq509tgQuReW7L2Ga5U3IbOmKOkWjsiAJWEyE++UsppWiKDzvjatmkZLmt5OXQRna2fGmrEBIrxEbr2/L8/Xydcvom0EEe6tpuo2E5dl2y3eYiVjuqZ4ri1TYw4djlHlqeDYrpijI8DVrjFJTynPsGWwmeHXULYlw72SF8JpdbyyIoyKFzpMaROs+vvcJG0tm1hl5k4ajYBNlhKE5cZOi2bamXlDGgaEl7X13kUQ6q/SwnKrZLzUTUUrXYe8AXW1pCODnGfUdfL167wzJIwMljxbYGgOgw1XFtVBnYAuJUsUsCNsJ0eAbx+4MNqMyg3aQ5zHqLyJ5xJZ2dLqc14Z55188YOeAeqTZWfuuwafm8BzcTBDNjcXMYK6aHQBIa0BT46/s9dgNRe7iB06Sfa3iaZWoMilcwkOscP+q8KRYaMxm1L3Ec6m7uHjbvoZXHtQbCVo1dC/64TDR/A1zHpE1J1y7svdnwmYo4CjaHmQMCRNW+/MFPYLU8glHVvPvtXdU1WCihu2k85AN7Lc2XYblRg5e4ODg+u9EEXQ0fMtejzOYTQ3QVOn6zJdz5l0e1WuCRX3rQdotwUXvMe6jVFFmsagEhbXkaPmg8kkr6nVafMm8ow2Pk537e4tMpCVrpHn8O83kgkvobdP225kh5HR4mlgJ+JR12Ijk+0adPgS4ZwlsZMm7E+H1s6FxxO4uzRcRHuFZTPsCVwOgtenQk9K0N+t5GCIKqt5KR06SVdGgPXX+2oQZgvVJ6yIwiUOJANTjItaQWdIFOVW9OveAh/EKN8H/T+GaEzAwDahx8typ+t1pGLPQjM78OvRNg+yDD/GpWEBFtfFSBGA8vYX0V+0BI1fi8PfyqMjgS1oRIrdMz2Li0rA4fKjGkOltyVgmKLCJ8xUmG+WFwCHCtQBvYRDnd88SNfMNcgDc0oHIr5omr4eEIwvLO0Sz28ACMiZpTIyEfk7Z5hWT9oIaWAp282/34Qty7cD1TcIfF3kPfJAF13wReD7/6b8HeiOAczBRZjBOAS3rrbbMYMXlkOoicacrBz6kS1F16jPFcKeUaM4v4+YLmQsYRrGbX6URDAuNjKaK/AB4A3a+BsoYtCk7N4T0MLF18gOxalymj43Ben2CQ/JdFs1ECXjBWUWYVPFiK6dQJ7ueVgCBRnNeLx4GJdV3x4wNO63FLtg0FF/4r/eefkAPKPRtn6mw+AB9RjnOH5Q4XrHmbm0DFLnnUle328HTN/doFFVCwxwUyon7ynrAtihkMC/i5PRvPO1DkctwhEv1C8v+tekzQersc8rR60FBbjooBdsRR2kUdsVEj0PSuQ2UCwVUNmFWnFiQcuw/OwGPux+M3KZot28iu70uBGj0/0byHYgG/2sAU1RWKouDHqzk/EGXfI61+807SOWB6VYitiaBRh43bJW3uXNglmWzubPX+MydOjv7svXAdsgqdGECX3a64zifCFeh2HNITOw5ImYg+OVdgX02fPUVWT+80INz0hd18xRHGkp8DRut+UU2RXDVIWpfEVJ7VIEC7KpqF61aLpbi5yfoB8OKZqgHpDFXLWR83o/7hhpUwrVYE81QVlL4igy+Xw9DOlccPeqQ3mGblUiLR36fgVw0qVfeK5ln4jwv4kOEb6XH10sAAAAAElFTkSuQmCC'


def app_icon_path():
    """يكتب الأيقونة في ملف مؤقت (أو يستخدم app.ico المجاور) ويعيد مسارها."""
    cached = getattr(app_icon_path, "_cached", "unset")
    if cached != "unset":
        return cached
    path = None
    here = ""
    try:
        here = os.path.dirname(os.path.abspath(__file__))
    except Exception:
        here = ""
    for cand in (os.path.join(here, "app.ico") if here else "",
                 os.path.join(getattr(sys, "_MEIPASS", "") or "", "app.ico")):
        try:
            if cand and os.path.isfile(cand) and os.path.getsize(cand) > 1500:
                path = cand
                break
        except Exception:
            continue
    if path is None:
        try:
            base = os.path.join(os.environ.get("LOCALAPPDATA")
                                or tempfile.gettempdir(), "GoldPortfolio")
            os.makedirs(base, exist_ok=True)
            p = os.path.join(base, "app.ico")
            if (not os.path.isfile(p)) or os.path.getsize(p) < 1500:
                with open(p, "wb") as f:
                    f.write(base64.b64decode(APP_ICON_ICO_B64))
            path = p
        except Exception:
            path = None
    app_icon_path._cached = path
    return path


class ScrollArea(tk.Frame):
    """منطقة تمرير عمودي: Canvas + إطار داخلي + شريط يظهر عند الحاجة فقط.
    الهدف: تظهر كل محتويات الصفحة مهما كان ارتفاع الشاشة، بلا أي قصّ.
    التمرير بمقدار ثابت بالبكسل (سلس) ويعمل بالعجلة وبلوحة المفاتيح."""

    def __init__(self, parent, bg="#10141b", bar_side="left", step=24):
        super().__init__(parent, bg=bg)
        self._bar_on = True
        try:
            sc = float(self.tk.call("tk", "scaling")) / 1.3333
        except Exception:
            sc = 1.0
        self._step = max(14, int(round(step * max(0.75, sc))))
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0,
                                takefocus=1, yscrollincrement=self._step)
        self.vbar = ttk.Scrollbar(self, orient="vertical",
                                  command=self.canvas.yview)
        try:
            self.canvas.configure(yscrollcommand=self._on_scroll)
        except Exception:
            pass
        self.vbar.pack(side=bar_side, fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.inner = tk.Frame(self.canvas, bg=bg)
        self._win = self.canvas.create_window((0, 0), window=self.inner,
                                              anchor="nw")
        self.inner.bind("<Configure>", self._on_inner)
        self.canvas.bind("<Configure>", self._on_canvas)
        self._bind_keys()

    def _bind_keys(self):
        """لوحة المفاتيح بعد النقر على المساحة الفارغة من الصفحة."""
        pairs = (("<Prior>", lambda e: self.scroll_pages(-1)),
                 ("<Next>", lambda e: self.scroll_pages(1)),
                 ("<Home>", lambda e: self.to_top()),
                 ("<End>", lambda e: self.to_bottom()),
                 ("<Up>", lambda e: self.scroll_units(-1)),
                 ("<Down>", lambda e: self.scroll_units(1)))
        for seq, fn in pairs:
            try:
                self.canvas.bind(seq, fn)
            except Exception:
                pass
        try:
            self.inner.bind("<Button-1>",
                            lambda e: self.canvas.focus_set(), add="+")
        except Exception:
            pass

    # ---------------- داخلي ----------------
    def _on_inner(self, _e=None):
        try:
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        except Exception:
            pass
        self._autohide()

    def _on_canvas(self, event=None):
        try:
            if event is not None:
                self.canvas.itemconfigure(self._win, width=max(1, event.width))
        except Exception:
            pass
        self._autohide()

    def _on_scroll(self, first, last):
        try:
            self.vbar.set(first, last)
        except Exception:
            pass
        try:
            need = not (float(first) <= 0.001 and float(last) >= 0.999)
        except Exception:
            need = True
        self._toggle_bar(need)

    def _toggle_bar(self, need):
        if need == self._bar_on:
            return
        self._bar_on = need
        try:
            if need:
                self.vbar.pack(side="left", fill="y", before=self.canvas)
            else:
                self.vbar.pack_forget()
        except Exception:
            pass

    def _autohide(self):
        try:
            y0, y1 = self.canvas.yview()
            need = (y1 - y0) < 0.999
        except Exception:
            return
        self._toggle_bar(need)

    # ---------------- واجهة ----------------
    def scroll_units(self, n, horizontal=False):
        try:
            if horizontal:
                self.canvas.xview_scroll(int(n), "units")
            else:
                self.canvas.yview_scroll(int(n), "units")
            return "break"
        except Exception:
            return None

    def scroll_pages(self, n):
        try:
            self.canvas.yview_scroll(int(n), "pages")
        except Exception:
            pass
        return "break"

    def to_top(self):
        try:
            self.canvas.yview_moveto(0)
        except Exception:
            pass
        return "break"

    def to_bottom(self):
        try:
            self.canvas.yview_moveto(1.0)
        except Exception:
            pass
        return "break"

    def refresh(self):
        self._on_inner()


class PriceWidget:
    """ويدجيت عائم صغير (دائماً في المقدمة) يعرض جنباً إلى جنب:
    سعر شراء الغرام بالريال السعودي | سعر الشراء بالدينار الأردني |
    القيمة الحالية للمحفظة. يُسحب بالماوس، يحفظ موضعه، ويُغلق بـ ✕ أو F9."""

    WIDTH = 380

    def __init__(self, app):
        self.app = app
        t = app.theme
        self.card = t.get("card", "#12161e")
        self.accent = t.get("accent", "#e9c684")
        self.text = t.get("text", "#e8edf7")
        self.muted = t.get("muted", "#9ba6ba")
        self.profit = t.get("profit", "#2ecc71")
        self.loss = t.get("loss", "#ff5c5c")
        self.border = t.get("border", "#252c39")
        self._dx = self._dy = 0
        try:
            self.width = app._px(self.WIDTH)
        except Exception:
            self.width = self.WIDTH

        self.win = tk.Toplevel(app.root)
        try:
            self.win.overrideredirect(True)
        except Exception:
            pass
        try:
            self.win.attributes("-topmost", True)
        except Exception:
            pass
        try:
            self.win.configure(bg=self.accent)
        except Exception:
            pass
        self._build()
        try:
            self.win.update_idletasks()
            h = int(self.win.winfo_reqheight() or 76)
            self.win.geometry("{}x{}".format(self.width, max(h, app._px(72))))
        except Exception:
            pass
        self._restore_geo()
        self._bind_drag()
        self.refresh()

    # ---------------- البناء ----------------
    def _mini(self, parent, text, cmd, size=10):
        return tk.Button(parent, text=text, command=cmd, bg=self.card,
                         fg=self.muted, activebackground=self.border,
                         activeforeground=self.accent, relief="flat", bd=0,
                         padx=6, pady=0, cursor="hand2",
                         font=("Segoe UI", size, "bold"))

    def _build(self):
        """شريط صغير: البيانات الثلاثة بجانب بعضها (أعمدة) لا فوق بعضها."""
        px = self.app._px
        outer = tk.Frame(self.win, bg=self.accent)
        outer.pack(fill="both", expand=True, padx=1, pady=1)
        body = tk.Frame(outer, bg=self.card)
        body.pack(fill="both", expand=True, padx=px(8), pady=px(4))
        self.body = body

        # -------- شريط علوي مضغوط: العنوان + العيار + تحديث + إغلاق --------
        head = tk.Frame(body, bg=self.card)
        head.pack(fill="x")
        tk.Label(head, text="الراجحي", bg=self.card, fg=self.accent,
                 font=("Segoe UI", 10, "bold")).pack(side="right")
        self._mini(head, "✕", self.app.close_widget).pack(side="left")
        self._mini(head, "⟳", self.app.refresh_all).pack(
            side="left", padx=(0, px(2)))
        self._mini(head, "+", lambda: self.app.widget_karat_step(1)).pack(
            side="left")
        self._mini(head, "−", lambda: self.app.widget_karat_step(-1)).pack(
            side="left", padx=(px(3), 0))
        self.lbl_karat = tk.Label(head, text="24", bg=self.card,
                                  fg=self.accent,
                                  font=("Segoe UI", 10, "bold"))
        self.lbl_karat.pack(side="left", padx=px(2))
        tk.Label(head, text="عيار", bg=self.card, fg=self.muted,
                 font=("Segoe UI", 10)).pack(side="left")

        tk.Frame(body, bg=self.border, height=1).pack(fill="x",
                                                     pady=(px(2), px(3)))

        # Two rows retain legible values without crowding four columns.
        strip=tk.Frame(body,bg=self.card);strip.pack(fill="x")
        self.v_sar=self._cell(strip,"شراء (ر.س)")
        self._vsep(strip)
        self.v_jod=self._cell(strip,"شراء (د.أ)")
        tk.Frame(body,bg=self.border,height=1).pack(fill="x",pady=px(4))
        portfolio=tk.Frame(body,bg=self.card);portfolio.pack(fill="x")
        self.v_val=self._cell(portfolio,"القيمة الحالية",color=self.accent,big=True)
        self._vsep(portfolio)
        self.v_pl=self._cell(portfolio,"الربح/الخسارة غير المحققة",big=True)
        self.cols=strip

    def _cell(self, parent, title, color=None, big=False):
        """عمود صغير: العنوان فوق القيمة، بداخل الشريط الأفقي."""
        px = self.app._px
        fr = tk.Frame(parent, bg=self.card)
        fr.pack(side="right",fill="both",expand=True,padx=px(4),pady=px(1))
        tk.Label(fr, text=title, bg=self.card, fg=self.muted,
                 font=("Segoe UI", 10)).pack()
        val = tk.Label(fr, text="--", bg=self.card, fg=color or self.text,
                       font=("Segoe UI", 13 if big else 12, "bold"))
        val.pack()
        return val

    def _vsep(self, parent):
        """فاصل رأسي رفيع بين الأعمدة."""
        px = self.app._px
        tk.Frame(parent, bg=self.border, width=1).pack(side="right",
                                                       fill="y",
                                                       padx=px(2))

    # ---------------- السحب وحفظ الموضع ----------------
    def _bind_drag(self):
        def walk(w):
            try:
                kids = w.winfo_children()
            except Exception:
                kids = []
            for c in kids:
                if isinstance(c, tk.Button):
                    continue
                for seq, fn in (("<ButtonPress-1>", self._drag_start),
                                ("<B1-Motion>", self._drag_move)):
                    try:
                        c.bind(seq, fn, add="+")
                    except Exception:
                        pass
                walk(c)
        walk(self.win)

    def _drag_start(self, e):
        try:
            self._dx = e.x_root - self.win.winfo_x()
            self._dy = e.y_root - self.win.winfo_y()
        except Exception:
            self._dx = self._dy = 0

    def _drag_move(self, e):
        try:
            x = e.x_root - self._dx
            y = e.y_root - self._dy
            sw = self.win.winfo_screenwidth()
            sh = self.win.winfo_screenheight()
            x = max(0,min(int(x),max(0,sw-self.win.winfo_width())))
            y = max(0,min(int(y),max(0,sh-self.win.winfo_height())))
            self.win.geometry("+{}+{}".format(x, y))
            self.app._widget_geo = "+{}+{}".format(x, y)
        except Exception:
            pass

    def _restore_geo(self):
        geo = str(getattr(self.app, "_widget_geo", "") or "")
        m = re.match(r"^([+-]?\d+)([+-]\d+)$", geo)
        if m:
            try:
                x, y = int(m.group(1)), int(m.group(2))
                sw = self.win.winfo_screenwidth()
                sh = self.win.winfo_screenheight()
                x = max(0,min(x,max(0,sw-self.win.winfo_width())))
                y = max(0,min(y,max(0,sh-self.win.winfo_height())))
                self.win.geometry("+{}+{}".format(x, y))
                return
            except Exception:
                pass
        try:
            sw = self.win.winfo_screenwidth()
            self.win.geometry("+{}+{}".format(
                max(10, sw - self.width - self.app._px(26)), self.app._px(72)))
        except Exception:
            pass

    # ---------------- التحديث ----------------
    def refresh(self):
        """يعرض: سعر الشراء بالريال، سعر الشراء بالدينار، والقيمة الحالية."""
        hide = self.app._hide
        try:
            d = self.app.widget_prices()
        except Exception:
            d = None
        if not d:
            self.v_sar.config(text="--", fg=self.text)
            self.v_jod.config(text="--", fg=self.text)
            self.v_val.config(text="--")
            self.v_pl.config(text="--",fg=self.text)
            self.lbl_karat.config(text=str(getattr(self.app, "karat", 24)))
            return
        delta=float(getattr(self.app,"_delta_oz",0) or 0)
        arrow="↑" if delta>1e-9 else "↓" if delta< -1e-9 else ""
        private=getattr(self.app,"privacy",False)
        color=self.text if private else self.profit if arrow=="↑" else self.loss if arrow=="↓" else self.text
        for key,lbl in (("sar",self.v_sar),("jod",self.v_jod)):
            value="{} {:,.3f}".format(arrow,d[key]["buy"]).strip()
            lbl.config(text="•••" if private else hide(value),fg=color)
        try:
            sym=CUR[self.app.cur]["sym"]
            val=self.app.portfolio_value_disp()
            pl=self.app.portfolio_profit_disp()
            self.v_val.config(text="•••" if private else hide("{}{:,.2f}".format(sym,val)))
            self.v_pl.config(text="•••" if private else hide("{}{:+,.2f}".format(sym,pl)),
                             fg=self.text if private else self.profit if pl>1e-9 else self.loss if pl< -1e-9 else self.text)
        except Exception:
            self.v_val.config(text="--")
            self.v_pl.config(text="--",fg=self.text)
        try:
            self.lbl_karat.config(text=str(getattr(self.app, "karat", 24)))
        except Exception:
            pass

    def copy_prices(self):
        d = self.app.widget_prices()
        if not d:
            return
        txt = ("سعر الغرام — عيار {k}\n"
               "الريال السعودي (الراجحي): بيع {s1:,.3f} / شراء {s2:,.3f}\n"
               "الدينار الأردني (الراجحي): بيع {j1:,.3f} / شراء {j2:,.3f}\n"
               "الأونصة: ${oz:,.2f}   ·   {t}").format(
            k=getattr(self.app, "karat", 24), t=self.app.last_update or "",
            s1=d["sar"]["sell"], s2=d["sar"]["buy"],
            j1=d["jod"]["sell"], j2=d["jod"]["buy"], oz=d["oz_usd"])
        try:
            self.win.clipboard_clear()
            self.win.clipboard_append(txt)
        except Exception:
            pass
        try:
            self.lbl_src.config(text="✔ نُسخ")
            self.app.root.after(1500, lambda: self.lbl_src.config(text=""))
        except Exception:
            pass

    # ---------------- الإغلاق ----------------
    def destroy(self):
        try:
            geo = str(self.win.geometry() or "")
            m = re.search(r"([+-]?\d+)([+-]\d+)$", geo)
            if m:
                self.app._widget_geo = "+{}+{}".format(int(m.group(1)),
                                                        int(m.group(2)))
        except Exception:
            pass
        try:
            self.app.price_widget = None
        except Exception:
            pass
        try:
            self.win.destroy()
        except Exception:
            pass



def _ps_quote(text):
    """تهيئة نص للاستخدام داخل أمر PowerShell بين علامتي اقتباس مفردتين."""
    return str(text).replace("'", "''").replace("\r", " ").replace("\n", " ")


def _win_toast(title, message):
    """إشعار ويندوز الأصلي (Toast) عبر PowerShell وبدون أي نافذة ظاهرة."""
    ps = (
        "$ErrorActionPreference='Stop';"
        "[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType=WindowsRuntime] > $null;"
        "$t=[Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent("
        "[Windows.UI.Notifications.ToastTemplateType]::ToastText02);"
        "$x=$t.GetElementsByTagName('text');"
        "$x.Item(0).AppendChild($t.CreateTextNode('{0}')) > $null;"
        "$x.Item(1).AppendChild($t.CreateTextNode('{1}')) > $null;"
        "$n=[Windows.UI.Notifications.ToastNotification]::new($t);"
        "$id='{{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}}\\WindowsPowerShell\\v1.0\\powershell.exe';"
        "[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($id).Show($n);"
    )
    try:
        ps = ps.format(_ps_quote(title), _ps_quote(message))
        flags = 0x08000000 if os.name == "nt" else 0
        subprocess.run(["powershell", "-NoProfile", "-NonInteractive",
                        "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden",
                        "-Command", ps],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                       creationflags=flags, timeout=15)
        return True
    except Exception:
        return False


def desktop_notify(title, message, app=None):
    """إشعار سطح المكتب: Toast على ويندوز (في الخلفية بلا تجميد) +
    بطاقة داخل التطبيق دائماً حتى لا يضيع التنبيه."""
    if app is not None:
        try:
            app.toast(title, message)
        except Exception:
            pass
    if os.name != "nt":
        return False
    try:
        threading.Thread(target=_win_toast, args=(title, message),
                         daemon=True).start()
        return True
    except Exception:
        return False


class PageBody(tk.Frame):
    """إطار محتوى صفحة داخل النافذة، يحاكي واجهة النافذة المنفصلة (Toplevel)
    حتى تعمل النوافذ القديمة كصفحات قابلة للتنقل."""

    def __init__(self, parent, app, key, title_lbl):
        super().__init__(parent, bg=app.theme["bg"])
        self.app = app
        self.page_key = key
        self._title_lbl = title_lbl
        self._back_cb = None

    # ----- واجهة تشبه Toplevel -----
    def title(self, *a):
        if a and self._title_lbl is not None:
            try:
                self._title_lbl.config(text=str(a[0]))
            except Exception:
                pass

    def geometry(self, *a):
        pass

    def transient(self, *a):
        pass

    def grab_set(self, *a):
        pass

    def grab_release(self, *a):
        pass

    def resizable(self, *a):
        pass

    def attributes(self, *a, **kw):
        pass

    def iconbitmap(self, *a, **kw):
        pass

    def protocol(self, name=None, fn=None):
        if fn is not None:
            self._back_cb = fn
        return self._back_cb

    def deiconify(self):
        self.app._show_page(self.page_key)

    def lift(self):
        self.app._show_page(self.page_key)

    def focus_set(self):
        pass

    def go_back(self):
        """زر الرجوع: ينفّذ معالج الإغلاق إن وُجد ثم يعود للرئيسية."""
        if callable(self._back_cb):
            self._back_cb()
        else:
            self.app._close_page(self.page_key)

    def destroy(self):
        if getattr(self.app, "_closing_page", False):
            return tk.Frame.destroy(self)
        self.app._close_page(self.page_key)


class GoldApp:
    def __init__(self, root):
        self.root = root
        self.spot_oz_usd = 0.0
        self.rates = dict(DEFAULT_RATES)
        self.cur = "USD"
        self.karat = 24             # العيار الافتراضي للعرض والاحتماب
        self.chart_range = "live"    # النطاق الزمني للمخطط: hour / day / week
        self.chart_metric = "gram"
        self.chart_currency = "USD"
        self.chart_price_mode = "market"
        self.chart_font_size = 12
        self.chart_viewport = None
        self.last_price_source = ""
        self._undo_snapshot = None
        self._connection_state = "waiting"
        self._live_chart_views = []
        self._chart_widgets = []
        self._pages = {}            # الصفحات الداخلية المفتوحة (زكاة، سجل، تنبيه، إعدادات، بيع)
        self._closing_page = False
        self.adjust = 5.22          # فرق الراجحي للبيع (يُطرح من سعر الغرام)
        self.buy_margin = 0.0       # فرق الراجحي للشراء (يُضاف لسعر الغرام)
        self.holdings = []          # عمليات الشراء المفتوحة (بالدولار داخلياً)
        self.sells = []             # عمليات البيع المحققة (سعر بيع مثبّت)
        self.show_sells = True      # إظهار/إخفاء صفوف البيع في جدول العمليات
        self.history = []           # سجل كل العمليات
        self.price_hist = []        # عينات السعر [epoch, oz_usd]
        self.alerts = []
        self.theme = dict(DEFAULT_THEME)
        self.titles = dict(DEFAULT_TITLES)
        self.sound_on = True
        self.source_mode = "auto"
        self.last_update = ""
        self.dark_mode = True
        self.theme = dict(DEFAULT_THEME)
        self.lock_minutes = 10
        self.locked = False
        self._delta_oz = 0.0
        self._ticks = 0
        self._fails = 0
        self._fx_last = 0.0
        self._fetch_lock = threading.Lock()
        self._fetch_busy = False
        self._fetch_results = queue.Queue()
        self._fetch_generation = 0
        self._fetch_serial = 0
        self._result_job = None
        self._price_workers = {}
        self._fetch_started = 0.0
        self._live_due = 0.0
        self._live_attempts = 0
        self._live_error = ""
        self._last_quote_source = None
        self._last_quote_timestamp = None
        self._fx_busy = False
        self._fx_retry_at = 0.0
        self.al_win = None
        self.row_map = {}
        self.price_widget = None   # الويدجيت العائم لسعر الغرام
        self._widget_on = False     # هل كان الويدجيت مفتوحاً (يُستعاد بعد إعادة البناء)
        self._widget_geo = ""       # موضع الويدجيت المحفوظ (+x+y)
        self._page = "main"
        self.user = None            # المستخدم الحالي بعد تسجيل الدخول
        self.data_file = DATA_FILE  # ملف بيانات المستخدم الحالي
        self.login_frame = None
        self._live_on = False
        # القفل التلقائي عند الخمول (0 = معطّل)
        self.dark_mode = True       # المظهر الداكن افتراضياً (يُحفظ لكل حساب)
        self.privacy = False        # وضع الخصوصية (إخفاء الأرقام) — للجلسة فقط
        self._last_backup = 0.0
        self.lock_minutes = 10
        self.locked = False
        self.lock_frame = None
        self._last_activity = time.monotonic()
        self._idle_job = None
        self._act_bound = False
        self._act_last = 0.0
        self.load()

        self.build_ui()
        if self.spot_oz_usd:
            self.render()
            self._set_update_caption()
        self.show_login()

    # ============================ المزامنة السحابية ============================
    def sync_snapshot(self):
        """البيانات التي تُزامن بين الأجهزة (بلا أسعار السوق اللحظية)."""
        d = {"holdings": getattr(self, "holdings", []),
             "sells": getattr(self, "sells", []),
             "history": getattr(self, "history", []),
             "alerts": getattr(self, "alerts", []),
             "theme": getattr(self, "theme", {}),
             "titles": getattr(self, "titles", {}),
             "cur": self.cur,
             "karat": self.karat,
             "chart_range": getattr(self, "chart_range", "day"),
             "chart_font_size": getattr(self,"chart_font_size",12),
             "adjust": getattr(self, "adjust", 5.22),
             "buy_margin": getattr(self, "buy_margin", 0.0),
             "sound_on": bool(getattr(self, "sound_on", True)),
             "source_mode": getattr(self, "source_mode", "auto"),
             "dark_mode": bool(getattr(self, "dark_mode", True)),
             "lock_minutes": int(getattr(self, "lock_minutes", 10) or 0)}
        try:
            return json.loads(json.dumps(d, ensure_ascii=False, default=str))
        except Exception:
            return d

    def cloud_hash(self):
        try:
            raw = json.dumps(self.sync_snapshot(), ensure_ascii=False, sort_keys=True)
            return hashlib.md5(raw.encode("utf-8")).hexdigest()
        except Exception:
            return ""

    def cloud_conf(self):
        if not hasattr(self, "_cloud_config"):
            self._cloud_config = load_cloud()
        return self._cloud_config

    def cloud_store(self):
        return save_cloud(self.cloud_conf())

    def cloud_server(self):
        """رابط القاعدة ومفتاح الويب: ما في الإعدادات، وإلا المضمّن في التطبيق."""
        conf = self.cloud_conf()
        url = FB_URL.strip().rstrip("/") or (conf.get("url") or "").strip().rstrip("/")
        key = FB_KEY.strip() or (conf.get("key") or "").strip()
        return url, key

    def cloud_builtin(self):
        """هل التطبيق مهيّأ بخادم مزامنة مضمّن؟ (دخول بالبريد فقط)"""
        return bool(FB_URL.strip() and FB_KEY.strip())

    def cloud_account(self, user=None):
        key = user if user is not None else (getattr(self, "user", None) or "")
        if not key:
            return {}
        return dict((self.cloud_conf().get("accounts") or {}).get(key) or {})

    def cloud_update_account(self, user, **kw):
        conf = self.cloud_conf()
        acct = dict((conf.get("accounts") or {}).get(user) or {})
        acct.update(kw)
        conf.setdefault("accounts", {})[user] = acct
        self.cloud_store()
        return acct

    def cloud_ready(self):
        acct = self.cloud_account()
        if self.cloud_provider() == "gmail":
            return bool(acct.get("email") and acct.get("gmail_pw"))
        return bool(acct.get("refresh_token") and self.cloud_server()[0])

    def cloud_provider(self):
        """طريقة المزامنة المرتبطة بهذا المستخدم: gmail أو firebase."""
        acct = self.cloud_account()
        p = acct.get("provider")
        if p:
            return p
        if acct.get("refresh_token"):
            return "firebase"
        if acct.get("gmail_pw"):
            return "gmail"
        return ""

    @staticmethod
    def _pw_wrap(pw):
        return base64.b64encode(
            "".join((pw or "").split()).encode("utf-8")).decode("ascii")

    @staticmethod
    def _pw_unwrap(v):
        try:
            return base64.b64decode(str(v or "").encode("ascii")).decode("utf-8")
        except Exception:
            return ""

    def gmail_creds(self, user=None):
        """(البريد، كلمة مرور التطبيق) للحساب المرتبط."""
        acct = self.cloud_account(user)
        return (acct.get("email", "").strip(), self._pw_unwrap(acct.get("gmail_pw")))

    def gmail_push(self):
        """يحفظ نسخة مشفّرة في بريد المستخدم نفسه."""
        email, pw = self.gmail_creds()
        if not email or not pw:
            return False, "لم يُربط البريد بعد."
        acct = self.cloud_account()
        rev = int(acct.get("rev") or 0) + 1
        payload = {"rev": rev, "updated_ts": time.time(),
                   "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
                   "device": socket.gethostname(),
                   "data": self.sync_snapshot()}
        try:
            blob = seal_sync(payload, pw, email)
            gmail_send(email, pw, blob)
        except CloudError as e:
            return False, "☁ {}".format(e)
        self.cloud_update_account(self.user, rev=rev,
                                  last_sync=payload["updated_at"],
                                  last_hash=self.cloud_hash(),
                                  last_device=payload["device"])
        self._gmail_last = time.time()
        return True, "☁ محفوظة في بريدك — {}".format(payload["updated_at"])

    def gmail_pull(self, user=None):
        """يجلب أحدث نسخة من بريد المستخدم (أو None)."""
        email, pw = self.gmail_creds(user)
        if not email or not pw:
            return None
        blobs = gmail_blobs(email, pw)
        best = None
        for b in blobs:
            try:
                d = open_sync(b, pw, email)
            except CloudError:
                continue
            if isinstance(d, dict) and int(d.get("rev") or 0) >= int((best or {}).get("rev") or 0):
                best = d
        return best

    # ---------- نسخة بملف واحد: أبسط طريقة نقل بين الأجهزة (بلا بريد ولا كلمة مرور) ----------
    def pocket_export(self, parent=None):
        """يحفظ كل بياناتك في ملف واحد صغير (.gold) لنقله بين الأجهزة."""
        par = parent or self.root
        try:
            who = re.sub(r"[^A-Za-z0-9._-]", "_", (self.user or "portfolio"))
            default = "GoldPortfolio-{}-{}.gold".format(
                who, datetime.now().strftime("%Y-%m-%d"))
            path = filedialog.asksaveasfilename(
                parent=par, title="حفظ نسخة في ملف",
                defaultextension=".gold", initialfile=default,
                filetypes=[("ملف محفظة الذهب", "*.gold"), ("كل الملفات", "*.*")])
            if not path:
                return
            if not path.lower().endswith(".gold"):
                path += ".gold"
            payload = {"app": "gold_portfolio", "kind": "pocket", "ver": 1,
                       "created": datetime.now().strftime("%Y-%m-%d %H:%M"),
                       "device": socket.gethostname(),
                       "user": self.user or "",
                       "data": self.sync_snapshot()}
            with open(path, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=1)
            kb = max(1, os.path.getsize(path) // 1024)
            messagebox.showinfo(
                "نسخة بملف",
                "تم حفظ نسخة كاملة من بياناتك في:\n{}\n({} كيلوبايت)\n\n"
                "أرسل الملف لنفسك (واتساب/تيليجرام) أو ضعه في مجلد Google Drive، "
                "وفي الجهاز الآخر اختر «استعادة من ملف».".format(os.path.basename(path), kb),
                parent=par)
        except Exception as e:
            messagebox.showerror("نسخة بملف", "تعذّر حفظ الملف:\n{}".format(e),
                                 parent=par)

    def pocket_import(self):
        """يستعيد البيانات من ملف .gold (بعد حفظ نسخة احتياطية من بيانات الجهاز)."""
        try:
            path = filedialog.askopenfilename(
                parent=self.root, title="استعادة من ملف",
                filetypes=[("ملف محفظة الذهب", "*.gold"), ("كل الملفات", "*.*")])
            if not path:
                return
            with open(path, "r", encoding="utf-8") as f:
                obj = json.load(f)
            data = obj.get("data") if isinstance(obj, dict) else None
            if not isinstance(data, dict) or not any(k in data for k in SYNC_KEYS):
                messagebox.showerror("استعادة من ملف",
                                     "هذا الملف غير صالح أو ليس ملف محفظة.",
                                     parent=self.root)
                return
            n = len(data.get("holdings") or [])
            when = obj.get("created") or "—"
            if not messagebox.askyesno(
                    "استعادة من ملف",
                    "سيتم استبدال بيانات هذا الجهاز ببيانات الملف:\n\n"
                    "الملف: {}\nالتاريخ: {}\nعدد العمليات: {}\n\n"
                    "تُحفظ نسخة من بياناتك الحالية قبل الاستبدال. متابعة؟".format(
                        os.path.basename(path), when, n),
                    parent=self.root):
                return
            try:
                self.auto_backup(force=True)
            except Exception:
                pass
            target = getattr(self, "data_file", None) or DATA_FILE
            self._cloud_skip = True
            try:
                base = {}
                try:
                    with open(target, "r", encoding="utf-8") as f:
                        base = json.load(f) or {}
                except Exception:
                    base = {}
                if not isinstance(base, dict):
                    base = {}
                base.update({k: data[k] for k in SYNC_KEYS if k in data})
                tmp = target + ".tmp"
                with open(tmp, "w", encoding="utf-8") as f:
                    json.dump(base, f, ensure_ascii=False)
                    f.flush()
                    os.fsync(f.fileno())
                os.replace(tmp, target)
            finally:
                self._cloud_skip = False
            self._reset_state()
            self.load()
            try:
                self.rebuild_ui()
                self.start_live_refresh()
            except Exception:
                pass
            # اجعل نسخة الجهاز هي الأحدث في السحابة (إن كان البريد مرتبطاً)
            try:
                if self.cloud_ready() and self.cloud_conf().get("auto", True):
                    threading.Thread(target=self.cloud_push, daemon=True).start()
            except Exception:
                pass
            messagebox.showinfo("استعادة من ملف",
                                "تم استرجاع {} عملية من الملف.".format(n),
                                parent=self.root)
        except Exception as e:
            messagebox.showerror("استعادة من ملف", "تعذّر فتح الملف:\n{}".format(e),
                                 parent=self.root)

    def cloud_status(self, text, kind="idle"):
        self._cloud_msg = text
        self._cloud_kind = kind
        lbl = getattr(self, "lbl_cloud", None)
        if lbl is not None:
            try:
                colors = {"ok": self.theme.get("profit", "#3ddc97"),
                          "err": self.theme.get("loss", "#ff6b6b"),
                          "busy": self.theme.get("accent", "#e9c684"),
                          "idle": self.theme.get("muted", "#888")}
                lbl.config(text=text, fg=colors.get(kind, colors["idle"]))
            except Exception:
                pass
        btn = getattr(self, "btn_cloud", None)
        if btn is not None:
            try:
                btn.config(text={"ok": "☁", "err": "⚠", "busy": "☁"}.get(kind, "☁"))
            except Exception:
                pass

    def cloud_db_base(self):
        url = self.cloud_server()[0]
        if not url:
            raise CloudError("لم يُضبط خادم المزامنة — راجع إعداد Firebase.")
        if not url.startswith("http"):
            url = "https://" + url
        return url

    def cloud_db_url(self, uid):
        if not uid:
            raise CloudError("لم يُربط حساب سحابي بعد.")
        return "{}/goldPortfolio/users/{}.json".format(self.cloud_db_base(), uid)

    def cloud_token(self, user=None, force=False):
        """يعيد رمز دخول صالحاً (ويجدّده تلقائياً عند انتهاء صلاحيته)."""
        uname = user if user is not None else (getattr(self, "user", None) or "")
        acct = self.cloud_account(uname)
        if not acct.get("refresh_token"):
            raise CloudError("لم يُربط حساب سحابي بهذا المستخدم.")
        if (not force) and acct.get("id_token") and \
                time.time() < float(acct.get("expires", 0) or 0) - 120:
            return acct["id_token"]
        key = self.cloud_server()[1]
        body = urllib.parse.urlencode({
            "grant_type": "refresh_token",
            "refresh_token": acct["refresh_token"]}).encode("utf-8")
        req = urllib.request.Request(
            FB_REFRESH.format(key=urllib.parse.quote(key)), data=body,
            headers={"Content-Type": "application/x-www-form-urlencoded",
                     "User-Agent": BROWSER_UA})
        try:
            with urllib.request.urlopen(req, timeout=15) as r:
                d = json.loads(r.read().decode("utf-8", "ignore"))
        except urllib.error.HTTPError as e:
            raise CloudError("انتهت صلاحية الجلسة — أعد ربط الحساب السحابي.", e.code)
        except Exception as e:
            raise CloudError("تعذّر الاتصال بالإنترنت: {}".format(e))
        tok = d.get("id_token") or d.get("idToken") or ""
        if not tok:
            raise CloudError("تعذّر تجديد جلسة الدخول — أعد الربط.")
        self.cloud_update_account(uname, id_token=tok,
                                  refresh_token=d.get("refresh_token",
                                                       acct["refresh_token"]),
                                  expires=time.time() + float(d.get("expires_in", 3600) or 3600),
                                  uid=acct.get("uid") or d.get("user_id", ""))
        return tok

    def cloud_auth(self, email, pw, url="", key="", mode="login"):
        """دخول/إنشاء حساب سحابي بالبريد وكلمة المرور فقط."""
        burl, bkey = self.cloud_server()
        url = (url or burl or "").strip().rstrip("/")
        if not url:
            raise CloudError("لم يُضبط خادم المزامنة بعد — راجع إعداد Firebase.")
        if not url.startswith("http"):
            url = "https://" + url
        key = (key or bkey or "").strip()
        if not key:
            raise CloudError("لم يُضبط مفتاح الخادم بعد — راجع إعداد Firebase.")
        email = (email or "").strip()
        if "@" not in email or "." not in email:
            raise CloudError("أدخل بريداً إلكترونياً صحيحاً.")
        if len(pw or "") < 6:
            raise CloudError("كلمة المرور يجب أن تكون ٦ أحرف على الأقل.")
        ep = FB_SIGNUP if mode == "register" else FB_SIGNIN
        data, _etag = cloud_http(
            "POST", ep.format(key=urllib.parse.quote(key)),
            {"email": email, "password": pw, "returnSecureToken": True}, timeout=20)
        if not isinstance(data, dict) or not data.get("idToken"):
            raise CloudError("لم يكتمل تسجيل الدخول — تحقق من إعدادات Firebase.")
        conf = self.cloud_conf()
        conf["last_email"] = email
        if not self.cloud_builtin():
            conf["url"] = url
            conf["key"] = key
        self.cloud_store()
        return {"uid": data.get("localId", ""), "email": email,
                "id_token": data.get("idToken", ""),
                "refresh_token": data.get("refreshToken", ""),
                "expires": time.time() + float(data.get("expiresIn", 3600) or 3600)}

    def cloud_reset_password(self, email):
        """يرسل رابط إعادة تعيين كلمة المرور إلى البريد."""
        email = (email or "").strip()
        if "@" not in email or "." not in email:
            raise CloudError("اكتب بريدك في الحقل أولاً ثم اضغط «نسيت كلمة المرور».")
        key = self.cloud_server()[1]
        if not key:
            raise CloudError("لم يُضبط خادم المزامنة بعد.")
        data, _etag = cloud_http(
            "POST", FB_RESET.format(key=urllib.parse.quote(key)),
            {"requestType": "PASSWORD_RESET", "email": email}, timeout=20)
        if isinstance(data, dict) and data.get("error"):
            raise CloudError("تعذّر إرسال الرابط — تأكد من البريد.")
        return "أرسلنا رابط إعادة تعيين كلمة المرور إلى {}".format(email)

    def cloud_login_device(self, info, pw):
        """ينشئ/يحدّث حساباً محلياً مرتبطاً بالحساب السحابي ويعيد اسمه."""
        email = info.get("email", "")
        users = load_users()
        target = None
        for u, a in (self.cloud_conf().get("accounts") or {}).items():
            if str(a.get("email", "")).lower() == email.lower():
                target = u
                break
        if target is None:
            base = re.sub(r"[^A-Za-z0-9._\-]", "_", email.split("@")[0]) or "user"
            cand, i = base, 2
            low = {k.lower() for k in users}
            while cand.lower() in low:
                cand = "{}{}".format(base, i)
                i += 1
            target = cand
        salt, h = hash_password(pw)
        rec = dict(users.get(target) or {})
        rec.update({"salt": salt, "hash": h, "email": email,
                    "created": rec.get("created") or
                    datetime.now().strftime("%Y-%m-%d %H:%M")})
        users[target] = rec
        if not save_users(users):
            raise CloudError("تعذّر إنشاء الحساب المحلي على هذا الجهاز.")
        old = self.cloud_account(target)
        acct = dict(old)
        same = (str(old.get("email", "")).strip().lower() == email.strip().lower())
        acct.update(info)
        if old.get("uid") and old.get("uid") != info.get("uid"):
            acct.update({"rev": 0, "etag": "", "last_hash": ""})
        if info.get("provider") == "gmail" and not same:
            acct.update({"rev": 0, "last_hash": ""})
        acct.setdefault("rev", 0)
        conf = self.cloud_conf()
        conf.setdefault("accounts", {})[target] = acct
        self.cloud_store()
        prefs = load_prefs()
        prefs["last_user"] = target
        save_prefs(prefs)
        return target

    def cloud_fetch(self, user=None):
        uname = user if user is not None else (getattr(self, "user", None) or "")
        acct = self.cloud_account(uname)
        if (acct.get("provider") or "") == "gmail":
            return self.gmail_pull(uname)
        token = self.cloud_token(uname)
        url = self.cloud_db_url(acct.get("uid")) + "?auth=" + urllib.parse.quote(token)
        data, _etag = cloud_http("GET", url, None, timeout=20)
        return data

    def cloud_download_to_file(self, user):
        """يجلب نسخة السحابة ويكتبها في ملف المستخدم (بنسخة احتياطية قبل الاستبدال)."""
        data = self.cloud_fetch(user)
        if not isinstance(data, dict) or not isinstance(data.get("data"), dict):
            return False
        payload = data["data"]
        path = user_data_file(user)
        if os.path.exists(path) and os.path.getsize(path) > 2:
            try:
                shutil.copyfile(path, path + ".before-sync")
            except Exception:
                pass
        try:
            base = {}
            if os.path.exists(path):
                with open(path, "r", encoding="utf-8") as f:
                    base = json.load(f) or {}
            if not isinstance(base, dict):
                base = {}
            base.update({k: payload[k] for k in SYNC_KEYS if k in payload})
            with open(path, "w", encoding="utf-8") as f:
                json.dump(base, f, ensure_ascii=False)
        except Exception as e:
            raise CloudError("تعذّر حفظ البيانات المسترجعة: {}".format(e))
        self.cloud_update_account(user, rev=int(data.get("rev", 0) or 0),
                                  last_sync=datetime.now().strftime("%Y-%m-%d %H:%M"),
                                  last_hash="")
        self._cloud_pulled = True
        return True

    def cloud_push(self, quiet=True):
        """يرفع نسخة الجهاز إلى السحابة. يعيد (نجاح، رسالة)."""
        if not getattr(self, "user", None) or not self.cloud_ready():
            return False, "لم يُربط حساب سحابي بعد."
        if self.cloud_provider() == "gmail":
            return self.gmail_push()
        acct = self.cloud_account()
        try:
            token = self.cloud_token()
        except CloudError as e:
            return False, "☁ {}".format(e)
        payload = {"rev": int(acct.get("rev") or 0) + 1,
                   "updated_ts": time.time(),
                   "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
                   "device": socket.gethostname(),
                   "data": self.sync_snapshot()}
        try:
            url = self.cloud_db_url(acct.get("uid")) + "?auth=" + urllib.parse.quote(token)
            cloud_http("PUT", url, payload, timeout=20)
        except CloudError as e:
            return False, "☁ تعذّرت المزامنة: {}".format(e)
        self.cloud_update_account(self.user, rev=payload["rev"],
                                  last_sync=payload["updated_at"],
                                  last_hash=self.cloud_hash(),
                                  last_device=payload["device"])
        return True, "☁ متزامن — {}".format(payload["updated_at"])

    def cloud_apply_remote(self, remote):
        """يستبدل بيانات الجهاز بنسخة السحابة (بعد نسخة احتياطية)."""
        payload = (remote or {}).get("data") or {}
        try:
            self.auto_backup(force=True)
        except Exception:
            pass
        path = getattr(self, "data_file", None) or DATA_FILE
        self._cloud_skip = True
        try:
            base = {}
            try:
                with open(path, "r", encoding="utf-8") as f:
                    base = json.load(f) or {}
            except Exception:
                base = {}
            if not isinstance(base, dict):
                base = {}
            base.update({k: payload[k] for k in SYNC_KEYS if k in payload})
            try:
                tmp = path + ".tmp"
                with open(tmp, "w", encoding="utf-8") as f:
                    json.dump(base, f, ensure_ascii=False)
                    f.flush()
                    os.fsync(f.fileno())
                os.replace(tmp, path)
            except Exception as e:
                raise CloudError("تعذّر حفظ البيانات: {}".format(e))
            self._reset_state()
            self.load()
            self.cloud_update_account(
                self.user, rev=int(remote.get("rev", 0) or 0),
                last_sync=datetime.now().strftime("%Y-%m-%d %H:%M"),
                last_hash=self.cloud_hash())
            self._cloud_pulled = True
        finally:
            self._cloud_skip = False
        try:
            self.rebuild_ui()
            self.start_live_refresh()
        except Exception:
            pass

    def cloud_touch(self):
        """بعد أي حفظ محلي: يرفع للسحابة بعد مهلة قصيرة (تجميع التعديلات)."""
        try:
            if getattr(self, "_cloud_skip", False):
                return
            if not self.cloud_ready() or not self.cloud_conf().get("auto", True):
                return
            if self.cloud_hash() == (self.cloud_account().get("last_hash") or ""):
                return
            job = getattr(self, "_cloud_job", None)
            if job:
                try:
                    self.root.after_cancel(job)
                except Exception:
                    pass
            delay = 4000
            if self.cloud_provider() == "gmail":
                # البريد ليس قاعدة بيانات: نجمّع التعديلات ونرسل رسالة كل 3 دقائق
                gap = 180 - (time.time() - float(getattr(self, "_gmail_last", 0) or 0))
                delay = max(delay, int(gap * 1000))
            self._cloud_job = self.root.after(delay, self._cloud_push_async)
        except Exception:
            pass

    def _cloud_push_async(self):
        self._cloud_job = None
        if not self.cloud_ready():
            return

        def work():
            ok, msg = self.cloud_push()
            try:
                self.root.after(0, lambda: self.cloud_status(msg, "ok" if ok else "err"))
            except Exception:
                pass

        threading.Thread(target=work, daemon=True).start()

    def cloud_sync_now(self, manual=True):
        """مزامنة كاملة: تنزيل الأحدث ثم رفع تعديلات الجهاز إن وُجدت."""
        if not getattr(self, "user", None):
            return
        if not self.cloud_ready():
            if manual:
                self.open_cloud_dialog()
            return
        self.cloud_status("☁ جارٍ المزامنة…", "busy")

        def work():
            try:
                remote = self.cloud_fetch()
            except CloudError as e:
                try:
                    self.root.after(0, lambda: self.cloud_status("☁ " + str(e), "err"))
                except Exception:
                    pass
                return
            try:
                self.root.after(0, lambda: self._cloud_after_fetch(remote, manual))
            except Exception:
                pass

        threading.Thread(target=work, daemon=True).start()

    def _cloud_after_fetch(self, remote, manual=True):
        acct = self.cloud_account()
        if not isinstance(remote, dict):
            self._cloud_push_async()
            return
        rrev = int(remote.get("rev") or 0)
        lrev = int(acct.get("rev") or 0)
        last_hash = acct.get("last_hash") or ""
        local_changed = (self.cloud_hash() != last_hash) if last_hash else \
            bool(self.holdings or self.history or self.sells)
        if rrev > lrev and local_changed:
            ask3 = getattr(messagebox, "askyesnocancel", None)
            q = ("توجد نسخة أحدث في السحابة، وبيانات هذا الجهاز تغيّرت أيضاً.\n\n"
                 "نعم = استخدم بيانات السحابة (تُحفظ نسخة من بياناتك الحالية)\n"
                 "لا = ارفع بيانات هذا الجهاز لتصبح هي النسخة السحابية\n"
                 "إلغاء = لاحقاً")
            ans = ask3("مزامنة", q) if ask3 else messagebox.askyesno("مزامنة", q)
            if ans is None:
                self.cloud_status("☁ المزامنة مؤجلة", "idle")
                return
            if ans:
                self.cloud_apply_remote(remote)
                self.cloud_status("☁ تم استرجاع بيانات السحابة", "ok")
            else:
                self._cloud_push_async()
            return
        if rrev > lrev:
            self.cloud_apply_remote(remote)
            self.cloud_status("☁ تم استرجاع بيانات السحابة", "ok")
            return
        if local_changed:
            self._cloud_push_async()
            return
        self.cloud_status("☁ متزامن — {}".format(
            remote.get("updated_at") or acct.get("last_sync") or ""), "ok")

    def cloud_unlink(self):
        if not getattr(self, "user", None):
            return
        if not messagebox.askyesno("فصل الحساب السحابي",
                                   "هل تريد فصل هذا الجهاز عن الحساب السحابي؟\n"
                                   "بياناتك تبقى على الجهاز والسحابة كما هي."):
            return
        conf = self.cloud_conf()
        (conf.get("accounts") or {}).pop(self.user, None)
        self.cloud_store()
        self.cloud_status("☁ غير مرتبط — البيانات محلية", "idle")
        try:
            self.open_settings()
        except Exception:
            pass

    def open_gmail_dialog(self, from_login=False):
        """ربط الحساب بالبريد فقط: بريد + كلمة مرور التطبيق (بلا قاعدة بيانات)."""
        t = self.theme
        BG, CARD = t["bg"], t["card"]
        TXT, MUT, GOLD = t["text"], t["muted"], t["accent"]
        acct = self.cloud_account()
        win = tk.Toplevel(self.root)
        try:
            win.title("المزامنة عبر بريدك")
            win.configure(bg=BG)
            win.transient(self.root)
            win.resizable(False, False)
            sw, sh = self._screen_size()
            w = min(560, max(420, int(sw) - 40))
            h = min(545, max(430, int(sh) - 80))   # لا تخرج عن الشاشة الصغيرة
            win.geometry("{}x{}+{}+{}".format(w, h, max(0, (sw - w) // 2),
                                              max(6, (sh - h) // 4)))
            win.grab_set()
        except Exception:
            pass

        tk.Label(win, text="☁  المزامنة عبر بريدك",
                 bg=BG, fg=GOLD, font=("Segoe UI", 15, "bold")).pack(pady=(14, 2))
        tk.Label(win, text=("سجّل ببريدك لاسترجاع محفظتك على هذا الجهاز"
                            if from_login else
                            "نسخة مشفّرة داخل بريدك — بلا قاعدة بيانات ولا رابط"),
                 bg=BG, fg=MUT, font=("Segoe UI", 10)).pack(pady=(0, 8))

        # شريط الأزرار أولاً ومثبّت في الأسفل: يبقى ظاهراً مهما طال المحتوى
        bar = tk.Frame(win, bg=BG)
        bar.pack(side="bottom", fill="x", padx=14, pady=(6, 12))
        btns = []

        body = tk.Frame(win, bg=CARD, highlightthickness=1,
                        highlightbackground=t.get("border", "#252c39"))
        body.pack(fill="both", expand=True, padx=14, pady=6)

        def fld(label, var, secret=False):
            tk.Label(body, text=label, bg=CARD, fg=MUT,
                     font=("Segoe UI", 10), anchor="e").pack(fill="x", padx=16, pady=(10, 2))
            e = tk.Entry(body, textvariable=var, justify="left",
                         bg=t.get("entry_bg", "#151a23"), fg=TXT,
                         insertbackground=TXT, relief="flat", bd=0,
                         highlightthickness=1, highlightbackground=t.get("border", "#252c39"),
                         highlightcolor=GOLD, font=("Segoe UI", 10))
            if secret:
                e.configure(show="•")
            e.pack(fill="x", padx=16, ipady=5)
            return e

        v_mail = tk.StringVar(value=acct.get("email", "") or
                              (self.cloud_conf().get("last_email") or ""))
        v_pw = tk.StringVar(value="")
        e_mail = fld("البريد الإلكتروني (Gmail)", v_mail)
        fld("كلمة مرور التطبيق (16 حرفاً)", v_pw, secret=True)
        try:
            e_mail.focus_set()
        except Exception:
            pass

        lbl = tk.Label(body, text="", bg=CARD, fg=MUT, wraplength=500,
                       justify="right", font=("Segoe UI", 10))
        lbl.pack(fill="x", padx=16, pady=(10, 2))
        tk.Label(body, text=("كلمة مرور التطبيق ليست كلمة مرور حسابك. يجب أولاً تفعيل "
                             "«التحقق بخطوتين» في حساب Google، وبعدها فقط تظهر صفحة "
                             "كلمات مرور التطبيقات. إن ظهرت رسالة "
                             "\"The setting you are looking for is not available\" "
                             "فمعناها أن التحقق بخطوتين غير مفعّل (أو الحساب يديره عمل/مدرسة)."),
                 bg=CARD, fg=MUT, wraplength=500, justify="right",
                 font=("Segoe UI", 10)).pack(fill="x", padx=16, pady=(0, 4))
        self._mkbtn(body, "طريقة الحصول عليها", lambda: webbrowser.open(
            "https://myaccount.google.com/apppasswords"), kind="ghost",
            font_size=10).pack(anchor="e", padx=14, pady=(0, 8))

        def run(mode):
            for b in btns:
                try:
                    b.config(state="disabled")
                except Exception:
                    pass
            lbl.config(text="جارٍ الاتصال ببريدك…", fg=GOLD)
            try:
                win.update_idletasks()
            except Exception:
                pass
            email = v_mail.get().strip()
            pw = "".join(v_pw.get().split())
            try:
                gmail_check(email, pw)
                info = {"provider": "gmail", "email": email,
                        "gmail_pw": self._pw_wrap(pw)}
                if from_login or mode == "pull":
                    name = self.cloud_login_device(info, pw)
                    got = False
                    try:
                        got = self.cloud_download_to_file(name)
                    except CloudError as e:
                        lbl.config(text="تم الدخول لكن تعذّر التنزيل: {}".format(e),
                                   fg=t.get("loss", "#ff6b6b"))
                    try:
                        win.destroy()
                    except Exception:
                        pass
                    self._finish_login(name)
                    self.cloud_status("☁ مرحباً {} — {}".format(
                        email, "تم استرجاع بياناتك" if got
                        else "لا نسخة في البريد بعد، سيُرفع ما على هذا الجهاز"), "ok")
                else:
                    old = self.cloud_account()
                    same = (str(old.get("email", "")).strip().lower() == email.lower())
                    self.cloud_update_account(
                        self.user, provider="gmail", email=email,
                        gmail_pw=self._pw_wrap(pw),
                        rev=int(old.get("rev") or 0) if same else 0,
                        last_hash=old.get("last_hash", "") if same else "")
                    try:
                        win.destroy()
                    except Exception:
                        pass
                    self.cloud_sync_now(True)
                    try:
                        self.open_settings()
                    except Exception:
                        pass
            except CloudError as e:
                lbl.config(text=str(e), fg=t.get("loss", "#ff6b6b"))
                for b in btns:
                    try:
                        b.config(state="normal")
                    except Exception:
                        pass
            except Exception as e:
                lbl.config(text="خطأ غير متوقع: {}".format(e),
                           fg=t.get("loss", "#ff6b6b"))
                for b in btns:
                    try:
                        b.config(state="normal")
                    except Exception:
                        pass

        main_txt = "✔ موافق ودخول" if from_login else "✔ دخول واسترجاع"
        btns.append(self._mkbtn(bar, main_txt, lambda: run("pull"),
                                kind="primary", font_size=11))
        btns[-1].pack(side="right", padx=4)
        btns.append(self._mkbtn(bar, "إرسال بيانات هذا الجهاز", lambda: run("push"),
                                kind="neutral", font_size=10))
        btns[-1].pack(side="right", padx=4)
        btns.append(self._mkbtn(bar, "إلغاء", win.destroy, kind="ghost", font_size=10))
        btns[-1].pack(side="left", padx=4)
        try:
            win.bind("<Return>", lambda e: run("pull"))
            win.bind("<Escape>", lambda e: win.destroy())
        except Exception:
            pass

        def _pocket_here():
            try:
                win.grab_release()
            except Exception:
                pass
            self.pocket_export(win)
            try:
                win.grab_set()
            except Exception:
                pass
        self._mkbtn(body, "أو 📦 نسخة بملف واحد (بلا بريد ولا كلمة مرور)",
                    _pocket_here, kind="ghost", font_size=10).pack(
                        anchor="center", pady=(0, 10))

    def open_cloud_dialog(self, from_login=False):
        """نافذة ربط الحساب: بالبريد (Gmail) افتراضياً، أو Firebase إن كان مضمّناً."""
        if not self.cloud_builtin():
            return self.open_gmail_dialog(from_login)
        """نافذة ربط الحساب السحابي / استرجاع البيانات من جهاز آخر."""
        t = self.theme
        BG, CARD = t["bg"], t["card"]
        TXT, MUT, GOLD = t["text"], t["muted"], t["accent"]
        conf = self.cloud_conf()
        acct = self.cloud_account()
        win = tk.Toplevel(self.root)
        try:
            win.title("المزامنة السحابية")
            win.configure(bg=BG)
            win.transient(self.root)
            win.resizable(False, False)
            w, h = 560, 520
            sw, sh = self._screen_size()
            win.geometry("{}x{}+{}+{}".format(w, h, max(0, (sw - w) // 2),
                                               max(0, (sh - h) // 3)))
            win.grab_set()
        except Exception:
            pass

        tk.Label(win, text=("☁  دخول بالبريد" if from_login else "☁  المزامنة السحابية"),
                 bg=BG, fg=GOLD, font=("Segoe UI", 15, "bold")).pack(pady=(14, 2))
        tk.Label(win, text=("أدخل بريدك وكلمة المرور — تُسترجع محفظتك تلقائياً على أي جهاز"
                            if from_login else
                            "اربط جهازك ببريدك لتُحدَّث بياناتك تلقائياً"),
                 bg=BG, fg=MUT, font=("Segoe UI", 10)).pack(pady=(0, 8))

        body = tk.Frame(win, bg=CARD, highlightthickness=1,
                        highlightbackground=t.get("border", "#252c39"))
        body.pack(fill="both", expand=True, padx=14, pady=6)

        def fld(parent, label, var, secret=False):
            tk.Label(parent, text=label, bg=CARD, fg=MUT,
                     font=("Segoe UI", 10), anchor="e").pack(fill="x", padx=16, pady=(8, 2))
            e = tk.Entry(parent, textvariable=var, justify="left",
                         bg=t.get("entry_bg", "#151a23"), fg=TXT,
                         insertbackground=TXT, relief="flat", bd=0,
                         highlightthickness=1, highlightbackground=t.get("border", "#252c39"),
                         highlightcolor=GOLD, font=("Segoe UI", 10))
            if secret:
                e.configure(show="•")
            e.pack(fill="x", padx=16, ipady=5)
            return e

        builtin = self.cloud_builtin()
        srv_url, srv_key = self.cloud_server()

        v_mail = tk.StringVar(value=acct.get("email", "") or
                              (self.cloud_conf().get("last_email") or ""))
        v_pw = tk.StringVar(value="")
        e_mail = fld(body, "البريد الإلكتروني", v_mail)
        e_pw = fld(body, "كلمة المرور (٦ أحرف على الأقل)", v_pw, secret=True)

        # حقول الخادم مخفيّة تماماً في الاستخدام العادي (دخول بالبريد فقط)
        v_url = tk.StringVar(value=srv_url)
        v_key = tk.StringVar(value=srv_key)
        adv = tk.Frame(body, bg=CARD)

        def show_adv():
            fld(adv, "رابط قاعدة البيانات (Database URL)", v_url)
            fld(adv, "مفتاح الويب (Web API key)", v_key)
            adv.pack(fill="x")
            btn_adv.pack_forget()

        lbl = tk.Label(body, text="", bg=CARD, fg=MUT, wraplength=500,
                       justify="right", font=("Segoe UI", 10))
        lbl.pack(fill="x", padx=16, pady=(8, 0))
        tip = ("بياناتك مرتبطة ببريدك: سجّل الدخول من أي جهاز لتجد محفظتك كما هي."
               if builtin else
               "أدخل بريدك وكلمة المرور لبدء المزامنة.")
        tk.Label(body, text=tip, bg=CARD, fg=MUT, wraplength=500, justify="right",
                 font=("Segoe UI", 10)).pack(fill="x", padx=16, pady=(2, 0))
        btn_adv = self._mkbtn(body, "⚙ إعداد الخادم", show_adv, kind="ghost",
                             font_size=10)
        if not builtin:
            btn_adv.pack(anchor="e", padx=14, pady=(2, 8))
        else:
            tk.Frame(body, bg=CARD, height=6).pack()

        bar = tk.Frame(win, bg=BG)
        bar.pack(fill="x", padx=14, pady=(4, 12))
        btns = []

        def run(mode):
            for b in btns:
                try:
                    b.config(state="disabled")
                except Exception:
                    pass
            lbl.config(text="جارٍ الاتصال بـ Firebase…", fg=GOLD)
            try:
                win.update_idletasks()
            except Exception:
                pass
            try:
                info = self.cloud_auth(v_mail.get(), v_pw.get(), v_url.get(),
                                       v_key.get(), mode)
                if from_login:
                    name = self.cloud_login_device(info, v_pw.get())
                    got = False
                    try:
                        got = self.cloud_download_to_file(name)
                    except CloudError as e:
                        lbl.config(text="تم الدخول لكن تعذّر التنزيل: {}".format(e),
                                   fg=t.get("loss", "#ff6b6b"))
                    try:
                        win.destroy()
                    except Exception:
                        pass
                    self._finish_login(name)
                    self.cloud_status("☁ مرحباً {} — {}".format(
                        info.get("email", ""), "تم استرجاع البيانات" if got
                        else "سيُرفع ما على هذا الجهاز"), "ok")
                    self.cloud_sync_now(True)
                else:
                    old = self.cloud_account()
                    rev = int(old.get("rev") or 0) if old.get("uid") == info.get("uid") else 0
                    self.cloud_update_account(
                        self.user, rev=rev, last_hash=old.get("last_hash", "")
                        if old.get("uid") == info.get("uid") else "", **info)
                    try:
                        win.destroy()
                    except Exception:
                        pass
                    self.cloud_sync_now(True)
                    try:
                        self.open_settings()
                    except Exception:
                        pass
            except CloudError as e:
                lbl.config(text=str(e), fg=t.get("loss", "#ff6b6b"))
                for b in btns:
                    try:
                        b.config(state="normal")
                    except Exception:
                        pass
            except Exception as e:
                lbl.config(text="خطأ غير متوقع: {}".format(e),
                           fg=t.get("loss", "#ff6b6b"))
                for b in btns:
                    try:
                        b.config(state="normal")
                    except Exception:
                        pass

        btns.append(self._mkbtn(bar, "دخول", lambda: run("login"),
                                kind="primary", font_size=11))
        btns[-1].pack(side="right", padx=4)
        btns.append(self._mkbtn(bar, "إنشاء حساب بالبريد", lambda: run("register"),
                                kind="neutral", font_size=10))
        btns[-1].pack(side="right", padx=4)

        def forgot():
            lbl.config(text="جارٍ إرسال رابط إعادة التعيين…", fg=GOLD)
            try:
                msg = self.cloud_reset_password(v_mail.get())
                lbl.config(text="✓ " + msg, fg=t.get("profit", "#3ddc97"))
            except CloudError as e:
                lbl.config(text=str(e), fg=t.get("loss", "#ff6b6b"))
            except Exception as e:
                lbl.config(text="تعذّر الإرسال: {}".format(e),
                           fg=t.get("loss", "#ff6b6b"))

        btns.append(self._mkbtn(bar, "نسيت كلمة المرور؟", forgot,
                                kind="ghost", font_size=10))
        btns[-1].pack(side="left", padx=4)
        btns.append(self._mkbtn(bar, "إلغاء", win.destroy, kind="ghost", font_size=10))
        btns[-1].pack(side="left", padx=4)


    # ============================ التخزين ============================
    def load(self):
        path = getattr(self, "data_file", None) or DATA_FILE
        try:
            with open(path, "r", encoding="utf-8") as f:
                d = json.load(f)
            self.holdings = d.get("holdings", [])
            self.sells = d.get("sells", [])
            # ترحيل قديم: مركز يحمل سعر بيع → عملية بيع
            migrated = []
            for h in list(self.holdings):
                if h.get("sell") is not None:
                    self.sells.append({
                        "name": h.get("name", "بيع ذهب"),
                        "qty": h.get("qty", 0),
                        "unit": h.get("unit", "g"),
                        "buy": h.get("buy", 0),
                        "sell": h["sell"],
                        "date": h.get("sell_date", ""),
                    })
                else:
                    migrated.append(h)
            self.holdings = migrated
            self.history = d.get("history", [])
            self.price_hist = d.get("price_hist", [])
            self.cur = d.get("cur", "USD")
            try:
                self.karat = int(d.get("karat", 24))
            except (TypeError, ValueError):
                self.karat = 24
            cr = d.get("chart_range", "live")
            self.chart_range = cr if d.get("live_chart_version") == 1 and cr in {r[0] for r in CHART_RANGES} else "live"
            self.chart_metric = "gram"
            self.chart_currency = d.get("chart_currency", self.cur)
            if self.chart_currency not in CUR:
                self.chart_currency = "USD"
            self.chart_price_mode = d.get("chart_price_mode", "market")
            if self.chart_price_mode not in ("market", "rajhi_sell", "rajhi_buy") or self.chart_currency != "SAR":
                self.chart_price_mode = "market"
            self.chart_font_size = max(10, min(18, int(d.get("chart_font_size", 12))))
            self.chart_viewport = None
            self.adjust = float(d.get("adjust", 5.22))
            self.buy_margin = float(d.get("buy_margin", 0.0))
            self.alerts = d.get("alerts", [])
            self.dark_mode = bool(d.get("dark_mode", True))
            self.theme = dict(self._base_theme())
            if d.get("theme_ver") == THEME_VER:
                self.theme.update(self._modern_saved_theme(d.get("theme", {})))
            self.titles.update(d.get("titles", {}))
            self.sound_on = bool(d.get("sound_on", True))
            self.source_mode = "auto"  # migrate old saved choices to fresh primary feed
            self.show_sells = bool(d.get("show_sells", True))
            self.spot_oz_usd = float(d.get("spot_oz_usd", 0.0) or 0.0)
            self.rates.update(d.get("rates", {}))
            self.last_update = d.get("last_update", "")
            self.last_price_source = self._short_price_source(d.get("last_price_source", ""))
            self._widget_on = bool(d.get("widget_on", False))
            self._widget_geo = str(d.get("widget_geo", "") or "")
            try:
                self.lock_minutes = max(0, int(d.get("lock_minutes", 10)))
            except (TypeError, ValueError):
                self.lock_minutes = 10
        except Exception:
            pass
        if self.cur not in CUR:
            self.cur = "USD"
        if self.karat not in KARATS:
            self.karat = 24

    def save(self):
        if not getattr(self, "user", None):
            return
        try:
            tmp = self.data_file + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump({
                    "holdings": self.holdings,
                    "sells": self.sells,
                    "history": self.history,
                    "price_hist": self.price_hist[-12000:],
                    "cur": self.cur,
                    "karat": self.karat,
                    "chart_range": self.chart_range,
                    "chart_font_size": getattr(self, "chart_font_size", 12),
                    "chart_metric": "gram",
                    "chart_currency": self.chart_currency,
                    "chart_price_mode": self.chart_price_mode,
                    "live_chart_version": 1,
                    "adjust": self.adjust,
                    "buy_margin": self.buy_margin,
                    "alerts": self.alerts,
                    "theme": self.theme,
                    "theme_ver": THEME_VER,
                    "titles": self.titles,
                    "sound_on": self.sound_on,
                    "source_mode": self.source_mode,
                    "show_sells": bool(getattr(self, "show_sells", True)),
                    "spot_oz_usd": self.spot_oz_usd,
                    "rates": self.rates,
                    "last_update": self.last_update,
                    "last_price_source": getattr(self,"last_price_source",""),
                    "lock_minutes": self.lock_minutes,
                    "dark_mode": self.dark_mode,
                    "widget_on": bool(getattr(self, "_widget_on", False)),
                    "widget_geo": getattr(self, "_widget_geo", ""),
                }, f, ensure_ascii=False)
                f.flush()
                os.fsync(f.fileno())
            os.replace(tmp, self.data_file)
            self.auto_backup()
            self.cloud_touch()
        except Exception:
            pass

    # ============================ النسخ الاحتياطي ============================
    def backup_dir(self):
        d = os.path.join(HOME_DIR, "gold_portfolio_backups")
        try:
            os.makedirs(d, exist_ok=True)
        except Exception:
            pass
        return d

    def _backup_glob(self):
        return sorted(glob.glob(os.path.join(self.backup_dir(),
                                             "gold_portfolio_data*.json")))

    def auto_backup(self, force=False):
        """نسخة احتياطية دورية (كل 10 دقائق كحد أقصى) مع الاحتفاظ بآخر 15 نسخة."""
        if not getattr(self, "user", None):
            return None
        try:
            if not os.path.exists(self.data_file):
                return None
            now = time.time()
            if not force and now - getattr(self, "_last_backup", 0.0) < 600:
                return None
            self._last_backup = now
            base = os.path.splitext(os.path.basename(self.data_file))[0]
            target = os.path.join(self.backup_dir(), "{}_{}.json".format(
                base, datetime.now().strftime("%Y%m%d_%H%M%S")))
            shutil.copyfile(self.data_file, target)
            files = self._backup_glob()
            for old in files[:-15]:
                try:
                    os.remove(old)
                except Exception:
                    pass
            return target
        except Exception:
            return None

    def latest_backup(self):
        files = self._backup_glob()
        return files[-1] if files else None

    def restore_latest_backup(self):
        src = self.latest_backup()
        if not src:
            messagebox.showinfo("استعادة نسخة",
                                "لا توجد نسخ احتياطية بعد.\nاضغط «نسخ احتياطي الآن» أولاً.")
            return
        if not messagebox.askyesno(
                "استعادة نسخة",
                "سيتم استبدال بياناتك الحالية بالنسخة:\n{}\n\nهل أنت متأكد؟".format(
                    os.path.basename(src)), parent=self.root):
            return
        try:
            tmp = self.data_file + ".tmp"
            shutil.copyfile(src, tmp)
            os.replace(tmp, self.data_file)
        except Exception as e:
            messagebox.showerror("استعادة نسخة", "تعذّرت الاستعادة:\n{}".format(e))
            return
        self._reset_state()
        self.load()
        self.rebuild_ui()
        messagebox.showinfo("استعادة نسخة", "تمت استعادة النسخة بنجاح ✔")

    # ============================ الحسابات وتسجيل الدخول ============================
    def _reset_state(self):
        """يعيد كل الحقول إلى الافتراضي قبل تحميل بيانات مستخدم آخر."""
        self._fetch_generation += 1
        self._last_quote_source = None
        self._last_quote_timestamp = None
        self._fx_last = 0.0
        self._fx_retry_at = 0.0
        self.holdings = []
        self.sells = []
        self.show_sells = True
        self.history = []
        self.price_hist = []
        self.alerts = []
        self.theme = dict(DEFAULT_THEME)
        self.titles = dict(DEFAULT_TITLES)
        self.cur = "USD"
        self.karat = 24
        self.chart_range = "live"
        self.adjust = 5.22
        self.buy_margin = 0.0
        self.sound_on = True
        self.source_mode = "auto"
        self.spot_oz_usd = 0.0
        self.rates = dict(DEFAULT_RATES)
        self.last_update = ""
        self._delta_oz = 0.0
        self._ticks = 0
        self._fails = 0
        self.al_win = None
        self.row_map = {}
        self.chart_metric = "gram"
        self.chart_currency = "USD"
        self.chart_price_mode = "market"
        self.chart_font_size = 12
        self.chart_viewport = None
        self.last_price_source = ""
        self._undo_snapshot = None
        self._connection_state = "waiting"
        self._pages = {}
        self._page = "main"

    def show_login(self, mode="login", error=""):
        """شاشة تسجيل الدخول / إنشاء حساب تغطي النافذة قبل استخدام البرنامج."""
        users = load_users()
        if not users:
            mode = "register"
        try:
            self._close_chooser()
        except Exception:
            pass
        self._lg_mode = mode
        t = self.theme
        BG, CARD = t["bg"], t["card"]
        GOLD, TXT, MUT = t["accent"], t["text"], t["muted"]

        old = getattr(self, "login_frame", None)
        if old is not None:
            try:
                old.destroy()
            except Exception:
                pass
        self.root.configure(bg=BG)
        ov = tk.Frame(self.root, bg=BG)
        ov.place(x=0, y=0, relwidth=1, relheight=1)
        self.login_frame = ov
        box = tk.Frame(ov, bg=CARD, highlightthickness=1,
                       highlightbackground=t.get("border", "#252c39"))
        box.place(relx=0.5, rely=0.5, anchor="center", width=470)
        inner = tk.Frame(box, bg=CARD)
        inner.pack(fill="both", expand=True, padx=36, pady=28)

        tk.Label(inner, text="◈", bg=CARD, fg=GOLD,
                 font=("Segoe UI", 20)).pack()
        tk.Label(inner, text="محفظة الذهب", bg=CARD, fg=TXT,
                 font=("Segoe UI", 19, "bold")).pack(pady=(2, 2))
        sub = ("سجّل الدخول لعرض بياناتك الخاصة" if mode == "login"
               else "أنشئ حسابك لتبدأ بمحفظة مستقلة")
        tk.Label(inner, text=sub, bg=CARD, fg=MUT,
                 font=("Segoe UI", 10)).pack(pady=(0, 14))
        tk.Frame(inner, bg=t.get("border", "#252c39"), height=1).pack(fill="x", pady=(0, 6))

        def field(label, secret=False):
            tk.Label(inner, text=label, bg=CARD, fg=MUT,
                     font=("Segoe UI", 11), anchor="e").pack(fill="x", pady=(9, 3))
            e = self._mkentry(inner, tk.StringVar(), width=22)
            e.configure(justify="right")
            if secret:
                e.configure(show="•")
            e.pack(fill="x", ipady=6)
            return e

        self.lg_user = field("اسم المستخدم أو البريد")
        self.lg_pw = field("كلمة المرور", secret=True)
        self.lg_pw2 = field("تأكيد كلمة المرور", secret=True) if mode == "register" else None

        last_user = load_prefs().get("last_user") or ""
        self.lg_remember = None
        if mode == "login":
            self.lg_remember = tk.BooleanVar(value=bool(last_user))
            tk.Checkbutton(inner, text="تذكّر اسم المستخدم",
                           variable=self.lg_remember, bg=CARD, fg=MUT,
                           selectcolor=t.get("entry_bg", "#151a23"),
                           activebackground=CARD, activeforeground=TXT,
                           bd=0, highlightthickness=0, cursor="hand2",
                           font=("Segoe UI", 10)).pack(anchor="e", pady=(10, 0))
            if last_user:
                self.lg_user.insert(0, last_user)
        self.lg_err = tk.Label(inner, text=error, bg=CARD, fg=t.get("loss", "#ff6b6b"),
                               font=("Segoe UI", 11), wraplength=380, justify="right")
        self.lg_err.pack(fill="x", pady=(10, 0))

        self._mkbtn(inner, "دخول" if mode == "login" else "إنشاء الحساب",
                    self._login_submit, kind="primary",
                    font_size=11).pack(fill="x", pady=(12, 6), ipady=4)
        alt = ("＋ إنشاء حساب جديد" if mode == "login"
               else "لديك حساب؟ سجّل الدخول")
        self._mkbtn(inner, alt,
                    lambda: self.show_login("register" if mode == "login" else "login"),
                    kind="ghost", font_size=10).pack(fill="x")
        self._mkbtn(inner, "☁  الدخول بالبريد واسترجاع بياناتي",
                    lambda: self.open_cloud_dialog(from_login=True),
                    kind="ghost", font_size=10).pack(fill="x", pady=(4, 0))

        if mode == "register" and os.path.exists(DATA_FILE) and not users:
            tk.Label(inner, text="لديك بيانات محفوظة على هذا الجهاز — ستُربط بحسابك تلقائياً.",
                     bg=CARD, fg=MUT, font=("Segoe UI", 11), wraplength=380,
                     justify="right").pack(fill="x", pady=(12, 0))

        for e in (self.lg_user, self.lg_pw, self.lg_pw2):
            if e is not None:
                e.bind("<Return>", lambda _e: self._login_submit())
        try:
            (self.lg_user if mode == "register" else self.lg_pw).focus_set()
        except Exception:
            pass

    def _login_error(self, msg):
        try:
            self.lg_err.config(text=msg)
        except Exception:
            pass

    def _remember_user(self, name):
        """يحفظ اسم المستخدم فقط إذا فعّل الخيار (كلمة المرور لا تُحفظ إطلاقاً)."""
        prefs = load_prefs()
        box = getattr(self, "lg_remember", None)
        remember = bool(box.get()) if box is not None else True
        if remember:
            prefs["last_user"] = name
        else:
            prefs.pop("last_user", None)
        save_prefs(prefs)

    def _login_submit(self):
        if getattr(self, "_lg_mode", "login") == "register":
            self._do_register()
        else:
            self._do_login()

    def _do_login(self):
        name = (self.lg_user.get() or "").strip()
        pw = self.lg_pw.get() or ""
        if not name or not pw:
            self._login_error("أدخل اسم المستخدم وكلمة المرور.")
            return
        users = load_users()
        rec = users.get(name)
        if rec is None:
            for k in users:
                if k.lower() == name.lower():
                    name, rec = k, users[k]
                    break
        if rec is None and "@" in name:
            for k in users:
                if str(users[k].get("email", "")).strip().lower() == name.lower():
                    name, rec = k, users[k]
                    break
        if rec is None or not verify_password(pw, rec.get("salt", ""), rec.get("hash", "")):
            self._login_error("اسم المستخدم أو كلمة المرور غير صحيحة.")
            return
        self._remember_user(name)
        self._finish_login(name)

    def _do_register(self):
        name = (self.lg_user.get() or "").strip()
        pw = self.lg_pw.get() or ""
        pw2 = (self.lg_pw2.get() or "") if self.lg_pw2 is not None else ""
        if len(name) < 3:
            self._login_error("اسم المستخدم يجب أن يكون 3 أحرف على الأقل.")
            return
        if not all(c.isalnum() or c in " ._-" for c in name):
            self._login_error("استخدم حروفاً وأرقاماً فقط (يُسمح بـ . _ - والمسافة).")
            return
        if len(pw) < 4:
            self._login_error("كلمة المرور يجب أن تكون 4 أحرف على الأقل.")
            return
        if pw != pw2:
            self._login_error("كلمتا المرور غير متطابقتين.")
            return
        users = load_users()
        if any(k.lower() == name.lower() for k in users):
            self._login_error("اسم المستخدم مستخدم بالفعل، اختر اسماً آخر.")
            return
        salt, h = hash_password(pw)
        users[name] = {"salt": salt, "hash": h,
                       "created": datetime.now().strftime("%Y-%m-%d %H:%M")}
        if not save_users(users):
            self._login_error("تعذّر حفظ الحساب، تحقق من صلاحيات الكتابة.")
            return
        prefs = load_prefs()
        prefs["last_user"] = name
        save_prefs(prefs)
        self._finish_login(name)

    def _finish_login(self, name):
        """يفتح جلسة المستخدم: يحمّل ملفه الخاص ويعيد بناء الواجهة."""
        self.user = name
        self.data_file = user_data_file(name)
        users = load_users()
        # أول حساب على جهاز فيه بيانات قديمة: يتبنّاها بدل أن تضيع
        if (not os.path.exists(self.data_file) and os.path.exists(DATA_FILE)
                and len(users) <= 1):
            try:
                shutil.copyfile(DATA_FILE, self.data_file)
            except Exception:
                pass
        self._reset_state()
        self.load()
        self.save()
        lf = getattr(self, "login_frame", None)
        if lf is not None:
            try:
                lf.destroy()
            except Exception:
                pass
        self.login_frame = None
        self._closing_page = True
        try:
            for w in list(self.root.winfo_children()):
                w.destroy()
        finally:
            self._closing_page = False
        self._live_on = False
        self.build_ui()
        self.locked = False
        self.lock_frame = None
        self._last_activity = time.monotonic()
        self._bind_activity()
        self._schedule_idle_check()
        self.render()
        self.start_live_refresh()
        if getattr(self, "_widget_on", False):
            self.open_widget()
        # مزامنة سحابية عند الدخول (استرجاع البيانات على أي جهاز)
        try:
            if getattr(self, "_cloud_pulled", False):
                self._cloud_pulled = False
                if self.cloud_ready():
                    self.cloud_update_account(self.user, last_hash=self.cloud_hash())
            elif self.cloud_ready():
                self.cloud_sync_now(True)
        except Exception:
            pass

    def logout(self):
        if not messagebox.askyesno("تسجيل الخروج",
                                   "هل تريد تسجيل الخروج من هذا الحساب؟",
                                   parent=self.root):
            return
        self.save()
        try:
            if self.cloud_ready() and self.cloud_conf().get("auto", True):
                self.cloud_push()
        except Exception:
            pass
        self._live_on = False
        self.locked = False
        self.close_widget()
        self.lock_frame = None
        self._closing_page = True
        try:
            for w in list(self.root.winfo_children()):
                w.destroy()
        finally:
            self._closing_page = False
        self._pages = {}
        self.user = None
        self.data_file = DATA_FILE
        self.login_frame = None
        self._reset_state()
        self.show_login("login")

    # ============================ القفل التلقائي عند الخمول ============================
    def _bind_activity(self):
        """يرصد أي تفاعل (كتابة/نقر/حركة مؤشر) لإعادة تصفير مؤقّت الخمول."""
        if getattr(self, "_act_bound", False):
            return
        self._act_bound = True
        for seq in ("<Any-KeyPress>", "<Any-ButtonPress>",
                    "<Motion>", "<MouseWheel>", "<Button-4>", "<Button-5>"):
            try:
                self.root.bind_all(seq, self._on_activity, add="+")
            except Exception:
                pass
        try:
            self.root.bind_all("<Control-l>", lambda e: self.lock_now(), add="+")
            self.root.bind_all("<Control-L>", lambda e: self.lock_now(), add="+")
            self.root.bind_all("<Control-h>", lambda e: self.toggle_privacy(), add="+")
            self.root.bind_all("<Control-H>", lambda e: self.toggle_privacy(), add="+")
            self.root.bind_all("<Control-n>", lambda e: self._shortcut_buy(), add="+")
            self.root.bind_all("<Control-s>", lambda e: self.open_sell(), add="+")
            self.root.bind_all("<Escape>", lambda e: self._shortcut_home(), add="+")
            self.root.bind_all("<F9>", lambda e: self.toggle_widget(), add="+")
            self.root.bind_all("<F5>", lambda e: self.refresh_all(), add="+")
        except Exception:
            pass

    def _on_activity(self, event=None):
        # منع التحديث المتكرر أثناء حركة المؤشر (throttle بسيط)
        now = time.monotonic()
        if now - getattr(self, "_act_last", 0.0) < 0.3:
            return
        self._act_last = now
        self._last_activity = now

    def _schedule_idle_check(self):
        if getattr(self, "_idle_job", None) is not None:
            return
        try:
            self._idle_job = self.root.after(5000, self._idle_tick)
        except Exception:
            self._idle_job = None

    def _idle_tick(self):
        self._idle_job = None
        try:
            if getattr(self, "user", None) and not getattr(self, "locked", False):
                mins = int(getattr(self, "lock_minutes", 0) or 0)
                if mins > 0 and (time.monotonic() - self._last_activity) >= mins * 60:
                    self.lock_now()
        except Exception:
            pass
        self._schedule_idle_check()

    def lock_now(self):
        """يقفل التطبيق بطبقة تغطي كل شيء وتطلب كلمة المرور للعودة."""
        if not getattr(self, "user", None) or getattr(self, "locked", False):
            return
        self.locked = True
        self._widget_restore = getattr(self, "price_widget", None) is not None
        self.close_widget()
        self.save()
        t = self.theme
        BG, CARD = t["bg"], t["card"]
        GOLD, TXT, MUT = t["accent"], t["text"], t["muted"]
        px = self._px
        old = getattr(self, "lock_frame", None)
        if old is not None:
            try:
                old.destroy()
            except Exception:
                pass
        ov = tk.Frame(self.root, bg=BG)
        ov.place(x=0, y=0, relwidth=1, relheight=1)
        ov.lift()
        self.lock_frame = ov

        card = tk.Frame(ov, bg=CARD, highlightthickness=1,
                        highlightbackground=t.get("border", "#252c39"))
        card.place(relx=0.5, rely=0.5, anchor="center", width=px(430))
        inner = tk.Frame(card, bg=CARD)
        inner.pack(fill="both", expand=True, padx=px(32), pady=px(26))

        tk.Label(inner, text="🔒", bg=CARD, fg=GOLD,
                 font=("Segoe UI", 26)).pack()
        tk.Label(inner, text="التطبيق مقفل", bg=CARD, fg=TXT,
                 font=("Segoe UI", 17, "bold")).pack(pady=(4, 2))
        tk.Label(inner, text="قُفل تلقائياً بعد {} دقيقة من الخمول".format(self.lock_minutes)
                 if self.lock_minutes else "أدخل كلمة المرور للمتابعة",
                 bg=CARD, fg=MUT, font=("Segoe UI", 10)).pack(pady=(0, 4))
        tk.Label(inner, text="الحساب: {}".format(self.user or ""), bg=CARD,
                 fg=GOLD, font=("Segoe UI", 11, "bold")).pack(pady=(0, 10))
        tk.Frame(inner, bg=t.get("border", "#252c39"), height=1).pack(fill="x", pady=(0, 8))

        v_pw = tk.StringVar()
        ent = self._mkentry(inner, v_pw, width=24)
        ent.configure(justify="right", show="•")
        ent.pack(fill="x", ipady=7, pady=(6, 6))
        err = tk.Label(inner, text="", bg=CARD, fg=t.get("loss", "#ff6b6b"),
                       font=("Segoe UI", 10))
        err.pack(fill="x")

        def do_unlock(event=None):
            users = load_users()
            rec = users.get(self.user) or {}
            if verify_password(v_pw.get(), rec.get("salt", ""), rec.get("hash", "")):
                self._close_lock()
            else:
                err.config(text="كلمة المرور غير صحيحة، حاول مرة أخرى.")
                v_pw.set("")
                try:
                    ent.focus_set()
                except Exception:
                    pass

        bar = tk.Frame(inner, bg=CARD)
        bar.pack(fill="x", pady=(10, 4))
        self._mkbtn(bar, "🔓  فتح", do_unlock, kind="primary",
                    font_size=11).pack(side="right", fill="x", expand=True, ipady=3)
        self._mkbtn(bar, "تسجيل الخروج", self._lock_logout, kind="ghost",
                    font_size=10).pack(side="right", padx=(0, 8))
        ent.bind("<Return>", do_unlock)
        try:
            ent.focus_set()
        except Exception:
            pass
        tk.Label(inner, text="Ctrl+L للقفل اليدوي في أي وقت", bg=CARD, fg=MUT,
                 font=("Segoe UI", 10)).pack(pady=(8, 0))

    def _close_lock(self):
        f = getattr(self, "lock_frame", None)
        self.lock_frame = None
        self.locked = False
        self._last_activity = time.monotonic()
        if getattr(self, "_widget_restore", False):
            self._widget_restore = False
            self.open_widget()
        if f is not None:
            try:
                f.destroy()
            except Exception:
                pass

    def _lock_logout(self):
        self._close_lock()
        self.logout()

    def open_account(self):
        """صفحة الحساب: تغيير كلمة المرور وتسجيل الخروج."""
        t = self.theme
        BG, CARD = t["bg"], t["card"]
        GOLD, TXT, MUT = t["accent"], t["text"], t["muted"]
        win = self._open_page("account", "الحساب")
        win.configure(bg=BG)
        win.geometry("520x460")
        win.transient(self.root)

        card = tk.Frame(win, bg=CARD, highlightthickness=1,
                        highlightbackground=t.get("border", "#252c39"))
        card.pack(fill="both", expand=True, padx=4, pady=8)

        head = tk.Frame(card, bg=CARD)
        head.pack(fill="x", padx=24, pady=(18, 12))
        tk.Label(head, text="👤", bg=CARD, fg=GOLD,
                 font=("Segoe UI", 20)).pack(side="right")
        info = tk.Frame(head, bg=CARD)
        info.pack(side="right", padx=12)
        tk.Label(info, text=self.user or "", bg=CARD, fg=TXT,
                 font=("Segoe UI", 14, "bold")).pack(anchor="e")
        tk.Label(info, text="بياناتك محفوظة في ملف خاص بهذا الحساب فقط",
                 bg=CARD, fg=MUT, font=("Segoe UI", 10)).pack(anchor="e")

        tk.Frame(card, bg=t.get("border", "#252c39"), height=1).pack(fill="x", padx=24)
        tk.Label(card, text="تغيير كلمة المرور", bg=CARD, fg=GOLD,
                 font=("Segoe UI", 11, "bold")).pack(anchor="e", padx=24, pady=(16, 6))

        v_old, v_new, v_conf = tk.StringVar(), tk.StringVar(), tk.StringVar()

        def prow(label, var, secret=True):
            fr = tk.Frame(card, bg=CARD)
            fr.pack(fill="x", padx=24, pady=5)
            tk.Label(fr, text=label, bg=CARD, fg=MUT, width=18, anchor="e",
                     font=("Segoe UI", 10)).pack(side="right")
            e = self._mkentry(fr, var, width=18)
            e.configure(justify="right")
            if secret:
                e.configure(show="•")
            e.pack(side="right", fill="x", expand=True, padx=8, ipady=4)

        prow("كلمة المرور الحالية", v_old)
        prow("كلمة المرور الجديدة", v_new)
        prow("تأكيد الجديدة", v_conf)

        def do_change():
            users = load_users()
            rec = users.get(self.user) or {}
            if not verify_password(v_old.get(), rec.get("salt", ""), rec.get("hash", "")):
                messagebox.showerror("تغيير كلمة المرور",
                                     "كلمة المرور الحالية غير صحيحة.", parent=win)
                return
            if len(v_new.get() or "") < 4:
                messagebox.showwarning("تغيير كلمة المرور",
                                       "كلمة المرور الجديدة قصيرة (4 أحرف على الأقل).",
                                       parent=win)
                return
            if v_new.get() != v_conf.get():
                messagebox.showwarning("تغيير كلمة المرور",
                                       "كلمتا المرور غير متطابقتين.", parent=win)
                return
            salt, h = hash_password(v_new.get())
            rec.update({"salt": salt, "hash": h})
            users[self.user] = rec
            if not save_users(users):
                messagebox.showerror("تغيير كلمة المرور", "تعذّر حفظ التغيير.", parent=win)
                return
            for v in (v_old, v_new, v_conf):
                v.set("")
            messagebox.showinfo("تم", "تم تغيير كلمة المرور بنجاح.", parent=win)

        bar = tk.Frame(card, bg=CARD)
        bar.pack(fill="x", padx=24, pady=(14, 10))
        self._mkbtn(bar, "✔  حفظ كلمة المرور", do_change, kind="primary",
                    font_size=10).pack(side="right")

        # ---- القفل التلقائي عند الخمول ----
        tk.Frame(card, bg=t.get("border", "#252c39"), height=1).pack(
            fill="x", padx=24, pady=(10, 0))
        tk.Label(card, text="🔒  القفل التلقائي عند الخمول", bg=CARD, fg=GOLD,
                 font=("Segoe UI", 11, "bold")).pack(anchor="e", padx=24, pady=(14, 2))
        tk.Label(card, text="عند ترك التطبيق دون استخدام يُقفل نفسه ويطلب كلمة المرور للعودة.",
                 bg=CARD, fg=MUT, font=("Segoe UI", 11), justify="right",
                 wraplength=520).pack(anchor="e", padx=24)

        v_on = tk.BooleanVar(value=int(getattr(self, "lock_minutes", 0) or 0) > 0)
        cur_mins = int(getattr(self, "lock_minutes", 0) or 0) or 10
        v_mins = tk.StringVar(value=str(cur_mins))
        row = tk.Frame(card, bg=CARD)
        row.pack(fill="x", padx=24, pady=(10, 2))
        tk.Label(row, text="دقيقة من دون استخدام", bg=CARD, fg=MUT,
                 font=("Segoe UI", 10)).pack(side="right", padx=(6, 8))
        om = tk.OptionMenu(row, v_mins, *["1", "2", "5", "10", "15", "30", "60"])
        om.configure(bg=t.get("entry_bg", "#151a23"), fg=TXT, bd=0,
                     highlightthickness=0, activebackground=CARD,
                     activeforeground=GOLD, font=("Segoe UI", 11),
                     relief="flat", cursor="hand2", width=4, anchor="center")
        om["menu"].configure(bg=t.get("card2", "#1a2029"), fg=TXT,
                             activebackground=t["accent"], activeforeground=t["accent_text"],
                             font=("Segoe UI", 11), bd=0)
        om.pack(side="right")
        tk.Checkbutton(row, text="مفعّل", variable=v_on, bg=CARD, fg=TXT,
                       selectcolor=t.get("entry_bg", "#151a23"),
                       activebackground=CARD, activeforeground=GOLD,
                       bd=0, highlightthickness=0, cursor="hand2",
                       font=("Segoe UI", 10)).pack(side="right", padx=(0, 10))

        st_lock = tk.Label(card, text="", bg=CARD, fg=t.get("profit", "#3ddc97"),
                           font=("Segoe UI", 10))
        st_lock.pack(anchor="e", padx=24, pady=(4, 0))

        def save_lock():
            try:
                mins = max(1, int(v_mins.get()))
            except (TypeError, ValueError):
                mins = 10
            self.lock_minutes = mins if v_on.get() else 0
            self.save()
            self._last_activity = time.monotonic()
            st_lock.config(text=("سيعمل القفل بعد {} دقيقة من الخمول ✔".format(mins)
                                 if self.lock_minutes else "القفل التلقائي معطّل"))

        self._mkbtn(card, "✔  حفظ إعداد القفل", save_lock, kind="primary",
                    font_size=10).pack(anchor="e", padx=24, pady=(8, 0))
        tk.Frame(card, bg=t.get("border", "#252c39"), height=1).pack(fill="x", padx=24, pady=(8, 0))
        foot = tk.Frame(card, bg=CARD)
        foot.pack(fill="x", padx=24, pady=(12, 16))
        self._mkbtn(foot, "تسجيل الخروج", self.logout, kind="danger",
                    font_size=10).pack(side="right")
        tk.Label(foot, text="عند الخروج تُغلق جلستك وتبقى بياناتك محفوظة.",
                 bg=CARD, fg=MUT, font=("Segoe UI", 10)).pack(side="right", padx=10)

    # ============================ حسابات ============================
    def rate(self):
        return self.rates.get(self.cur, 1.0)

    def money(self, usd_amount):
        return "{}{:,.2f}".format(CUR[self.cur]["sym"], usd_amount * self.rate())

    def gram_market_disp(self):
        """سعر الغرام السوقي (بدون طرح) بالعملة المعروضة."""
        return (self.spot_oz_usd / OZ_TO_G) * self.rate()

    def gram_eff_disp(self):
        """سعر الراجحي للبيع (السوقي ناقص فرق البيع) للغرام."""
        if self.cur != "SAR":
            # فرق الراجحي بالريال السعودي فقط — لغير الريال نعرض السوقي
            return self.gram_market_disp()
        return self.gram_market_disp() - self.adjust

    def gram_buy_disp(self):
        """سعر الراجحي للشراء (السوقي زائد فرق الشراء) للغرام."""
        if self.cur != "SAR":
            return self.gram_market_disp()
        return self.gram_market_disp() + self.buy_margin

    def portfolio_value_disp(self):
        """القيمة الحالية للمحفظة بعملة العرض (وزن كل عملية × سعر عيارها)."""
        total = 0.0
        for h in (getattr(self, "holdings", None) or []):
            try:
                total += float(h["qty"]) * self.current_per_unit_disp(
                    h.get("unit", "g"), h.get("karat", 24))
            except Exception:
                continue
        return total

    def portfolio_profit_disp(self):
        """Unrealized P/L of remaining holdings, in the portfolio display currency."""
        cost=0.0
        for holding in getattr(self,"holdings",[]):
            try:
                cost+=float(holding["qty"])*float(holding["buy"])*self.rate()
            except (KeyError,ValueError,TypeError):
                continue
        return self.portfolio_value_disp()-cost

    def current_per_unit_disp(self, unit, karat=24):
        """السعر الحالي للوحدة بعيار محدد (الراجحي للبيع إن كانت العملة سعودية، وإلا السوقي)."""
        g = self.gram_eff_disp() * karat_factor(karat)
        return g * OZ_TO_G if unit == "oz" else g

    def widget_prices(self):
        """سعر الغرام (بيع/شراء) بالريال السعودي والدينار الأردني حسب الراجحي.
        فرق الراجحي بالريال السعودي يُحوّل للدينار بنفس سعر الصرف."""
        oz = float(getattr(self, "spot_oz_usd", 0.0) or 0.0)
        if oz <= 0:
            return None
        kf = karat_factor(getattr(self, "karat", 24))
        sar = float(self.rates.get("SAR", 3.75) or 3.75)
        jod = float(self.rates.get("JOD", 0.709) or 0.709)
        g_usd = oz / OZ_TO_G
        sar_mkt = g_usd * sar
        jod_mkt = g_usd * jod
        sar_per_jod = (sar / jod) if jod else 1.0
        adjust = float(getattr(self, "adjust", 0.0) or 0.0)
        margin = float(getattr(self, "buy_margin", 0.0) or 0.0)
        adj_jod = adjust / sar_per_jod if sar_per_jod else 0.0
        mar_jod = margin / sar_per_jod if sar_per_jod else 0.0
        return {
            "sar": {"mkt": sar_mkt * kf, "sell": (sar_mkt - adjust) * kf,
                    "buy": (sar_mkt + margin) * kf},
            "jod": {"mkt": jod_mkt * kf, "sell": (jod_mkt - adj_jod) * kf,
                    "buy": (jod_mkt + mar_jod) * kf},
            "oz_usd": oz, "adjust_sar": adjust,
        }

    def total_open_grams(self):
        return sum(h["qty"] * (OZ_TO_G if h["unit"] == "oz" else 1)
                   for h in self.holdings)

    def avg_buy_per_gram_usd(self):
        g = self.total_open_grams()
        if not g:
            return 0.0
        return sum(h["qty"] * h["buy"] for h in self.holdings) / g

    # ============================ مساعدات واجهة ============================
    def _mkbtn(self, parent, text, cmd, kind="primary", font_size=11, **kw):
        t = self.theme
        if text not in ("🙈", "👁", "🌓", "☁"):
            text = re.sub(r"[\U0001F300-\U0001FAFF\u2700-\u27FF\u2190-\u21FF\u2B00-\u2BFF]", "", text).strip()
        palettes = {
            "primary": (t["accent"], t["accent_text"], t.get("hover", "#f3d8a5")),
            "success": (t["card2"], t["profit"], t["border"]),
            "danger": (t["card2"], t["loss"], t["border"]),
            "neutral": (t["card2"], t["text"], t.get("border", "#252c39")),
            "ghost": (t["bg"], t["accent"], t.get("card2", "#1a2029")),
        }
        bg, fg, hv = palettes.get(kind, palettes["primary"])
        btn = tk.Button(parent, text=text, command=cmd, bg=bg, fg=fg,
                        activebackground=hv, activeforeground=fg, relief="flat",
                        bd=0, padx=12, pady=7, cursor="hand2",
                        highlightthickness=1,
                        highlightbackground=t.get("border", "#252c39"),
                        highlightcolor=t["accent"],
                        font=("Segoe UI", font_size, "bold"), **kw)

        def _on_enter(_e=None):
            try:
                btn.configure(bg=hv)
            except Exception:
                pass

        def _on_leave(_e=None):
            try:
                btn.configure(bg=bg)
            except Exception:
                pass

        btn.bind("<Enter>", _on_enter)
        btn.bind("<Leave>", _on_leave)
        tips = {"🙈": "إخفاء الأرقام لحماية الخصوصية", "👁": "إظهار الأرقام",
                "🌓": "تبديل المظهر الداكن والفاتح", "☁": "مزامنة الحساب الآن"}
        if text in tips:
            self._attach_tooltip(btn, tips[text])
        return btn

    def _mkentry(self, parent, var, width=10):
        t = self.theme
        return tk.Entry(parent, textvariable=var, width=width, justify="center",
                        bg=t.get("entry_bg", "#151a23"), fg=t["text"],
                        insertbackground=t["text"], relief="flat", bd=0,
                        highlightthickness=1,
                        highlightbackground=t.get("border", "#252c39"),
                        highlightcolor=t.get("accent", "#e9c684"),
                        font=("Segoe UI", 10))

    def _rounded_surface(self, parent, background):
        surface = tk.Frame(parent, bg=parent.cget("bg"))
        canvas = tk.Canvas(surface, bg=parent.cget("bg"), bd=0, highlightthickness=0)
        canvas.place(x=0, y=0, relwidth=1, relheight=1)
        content = tk.Frame(surface, bg=background)
        content.pack(fill="both", expand=True, padx=8, pady=6)
        def draw(event):
            w, h = event.width - 1, event.height - 1
            r = min(self._px(12), w // 2, h // 2)
            points = [r, 1, w-r, 1, w, 1, w, r, w, h-r, w, h,
                      w-r, h, r, h, 1, h, 1, h-r, 1, r, 1, 1]
            canvas.delete("surface")
            canvas.create_polygon(points, smooth=True, splinesteps=24,
                                  fill=background, outline=self.theme["border"],
                                  width=1, tags="surface")
        surface.bind("<Configure>", draw)
        return surface, content

    def _price_card(self, parent, title, attr, color, col):
        t = self.theme
        surface, card = self._rounded_surface(parent, t["card"])
        surface.grid(row=0, column=col, padx=7, pady=(0, 4), sticky="nsew")
        tk.Frame(card, bg=t["accent"] if "rajhi" in attr else t["border"], height=2).pack(fill="x")
        title_lbl = tk.Label(card, text=title, bg=t["card"], fg=t["muted"],
                             font=("Segoe UI", 10))
        title_lbl.pack(pady=(5, 2))
        lbl = tk.Label(card, text="--", bg=t["card"], fg=color,
                       font=("Segoe UI", 20, "bold"))
        lbl.pack(pady=(0, 7))
        setattr(self, attr, lbl)
        setattr(self, attr + "_title", title_lbl)
        self._card_base_titles = getattr(self, "_card_base_titles", {})
        self._card_base_titles[attr] = title
        parent.grid_columnconfigure(col, weight=1)
        return surface

    # ============================ بناء الواجهة ============================
    def _modern_saved_theme(self, saved):
        """Refresh stock colors only; retain all user-customized colors."""
        legacy = LEGACY_DARK_THEME if getattr(self, "dark_mode", True) else LEGACY_LIGHT_THEME
        modern = self._base_theme()
        return {key: modern.get(key, value) if value == legacy.get(key) else value
                for key, value in (saved or {}).items()}

    def _attach_tooltip(self, widget, text):
        state = {"timer": None, "window": None}
        def hide(_event=None):
            if state["timer"] is not None:
                try: widget.after_cancel(state["timer"])
                except tk.TclError: pass
                state["timer"] = None
            if state["window"] is not None:
                try: state["window"].destroy()
                except tk.TclError: pass
                state["window"] = None
        def show():
            state["timer"] = None
            if not widget.winfo_exists(): return
            t = self.theme
            tip = tk.Toplevel(widget)
            state["window"] = tip
            tip.overrideredirect(True)
            tk.Label(tip, text=text, bg=t["card2"], fg=t["text"],
                     font=("Segoe UI", 10), padx=10, pady=6, relief="solid", bd=1).pack()
            tip.update_idletasks()
            x = min(widget.winfo_rootx(), tip.winfo_screenwidth() - tip.winfo_reqwidth() - 12)
            y = widget.winfo_rooty() + widget.winfo_height() + 6
            tip.geometry("+{}+{}".format(max(0, x), max(0, y)))
        def schedule(_event=None):
            hide()
            state["timer"] = widget.after(550, show)
        widget.bind("<Enter>", schedule, add="+")
        for event in ("<Leave>", "<ButtonPress>", "<Destroy>"):
            widget.bind(event, hide, add="+")

    def _base_theme(self):
        """لوحة الألوان الأساسية حسب المظهر المختار."""
        return LIGHT_THEME if not getattr(self, "dark_mode", True) else DEFAULT_THEME

    def _hide(self, txt):
        """وضع الخصوصية: يستبدل الأرقام برموز عند تفعيله."""
        out = "" if txt is None else str(txt)
        if not getattr(self, "privacy", False):
            return out
        return re.sub(r"\d[\d,\.]*", "•••", out)

    # ==================== ويدجيت سعر الغرام ====================
    def _refresh_widget(self):
        w = getattr(self, "price_widget", None)
        if w is None:
            return
        try:
            w.refresh()
        except Exception:
            pass

    def toggle_widget(self):
        """إظهار/إخفاء ويدجيت سعر الغرام (F9)."""
        if getattr(self, "price_widget", None) is not None:
            self.close_widget()
        else:
            self.open_widget()

    def open_widget(self):
        if not getattr(self, "user", None):
            return
        if getattr(self, "price_widget", None) is not None:
            try:
                self.price_widget.win.lift()
            except Exception:
                pass
            return
        try:
            self.price_widget = PriceWidget(self)
        except Exception as e:
            self.price_widget = None
            messagebox.showerror("ويدجيت السعر",
                                 "تعذّر فتح الويدجيت: {}".format(e))
            return
        self._widget_on = True
        self._paint_widget_btn()
        self.save()
        self._refresh_widget()

    def close_widget(self):
        w = getattr(self, "price_widget", None)
        self.price_widget = None
        self._widget_on = False
        if w is not None:
            try:
                w.destroy()
            except Exception:
                pass
        self._paint_widget_btn()
        self.save()

    def _paint_widget_btn(self):
        btn = getattr(self, "btn_widget", None)
        if btn is None:
            return
        try:
            btn.config(text="🪟 إخفاء"
                       if getattr(self, "price_widget", None) is not None
                       else "🪟 ويدجيت")
        except Exception:
            pass

    def widget_karat_step(self, step):
        """تبديل عيار الويدجيت (يُزامن عيار التطبيق كله)."""
        ks = list(KARATS)
        try:
            i = ks.index(int(getattr(self, "karat", 24)))
        except (TypeError, ValueError):
            i = 0
        k = ks[(i + step) % len(ks)]
        try:
            self.karat_var.set("{} قيراط".format(k))
        except Exception:
            self.karat = k
        self._refresh_widget()

    def toggle_theme(self):
        """تبديل المظهر بين الداكن والفاتح (يُحفظ لكل حساب)."""
        page = getattr(self, "_page", "main")
        self.dark_mode = not getattr(self, "dark_mode", True)
        self.theme = dict(self._base_theme())
        self.save()
        self.rebuild_ui()
        self._show_page(page if page in ("main", "chart") else "main")

    def toggle_privacy(self):
        """إخفاء/إظهار كل الأرقام بضغطة واحدة (Ctrl+H)."""
        self.privacy = not getattr(self, "privacy", False)
        btn = getattr(self, "btn_privacy", None)
        if btn is not None:
            try:
                btn.configure(text="👁" if self.privacy else "🙈")
            except Exception:
                pass
        try:
            self.render()
            self.update_price_labels()
        except Exception:
            pass

    def _dash_card(self, parent, title, attr, color, col):
        t = self.theme
        surface, card = self._rounded_surface(parent, t["card2"])
        surface.grid(row=0, column=col, padx=6, sticky="nsew")
        tk.Label(card, text=title, bg=t.get("card2", t["card"]), fg=t["muted"],
                 font=("Segoe UI", 10)).pack(pady=(2, 0))
        lbl = tk.Label(card, text="--", bg=t.get("card2", t["card"]), fg=color,
                       font=("Segoe UI", 16, "bold"))
        lbl.pack(pady=(2, 3))
        setattr(self, attr, lbl)
        parent.grid_columnconfigure(col, weight=1)
        return surface
        return card

    def _update_dashboard(self, qty_g, cost, val, pl):
        if not hasattr(self, "dash_qty"):
            return
        t = self.theme
        sym = CUR[self.cur]["sym"]
        self.dash_qty.config(text=self._hide("{:,.2f} غ".format(qty_g)))
        self.dash_cost.config(text=self._hide("{}{:,.2f}".format(sym, cost)))
        self.dash_val.config(text=self._hide("{}{:,.2f}".format(sym, val)))
        color = t["profit"] if pl > 0 else (t["loss"] if pl < 0 else t["total"])
        self.dash_pl.config(text=self._hide("{}{}{:,.2f}".format(
            "+" if pl > 0 else "", sym, pl)), fg=color)

    def _px(self, n):
        """يحوّل مقاساً منطقياً إلى بكسل حسب مقياس الشاشة الحالي."""
        return max(2, int(round(n * getattr(self, "ui_scale", 1.0))))

    def _screen_size(self):
        try:
            return (self.root.winfo_screenwidth(), self.root.winfo_screenheight())
        except Exception:
            return (1366, 768)

    # ------------------------- الأيقونة والتمرير -------------------------
    def _apply_icon(self):
        """أيقونة النافذة وشريط المهام (مكتوبة من داخل البرنامج نفسه)."""
        p = app_icon_path()
        if not p:
            return
        try:
            self.root.iconbitmap(default=p)
            return
        except Exception:
            pass
        try:
            self.root.iconbitmap(p)
        except Exception:
            pass

    def _bind_wheel(self):
        """تمرير موحّد بعجلة الفأرة: يمرّر العنصر تحت المؤشر أو الصفحة."""
        if getattr(self, "_wheel_bound", False):
            return
        self._wheel_bound = True
        try:
            self.root.bind_all("<MouseWheel>", self._on_wheel, add="+")
            self.root.bind_all("<Button-4>",
                               lambda e: self._on_wheel(e, -1), add="+")
            self.root.bind_all("<Button-5>",
                               lambda e: self._on_wheel(e, 1), add="+")
        except Exception:
            pass

    @staticmethod
    def _widget_can_scroll(w, notches, horizontal, cls):
        """يمرّر جدولاً/قائمة إن كانت قابلة للتمرير في هذا الاتجاه."""
        try:
            if horizontal and cls == "Treeview":
                w.xview_scroll(int(notches), "units")
                return True
        except Exception:
            pass
        try:
            first, last = w.yview()
            first, last = float(first), float(last)
        except Exception:
            return False
        if (last - first) >= 0.999:
            return False
        if notches < 0 and first <= 0.001:
            return False
        if notches > 0 and last >= 0.999:
            return False
        try:
            w.yview_scroll(int(notches), "units")
            return True
        except Exception:
            return False

    def _on_wheel(self, event, direction=None):
        """يمرّر أعمق عنصر قابل للتمرير تحت المؤشر، وإلا يمرّر صفحته."""
        try:
            node = self.root.winfo_containing(event.x_root, event.y_root)
        except Exception:
            node = None
        if direction is None:
            delta = getattr(event, "delta", 0) or 0
            if delta == 0:
                return None
            notches = -int(round(float(delta) / 120.0))
            if notches == 0:
                notches = -1 if delta > 0 else 1
        else:
            notches = 3 * int(direction)
        try:
            shift = bool(int(getattr(event, "state", 0)) & 0x0001)
        except Exception:
            shift = False

        target = None
        while node is not None:
            try:
                cls = node.winfo_class()
            except Exception:
                cls = ""
            if cls in ("TCombobox", "Spinbox"):
                return "break"          # لا نغيّر قيمتها بالعجلة سهواً
            if cls in ("Treeview", "Listbox", "Text"):
                if self._widget_can_scroll(node, notches, shift, cls):
                    return "break"
            if isinstance(node, ScrollArea):
                target = node
                break
            node = getattr(node, "master", None)
        if target is not None:
            try:
                target.scroll_units(notches * 3, horizontal=shift)
                return "break"
            except Exception:
                return None
        return None

    # ------------------------- إشعارات سطح المكتب -------------------------
    def notify(self, title, message):
        """إشعار سطح مكتب حقيقي + بطاقة داخل التطبيق."""
        return desktop_notify(title, message, app=self)

    def toast(self, title, message, system_ok=False):
        """بطاقة إشعار صغيرة أسفل الشاشة تختفي وحدها بعد ثوانٍ."""
        t = self.theme
        old = getattr(self, "_toast", None)
        if old is not None:
            try:
                old.destroy()
            except Exception:
                pass
        try:
            win = tk.Toplevel(self.root)
            win.overrideredirect(True)
            try:
                win.attributes("-topmost", True)
            except Exception:
                pass
            bg = t.get("card2", t["card"])
            box = tk.Frame(win, bg=bg, highlightthickness=2,
                           highlightbackground=t.get("accent", "#e9c684"))
            box.pack(fill="both", expand=True)
            tk.Label(box, text="🔔  " + str(title), bg=bg,
                     fg=t.get("accent", "#e9c684"),
                     font=("Segoe UI", 12, "bold"), anchor="e").pack(
                fill="x", padx=16, pady=(12, 2))
            tk.Label(box, text=str(message), bg=bg, fg=t.get("text", "#fff"),
                     font=("Segoe UI", 10), anchor="e", justify="right",
                     wraplength=360).pack(fill="x", padx=16, pady=(0, 12))
            win.bind("<Button-1>", lambda e: win.destroy())
            for child in box.winfo_children():
                child.bind("<Button-1>", lambda e: win.destroy())
            win.update_idletasks()
            sw, sh = self._screen_size()
            ww = min(max(300, win.winfo_reqwidth()), max(320, sw - 60))
            wh = win.winfo_reqheight()
            win.geometry("{}x{}+{}+{}".format(ww, wh, max(20, sw - ww - 26),
                                              max(20, sh - wh - 80)))
            self._toast = win

            def _bye(w=win):
                try:
                    if w.winfo_exists():
                        w.destroy()
                except Exception:
                    pass

            win.after(9000, _bye)
        except Exception:
            pass

    def _apply_screen_layout(self):
        """قياس واضح للخطوط + نافذة مناسبة للشاشة. لا نصغّر الخطوط أبداً؛
        وإذا ضاق المكان فالتمرير يكشف بقية المحتوى بدل قصّه."""
        sw, sh = self._screen_size()
        # 1.0 = حد أدنى (خط واضح دائماً)، والتكبير فقط على الشاشات الكبيرة
        dpi_k = max(1.0, min(1.6, (globals().get("SYSTEM_DPI", 96) or 96) / 96.0))
        scale = dpi_k
        if sw >= 1900 and sh >= 1040 and scale < 1.1:
            scale = 1.10
        elif sw >= 1600 and sh >= 900 and scale < 1.05:
            scale = 1.05
        self.ui_scale = scale
        self.screen_w, self.screen_h = sw, sh
        try:
            self.root.tk.call("tk", "scaling", 1.3333 * scale)
        except Exception:
            pass
        w = int(min(1280 * scale, max(880, sw - 60)))
        h = int(min(880 * scale, max(600, sh - 90)))
        self.root.geometry("{}x{}".format(w, h))
        self.root.minsize(880, 560)

    def _build_sidebar(self):
        t = self.theme
        self._nav_buttons = {}
        self._nav_labels = {}
        tk.Label(self.sidebar, text="التنقل", bg=t["card"], fg=t["muted"],
                 font=("Segoe UI", 10)).pack(fill="x", pady=(18, 12))
        items = [("main", "المحفظة", lambda: self._show_page("main")),
                 ("chart", "المخطط", self.show_chart),
                 ("history", "سجل العمليات", self.open_history),
                 ("alerts", "التنبيهات", self.open_alerts),
                 ("zakat", "الزكاة", self.open_zakat),
                 ("settings", "الإعدادات", self.open_settings)]
        self._nav_images = {}
        for key, label, command in items + [("account", "الحساب", self.open_account)]:
            button = self._mkbtn(self.sidebar, label, command, kind="neutral", font_size=11)
            button.configure(anchor="e", padx=12, pady=11, bg=t["card"], relief="flat")
            try:
                folder = getattr(sys, "_MEIPASS", None) or os.path.dirname(os.path.abspath(__file__))
                mode = "dark" if getattr(self, "dark_mode", True) else "light"
                image = tk.PhotoImage(file=os.path.join(folder, "nav_{}_{}.png".format(mode, key)))
                self._nav_images[key] = image
                button.configure(image=image, compound="right")
            except (tk.TclError, OSError):
                pass
            if key == "account":
                button.pack(side="bottom", fill="x", padx=8, pady=16)
            else:
                button.pack(fill="x", padx=8, pady=3)
            self._nav_buttons[key] = button
            self._nav_labels[key] = label
            self._attach_tooltip(button, label)
            button.bind("<Leave>", lambda event: self._paint_navigation(getattr(self, "_page", "main")), add="+")
        self.btn_settings = self._nav_buttons["settings"]
        self._sidebar_compact = None

    def _paint_navigation(self, page):
        t = self.theme
        active = "main" if page in ("addbuy", "sell", "editholding") else page
        for key, button in getattr(self, "_nav_buttons", {}).items():
            button.configure(bg=t["total_bg"] if key == active else t["card"],
                             fg=t["accent"] if key == active else t["text"])
        title = getattr(self, "page_title", None)
        if title is not None:
            title.configure(text={"main":"المحفظة", "chart":"المخطط", "history":"سجل العمليات",
                                  "alerts":"التنبيهات", "zakat":"الزكاة", "settings":"الإعدادات",
                                  "account":"الحساب", "addbuy":"إضافة شراء", "sell":"أمر بيع",
                                  "editholding":"تعديل المقتنى"}.get(page, "المحفظة"))

    def _responsive_shell(self, event=None):
        if event is not None and event.widget is not self.app_shell:
            return
        compact = self.app_shell.winfo_width() < self._px(1120)
        if compact == getattr(self, "_sidebar_compact", None):
            return
        self._sidebar_compact = compact
        self.sidebar.configure(width=self._px(64 if compact else 174))
        for key, button in self._nav_buttons.items():
            button.configure(text="" if compact and key in self._nav_images else self._nav_labels[key],
                             anchor="center" if compact else "e", padx=6 if compact else 12)
        self._paint_navigation(getattr(self, "_page", "main"))

    def _responsive_cards(self, event=None):
        summaries = getattr(self, "_summary_cards", [])
        prices = getattr(self, "_price_cards", [])
        if not summaries or not prices:
            return
        width = self.main_wrap.canvas.winfo_width()
        cols = 2 if width < self._px(860) else 4
        key = (cols, self.cur == "SAR")
        if key == getattr(self, "_cards_layout_key", None):
            return
        self._cards_layout_key = key
        for cards in (summaries, prices):
            parent = cards[0].master
            for column in range(4): parent.grid_columnconfigure(column, weight=0, uniform="", minsize=0)
            visible = cards if cards is summaries or self.cur == "SAR" else cards[:2]
            for card in cards: card.grid_remove()
            for index, card in enumerate(visible):
                # Main financial hierarchy begins on the right.
                column = cols - 1 - index % cols
                card.grid(row=index // cols, column=column, sticky="nsew", padx=6, pady=4)
                parent.grid_columnconfigure(column, weight=1, uniform="cards")

    def _sync_selection_actions(self, event=None):
        tree = getattr(self, "tree", None)
        selected = tree.selection() if tree is not None and hasattr(tree, "selection") else ()
        mapped = getattr(self, "row_map", {}).get(selected[0]) if selected else None
        delete = getattr(self, "btn_delete", None)
        edit = getattr(self, "btn_edit", None)
        if delete is not None: delete.configure(state="normal" if mapped else "disabled")
        if edit is not None: edit.configure(state="normal" if mapped and mapped[0] == "hold" else "disabled")

    def _operations_more(self):
        t = self.theme
        menu = tk.Menu(self.root, tearoff=False, bg=t["card"], fg=t["text"],
                       activebackground=t["card2"], activeforeground=t["accent"],
                       font=("Segoe UI", 11))
        menu.add_command(label="مسح جميع العمليات…", command=self.clear_all)
        try:
            menu.tk_popup(self.btn_more.winfo_rootx(), self.btn_more.winfo_rooty() + self.btn_more.winfo_height())
        finally:
            menu.grab_release()

    def _form_section(self, parent, title):
        t = self.theme
        frame = tk.LabelFrame(parent, text="  " + title + "  ", bg=t["card"], fg=t["accent"],
                              font=("Segoe UI", 11, "bold"), bd=0,
                              highlightthickness=1, highlightbackground=t["border"], padx=12, pady=10)
        frame.pack(fill="x", padx=8, pady=8)
        return frame

    def _form_actions(self, key):
        footer = self._pages[key]["footer"]
        for child in footer.winfo_children(): child.destroy()
        actions = tk.Frame(footer, bg=self.theme["bg"])
        actions.pack(fill="x", pady=4)
        return actions

    def build_ui(self):
        self._live_chart_views = []
        self._chart_widgets = []
        t = self.theme
        BG, CARD, CARD2 = t["bg"], t["card"], t["card2"]
        GOLD, ATXT, TXT, MUT = t["accent"], t["accent_text"], t["text"], t["muted"]
        self.root.configure(bg=BG)
        self.root.title(self.titles.get("window", DEFAULT_TITLES["window"]) + " — " + APP_VERSION)
        self._apply_screen_layout()
        p = self._px

        # حاوية الصفحات: الرئيسية + المخطط (تنقّل داخل نفس النافذة)
        self.app_shell = tk.Frame(self.root, bg=BG)
        self.app_shell.pack(fill="both", expand=True)
        self.sidebar = tk.Frame(self.app_shell, bg=CARD, width=p(174))
        self.sidebar.pack(side="right", fill="y", padx=(p(8), 0))
        self.sidebar.pack_propagate(False)
        self.page_holder = tk.Frame(self.app_shell, bg=BG)
        self.page_holder.pack(side="left", fill="both", expand=True)
        # الصفحة الرئيسية داخل منطقة تمرير: لا يُقصّ أي محتوى مهما صغُر الارتفاع
        self.main_wrap = ScrollArea(self.page_holder, BG)
        self.main_wrap.pack(fill="both", expand=True)
        self.main_page = self.main_wrap.inner
        try:
            self._apply_icon()
        except Exception:
            pass
        if not getattr(self, "_wheel_bound", False):
            self._bind_wheel()
            self._wheel_bound = True

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Treeview", background=CARD, fieldbackground=CARD,
                        foreground=TXT, rowheight=self._px(37),
                        font=("Segoe UI", 11), borderwidth=0,
                        highlightthickness=0)
        style.configure("Treeview.Heading", background=CARD2, foreground=GOLD,
                        font=("Segoe UI", 11, "bold"), relief="flat",
                        padding=(p(9), p(8)))
        style.map("Treeview",
                  background=[("selected", t.get("border", "#2c3a52"))],
                  foreground=[("selected", t["text"])])
        style.map("Treeview.Heading",
                  background=[("active", t.get("border", "#252c39"))])
        try:
            style.configure("TMenubutton", background=CARD2, foreground=TXT,
                            relief="flat", padding=(12, 5),
                            font=("Segoe UI", 10))
            style.map("TMenubutton",
                      background=[("active", t.get("border", "#252c39"))])
        except Exception:
            pass

        # رأس مختصر مع هوية بصرية ثابتة
        hdr = tk.Frame(self.root, bg=BG)
        hdr.pack(fill="x", padx=p(20), pady=(p(12), p(8)), before=self.app_shell)
        hdr.grid_columnconfigure(0, weight=1)
        hdr.grid_columnconfigure(1, weight=1)
        hdr.grid_columnconfigure(2, weight=1)
        brand = tk.Frame(hdr, bg=BG)
        brand.grid(row=0, column=1)
        try:
            asset_dir = getattr(sys, "_MEIPASS", None) or os.path.dirname(os.path.abspath(__file__))
            image = tk.PhotoImage(file=os.path.join(asset_dir, "brand_badge.png"))
            self.root.iconphoto(True, image)
            self._brand_image = image.subsample(2, 2)
            tk.Label(brand, image=self._brand_image, bg=BG).pack(side="right", padx=6)
        except (tk.TclError, OSError):
            self._brand_image = None
        tk.Label(brand, text="محفظة الذهب", bg=BG, fg=GOLD,
                 font=("Segoe UI", 14, "bold")).pack(side="right")
        self.page_title = tk.Label(hdr, text="المحفظة", bg=BG, fg=TXT,
                                   font=("Segoe UI", 14, "bold"))
        self.page_title.grid(row=0, column=2, sticky="e")
        if getattr(self, "user", None):
            chip = tk.Frame(hdr, bg=BG)
            chip.grid(row=0, column=0, sticky="w")
            self.btn_privacy = self._mkbtn(chip, "👁" if getattr(self, "privacy", False) else "🙈",
                                          self.toggle_privacy, kind="ghost", font_size=10)
            self.btn_privacy.pack(side="left", padx=2)
            self._mkbtn(chip, "🌓", self.toggle_theme, kind="ghost", font_size=10).pack(side="left", padx=2)
            self.btn_cloud = self._mkbtn(chip, "☁", lambda: self.cloud_sync_now(True), kind="ghost", font_size=10)
            self.btn_cloud.pack(side="left", padx=2)
            tk.Label(chip, text=self.user, bg=BG, fg=MUT, font=("Segoe UI", 10)).pack(side="left", padx=6)
        self._build_sidebar()
        self.app_shell.bind("<Configure>", self._responsive_shell, add="+")
        tk.Frame(self.root, bg=t["border"], height=1).pack(fill="x", before=self.app_shell)

        # بطاقات الأسعار (تُحزَم لاحقاً أسفل سطر العملة والعيار)
        prices = tk.Frame(self.main_page, bg=BG)
        self._price_card(prices, "الأونصة (عالمي)", "lbl_oz", GOLD, 0)
        self._price_card(prices, "الغرام (عالمي)", "lbl_gr", GOLD, 1)
        self.card_rajhi_buy = self._price_card(
            prices, "الراجحي — شراء/غرام", "lbl_gr_buy", t["loss"], 2)
        self.card_rajhi_sell = self._price_card(
            prices, "الراجحي — بيع/غرام", "lbl_gr_adj", t["profit"], 3)

        # الإعدادات (سطر العملة والعيار — أول عناصر الصفحة)
        ctrl = tk.Frame(self.main_page, bg=BG)
        ctrl.pack(fill="x", padx=p(20), pady=(p(12), p(4)))
        tk.Label(ctrl, text="العملة:", bg=BG, fg=TXT,
                 font=("Segoe UI", 10)).pack(side="right", padx=(10, 2))
        self.cur_var = tk.StringVar(value="{} ({})".format(CUR[self.cur]["name"], self.cur))
        cur_menu = ttk.OptionMenu(ctrl, self.cur_var, self.cur_var.get(),
                                  *["{} ({})".format(v["name"], k) for k, v in CUR.items()])
        cur_menu.pack(side="right", ipady=2)

        # اختيار العيار (24 / 22 / 21 / 18)
        tk.Label(ctrl, text="العيار:", bg=BG, fg=TXT,
                 font=("Segoe UI", 10)).pack(side="right", padx=(14, 2))
        self.karat_var = tk.StringVar(value="{} قيراط".format(self.karat))
        karat_menu = ttk.OptionMenu(
            ctrl, self.karat_var, self.karat_var.get(),
            *["{} قيراط".format(k) for k in KARATS])
        karat_menu.pack(side="right", ipady=2)
        self.karat_var.trace_add("write", lambda *a: self.on_karat_change())

        # فروق الراجحي (تُخفى تلقائياً عند اختيار عملة غير الريال السعودي)
        market_options = tk.Frame(self.main_page, bg=BG)
        market_options.pack(fill="x", padx=p(20), pady=(0, p(4)))
        self.rajhi_frame = tk.Frame(market_options, bg=BG)
        self.rajhi_frame.pack(side="right")

        lbl_adj = tk.Label(self.rajhi_frame, text="فرق البيع/غرام:", bg=BG, fg=TXT,
                           font=("Segoe UI", 10))
        self.adj_var = tk.StringVar(value=str(self.adjust))
        self.adj_e = self._mkentry(self.rajhi_frame, self.adj_var, width=6)
        lbl_buy = tk.Label(self.rajhi_frame, text="فرق الشراء/غرام:", bg=BG, fg=TXT,
                           font=("Segoe UI", 10))
        self.buy_var = tk.StringVar(value=str(self.buy_margin))
        self.buy_e = self._mkentry(self.rajhi_frame, self.buy_var, width=6)
        # ترتيب ثابت (من اليمين): فرق البيع، قيمته، فرق الشراء، قيمته
        self._rajhi_kids = [
            (lbl_adj, {"side": "right", "padx": (12, 2)}),
            (self.adj_e, {"side": "right", "ipady": 4}),
            (lbl_buy, {"side": "right", "padx": (12, 2)}),
            (self.buy_e, {"side": "right", "ipady": 4}),
        ]
        for w, o in self._rajhi_kids:
            w.pack(**o)
        self.adj_var.trace_add("write", lambda *a: self.on_adjust_change())
        self.buy_var.trace_add("write", lambda *a: self.on_buy_margin_change())

        # ربط تغيير العملة بعد جهوزية عناصر الواجهة كلها
        self.cur_var.trace_add("write", lambda *a: self.on_currency_change())

        self._mkbtn(ctrl, "🧮 حاسبة التكلفة", self.open_calc, kind="neutral",
                    font_size=10).pack(side="left", padx=6)

        self._price_cards = [self.lbl_oz.master.master, self.lbl_gr.master.master,
                             self.card_rajhi_buy, self.card_rajhi_sell]
        self._cards_layout_key = None
        # بطاقات الأسعار تحت خيارات العملة والعيار
        prices.pack(fill="x", padx=p(16), pady=(p(8), p(8)))

        # ملاحظة تظهر عند اختيار عملة غير الريال السعودي
        self.rajhi_note = tk.Label(
            market_options, text="ℹ أسعار الراجحي متاحة بالريال السعودي فقط",
            bg=BG, fg=MUT, font=("Segoe UI", 10))

        # لوحة ملخص سريعة (تُقرأ بنظرة واحدة)
        self._summary_cards = []
        dash = tk.Frame(self.main_page, bg=BG)
        dash.pack(fill="x", padx=p(16), pady=(p(12), p(8)), before=ctrl)
        self._summary_cards.append(self._dash_card(dash, "إجمالي الغرامات", "dash_qty", TXT, 0))
        self._summary_cards.append(self._dash_card(dash, "التكلفة", "dash_cost", TXT, 1))
        self._summary_cards.append(self._dash_card(dash, "القيمة الحالية", "dash_val", GOLD, 2))
        self._summary_cards.append(self._dash_card(dash, "الربح غير المحقق", "dash_pl", t["profit"], 3))

        # Caption and price change occupy different rows: neither can cover the other.
        updfr = tk.Frame(self.main_page, bg=BG)
        updfr.pack(fill="x", padx=20, pady=(2,4))
        self.lbl_upd = tk.Label(updfr, text="", bg=BG, fg=MUT,
                                font=("Segoe UI", 10), cursor="hand2",anchor="e",justify="right")
        self.lbl_upd.pack(side="right",fill="x",expand=True)
        self.lbl_upd.bind("<Button-1>",lambda e:self.refresh_all())
        updfr.bind("<Configure>",lambda e:self.lbl_upd.config(wraplength=max(180,e.width-140)))
        self._mkbtn(updfr, "🔍 تشخيص", self.price_diag, kind="ghost",
                    font_size=10).pack(side="left", padx=(8,0))
        changefr=tk.Frame(self.main_page,bg=BG)
        changefr.pack(fill="x",padx=20,pady=(0,4))
        tk.Label(changefr,text="تغير السعر:",bg=BG,fg=MUT,font=("Segoe UI",10)).pack(side="right",padx=6)
        self.lbl_delta = tk.Label(changefr, text="", bg=BG, fg=MUT,
                                  font=("Segoe UI", 11, "bold"))
        self.lbl_delta.pack(side="right",padx=4)
        self._set_update_caption()

        tk.Label(self.main_page, text="العمليات", bg=BG, fg=TXT,
                 font=("Segoe UI", 13, "bold"), anchor="e").pack(fill="x", padx=p(20), pady=(p(10), p(4)))
        # شريط الإضافة: العمليات
        addbar = tk.Frame(self.main_page, bg=BG)
        addbar.pack(fill="x", padx=p(18), pady=(p(4), p(2)))
        self.btn_add = self._mkbtn(addbar, "+  إضافة عملية (شراء / بيع)",
                                   self.open_add_chooser, kind="primary",
                                   font_size=12)
        self.btn_add.pack(side="right", pady=2)
        self._attach_tooltip(self.btn_add, "اختر «شراء» أو «بيع» ثم أدخل الوزن والعيار والسعر")
        # إظهار/إخفاء صفوف البيع في قائمة العمليات
        self.btn_sells = self._mkbtn(addbar, "👁 إخفاء البيع", self.toggle_sells,
                                     kind="ghost", font_size=10)
        self.btn_sells.pack(side="left", pady=2)
        self._paint_sells_btn()
        edits=tk.Frame(self.main_page,bg=BG)
        edits.pack(fill="x",padx=p(18),pady=2)
        self.btn_edit = self._mkbtn(edits, "تعديل المحدد", self.open_edit_holding, kind="neutral", font_size=10)
        self.btn_edit.pack(side="right", padx=4)
        self._mkbtn(edits,"↶ التراجع",self.undo_last_change,kind="neutral",font_size=10).pack(side="right",padx=4)
        self.btn_delete = self._mkbtn(edits, "حذف المحدد", self.del_item, kind="danger", font_size=10)
        self.btn_delete.pack(side="right", padx=4)
        self._mkbtn(edits, "Excel", self.export_excel, kind="neutral", font_size=10).pack(side="left", padx=4)
        self.btn_more = self._mkbtn(edits, "المزيد ▾", self._operations_more, kind="ghost", font_size=10)
        self.btn_more.pack(side="left", padx=4)
        self.root.bind("<Control-z>",self._undo_shortcut)
        self.root.bind("<Control-Z>",self._undo_shortcut)

        # الجدول
        cols = ("name", "status", "qty", "buy", "cost", "cur", "val",
                "sellprice", "pl", "pct")
        heads = ("الاسم", "الحالة", "الكمية", "سعر الشراء", "التكلفة",
                 "السعر الحالي", "القيمة", "سعر البيع", "الربح/الخسارة", "%")
        tree_wrap = tk.Frame(self.main_page, bg=BG)
        tree_wrap.pack(fill="x", padx=p(18), pady=p(6))
        self.tree = ttk.Treeview(tree_wrap, columns=cols, show="headings",
                                 height=0,
                                 displaycolumns=("pct", "sellprice", "buy", "cur", "pl", "val", "cost", "qty", "status", "name"))
        for c, h in zip(cols, heads):
            self.tree.heading(c, text=h)
            self.tree.column(c, anchor="center", width=110)
        self.tree.column("name", width=150, anchor="e")
        self.tree.column("status", width=110)
        self.tree.column("sellprice", width=125)
        self.tree.column("pct", width=80)
        for column in ("qty", "buy", "cost", "cur", "val", "sellprice", "pl", "pct"):
            self.tree.column(column, anchor="e")
        self.tree.bind("<Configure>", lambda event: self.tree.xview_moveto(1.0), add="+")
        self.tree.bind("<<TreeviewSelect>>", self._sync_selection_actions, add="+")
        self.tree_x = ttk.Scrollbar(tree_wrap, orient="horizontal",
                                    command=self.tree.xview)
        self.tree.configure(xscrollcommand=self.tree_x.set)
        self.tree.tag_configure("pos", foreground=t["profit"])
        self.tree.tag_configure("neg", foreground=t["loss"])
        self.tree.tag_configure("odd", background=t.get("row_alt", CARD))
        self.tree.tag_configure("even", background=CARD)
        self.tree.tag_configure("total", background=t["total_bg"],
                                foreground=t["total"], font=("Segoe UI", 11, "bold"))
        self.tree.tag_configure("totalpos", background=t["total_bg"],
                                foreground=t["profit"], font=("Segoe UI", 11, "bold"))
        self.tree.tag_configure("totalneg", background=t["total_bg"],
                                foreground=t["loss"], font=("Segoe UI", 11, "bold"))
        self.tree.tag_configure("soldgreen", background=t.get("sold_bg", "#12301f"),
                                foreground=t["profit"], font=("Segoe UI", 11, "bold"))
        self.tree.pack(side="top", fill="x")
        self.tree_x.pack(side="bottom", fill="x", pady=(2, 0))
        # نقر مزدوج على أي عملية شراء => فتح أمر البيع عليها مباشرة
        try:
            self.tree.bind("<Double-1>", self._tree_dblclick)
        except Exception:
            pass

        self.empty_operations = tk.Label(self.main_page, text="لا توجد عمليات لعرضها", bg=BG,
                                         fg=MUT, font=("Segoe UI", 10))

        # أدوات ثانوية أسفل الجدول؛ بقية الصفحات في شريط التنقل.
        btns = tk.Frame(self.main_page, bg=BG)
        btns.pack(fill="x", padx=p(18), pady=(p(4), p(2)))
        self.btn_widget = self._mkbtn(btns, "ويدجيت", self.toggle_widget, kind="neutral", font_size=10)
        self.btn_widget.pack(side="left", padx=2)
        self._paint_widget_btn()
        self.operations_toolbar = btns
        self._sync_selection_actions()
        self.main_wrap.canvas.bind("<Configure>", self._responsive_cards, add="+")

        # الملخص
        self.lbl_sum = tk.Label(self.main_page, text="", bg=BG, fg=t["total"],
                                font=("Segoe UI", 11, "bold"))
        self.lbl_sum.pack(fill="x", padx=p(20), pady=p(10))
        self._update_rajhi_visibility()
        self._responsive_cards()

        # صفحة المخطط (مخفية حتى يُطلب عرضها)
        self.chart_page = ScrollArea(self.page_holder, BG)
        self._build_chart_page()
        self._show_page(getattr(self, "_page", "main"))


    # ============================ أحداث الإعدادات ============================
    def on_currency_change(self):
        txt = self.cur_var.get()
        for k, v in CUR.items():
            if txt == "{} ({})".format(v["name"], k):
                self.cur = k
                break
        self._update_rajhi_visibility()
        self.save()
        self.render()

    def _update_rajhi_visibility(self):
        """إظهار أسعار الراجحي وفروقه فقط عند اختيار الريال السعودي."""
        show = (self.cur == "SAR")
        for card in (getattr(self, "card_rajhi_buy", None),
                     getattr(self, "card_rajhi_sell", None)):
            if card is None:
                continue
            try:
                grid = card.grid_info()
                if show and not grid:
                    card.grid()
                elif not show and grid:
                    card.grid_remove()
            except Exception:
                pass
        note = getattr(self, "rajhi_note", None)
        if note is not None:
            try:
                if show and note.winfo_manager():
                    note.pack_forget()
                elif not show and not note.winfo_manager():
                    note.pack(side="left", padx=6)
            except Exception:
                pass
        kids = getattr(self, "_rajhi_kids", None)
        if kids:
            visible = bool(kids[0][0].winfo_manager())
            if show and not visible:
                for w, o in kids:
                    w.pack(**o)
            elif not show and visible:
                for w, _o in kids:
                    w.pack_forget()
        self._responsive_cards()

    def on_karat_change(self):
        txt = str(self.karat_var.get()).split()[0]
        try:
            k = int(txt)
        except (TypeError, ValueError):
            return
        if k not in KARATS or k == self.karat:
            return
        self.karat = k
        fv = getattr(self, "form_karat_var", None)
        if fv is not None:
            fv.set("{} قيراط".format(k))
        self.save()
        self.render()

    def on_adjust_change(self):
        try:
            self.adjust = float(self.adj_var.get() or 0)
        except ValueError:
            return
        self.save()
        self.render()

    def on_buy_margin_change(self):
        try:
            self.buy_margin = float(self.buy_var.get() or 0)
        except ValueError:
            return
        self.save()
        self.render()

    # ==================== إضافة عملية (شراء / بيع) ====================
    def _close_chooser(self):
        win = getattr(self, "_choice_win", None)
        self._choice_win = None
        if win is not None:
            try:
                win.destroy()
            except Exception:
                pass

    def open_add_chooser(self):
        """نافذة صغيرة: شراء أم بيع؟ ثم تنتقل لصفحة التفاصيل."""
        if not getattr(self, "user", None):
            return
        t = self.theme
        self._close_chooser()
        win = tk.Toplevel(self.root)
        self._choice_win = win
        try:
            win.title("إضافة عملية")
            win.configure(bg=t["bg"])
            win.transient(self.root)
            win.resizable(False, False)
        except Exception:
            pass
        w, h = self._px(470), self._px(270)
        sw, sh = self._screen_size()
        try:
            win.geometry("{}x{}+{}+{}".format(w, h, max(0, (sw - w) // 2),
                                               max(0, (sh - h) // 3)))
        except Exception:
            pass
        try:
            win.grab_set()
        except Exception:
            pass

        tk.Label(win, text="إضافة عملية جديدة", bg=t["bg"], fg=t["accent"],
                 font=("Segoe UI", 15, "bold")).pack(pady=(self._px(18), 2))
        tk.Label(win, text="اختر نوع العملية للمتابعة إلى صفحة التفاصيل",
                 bg=t["bg"], fg=t["muted"],
                 font=("Segoe UI", 10)).pack()

        row = tk.Frame(win, bg=t["bg"])
        row.pack(fill="both", expand=True, padx=self._px(14),
                 pady=self._px(14))

        def choose(kind):
            self._close_chooser()
            if kind == "buy":
                self.open_add_buy()
            else:
                self.open_sell()

        self._mkbtn(row, "🟢  شراء (إضافة رصيد)",
                    lambda: choose("buy"), kind="success",
                    font_size=12).pack(side="right", fill="both",
                                       expand=True, padx=self._px(6),
                                       ipady=self._px(22))
        self._mkbtn(row, "🔴  بيع (من الرصيد)",
                    lambda: choose("sell"), kind="danger",
                    font_size=12).pack(side="right", fill="both",
                                       expand=True, padx=self._px(6),
                                       ipady=self._px(22))
        self._mkbtn(win, "إلغاء", self._close_chooser, kind="neutral",
                    font_size=10).pack(pady=(0, self._px(12)))
        try:
            win.bind("<Escape>", lambda e: self._close_chooser())
            win.protocol("WM_DELETE_WINDOW", self._close_chooser)
        except Exception:
            pass

    def _addbuy_total(self):
        """يعرض إجمالي التكلفة (الوزن × سعر الوحدة) أثناء الكتابة."""
        lbl = getattr(self, "ab_total", None)
        if lbl is None:
            return
        try:
            qty = float(str(self.ab_qty.get()).strip() or 0)
            price = float(str(self.ab_price.get()).strip() or 0)
        except (TypeError, ValueError):
            qty = price = 0.0
        try:
            lbl.config(text=self._hide("{}{:,.2f}".format(
                CUR[self.cur]["sym"], qty * price)))
        except Exception:
            pass

    def _addbuy_use_price(self):
        """يملأ السعر بسعر الراجحي/السوق الحالي للعيار والوحدة المختارين."""
        try:
            unit = UNIT_MAP.get(self.ab_unit.get(), "g")
            k = int(str(self.ab_karat.get()).split()[0])
        except (TypeError, ValueError):
            unit, k = "g", 24
        val = self.current_per_unit_disp(unit, k)
        if not val:
            messagebox.showinfo("السعر الحالي",
                                "لم يُجلب سعر السوق بعد — اضغط ⟳ في الأعلى أولاً.")
            return
        self.ab_price.set("{:.2f}".format(val))
        self._addbuy_total()

    def _addbuy_clear(self):
        for v in (getattr(self, "ab_name", None), getattr(self, "ab_qty", None),
                  getattr(self, "ab_price", None)):
            if v is not None:
                try:
                    v.set("")
                except Exception:
                    pass
        self._addbuy_total()

    def open_add_buy(self):
        """صفحة تفاصيل عملية الشراء: الاسم والوزن والوحدة والعيار وسعر الشراء."""
        if not getattr(self, "user", None):
            return
        t = self.theme
        CARD, MUT, GOLD = t["card"], t["muted"], t["accent"]
        body = self._open_page("addbuy", "+ إضافة عملية شراء")
        body.title("+ إضافة عملية شراء")

        self.ab_name = tk.StringVar()
        self.ab_qty = tk.StringVar()
        self.ab_price = tk.StringVar()
        self.ab_unit = tk.StringVar(value="غرام")
        self.ab_karat = tk.StringVar(value="{} قيراط".format(self.karat))

        card = tk.Frame(body, bg=CARD, highlightthickness=1,
                        highlightbackground=t.get("border", "#252c39"))
        card.pack(fill="x", padx=self._px(4), pady=(0, self._px(6)))
        tk.Label(card, text="تفاصيل العملية", bg=CARD, fg=GOLD, anchor="e",
                 font=("Segoe UI", 12, "bold")).pack(
            fill="x", padx=self._px(16), pady=(self._px(12), 0))
        tk.Frame(card, bg=t.get("border", "#252c39"), height=1).pack(
            fill="x", padx=self._px(14), pady=(self._px(8), 0))

        sections = [self._form_section(card, title) for title in
                    ("بيانات العملية", "الوزن والعيار", "السعر والقيمة")]
        for section in sections:
            for c in (0, 1, 2): section.columnconfigure(c, weight=1, uniform="ab")

        def entry(parent, var, fsize=12):
            return tk.Entry(parent, textvariable=var, justify="center",
                            bg=t.get("entry_bg", "#151a23"), fg=t["text"],
                            insertbackground=t["text"], relief="flat", bd=0,
                            highlightthickness=1,
                            highlightbackground=t.get("border", "#252c39"),
                            highlightcolor=t["accent"],
                            font=("Segoe UI", fsize))

        def cell(row, col, label, builder, hint=""):
            box = tk.Frame(sections[row], bg=CARD)
            box.grid(row=0, column=col, sticky="ew", padx=self._px(8),
                     pady=self._px(5))
            tk.Label(box, text=label, bg=CARD, fg=MUT, anchor="e",
                     font=("Segoe UI", 10)).pack(fill="x")
            wdg = builder(box)
            wdg.pack(fill="x", ipady=self._px(5), pady=(self._px(4), 0))
            if hint:
                tk.Label(box, text=hint, bg=CARD, fg=MUT, anchor="e",
                         font=("Segoe UI", 10)).pack(fill="x")
            return wdg

        self.ab_name_e = cell(0, 1, "اسم العملية",
                              lambda p: entry(p, self.ab_name),
                              hint="مثال: سبيكة 10 غرام")
        self.ab_qty_e = cell(1, 0, "الوزن", lambda p: entry(p, self.ab_qty),
                             hint="بالوحدة المختارة")
        cell(1, 1, "الوحدة",
             lambda p: ttk.OptionMenu(p, self.ab_unit, self.ab_unit.get(),
                                      "غرام", "أونصة"))
        cell(1, 2, "العيار",
             lambda p: ttk.OptionMenu(p, self.ab_karat, self.ab_karat.get(),
                                      *["{} قيراط".format(k) for k in KARATS]))
        self.ab_price_e = cell(2, 1, "سعر الشراء للوحدة",
                               lambda p: entry(p, self.ab_price),
                               hint="بالعملة الحالية: {}".format(
                                   CUR[self.cur]["name"]))
        tot_box = tk.Frame(sections[2], bg=CARD)
        tot_box.grid(row=0, column=0, sticky="ew", padx=self._px(8),
                     pady=self._px(5))
        tk.Label(tot_box, text="إجمالي التكلفة", bg=CARD, fg=MUT, anchor="e",
                 font=("Segoe UI", 10)).pack(fill="x")
        self.ab_total = tk.Label(tot_box, text="--", bg=CARD, fg=GOLD,
                                 anchor="e", font=("Segoe UI", 16, "bold"))
        self.ab_total.pack(fill="x", pady=(self._px(4), 0))

        for v in (self.ab_qty, self.ab_price):
            try:
                v.trace_add("write", lambda *a: self._addbuy_total())
            except Exception:
                pass

        bar = self._form_actions("addbuy")
        self._mkbtn(bar, "💾  حفظ العملية", self.add_item, kind="primary",
                    font_size=12).pack(side="right")
        self._mkbtn(bar, "⟲  السعر الحالي", self._addbuy_use_price,
                    kind="neutral", font_size=10).pack(side="right",
                                                       padx=self._px(8))
        self._mkbtn(bar, "إلغاء", body.go_back, kind="neutral", font_size=10).pack(side="left", padx=4)
        self._mkbtn(bar, "🧹  تفريغ الحقول", self._addbuy_clear,
                    kind="ghost", font_size=10).pack(side="left")

        tk.Label(body, text="💡 تُخزَّن العملية داخلياً بالدولار وتُعرض بعملتك الحالية،"
                            " ويمكنك بيعها لاحقاً من «أمر بيع».",
                 bg=t["bg"], fg=MUT, anchor="e", justify="right",
                 font=("Segoe UI", 10)).pack(fill="x", padx=self._px(10))

        self._addbuy_total()
        try:
            self.ab_name_e.focus_set()
        except Exception:
            pass

    # ============================ عمليات الشراء ============================
    def add_item(self):
        """حفظ عملية الشراء من صفحة التفاصيل (بعد التحقق من المدخلات)."""
        name = str(getattr(self, "ab_name", tk.StringVar()).get()).strip()
        try:
            qty = float(str(self.ab_qty.get()).strip())
            buy_disp = float(str(self.ab_price.get()).strip())
        except (TypeError, ValueError, AttributeError):
            messagebox.showwarning("تنبيه", "أدخل الوزن وسعر الشراء كأرقام صحيحة")
            return
        if qty <= 0 or buy_disp <= 0:
            messagebox.showwarning("تنبيه", "الوزن والسعر يجب أن يكونا أكبر من صفر")
            return
        buy_usd = buy_disp / self.rate()
        unit = UNIT_MAP.get(self.ab_unit.get(), "g")
        try:
            item_karat = int(str(self.ab_karat.get()).split()[0])
        except (TypeError, ValueError):
            item_karat = 24
        if item_karat not in KARATS:
            item_karat = 24
        name = name or "بدون اسم"
        self._checkpoint_undo("إضافة شراء")
        self.holdings.append({"name": name, "qty": qty, "unit": unit,
                              "buy": buy_usd, "karat": item_karat})
        factor = OZ_TO_G if unit == "oz" else 1
        self.history.append({
            "type": "buy", "name": name,
            "qty_g": qty * factor, "price_usd": buy_usd / factor,
            "amount_usd": qty * buy_usd, "karat": item_karat,
            "date": datetime.now().strftime("%Y-%m-%d %H:%M"),
        })
        self.save()
        self.render()
        self._close_page("addbuy")
        try:
            self.toast("تمت إضافة العملية",
                       "{} — {:g} {} (عيار {}) بسعر {}{:,.2f}".format(
                           name, qty, "أونصة" if unit == "oz" else "غرام",
                           item_karat, CUR[self.cur]["sym"], buy_disp))
        except Exception:
            pass

    def _edit_holding(self, index, name, qty, buy_disp, karat):
        if getattr(self,"locked",False) or not getattr(self,"user",None):
            raise ValueError("سجل الدخول وافتح الحساب أولاً.")
        if not 0<=index<len(self.holdings):
            raise ValueError("المقتنى لم يعد متاحاً.")
        try:
            qty=float(qty); buy_disp=float(buy_disp); karat=int(karat)
            rate=float(self.rate())
        except (ValueError,TypeError):
            raise ValueError("أدخل أرقاماً صحيحة.")
        if not all(math.isfinite(v) and v>0 for v in (qty,buy_disp,rate)) or karat not in KARATS:
            raise ValueError("الوزن والسعر أكبر من صفر، والعيار من العيارات المتاحة.")
        self._checkpoint_undo("تعديل مقتنى")
        self.holdings[index].update(name=str(name).strip() or "بدون اسم",qty=qty,buy=buy_disp/rate,karat=karat)
        self.save()

    def open_edit_holding(self):
        selection=self.tree.selection()
        selected=self.row_map.get(selection[0]) if selection else None
        if not selected or selected[0]!="hold":
            messagebox.showinfo("تعديل مقتنى","حدد مقتنى متبقياً من الجدول أولاً. تعديل المبيعات المحققة غير متاح.",parent=self.root)
            return
        index=selected[1]; holding=self.holdings[index]
        win=self._open_page("editholding", "تعديل المقتنى المتبقي")
        win.configure(bg=self.theme["bg"])
        sections = [self._form_section(win, title) for title in
                    ("بيانات العملية", "الوزن والعيار", "السعر والقيمة")]
        # Reject a stale editor after undo/restore/account switch.
        owner=self.user
        variables=[]
        labels=("الاسم","الكمية ("+("أونصة" if holding.get("unit")=="oz" else "غرام")+")",
                "تكلفة الشراء للوحدة ("+self.cur+")","العيار")
        values=(holding.get("name",""),holding["qty"],float(holding["buy"])*self.rate(),holding.get("karat",24))
        currency=self.cur
        for index_field, (label,value) in enumerate(zip(labels,values)):
            section = sections[0 if index_field == 0 else (2 if index_field == 2 else 1)]
            tk.Label(section, text=label, bg=self.theme["card"], fg=self.theme["muted"],
                     font=("Segoe UI",10), anchor="e").pack(fill="x", pady=4)
            var=tk.StringVar(value=str(value));variables.append(var)
            self._mkentry(section, var, width=28).pack(fill="x", padx=8, pady=4, ipady=5)
        tk.Label(win,text="التعديل يخص الرصيد المتبقي؛ لا يعيد كتابة سجل الشراء الأصلي أو المبيعات السابقة.\nيمكن التراجع عنه لخطوة واحدة.",bg=self.theme["card"],fg=self.theme["muted"],wraplength=380).pack(padx=20,pady=8)
        def apply():
            if self.user!=owner or self.cur!=currency or index>=len(self.holdings) or self.holdings[index] is not holding:
                messagebox.showwarning("تعديل مقتنى","تغير الحساب أو المقتنى؛ أعد فتح التعديل.",parent=win);return
            try:
                self._edit_holding(index,*(v.get() for v in variables))
            except ValueError as exc:
                messagebox.showwarning("تعديل مقتنى",str(exc),parent=win);return
            win.destroy();self.render()
        actions = self._form_actions("editholding")
        self._mkbtn(actions, "حفظ التعديل", apply).pack(side="right", padx=4)
        self._mkbtn(actions, "إلغاء", win.go_back, kind="neutral").pack(side="left", padx=4)

    def _checkpoint_undo(self, label):
        """One deep snapshot of portfolio operations, isolated to this session/account."""
        if not getattr(self,"user",None):
            return
        self._undo_snapshot={"user":self.user,"label":label,
            "holdings":copy.deepcopy(self.holdings),
            "sells":copy.deepcopy(self.sells),"history":copy.deepcopy(self.history)}

    def _restore_undo(self):
        snapshot=getattr(self,"_undo_snapshot",None)
        if (not snapshot or not getattr(self,"user",None)
                or snapshot["user"]!=self.user or getattr(self,"locked",False)):
            return False
        self.holdings=copy.deepcopy(snapshot["holdings"])
        self.sells=copy.deepcopy(snapshot["sells"])
        self.history=copy.deepcopy(snapshot["history"])
        self._undo_snapshot=None
        self.save()
        return True

    def undo_last_change(self):
        if getattr(self,"locked",False) or not getattr(self,"user",None):
            return
        snapshot=getattr(self,"_undo_snapshot",None)
        if not snapshot or snapshot["user"]!=self.user:
            messagebox.showinfo("التراجع","لا يوجد تغيير يمكن التراجع عنه في هذه الجلسة.",parent=self.root)
            return
        if not messagebox.askyesno("التراجع عن آخر تغيير",
                "التراجع عن «{}»؟\nستُستعاد المقتنيات والمبيعات والسجل قبل هذا التغيير.\nلن تتغير أسعار السوق أو إعداداتك.".format(snapshot["label"]),parent=self.root):
            return
        page=getattr(self,"_page","main")
        if self._restore_undo():
            # Rebuild closes forms that held indices from the previous portfolio.
            self.rebuild_ui()
            self._show_page("chart" if page=="chart" else "main")
            self.toast("تم التراجع","تم استعادة العمليات السابقة. التراجع متاح لخطوة واحدة فقط.")

    def _undo_shortcut(self, event=None):
        widget=getattr(event,"widget",None)
        try:
            if widget is not None and widget.winfo_class() in ("Entry","TEntry","Text","Spinbox","TSpinbox"):
                return None
        except Exception:
            pass
        self.undo_last_change()
        return "break"

    def del_item(self):
        sel = self.tree.selection()
        if not sel:
            return
        kind_idx = self.row_map.get(sel[0])
        if not kind_idx:
            return
        kind, idx = kind_idx
        if kind == "hold" and 0 <= idx < len(self.holdings):
            self._checkpoint_undo("حذف مقتنى")
            del self.holdings[idx]
            self.save()
            self.render()
        elif kind == "sell" and 0 <= idx < len(self.sells):
            if messagebox.askyesno("تراجع عن بيع",
                                   "حذف عملية البيع؟ ستُعاد الكمية إلى الرصيد."):
                self._checkpoint_undo("حذف بيع وإعادة كميته")
                s = self.sells.pop(idx)
                self.holdings.append({"name": s.get("name", "مُسترد"),
                                      "qty": s["qty"], "unit": s.get("unit", "g"),
                                      "buy": s.get("buy", 0),
                                      "karat": s.get("karat", 24)})
                self.save()
                self.render()

    def clear_all(self):
        if (self.holdings or self.sells) and messagebox.askyesno(
                "تأكيد", "مسح جميع المقتنيات وعمليات البيع؟"):
            self._checkpoint_undo("مسح المقتنيات والمبيعات")
            self.holdings = []
            self.sells = []
            self.save()
            self.render()

    # ============================ أمر البيع ============================
    def _tree_dblclick(self, event=None):
        """نقر مزدوج على سطر عملية شراء => أمر بيع عليها."""
        iid = ""
        try:
            if event is not None:
                iid = self.tree.identify_row(event.y)
        except Exception:
            iid = ""
        if not iid:
            try:
                sel = self.tree.selection()
                iid = sel[0] if sel else ""
            except Exception:
                iid = ""
        kind_idx = getattr(self, "row_map", {}).get(iid)
        if not kind_idx or kind_idx[0] != "hold":
            return
        self.open_sell(preselect=kind_idx[1])

    def open_sell(self, preselect=None):
        """صفحة أمر البيع: تختار عملية الشراء + الكمية + طريقة السعر
        (السعر الحالي للراجحي أو سعر تُدخله أنت)."""
        if not self.holdings:
            messagebox.showinfo("أمر بيع", "لا يوجد رصيد مفتوح لبيعه")
            return
        t = self.theme
        BG, CARD = t["bg"], t["card"]
        CARD2 = t.get("card2", "#1a2029")
        GOLD, TXT, MUT = t["accent"], t["text"], t["muted"]
        PROF = t.get("profit", "#3ddc97")
        LOSS = t.get("loss", "#ff6b6b")
        BORD = t.get("border", "#252c39")
        EBG = t.get("entry_bg", "#151a23")
        sym = CUR[self.cur]["sym"]
        uname = CUR[self.cur]["name"]
        px = self._px

        win = self._open_page("sell", "أمر بيع — اختر العملية وسعر البيع")
        win.configure(bg=BG)
        win.transient(self.root)
        win.grab_set()

        def mkcard(title):
            f = tk.LabelFrame(win, text="  {}  ".format(title), bg=CARD, fg=GOLD,
                              font=("Segoe UI", 11, "bold"), bd=0,
                              highlightthickness=1, highlightbackground=BORD)
            f.pack(fill="x", padx=px(4), pady=(px(6), px(4)))
            return f

        def lot_grams(h):
            return float(h.get("qty", 0)) * (OZ_TO_G if h.get("unit") == "oz" else 1)

        def lot_unit(h):
            return "أونصة" if h.get("unit") == "oz" else "غرام"

        def lot_price_now(h):
            try:
                return float(self.current_per_unit_disp(h.get("unit", "g"),
                                                        h.get("karat", 24)))
            except Exception:
                return 0.0

        # ---------------- ١) اختيار عملية الشراء ----------------
        c1 = mkcard("١) اختر عملية الشراء التي تريد بيعها")
        list_fr = tk.Frame(c1, bg=CARD)
        list_fr.pack(fill="x", padx=px(10), pady=(px(8), px(2)))
        sb = tk.Scrollbar(list_fr, width=px(14))
        sb.pack(side="left", fill="y")
        lb = tk.Listbox(list_fr, height=6, bg=EBG, fg=TXT,
                        selectbackground=t.get("select", "#2c3a52"),
                        selectforeground="#ffffff", relief="flat", bd=0,
                        highlightthickness=1, highlightbackground=BORD,
                        font=("Segoe UI", 10), yscrollcommand=sb.set,
                        activestyle="none", exportselection=False)
        lb.pack(side="right", fill="x", expand=True)
        sb.config(command=lb.yview)
        for i, h in enumerate(self.holdings):
            lb.insert("end", "{n}.  {nm}  —  {q:g} {u}  •  عيار {k}  •  شراء {s}{p:,.2f}".format(
                n=i + 1, nm=h.get("name", "بدون اسم"), q=float(h.get("qty", 0)),
                u=lot_unit(h), k=h.get("karat", 24), s=sym,
                p=float(h.get("buy", 0)) * self.rate()))
        try:
            first = int(preselect) if preselect is not None else 0
        except Exception:
            first = 0
        if not (0 <= first < len(self.holdings)):
            first = 0
        lb.selection_set(first)
        lb.see(first)

        lbl_info = tk.Label(c1, text="", bg=CARD, fg=MUT, justify="right",
                            anchor="e", font=("Segoe UI", 10))
        lbl_info.pack(fill="x", padx=px(12), pady=(px(2), px(8)))

        # ---------------- ٢) الكمية ----------------
        c2 = mkcard("٢) الكمية المباعة")
        r2 = tk.Frame(c2, bg=CARD)
        r2.pack(fill="x", padx=px(10), pady=(px(10), px(4)))
        var_qty = tk.StringVar()
        e_qty = self._mkentry(r2, var_qty, width=13)
        e_qty.pack(side="right", padx=px(6), ipady=px(4))
        tk.Label(r2, text="الكمية:", bg=CARD, fg=TXT,
                 font=("Segoe UI", 10)).pack(side="right")
        lbl_max = tk.Label(r2, text="", bg=CARD, fg=MUT, font=("Segoe UI", 10))
        lbl_max.pack(side="right", padx=px(6))
        bar_q = tk.Frame(c2, bg=CARD)
        bar_q.pack(fill="x", padx=px(10), pady=(0, px(8)))

        # ---------------- ٣) سعر البيع ----------------
        c3 = mkcard("٣) سعر البيع — اختر الطريقة")
        r3 = tk.Frame(c3, bg=CARD)
        r3.pack(fill="x", padx=px(10), pady=(px(10), px(4)))
        mode = tk.StringVar(value="current")
        rb_kw = dict(bg=CARD, fg=TXT, selectcolor=CARD2, activebackground=CARD,
                     activeforeground=GOLD, font=("Segoe UI", 10),
                     highlightthickness=0, bd=0, anchor="e", justify="right")
        rb_cur = tk.Radiobutton(r3, text="السعر الحالي (الراجحي للبيع)",
                                variable=mode, value="current", **rb_kw)
        rb_mine = tk.Radiobutton(r3, text="سعر أحدده بنفسي",
                                 variable=mode, value="manual", **rb_kw)
        rb_cur.pack(side="right", padx=px(2))
        rb_mine.pack(side="right", padx=px(10))

        r4 = tk.Frame(c3, bg=CARD)
        r4.pack(fill="x", padx=px(10), pady=(px(2), px(2)))
        var_price = tk.StringVar()
        price_e = tk.Entry(r4, textvariable=var_price, width=14, justify="center",
                           bg=EBG, fg=GOLD, insertbackground=TXT, relief="flat",
                           bd=0, highlightthickness=1, highlightbackground=BORD,
                           highlightcolor=GOLD, font=("Segoe UI", 12, "bold"),
                           readonlybackground=CARD2)
        lbl_price_title = tk.Label(r4, text="سعر البيع:", bg=CARD, fg=TXT,
                                   font=("Segoe UI", 10))
        lbl_price_title.pack(side="right")
        lbl_unit_hint = tk.Label(r4, text="", bg=CARD, fg=MUT,
                                 font=("Segoe UI", 10))
        lbl_unit_hint.pack(side="right", padx=px(4))
        price_e.pack(side="right", padx=px(6), ipady=px(4))
        btn_ref = self._mkbtn(r4, "⟳  تحديث", lambda: refill_price(False),
                              kind="neutral", font_size=10)
        btn_ref.pack(side="left", padx=px(4))
        lbl_ind = tk.Label(c3, text="", bg=CARD, fg=MUT, anchor="e",
                           justify="right", font=("Segoe UI", 10))
        lbl_ind.pack(fill="x", padx=px(12), pady=(0, px(8)))

        # ---------------- المعاينة والتأكيد ----------------
        lbl_prev = tk.Label(win, text="", bg=BG, fg=GOLD, justify="right",
                            anchor="e", font=("Segoe UI", 12, "bold"))
        lbl_prev.pack(fill="x", padx=px(16), pady=(px(4), px(2)))

        state = {"after": None, "open_price": None}

        def sel_index():
            s = lb.curselection()
            return int(s[0]) if s else 0

        def refill_price(quiet=True):
            i = sel_index()
            if i >= len(self.holdings):
                return
            p = lot_price_now(self.holdings[i])
            if not p:
                if not quiet:
                    messagebox.showinfo("السعر الحالي",
                                        "لم يُجلب سعر السوق بعد — اضغط ⟳ أعلى "
                                        "النافذة الرئيسية ثم أعد المحاولة.", parent=win)
                return
            var_price.set("{:,.2f}".format(p))
            base = state.get("open_price")
            if base is None:
                state["open_price"] = p
            if base and abs(p - base) > 0.004:
                d = p - base
                try:
                    lbl_ind.config(
                        fg=PROF if d > 0 else LOSS,
                        text="{sg}{s}{d:,.2f}  ({pc:+.1f}%)  منذ فتح الصفحة".format(
                            sg="▲ " if d > 0 else "▼ ", s=sym, d=abs(d),
                            pc=(d / base * 100) if base else 0))
                except Exception:
                    pass
            else:
                try:
                    lbl_ind.config(fg=MUT,
                                   text="سعر لحظي — يُحدَّث تلقائياً كل ٥ ثوانٍ")
                except Exception:
                    pass

        def preview(*_a):
            i = sel_index()
            if i >= len(self.holdings):
                lbl_prev.config(text="")
                return
            h = self.holdings[i]
            try:
                q = float(str(var_qty.get()).replace(",", "").strip() or 0)
                p = float(str(var_price.get()).replace(",", "").strip() or 0)
            except ValueError:
                lbl_prev.config(text="")
                return
            if q <= 0 or p <= 0:
                lbl_prev.config(text="")
                return
            val = q * p
            cost = q * float(h.get("buy", 0)) * self.rate()
            pl = val - cost
            pct = (pl / cost * 100.0) if cost else 0.0
            lbl_prev.config(
                fg=GOLD,
                text="قيمة البيع: {s}{v:,.2f}   •   التكلفة: {s}{c:,.2f}   •   "
                     "الربح/الخسارة: {sg}{s}{p:,.2f} ({pc:+.1f}%)".format(
                         s=sym, v=val, c=cost, p=pl, pc=pct,
                         sg="+" if pl > 0 else ("−" if pl < 0 else "")))

        def on_select(*_a):
            i = sel_index()
            if i >= len(self.holdings):
                return
            h = self.holdings[i]
            qty = float(h.get("qty", 0))
            var_qty.set("{:g}".format(qty))
            lbl_max.config(text="(المتاح {:g} {})".format(qty, lot_unit(h)))
            lbl_unit_hint.config(text="للـ{}".format(lot_unit(h)))
            now_p = lot_price_now(h)
            try:
                val_now = qty * now_p
                cost = qty * float(h.get("buy", 0)) * self.rate()
                pl = val_now - cost
                pct = (pl / cost * 100.0) if cost else 0.0
                lbl_info.config(
                    fg=TXT,
                    text="التكلفة: {s}{c:,.2f}  •  قيمة اليوم: {s}{v:,.2f}  •  "
                         "ربح غير محقق: {sg}{s}{p:,.2f} ({pc:+.1f}%)".format(
                             s=sym, c=cost, v=val_now, p=pl, pc=pct,
                             sg="+" if pl > 0 else ("−" if pl < 0 else "")))
            except Exception:
                lbl_info.config(text="")
            if state.get("open_price") is None:
                state["open_price"] = now_p or None
            if mode.get() == "current":
                refill_price(True)
            preview()

        def apply_mode(*_a):
            if mode.get() == "current":
                try:
                    price_e.config(state="readonly", fg=GOLD,
                                   readonlybackground=CARD2)
                    btn_ref.config(state="normal")
                except Exception:
                    pass
                refill_price(True)
            else:
                try:
                    price_e.config(state="normal", fg=GOLD)
                    btn_ref.config(state="disabled")
                    lbl_ind.config(fg=MUT,
                                   text="اكتب سعر البيع الذي تريده بالـ" + uname +
                                        " (يُثبَّت كما هو عند التأكيد)")
                    price_e.focus_set()
                    price_e.selection_range(0, "end")
                except Exception:
                    pass
            preview()

        def fill_all():
            i = sel_index()
            if i < len(self.holdings):
                var_qty.set("{:g}".format(float(self.holdings[i].get("qty", 0))))

        def fill_half():
            i = sel_index()
            if i < len(self.holdings):
                var_qty.set("{:g}".format(round(float(self.holdings[i].get("qty", 0)) / 2.0, 3)))

        self._mkbtn(bar_q, "الكل", fill_all, kind="neutral",
                    font_size=10).pack(side="right", padx=px(4))
        self._mkbtn(bar_q, "النصف", fill_half, kind="neutral",
                    font_size=10).pack(side="right", padx=px(4))

        lb.bind("<<ListboxSelect>>", on_select)
        var_qty.trace_add("write", preview)
        var_price.trace_add("write", preview)
        mode.trace_add("write", apply_mode)

        def key_open():
            try:
                info = self._pages.get("sell")
                return bool(info) and info.get("body") is win and win.winfo_exists()
            except Exception:
                return False

        def tick():
            state["after"] = None
            if not key_open():
                return
            if mode.get() == "current":
                refill_price(True)
            try:
                state["after"] = self.root.after(5000, tick)
            except Exception:
                pass

        def stop_tick():
            a = state.get("after")
            if a:
                try:
                    self.root.after_cancel(a)
                except Exception:
                    pass
                state["after"] = None

        def do_sell():
            i = sel_index()
            if i >= len(self.holdings):
                return
            h = self.holdings[i]
            try:
                qty = float(str(var_qty.get()).replace(",", "").strip())
                price_disp = float(str(var_price.get()).replace(",", "").strip())
            except (TypeError, ValueError):
                messagebox.showwarning("تنبيه", "أدخل كمية وسعر بيع صحيحين", parent=win)
                return
            if qty <= 0 or price_disp <= 0:
                messagebox.showwarning("تنبيه", "الكمية والسعر يجب أن يكونا أكبر من صفر",
                                       parent=win)
                return
            if qty > float(h.get("qty", 0)) + 1e-9:
                messagebox.showwarning("تنبيه", "الكمية أكبر من المتاح في هذه العملية",
                                       parent=win)
                return
            sell_usd = price_disp / self.rate()
            sell_date = datetime.now().strftime("%Y-%m-%d %H:%M")
            self._checkpoint_undo("تسجيل بيع")
            self.sells.append({
                "name": h.get("name", "بدون اسم"),
                "qty": qty,
                "unit": h.get("unit", "g"),
                "buy": float(h.get("buy", 0)),
                "sell": sell_usd,
                "karat": h.get("karat", 24),
                "date": sell_date,
                "price_mode": mode.get(),
            })
            factor = OZ_TO_G if h.get("unit") == "oz" else 1
            self.history.append({
                "type": "sell", "name": h.get("name", "بدون اسم"),
                "qty_g": qty * factor, "price_usd": sell_usd / factor,
                "amount_usd": qty * sell_usd,
                "cost_usd": qty * float(h.get("buy", 0)),
                "pl_usd": qty * (sell_usd - float(h.get("buy", 0))),
                "karat": h.get("karat", 24),
                "date": sell_date,
            })
            if abs(qty - float(h.get("qty", 0))) < 1e-9:
                del self.holdings[i]
            else:
                h["qty"] = float(h.get("qty", 0)) - qty
            self.save()
            self.render()
            stop_tick()
            try:
                self.toast("تم تسجيل البيع", "{:g} {} بـ{}{:,.2f} للوحدة".format(
                    qty, lot_unit(h), sym, price_disp))
            except Exception:
                pass
            self._close_page("sell")

        bar = self._form_actions("sell")
        bar.pack(fill="x", padx=px(12), pady=(px(4), px(10)))
        self._mkbtn(bar, "تأكيد البيع", do_sell, kind="primary",
                    font_size=11).pack(side="right", padx=px(6))
        self._mkbtn(bar, "إلغاء", lambda: (stop_tick(), self._close_page("sell")),
                    kind="neutral", font_size=10).pack(side="right", padx=px(6))
        tk.Label(bar, text="السعر يُثبَّت لحظة التأكيد — لا يتغيّر بعدها مع تغيّر السوق.\n"
                          "تلميح: النقر المزدوج على أي عملية شراء في الجدول يفتحها هنا مباشرة.",
                 bg=BG, fg=MUT, justify="right",
                 font=("Segoe UI", 9)).pack(side="left", padx=px(6))

        on_select()
        apply_mode()
        try:
            state["open_price"] = lot_price_now(self.holdings[sel_index()]) or None
            state["after"] = self.root.after(5000, tick)
        except Exception:
            pass

    # ============================ سجل العمليات ============================
    def open_history(self):
        t = self.theme
        BG, CARD = t["bg"], t["card"]
        GOLD, TXT, MUT = t["accent"], t["text"], t["muted"]
        sym = CUR[self.cur]["sym"]
        win = self._open_page("history", "سجل العمليات")
        win.configure(bg=BG)
        win.geometry("800x460")
        win.transient(self.root)

        cols = ("date", "type", "name", "qty", "price", "amount", "pl")
        heads = ("التاريخ", "النوع", "الاسم", "الكمية (غ)",
                 "السعر/غ", "القيمة", "الربح/الخسارة")
        tv = ttk.Treeview(win, columns=cols, show="headings", height=15)
        for c, h in zip(cols, heads):
            tv.heading(c, text=h)
            tv.column(c, anchor="center", width=100)
        tv.column("date", width=120)
        tv.column("name", width=120)
        tv.tag_configure("buy", foreground=t.get("accent", "#e3b23c"))
        tv.tag_configure("sellpos", foreground=t.get("profit", "#2ecc71"))
        tv.tag_configure("sellneg", foreground=t.get("loss", "#ff5c5c"))
        tv.pack(fill="both", expand=True, padx=4, pady=(10, 8))

        tot_buy = tot_sell = tot_pl = 0.0
        for e in self.history:
            amount = e.get("amount_usd", 0) * self.rate()
            price = e.get("price_usd", 0) * self.rate()
            if e.get("type") == "buy":
                tot_buy += amount
                tv.insert("", "end", tags=("buy",), values=(
                    e.get("date", ""),
                    "▲ شراء ({})".format(e.get("karat", 24)) if e.get("karat") else "▲ شراء",
                    e.get("name", ""),
                    "{:g}".format(round(e.get("qty_g", 0), 3)),
                    "{}{:,.2f}".format(sym, price),
                    "{}{:,.2f}".format(sym, amount), "-"))
            else:
                pl = e.get("pl_usd", 0) * self.rate()
                tot_sell += amount
                tot_pl += pl
                tag = "sellpos" if pl >= 0 else "sellneg"
                tv.insert("", "end", tags=(tag,), values=(
                    e.get("date", ""),
                    "▼ بيع ({})".format(e.get("karat", 24)) if e.get("karat") else "▼ بيع",
                    e.get("name", ""),
                    "{:g}".format(round(e.get("qty_g", 0), 3)),
                    "{}{:,.2f}".format(sym, price),
                    "{}{:,.2f}".format(sym, amount),
                    "{}{}{:,.2f}".format("+" if pl > 0 else "", sym, pl)))

        info = ("إجمالي الشراء: {s}{b:,.2f}    |    إجمالي البيع: {s}{sl:,.2f}    |    "
                "الربح المحقق: {sg}{s}{p:,.2f}").format(
            s=sym, b=tot_buy, sl=tot_sell, p=tot_pl, sg="+" if tot_pl > 0 else "")
        tk.Label(win, text=info if self.history else "لا توجد عمليات بعد.",
                 bg=BG, fg=TXT, font=("Segoe UI", 11, "bold")).pack(anchor="e",
                                                                     padx=14, pady=(0, 6))

        def clear_hist():
            if self.history and messagebox.askyesno(
                    "تأكيد", "مسح سجل العمليات؟ (لا يؤثر على الرصيد أو الأرباح)", parent=win):
                self._checkpoint_undo("مسح سجل العمليات")
                self.history = []
                self.save()
                win.destroy()
        bar = tk.Frame(win, bg=BG)
        bar.pack(fill="x", padx=12, pady=(0, 10))
        self._mkbtn(bar, "مسح السجل", clear_hist, kind="danger",
                    font_size=10).pack(side="left")
        self._mkbtn(bar, "إغلاق", win.destroy, kind="neutral",
                    font_size=10).pack(side="left", padx=6)

    # ============================ حاسبة الزكاة ============================
    def open_zakat(self):
        t = self.theme
        BG, CARD = t["bg"], t["card"]
        GOLD, TXT, MUT = t["accent"], t["text"], t["muted"]
        sym = CUR[self.cur]["sym"]
        NISAB_G = 85.0
        ZAKAT_RATE = 0.025
        grams = self.total_open_grams()
        gram_val = self.gram_market_disp() * karat_factor(self.karat)
        # القيمة حسب عيار كل عملية، والنصاب يُقاس بالغرام الصافي (مكافئ عيار 24)
        oz_g = self.gram_market_disp()
        total_val = 0.0
        pure_g = 0.0
        for h in self.holdings:
            f = OZ_TO_G if h["unit"] == "oz" else 1
            kf = karat_factor(h.get("karat", 24))
            total_val += h["qty"] * f * oz_g * kf
            pure_g += h["qty"] * f * kf
        zakat_val = total_val * ZAKAT_RATE
        zakat_g = pure_g * ZAKAT_RATE
        reached = pure_g >= NISAB_G

        win = self._open_page("zakat", "حاسبة الزكاة")
        win.configure(bg=BG)
        win.geometry("470x380")
        win.transient(self.root)
        card = tk.Frame(win, bg=CARD, highlightthickness=1,
                        highlightbackground=t.get("border", "#252c39"))
        card.pack(fill="both", expand=True, padx=6, pady=8)

        def row(label, value, color=None):
            fr = tk.Frame(card, bg=CARD)
            fr.pack(fill="x", padx=22, pady=11)
            tk.Label(fr, text=label, bg=CARD, fg=MUT,
                     font=("Segoe UI", 11)).pack(side="right")
            tk.Label(fr, text=value, bg=CARD, fg=color or TXT,
                     font=("Segoe UI", 12, "bold")).pack(side="left")

        row("الرصيد المملوك:", "{:,.3f} غرام (صافي {:,.3f} غ عيار 24)".format(
            grams, pure_g))
        row("سعر الغرام الحالي:", "{}{:,.2f}".format(sym, gram_val))
        row("القيمة الإجمالية:", "{}{:,.2f}".format(sym, total_val))
        row("النصاب (85غ):", "✔ بلغ النصاب" if reached else "✖ دون النصاب",
            t.get("profit") if reached else t.get("loss"))
        tk.Frame(card, bg=t.get("border", "#252c39"), height=1).pack(fill="x", padx=22, pady=10)
        row("الزكاة الواجبة (2.5%):",
            "{}{:,.2f}".format(sym, zakat_val) if reached else "لا تجب",
            GOLD if reached else MUT)
        if reached:
            row("أو ما يعادل:", "{:,.3f} غرام ذهب".format(zakat_g), GOLD)

        tk.Label(win, text="ملاحظة: تجب الزكاة بعد مرور الحول (سنة هجرية) على بلوغ النصاب.\n"
                          "الحساب تقديري للاستئناس فقط.",
                 bg=BG, fg=MUT, font=("Segoe UI", 11), justify="right",
                 wraplength=520).pack(anchor="e", padx=24, pady=(8, 14))

    # ============================ مخطط سعر الغرام ============================
    def _record_price_sample(self, oz_usd):
        now = time.time()
        val = _valid_price(oz_usd)
        points = clean_points(self.price_hist, now, 32*86400)
        if points and now-points[-1][0] < 0.5:
            return
        points.append((now, val))
        recent, bins = [], {}
        for ts, value in points:
            if ts >= now-7200:
                recent.append([ts, value])
            else:
                interval = 300 if ts >= now-2*86400 else 900
                key = (interval, int(ts//interval))
                bins.setdefault(key, []).append([ts, value])
        history = []
        for rows in bins.values():
            # Preserve low/high in each historical bin rather than just its end.
            lo = min(rows, key=lambda r:r[1])
            hi = max(rows, key=lambda r:r[1])
            history.extend([lo] if lo == hi else sorted([lo,hi]))
        self.price_hist = sorted(history+recent)[-12000:]

    def _create_live_chart(self, parent):
        t = self.theme
        outer = tk.Frame(parent, bg=t["card"], highlightthickness=1,
                         highlightbackground=t.get("border", "#252c39"))
        outer.pack(fill="both", expand=True, padx=self._px(18), pady=self._px(8))
        head = tk.Frame(outer, bg=t["card"])
        head.pack(fill="x", padx=14, pady=(12, 6))
        tk.Label(head, text="● مخطط سعر الغرام اللحظي", bg=t["card"],
                 fg=t["accent"], font=("Segoe UI", 16, "bold")).pack(side="right")
        self._mkbtn(head, "🏠 الرئيسية", lambda:self._show_page("main"),
                    kind="ghost", font_size=10).pack(side="left")
        currency_row = tk.Frame(outer, bg=t["card"])
        currency_row.pack(fill="x", padx=14, pady=4)
        tk.Label(currency_row, text="عملة المخطط:", bg=t["card"],fg=t["text"],
                 font=("Segoe UI",11)).pack(side="right",padx=(8,2))
        choices = ["{} ({})".format(v["name"],k) for k,v in CUR.items()]
        selected = "{} ({})".format(CUR[self.chart_currency]["name"],self.chart_currency)
        self.chart_currency_var = tk.StringVar(value=selected)
        ttk.OptionMenu(currency_row,self.chart_currency_var,selected,*choices).pack(side="right",padx=4)
        self.chart_currency_var.trace_add("write",lambda *args:self._on_chart_currency())
        tk.Label(currency_row,text="الغرام فقط • العيار يتبع اختيار المحفظة",bg=t["card"],
                 fg=t["muted"],font=("Segoe UI",10)).pack(side="left",padx=4)
        self.chart_modes_frame = tk.Frame(outer,bg=t["card"])
        # pack positioning is updated after all rows exist.
        for mode,title in (("market","السعر العالمي"),("rajhi_sell","الراجحي بيع"),("rajhi_buy","الراجحي شراء")):
            button=tk.Label(self.chart_modes_frame,text=title,padx=14,pady=5,cursor="hand2",
                            bg=t.get("card2",t["card"]),fg=t["muted"],font=("Segoe UI",11,"bold"))
            button.pack(side="right",padx=3)
            button.bind("<Button-1>",lambda e,k=mode:self._set_chart_price_mode(k))
            self._chart_widgets.append((button,"mode",mode))
        self.chart_range_frame = tk.Frame(outer,bg=t["card"])
        self.chart_range_frame.pack(fill="x",padx=14,pady=5)
        for key,title,seconds in CHART_RANGES:
            button=tk.Label(self.chart_range_frame,text=title,padx=12,pady=5,cursor="hand2",
                            bg=t.get("card2",t["card"]),fg=t["muted"],font=("Segoe UI",11))
            button.pack(side="right",padx=3)
            button.bind("<Button-1>",lambda e,k=key:self._set_chart_range(k))
            self._chart_widgets.append((button,"range",key))
        tools = tk.Frame(outer, bg=t["card"])
        tools.pack(fill="x", padx=14, pady=4)
        for title, action in (("نطاق مخصص", self._custom_chart_range),
                              ("تكبير +", lambda:self._zoom_chart(0.5)),
                              ("تصغير −", lambda:self._zoom_chart(2)),
                              ("← أقدم", lambda:self._pan_chart(-0.25)),
                              ("أحدث →", lambda:self._pan_chart(0.25)),
                              ("مباشر", lambda:self._set_chart_range(self.chart_range)),
                              ("داكن / فاتح", self.toggle_theme)):
            self._mkbtn(tools,title,action,kind="ghost",font_size=10).pack(side="right",padx=2)
        fonts = tk.Frame(outer, bg=t["card"])
        fonts.pack(fill="x", padx=14, pady=2)
        self._mkbtn(fonts,"↶ التراجع",self.undo_last_change,kind="ghost",font_size=10).pack(side="left",padx=4)
        tk.Label(fonts,text="حجم خط المخطط:",bg=t["card"],fg=t["text"],font=("Segoe UI",11)).pack(side="right")
        for size in (10,12,14,16,18):
            self._mkbtn(fonts,str(size),lambda n=size:self._set_chart_font(n),kind="ghost",font_size=11).pack(side="right",padx=2)
        tk.Label(fonts,text="اسحب الرسم للتحريك • العجلة للتكبير • أخضر: شراء / أحمر: بيع",bg=t["card"],fg=t["muted"],font=("Segoe UI",10)).pack(side="left")
        info=tk.Label(outer,text="بانتظار السعر…",bg=t["card"],fg=t["muted"],
                      anchor="e",justify="right",font=("Segoe UI",11))
        info.pack(fill="x",padx=14,pady=6)
        canvas=tk.Canvas(outer,bg=t["card"],height=self._px(420),highlightthickness=0)
        canvas.pack(fill="both",expand=True,padx=10,pady=(4,8))
        note=tk.Label(outer,text="",bg=t["card"],fg=t["muted"],anchor="e",justify="right",font=("Segoe UI",10))
        note.pack(fill="x",padx=14,pady=(2,10))
        view={"canvas":canvas,"info":info,"note":note,"layout":None}
        self._live_chart_views.append(view)
        canvas.bind("<Configure>",lambda e:self.draw_chart())
        canvas.bind("<Motion>",lambda e,v=view:hover_chart(v["canvas"],v["layout"],e,self.theme))
        canvas.bind("<Leave>",lambda e,c=canvas:c.delete("hover"))
        canvas.bind("<MouseWheel>",lambda e:self._chart_wheel(e))
        canvas.bind("<Button-4>",lambda e:self._zoom_chart(0.8))
        canvas.bind("<Button-5>",lambda e:self._zoom_chart(1.25))
        canvas.bind("<ButtonPress-1>",lambda e,v=view:self._chart_drag_start(e,v))
        canvas.bind("<B1-Motion>",lambda e,v=view:self._chart_drag_move(e,v))
        self._update_chart_mode_visibility()
        self._paint_range_buttons()
        return view

    def _build_chart_page(self):
        self._create_live_chart(self.chart_page.inner)

    # ------------------------- التنقل بين الصفحات -------------------------
    def _all_pages(self):
        pages = {"main": getattr(self, "main_wrap", None),
                 "chart": getattr(self, "chart_page", None)}
        for key, info in getattr(self, "_pages", {}).items():
            pages[key] = info["container"]
        return pages

    def _show_page(self, name):
        """التبديل بين الرئيسية والمخطط وأي صفحة داخلية أخرى بنفس النافذة."""
        pages = self._all_pages()
        if name not in pages or pages[name] is None:
            name = "main"
        self._page = name
        self._paint_navigation(name)
        for page in pages.values():
            if page is not None and page.winfo_manager():
                page.pack_forget()
        pages[name].pack(fill="both", expand=True)
        # إعادة أي صفحة إلى أعلاها عند فتحها (لا يظهر منتصفها)
        info = getattr(self, "_pages", {}).get(name) or {}
        scroller = info.get("scroll")
        if name == "main":
            scroller = getattr(self, "main_wrap", None)
        elif name == "chart":
            scroller = getattr(self,"chart_page",None)
        if scroller is not None:
            try:
                scroller.refresh()
                scroller.to_top()
            except Exception:
                pass
        if name == "chart":
            self.root.after(30, self.draw_chart)

    def _open_page(self, key, title):
        """ينشئ (أو يعيد إنشاء) صفحة داخلية ويظهرها، ويعيد إطار المحتوى."""
        t = self.theme
        BG = t["bg"]
        if key in getattr(self, "_pages", {}):
            self._close_page(key)
        container = tk.Frame(self.page_holder, bg=BG)
        px = self._px
        header = tk.Frame(container, bg=BG)
        header.pack(fill="x", padx=px(24), pady=(px(16), px(8)))
        title_lbl = tk.Label(header, text=title, bg=BG, fg=t["accent"],
                             font=("Segoe UI", 20, "bold"))
        title_lbl.pack(side="right")
        # محتوى الصفحة داخل منطقة تمرير حتى لا يُقصّ أي جزء على الشاشات الصغيرة
        scroll = ScrollArea(container, BG)
        body = PageBody(scroll.inner, self, key, title_lbl)
        self._mkbtn(header, "🏠 الرئيسية", lambda b=body: b.go_back(),
                    kind="ghost", font_size=10).pack(side="left")
        # خط فاصل رقيق أسفل شريط عنوان الصفحة
        tk.Frame(container, bg=t.get("border", "#252c39"),
                 height=1).pack(fill="x", padx=px(22))
        # شريط سفلي بزر رجوع واضح (يُحجز قبل المحتوى حتى لا يُزاحم)
        foot = tk.Frame(container, bg=BG)
        foot.pack(side="bottom", fill="x", padx=px(24), pady=(px(6), px(12)))
        self._mkbtn(foot, "←  رجوع إلى الرئيسية",
                    lambda b=body: b.go_back(), kind="neutral",
                    font_size=10).pack(anchor="center")
        scroll.pack(fill="both", expand=True, padx=px(8), pady=(px(4), px(2)))
        body.pack(fill="both", expand=True, padx=px(18), pady=(px(10), px(16)))
        maxw = {"zakat": 700, "alerts": 820, "settings": 780,
                "sell": 880, "account": 700, "history": 1400,
                "calc": 900}.get(key, 1040)

        def _fit(event=None, b=body, m=maxw):
            # توسيط المحتوى بعرض مريح داخل نافذة عريضة بدل شده على الطول
            try:
                avail = b.master.winfo_width()
                if avail < 80:
                    return
                side = max(0, int((avail - m) / 2))
                edge = px(18)
                b.pack_configure(padx=(edge + side, edge + side))
            except Exception:
                pass

        scroll.canvas.bind("<Configure>", _fit, add="+")
        self.root.after(80, _fit)
        self._pages[key] = {"container": container, "header": header,
                            "body": body, "scroll": scroll, "footer": foot}
        self._show_page(key)
        return body

    def _close_page(self, key):
        """يغلق صفحة داخلية ويعود للصفحة الرئيسية."""
        info = getattr(self, "_pages", {}).pop(key, None)
        if info is None:
            self._show_page("main")
            return
        self._closing_page = True
        try:
            info["container"].destroy()
        except Exception:
            pass
        finally:
            self._closing_page = False
        if key == "alerts":
            self.al_win = None
            self.al_list = None
            self.al_price = None
            self.al_cur_lbl = None
            self.al_metric = None
            self.al_op = None
        if getattr(self, "_page", "main") == key:
            self._show_page("main")

    def _goto_back(self, win):
        """زر الرجوع أعلى الصفحات الداخلية."""
        try:
            win.go_back()
        except Exception:
            self._close_page(getattr(win, "page_key", ""))

    def rebuild_ui(self):
        try:
            self._close_chooser()
        except Exception:
            pass
        self._closing_page = True
        self._pages = {}
        was_widget = getattr(self, "price_widget", None) is not None
        self.price_widget = None
        try:
            for w in list(self.root.winfo_children()):
                w.destroy()
        finally:
            self._closing_page = False
        self._page = "main"
        self.build_ui()
        self.render()
        if getattr(self, "user", None):
            self.start_live_refresh()
        if was_widget and getattr(self, "user", None):
            self.open_widget()

    def show_chart(self):
        self._show_page("chart")

    def _set_chart_range(self, key):
        if key not in {r[0] for r in CHART_RANGES}:
            return
        self.chart_range = key
        self.chart_viewport = None
        self.save()
        self._paint_range_buttons()
        self.draw_chart()

    def _on_chart_currency(self):
        selected = self.chart_currency_var.get()
        for key,value in CUR.items():
            if selected == "{} ({})".format(value["name"],key):
                self._set_chart_currency(key)
                break

    def _set_chart_currency(self, currency):
        if currency not in CUR:
            return
        self.chart_currency = currency
        if currency != "SAR":
            self.chart_price_mode = "market"
        self._update_chart_mode_visibility()
        self._paint_range_buttons()
        self.save()
        self.draw_chart()

    def _set_chart_price_mode(self, mode):
        if mode not in ("market","rajhi_sell","rajhi_buy"):
            return
        self.chart_price_mode = mode if self.chart_currency == "SAR" else "market"
        self._paint_range_buttons()
        self.save()
        self.draw_chart()

    def _update_chart_mode_visibility(self):
        frame = getattr(self,"chart_modes_frame",None)
        if frame is None:
            return
        if self.chart_currency == "SAR":
            frame.pack(fill="x",padx=14,pady=4,before=self.chart_range_frame)
        else:
            frame.pack_forget()

    def _paint_range_buttons(self):
        for widget,kind,key in getattr(self,"_chart_widgets",[]):
            selected=key == (self.chart_price_mode if kind=="mode" else self.chart_range)
            try:
                widget.config(bg=self.theme["accent"] if selected else self.theme.get("card2",self.theme["card"]),
                              fg=self.theme.get("accent_text","#000") if selected else self.theme["muted"])
            except tk.TclError:
                pass

    def _range_seconds(self):
        return next((r[2] for r in CHART_RANGES if r[0]==self.chart_range),60)

    def _range_label(self):
        return next((r[1] for r in CHART_RANGES if r[0]==self.chart_range),"دقيقة")

    def _chart_factor(self):
        rate=float(self.rates.get(self.chart_currency,DEFAULT_RATES.get(self.chart_currency,1)))
        return rate*karat_factor(self.karat)/OZ_TO_G

    def _chart_offset(self):
        if self.chart_currency != "SAR":
            return 0.0
        kf=karat_factor(self.karat)
        return -float(self.adjust)*kf if self.chart_price_mode=="rajhi_sell" else float(self.buy_margin)*kf if self.chart_price_mode=="rajhi_buy" else 0.0

    def _chart_bounds(self):
        viewport = getattr(self,"chart_viewport",None)
        if viewport:
            return viewport
        now=time.time()
        return now-self._range_seconds(), now

    def _chart_points(self):
        start,end=self._chart_bounds()
        points=clean_points(self.price_hist,end,end-start,self._chart_factor())
        offset=self._chart_offset()
        return [(ts,val+offset) for ts,val in points if val+offset>0 and ts<=end]

    def _chart_average(self):
        # Convert mixed karats to pure-gold equivalent, then the selected karat.
        pure_grams=cost=0.0
        for h in getattr(self,"holdings",[]):
            try:
                qty=float(h["qty"]); factor=OZ_TO_G if h.get("unit")=="oz" else 1
                buy=float(h["buy"]); kf=karat_factor(h.get("karat",24))
                if qty<=0 or buy<=0 or not math.isfinite(qty) or not math.isfinite(buy):
                    continue
                pure_grams+=qty*factor*kf
                cost+=qty*buy
            except (ValueError,TypeError,KeyError):
                continue
        rate=float(self.rates.get(self.chart_currency,1))
        return cost/pure_grams*karat_factor(self.karat)*rate if pure_grams>0 else None

    def _chart_transactions(self):
        start,end=self._chart_bounds(); rows=[]
        rate=float(self.rates.get(self.chart_currency,1))
        for h in getattr(self,"history",[]):
            try:
                if h.get("type") not in ("buy","sell"):
                    continue
                ts=datetime.fromisoformat(h["date"]).timestamp()
                original_karat=int(h.get("karat",24))
                original_price=float(h["price_usd"])*rate
                price=original_price/karat_factor(original_karat)*karat_factor(self.karat)
                if start<=ts<=end and price>0 and math.isfinite(price):
                    rows.append({"ts":ts,"value":price,"type":h["type"],"label":"{} • {} • {} • {:g} g • عيار {} • السعر الفعلي {:,.3f} {} / g • المكافئ لعيار {}: {:,.3f}".format(
                        "شراء" if h["type"]=="buy" else "بيع",h.get("name",""),h["date"],float(h.get("qty_g",0)),original_karat,original_price,self.chart_currency,self.karat,price)})
            except (ValueError,TypeError,KeyError,OverflowError):
                continue
        return rows

    def _set_chart_font(self, size):
        self.chart_font_size=max(10,min(18,int(size)))
        self.save(); self.draw_chart()

    def _custom_chart_range(self):
        win=tk.Toplevel(self.root); win.title("نطاق زمني مخصص")
        win.transient(self.root); win.configure(bg=self.theme["card"])
        start,end=self._chart_bounds(); variables=[]
        for title,ts in (("من",start),("إلى",end)):
            tk.Label(win,text=title,bg=self.theme["card"],fg=self.theme["text"],font=("Segoe UI",12)).pack(pady=4)
            var=tk.StringVar(value=datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M:%S")); variables.append(var)
            tk.Entry(win,textvariable=var,width=26,font=("Segoe UI",12)).pack(padx=20,pady=4)
        tk.Label(win,text="YYYY-MM-DD HH:MM:SS • توقيت الجهاز المحلي\nيعرض البيانات المسجلة فقط؛ الاحتفاظ بآخر ٣٢ يوماً.",bg=self.theme["card"],fg=self.theme["muted"]).pack(padx=20,pady=8)
        def apply():
            try:
                self._apply_custom_chart_range(*(v.get() for v in variables))
            except ValueError as exc:
                messagebox.showwarning("النطاق الزمني",str(exc),parent=win); return
            win.destroy()
        self._mkbtn(win,"عرض",apply).pack(pady=12)

    def _apply_custom_chart_range(self, start, end):
        try:
            first=datetime.fromisoformat(start).timestamp(); last=datetime.fromisoformat(end).timestamp()
        except (ValueError,OverflowError):
            raise ValueError("أدخل تاريخاً ووقتاً صحيحين بالصيغة الموضحة.")
        if last<=first or last-first<10:
            raise ValueError("النهاية يجب أن تكون بعد البداية بعشر ثوانٍ على الأقل.")
        if last>time.time()+1:
            raise ValueError("لا يمكن اختيار وقت مستقبلي.")
        self.chart_viewport=(first,last); self.draw_chart()

    def _zoom_chart(self, factor):
        start,end=self._chart_bounds()
        if not getattr(self,"chart_viewport",None):
            for view in getattr(self,"_live_chart_views",[]):
                layout=view.get("layout")
                if layout:
                    start,end=layout["oldest"],layout["newest"]; break
        span=max(10,min(32*86400,(end-start)*factor)); center=(start+end)/2
        last=min(time.time(),center+span/2)
        self.chart_viewport=(last-span,last); self.draw_chart()

    def _pan_chart(self, fraction):
        start,end=self._chart_bounds(); span=end-start
        last=min(time.time(),end+span*fraction)
        self.chart_viewport=(last-span,last); self.draw_chart()

    def _chart_wheel(self, event):
        self._zoom_chart(0.8 if event.delta>0 else 1.25)
        return "break"

    def _chart_drag_start(self, event, view):
        layout=view.get("layout")
        if layout:
            self._chart_drag=(event.x,layout["oldest"],layout["newest"],layout["bounds"][2]-layout["bounds"][0])

    def _chart_drag_move(self, event, view):
        state=getattr(self,"_chart_drag",None)
        if state:
            x,start,end,width=state; span=end-start
            last=min(time.time(),end-(event.x-x)/max(width,1)*span)
            self.chart_viewport=(last-span,last); self.draw_chart()

    def _chart_mode_label(self):
        return {"market":"السعر العالمي","rajhi_sell":"الراجحي بيع","rajhi_buy":"الراجحي شراء"}.get(self.chart_price_mode,"السعر العالمي")

    def draw_chart(self):
        if not getattr(self,"_live_chart_views",[]):
            return
        now=time.time()
        points=self._chart_points()
        sym=CUR[self.chart_currency]["sym"]
        unit="XAU / {} / GRAM {}K".format(self.chart_currency,self.karat)
        mode=self._chart_mode_label()
        for view in getattr(self,"_live_chart_views",[]):
            canvas=view["canvas"]
            try:
                if not canvas.winfo_exists() or not canvas.winfo_ismapped():
                    continue
                start,end=self._chart_bounds()
                view["layout"]=render_chart(canvas,points,end,end-start,self.theme,unit,getattr(self,"privacy",False),
                    average=self._chart_average(),transactions=self._chart_transactions(),
                    viewport=getattr(self,"chart_viewport",None),font_size=getattr(self,"chart_font_size",12))
                if getattr(self,"privacy",False):
                    summary="وضع الخصوصية — الرسم والأسعار مخفية"
                    note=""
                elif points:
                    values=[v for ts,v in points]
                    change=values[-1]-values[0]
                    pct=change/values[0]*100 if values[0] else 0
                    summary="{} • {} • عيار {} • آخر عينة {}{:,.3f} • {:+.3f} ({:+.2f}%)\nالافتتاح {:,.3f} • أعلى {:,.3f} • أدنى {:,.3f} • {} عينات".format(self._range_label(),mode,self.karat,sym,values[-1],change,pct,values[0],max(values),min(values),len(points))
                    begin=datetime.fromtimestamp(points[0][0]).strftime("%m-%d %H:%M:%S")
                    end=datetime.fromtimestamp(points[-1][0]).strftime("%m-%d %H:%M:%S")
                    note="البيانات المتاحة: {} — {} • النطاق المطلوب: {} • تحديث نحو ٥ ثوانٍ".format(begin,end,self._range_label())
                    age=max(0,now-points[-1][0])
                    if age>15:
                        note += " • لا توجد عينة حديثة منذ {:.0f} ثانية".format(age)
                    if self.chart_currency=="SAR" and self.chart_price_mode!="market":
                        note += "\nأسعار الراجحي محسوبة بفروق البرنامج الحالية، وليست تغذية مباشرة من البنك."
                else:
                    summary="لا توجد عينات سوقية في هذا النطاق — تُعرض العمليات المسجلة إن وجدت"
                    note="لا توجد بيانات لهذا النطاق بعد؛ يبدأ التاريخ بالتراكم أثناء تشغيل البرنامج."
                if not getattr(self,"privacy",False):
                    note += "\n"+self._connection_label()
                    if getattr(self,"chart_viewport",None):
                        note += " • عرض تاريخي ثابت — زر مباشر للعودة"
                    note += "\nالخط المتقطع الذهبي: متوسط التكلفة الحالي المكافئ للعيار المختار، وليس متوسطاً تاريخياً. علامات كل العمليات بأسعار مكافئة للعيار المختار؛ المؤشر يعرض التاريخ والعيار والسعر الفعلي."
                view["info"].config(font=("Segoe UI",getattr(self,"chart_font_size",12)),text=summary,wraplength=max(200,canvas.winfo_width()-24))
                if "note" in view:
                    view["note"].config(font=("Segoe UI",max(10,getattr(self,"chart_font_size",12)-1)),text=note,wraplength=max(200,canvas.winfo_width()-24))
            except tk.TclError as exc:
                self._live_error=str(exc)[:160]

    def _refresh_chart_if_visible(self):
        self.draw_chart()

    # ============================ التنبيهات ============================
    def open_alerts(self):
        if self.al_win is not None:
            try:
                self.al_win.deiconify()
                self.al_win.lift()
                return
            except Exception:
                self.al_win = None
        t = self.theme
        BG, CARD, CARD2 = t["bg"], t["card"], t["card2"]
        GOLD, ATXT, TXT, MUT = t["accent"], t["accent_text"], t["text"], t["muted"]
        win = self._open_page("alerts", "تنبيهات الأسعار")
        self.al_win = win
        win.configure(bg=BG)
        win.geometry("560x380")
        win.transient(self.root)

        al = tk.LabelFrame(win, text="  ➕ إضافة تنبيه جديد  ", bg=CARD, fg=GOLD,
                           font=("Segoe UI", 11, "bold"), bd=0,
                           highlightthickness=1,
                           highlightbackground=t.get("border", "#252c39"))
        al.pack(fill="both", expand=True, padx=4, pady=8)

        row = tk.Frame(al, bg=CARD)
        row.pack(fill="x", padx=14, pady=11)
        tk.Label(row, text="نبّهني عندما يصبح سعر", bg=CARD, fg=TXT,
                 font=("Segoe UI", 10)).pack(side="right")
        self.al_metric = tk.StringVar(value="g")
        ttk.OptionMenu(row, self.al_metric, "g", "g", "oz").pack(side="right", padx=4)
        self.al_op = tk.StringVar(value=">=")
        ttk.OptionMenu(row, self.al_op, ">=", ">=", "<=").pack(side="right", padx=4)
        self.al_price = self._mkentry(row, tk.StringVar(), width=10)
        self.al_price.pack(side="right", padx=6, ipady=4)
        self.al_cur_lbl = tk.Label(row, text="(" + CUR[self.cur]["name"] + ")",
                                   bg=CARD, fg=MUT, font=("Segoe UI", 10))
        self.al_cur_lbl.pack(side="right")
        self._mkbtn(row, "＋ إضافة", self.add_alert, kind="primary",
                    font_size=10).pack(side="right", padx=10)
        tk.Label(al, text="(السعر = سعر الراجحي للبيع بعد الطرح / بالوحدة المختارة)",
                 bg=CARD, fg=MUT, font=("Segoe UI", 10)).pack(anchor="e", padx=16)

        row2 = tk.Frame(al, bg=CARD)
        row2.pack(fill="both", expand=True, padx=14, pady=(4, 12))
        self.al_list = tk.Listbox(row2, height=7, bg=CARD2, fg=TXT,
                                  selectbackground=t.get("border", "#2c3a52"),
                                  selectforeground="#ffffff", relief="flat",
                                  bd=0, highlightthickness=1,
                                  highlightbackground=t.get("border", "#252c39"),
                                  font=("Segoe UI", 10))
        self.al_list.pack(side="left", fill="both", expand=True)
        self._mkbtn(row2, "🗑 حذف", self.del_alert, kind="danger",
                    font_size=10).pack(side="left", padx=(8, 2))

        row3 = tk.Frame(al, bg=CARD)
        row3.pack(fill="x", padx=14, pady=(2, 10))
        self.sound_var = tk.BooleanVar(value=self.sound_on)
        tk.Checkbutton(row3, text="تفعيل التنبيه الصوتي 🔊", variable=self.sound_var,
                       command=self.on_sound_toggle, bg=CARD, fg=TXT,
                       selectcolor=t.get("entry_bg", "#151a23"),
                       activebackground=CARD, activeforeground=TXT,
                       bd=0, highlightthickness=0,
                       font=("Segoe UI", 10)).pack(side="right")
        self._mkbtn(row3, "🔊 تجربة", self._test_sound, kind="ghost",
                    font_size=10).pack(side="right", padx=10)

        def on_close():
            self.al_win = None
            self.al_list = None
            self.al_price = None
            self.al_cur_lbl = None
            self.al_metric = None
            self.al_op = None
            win.destroy()
        win.protocol("WM_DELETE_WINDOW", on_close)
        self.refresh_alert_list()

    def add_alert(self):
        if not getattr(self, "al_price", None):
            return
        try:
            price = float(self.al_price.get())
        except ValueError:
            messagebox.showwarning("تنبيه", "الرجاء إدخال سعر صحيح للتنبيه")
            return
        if price <= 0:
            messagebox.showwarning("تنبيه", "السعر يجب أن يكون أكبر من صفر")
            return
        self.alerts.append({"metric": self.al_metric.get(),
                            "op": self.al_op.get(), "price": price,
                            "cur": self.cur, "fired": False})
        self.save()
        self.refresh_alert_list()
        if getattr(self, "al_price", None):
            self.al_price.delete(0, "end")

    def del_alert(self):
        if not getattr(self, "al_list", None):
            return
        s = self.al_list.curselection()
        if not s:
            return
        idx = s[0]
        if 0 <= idx < len(self.alerts):
            del self.alerts[idx]
            self.save()
            self.refresh_alert_list()

    def refresh_alert_list(self):
        if not getattr(self, "al_list", None):
            return
        self.al_list.delete(0, "end")
        for a in self.alerts:
            unit = "غرام" if a["metric"] == "g" else "أونصة"
            self.al_list.insert("end", "{} {} {}{:,.2f} ({})".format(
                unit, a["op"], CUR.get(a.get("cur", "USD"), {}).get("sym", ""),
                a["price"], "✔" if a.get("fired") else "…"))

    def on_sound_toggle(self):
        self.sound_on = bool(self.sound_var.get())
        self.save()

    def _test_sound(self):
        self._beep()

    def _beep(self):
        if not self.sound_on:
            return
        try:
            import winsound
            winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
        except Exception:
            try:
                self.root.bell()
            except Exception:
                pass

    def check_alerts(self):
        if not self.spot_oz_usd:
            return
        changed = False
        fired_msgs = []
        for a in self.alerts:
            ccy = a.get("cur", self.cur)
            gram = (self.spot_oz_usd / OZ_TO_G) * self.rates.get(ccy, 1.0)
            if ccy == "SAR":
                gram -= self.adjust
            val = gram * OZ_TO_G if a["metric"] == "oz" else gram
            ok = (val >= a["price"]) if a["op"] == ">=" else (val <= a["price"])
            if ok and not a.get("fired"):
                a["fired"] = True
                changed = True
                self._beep()
                unit = "غرام" if a["metric"] == "g" else "أونصة"
                sym = CUR.get(ccy, {}).get("sym", "")
                sign = "≥" if a["op"] == ">=" else "≤"
                fired_msgs.append("{} الذهب {} {:.2f} — الآن {:.2f} {}".format(
                    unit, sign, a["price"], val, sym))
                # Alerts use notifications; never overwrite the compact source/time caption.
        if fired_msgs:
            # إشعار سطح المكتب (Windows toast) + بطاقة داخل التطبيق
            self.notify("تنبيه سعر الذهب", "\n".join(fired_msgs[:3]))
        if changed:
            self.save()
            self.refresh_alert_list()

    # ============================ الإعدادات ============================
    # ============================ حاسبة تكلفة الشراء ============================
    def open_calc(self):
        """حاسبة التكلفة الفعلية: سعر الذهب + المصنعية + ضريبة القيمة المضافة."""
        t = self.theme
        BG, CARD = t["bg"], t["card"]
        GOLD, TXT, MUT = t["accent"], t["text"], t["muted"]
        win = self._open_page("calc", "حاسبة تكلفة الشراء")
        win.configure(bg=BG)
        win.geometry("760x520")
        sym = CUR[self.cur]["sym"]

        card = tk.Frame(win, bg=CARD, highlightthickness=1,
                        highlightbackground=t.get("border", "#252c39"))
        card.pack(fill="both", expand=True, padx=4, pady=8)
        tk.Label(card, text="🧮  احسب ما ستدفعه فعلياً قبل الشراء",
                 bg=CARD, fg=GOLD, font=("Segoe UI", 13, "bold")).pack(
            anchor="e", padx=24, pady=(16, 2))
        tk.Label(card, text="سعر الذهب + المصنعية + ضريبة القيمة المضافة = التكلفة الفعلية",
                 bg=CARD, fg=MUT, font=("Segoe UI", 10)).pack(anchor="e", padx=24)
        tk.Frame(card, bg=t.get("border", "#252c39"), height=1).pack(
            fill="x", padx=24, pady=(10, 6))

        form = tk.Frame(card, bg=CARD)
        form.pack(fill="x", padx=24)
        for c in range(3):
            form.grid_columnconfigure(c, weight=1)

        v_qty = tk.StringVar(value="10")
        v_price = tk.StringVar(value="{:,.3f}".format(
            self.current_per_unit_disp("g", self.karat)))
        v_mk = tk.StringVar(value="3")
        v_vat = tk.StringVar(value="15")
        v_karat = tk.StringVar(value="{} قيراط".format(self.karat))

        def field(row, col, label, var, width=12):
            box = tk.Frame(form, bg=CARD)
            box.grid(row=row, column=col, sticky="ew", padx=6, pady=(6, 2))
            tk.Label(box, text=label, bg=CARD, fg=MUT,
                     font=("Segoe UI", 11), anchor="e").pack(fill="x")
            e = self._mkentry(box, var, width=width)
            e.configure(justify="center")
            e.pack(fill="x", ipady=5)
            return e

        field(0, 0, "الكمية (غرام)", v_qty)
        field(0, 1, "سعر الغرام", v_price)
        box_k = tk.Frame(form, bg=CARD)
        box_k.grid(row=0, column=2, sticky="ew", padx=6, pady=(6, 2))
        tk.Label(box_k, text="العيار", bg=CARD, fg=MUT,
                 font=("Segoe UI", 11), anchor="e").pack(fill="x")
        ttk.OptionMenu(box_k, v_karat, v_karat.get(),
                       *["{} قيراط".format(k) for k in KARATS]).pack(
            fill="x", ipady=3)
        field(1, 0, "المصنعية %", v_mk)
        field(1, 1, "ضريبة القيمة المضافة %", v_vat)

        res = tk.Frame(card, bg=t.get("card2", CARD), highlightthickness=1,
                       highlightbackground=t.get("border", "#252c39"))
        res.pack(fill="x", padx=24, pady=(14, 4))
        out = {}
        for i, (key, label) in enumerate((("gold", "قيمة الذهب"),
                                          ("mk", "المصنعية"),
                                          ("vat", "الضريبة"),
                                          ("total", "الإجمالي المطلوب"),
                                          ("per", "سعر الغرام الفعلي"))):
            tk.Label(res, text=label, bg=t.get("card2", CARD), fg=MUT,
                     font=("Segoe UI", 11), anchor="e").grid(
                row=i, column=0, sticky="e", padx=(18, 10), pady=3)
            lbl = tk.Label(res, text="--", bg=t.get("card2", CARD),
                           fg=GOLD if key in ("total", "per") else TXT,
                           font=("Segoe UI", 12 if key == "total" else 10, "bold"),
                           anchor="w")
            lbl.grid(row=i, column=1, sticky="w", pady=3)
            out[key] = lbl
        res.grid_columnconfigure(1, weight=1)

        def calc(*_a):
            try:
                qty = float(str(v_qty.get()).replace(",", ""))
                price = float(str(v_price.get()).replace(",", ""))
                mk = float(str(v_mk.get()).replace(",", "") or 0)
                vat = float(str(v_vat.get()).replace(",", "") or 0)
            except (TypeError, ValueError):
                for lbl in out.values():
                    lbl.config(text="—")
                return
            if qty <= 0 or price < 0:
                return
            gold = qty * price
            making = gold * mk / 100.0
            sub = gold + making
            tax = sub * vat / 100.0
            total = sub + tax
            out["gold"].config(text="{}{:,.2f}".format(sym, gold))
            out["mk"].config(text="{}{:,.2f}".format(sym, making))
            out["vat"].config(text="{}{:,.2f}".format(sym, tax))
            out["total"].config(text="{}{:,.2f}".format(sym, total))
            out["per"].config(text="{}{:,.2f}".format(sym, total / qty))

        for var in (v_qty, v_price, v_mk, v_vat):
            var.trace_add("write", calc)

        def refresh_price(*_a):
            try:
                k = int(str(v_karat.get()).split()[0])
            except (TypeError, ValueError):
                k = 24
            v_price.set("{:,.3f}".format(self.current_per_unit_disp("g", k)))

        v_karat.trace_add("write", refresh_price)

        bar = tk.Frame(card, bg=CARD)
        bar.pack(fill="x", padx=24, pady=(10, 16))

        def use_in_form():
            try:
                qty = float(str(v_qty.get()).replace(",", ""))
                total_per = float(str(v_price.get()).replace(",", ""))
                mk = float(str(v_mk.get()).replace(",", "") or 0)
                vat = float(str(v_vat.get()).replace(",", "") or 0)
            except (TypeError, ValueError):
                messagebox.showinfo("حاسبة التكلفة", "أدخل كمية وسعراً صحيحين أولاً.",
                                    parent=win)
                return
            if qty <= 0:
                messagebox.showinfo("حاسبة التكلفة", "الكمية يجب أن تكون أكبر من صفر.",
                                    parent=win)
                return
            per_gram = total_per * (1 + mk / 100.0) * (1 + vat / 100.0)
            try:
                k = int(str(v_karat.get()).split()[0])
            except (TypeError, ValueError):
                k = 24
            win.go_back()
            self.open_add_buy()
            try:
                self.ab_name.set("سبيكة عيار {}".format(k))
                self.ab_qty.set("{:g}".format(qty))
                self.ab_price.set("{:.2f}".format(per_gram))
                self.ab_unit.set("غرام")
                self.ab_karat.set("{} قيراط".format(k))
                self._addbuy_total()
            except Exception:
                pass

        self._mkbtn(bar, "🛒  تعبئة نموذج الشراء", use_in_form, kind="primary",
                    font_size=11).pack(side="right", padx=(0, 10))
        self._mkbtn(bar, "↺  إعادة ضبط",
                    lambda: (v_qty.set("1"), v_mk.set("0"), v_vat.set("15")),
                    kind="neutral", font_size=10).pack(side="right")
        calc()

    def _do_backup_now(self, parent=None):
        path = self.auto_backup(force=True)
        if path:
            messagebox.showinfo("نسخ احتياطي",
                                "تم إنشاء نسخة:\n{}".format(os.path.basename(path)),
                                parent=parent)
        else:
            messagebox.showwarning("نسخ احتياطي", "تعذّر إنشاء نسخة احتياطية.",
                                   parent=parent)

    def _shortcut_buy(self):
        if getattr(self, "locked", False) or not getattr(self, "user", None):
            return
        self.open_add_buy()

    def _shortcut_home(self):
        if getattr(self, "locked", False):
            return
        if getattr(self, "_page", "main") != "main":
            self._show_page("main")

    def open_settings(self):
        t = self.theme
        win = self._open_page("settings", "الإعدادات")
        win.configure(bg=t["bg"])
        win.geometry("480x520")
        win.transient(self.root)
        PAD = {"padx": 22, "pady": 7}
        tk.Label(win, text="العناوين", bg=t["bg"], fg=t["accent"],
                 font=("Segoe UI", 13, "bold")).pack(anchor="e", **PAD)
        tvars = {}
        for key, lbl in (("title", "العنوان"), ("subtitle", "العنوان الفرعي"),
                         ("window", "عنوان النافذة")):
            fr = tk.Frame(win, bg=t["bg"])
            fr.pack(fill="x", **PAD)
            tk.Label(fr, text=lbl, bg=t["bg"], fg=t["text"], width=14,
                     anchor="e", font=("Segoe UI", 10)).pack(side="right")
            v = tk.StringVar(value=self.titles.get(key, ""))
            e = self._mkentry(fr, v, width=24)
            e.configure(justify="right")
            e.pack(side="right", fill="x", expand=True, padx=6, ipady=4)
            tvars[key] = v

        tk.Label(win, text="الألوان", bg=t["bg"], fg=t["accent"],
                 font=("Segoe UI", 13, "bold")).pack(anchor="e", **PAD)
        color_fields = [("bg", "خلفية التطبيق"), ("card", "البطاقات"),
                        ("accent", "اللون المميز"), ("text", "لون النص"),
                        ("profit", "لون الربح"), ("loss", "لون الخسارة"),
                        ("total", "لون سطر الإجمالي")]
        tmp = dict(self.theme)

        def make_picker(k, swatch):
            def pick():
                c = colorchooser.askcolor(color=tmp.get(k), title="اختر لوناً", parent=win)
                if c and c[1]:
                    tmp[k] = c[1]
                    swatch.config(bg=c[1])
            return pick

        for key, lbl in color_fields:
            fr = tk.Frame(win, bg=t["bg"])
            fr.pack(fill="x", **PAD)
            tk.Label(fr, text=lbl, bg=t["bg"], fg=t["text"], width=16,
                     anchor="e").pack(side="right")
            sw = tk.Label(fr, text="        ", bg=tmp.get(key, "#000"),
                          relief="flat", bd=0, highlightthickness=1,
                          highlightbackground=t.get("border", "#252c39"))
            sw.pack(side="right", padx=6)
            self._mkbtn(fr, "اختيار لون", make_picker(key, sw),
                        kind="neutral", font_size=8).pack(side="right", padx=4)

        bkfr = tk.Frame(win, bg=t["bg"])
        bkfr.pack(fill="x", padx=22, pady=(14, 0))
        tk.Label(bkfr, text="💾 النسخ الاحتياطي", bg=t["bg"], fg=t["accent"],
                 font=("Segoe UI", 11, "bold")).pack(side="right", padx=(0, 10))
        self._mkbtn(bkfr, "نسخ احتياطي الآن",
                    lambda: self._do_backup_now(win), kind="neutral",
                    font_size=10).pack(side="right", padx=4)
        self._mkbtn(bkfr, "♻ استعادة أحدث نسخة", self.restore_latest_backup,
                    kind="danger", font_size=10).pack(side="right", padx=4)

        tk.Frame(win, bg=t.get("border", "#252c39"), height=1).pack(
            fill="x", padx=22, pady=(14, 0))
        clfr = tk.Frame(win, bg=t["bg"])
        clfr.pack(fill="x", padx=22, pady=(8, 0))
        tk.Label(clfr, text="☁ المزامنة عبر بريدك", bg=t["bg"], fg=t["accent"],
                 font=("Segoe UI", 11, "bold")).pack(side="right")
        acct = self.cloud_account()
        connected = bool(self.cloud_ready())
        if connected:
            ctext = "مرتبط: {}  •  آخر مزامنة: {}".format(
                acct.get("email", "") or acct.get("sync_name", ""),
                acct.get("last_sync", "—"))
            ccolor = t.get("profit", "#3ddc97")
        else:
            ctext = ("غير مرتبط — إما البريد (كلمة مرور التطبيق)، أو «نسخة بملف» "
                     "بدون أي إعداد: ملف واحد تنقله بين أجهزتك.")
            ccolor = t.get("muted", "#888")
        self.lbl_cloud = tk.Label(clfr, text=getattr(self, "_cloud_msg", ctext),
                                  bg=t["bg"], fg=ccolor, anchor="e",
                                  justify="right", wraplength=430,
                                  font=("Segoe UI", 10))
        self.lbl_cloud.pack(fill="x", pady=(2, 0))
        row_c = tk.Frame(win, bg=t["bg"])
        row_c.pack(fill="x", padx=22, pady=(6, 0))
        self._mkbtn(row_c, "🔗 ربط بالبريد", self.open_cloud_dialog,
                    kind="primary", font_size=10).pack(side="right", padx=3)
        self._mkbtn(row_c, "⟳ مزامنة الآن", lambda: self.cloud_sync_now(True),
                    kind="neutral", font_size=10).pack(side="right", padx=3)
        if connected:
            self._mkbtn(row_c, "فصل", self.cloud_unlink, kind="danger",
                        font_size=10).pack(side="left", padx=3)
        row_p = tk.Frame(win, bg=t["bg"])
        row_p.pack(fill="x", padx=22, pady=(4, 0))
        tk.Label(row_p, text="أبسط بديل بلا بريد ولا إعداد:", bg=t["bg"],
                 fg=t.get("muted"), font=("Segoe UI", 10)).pack(side="right", padx=(0, 8))
        self._mkbtn(row_p, "📦 حفظ نسخة في ملف", self.pocket_export,
                    kind="ghost", font_size=10).pack(side="right", padx=3)
        self._mkbtn(row_p, "📂 استعادة من ملف", self.pocket_import,
                    kind="neutral", font_size=10).pack(side="right", padx=3)
        auto_v = tk.BooleanVar(value=bool(self.cloud_conf().get("auto", True)))

        def toggle_auto():
            conf = self.cloud_conf()
            conf["auto"] = bool(auto_v.get())
            self.cloud_store()
        tk.Checkbutton(row_c, text="تلقائي", variable=auto_v, command=toggle_auto,
                       bg=t["bg"], fg=t.get("muted"), selectcolor=t.get("card", "#151a23"),
                       activebackground=t["bg"], activeforeground=t.get("text"),
                       bd=0, highlightthickness=0, font=("Segoe UI", 10)).pack(
                           side="left", padx=6)

        btnfr = tk.Frame(win, bg=t["bg"])
        btnfr.pack(fill="x", padx=22, pady=(16, 10))

        def do_save():
            for k, v in tvars.items():
                self.titles[k] = v.get().strip() or DEFAULT_TITLES.get(k, "")
            self.theme.update(tmp)
            self.save()
            win.destroy()
            self.rebuild_ui()

        def do_reset():
            self.theme = dict(DEFAULT_THEME)
            self.titles = dict(DEFAULT_TITLES)
            self.source_mode = "auto"
            self.save()
            win.destroy()
            self.rebuild_ui()

        self._mkbtn(btnfr, "✔  حفظ وتطبيق", do_save, kind="primary",
                    font_size=11).pack(side="right", padx=10)
        self._mkbtn(btnfr, "استعادة الافتراضي", do_reset, kind="neutral",
                    font_size=11).pack(side="right")

    # ============================ تصدير Excel ============================
    def export_excel(self):
        if not self.holdings and not self.sells:
            messagebox.showinfo("تصدير Excel", "لا توجد عمليات لتصديرها.")
            return
        default = "محفظة_الذهب_{}.xlsx".format(datetime.now().strftime("%Y%m%d_%H%M"))
        path = filedialog.asksaveasfilename(
            title="حفظ ملف Excel", defaultextension=".xlsx",
            initialfile=default, filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        ccy = CUR[self.cur]["name"]
        headers = ["الاسم", "الحالة", "الكمية", "الوحدة", "العيار",
                   "سعر الشراء/وحدة ({})".format(ccy),
                   "التكلفة ({})".format(ccy),
                   "السعر الحالي/البيع ({})".format(ccy),
                   "القيمة ({})".format(ccy),
                   "الربح/الخسارة ({})".format(ccy), "النسبة %"]
        rows = []
        t_cost = t_val = 0.0
        t_qty_g = 0.0
        realized = 0.0
        for h in self.holdings:
            unit_txt = "أونصة" if h["unit"] == "oz" else "غرام"
            cost_disp = h["qty"] * h["buy"] * self.rate()
            karat = h.get("karat", 24)
            cur_disp = self.current_per_unit_disp(h["unit"], karat)
            value_disp = h["qty"] * cur_disp
            pl = value_disp - cost_disp
            pct = (pl / cost_disp * 100) if cost_disp else 0
            t_cost += cost_disp
            t_val += value_disp
            t_qty_g += h["qty"] * (OZ_TO_G if h["unit"] == "oz" else 1)
            rows.append([h["name"], "مفتوح", round(h["qty"], 6), unit_txt, karat,
                         round(h["buy"] * self.rate(), 3), round(cost_disp, 3),
                         round(cur_disp, 3), round(value_disp, 3),
                         round(pl, 3), round(pct, 2)])
        for s in self.sells:
            unit_txt = "أونصة" if s["unit"] == "oz" else "غرام"
            cost_disp = s["qty"] * s["buy"] * self.rate()
            sell_disp = s["sell"] * self.rate()
            s_karat = s.get("karat", 24)
            value_disp = s["qty"] * sell_disp
            pl = value_disp - cost_disp
            pct = (pl / cost_disp * 100) if cost_disp else 0
            realized += pl
            rows.append(["{} ({})".format(s.get("name", "بيع"), s.get("date", "")),
                         "مباع", round(s["qty"], 6), unit_txt, s_karat,
                         round(s["buy"] * self.rate(), 3), round(cost_disp, 3),
                         round(sell_disp, 3), round(value_disp, 3),
                         round(pl, 3), round(pct, 2)])
        t_pl = t_val - t_cost
        t_pct = (t_pl / t_cost * 100) if t_cost else 0
        rows.append(["∑ الرصيد المتبقي (غير محقق)", "", round(t_qty_g, 3), "غرام", "",
                     "", round(t_cost, 3), "", round(t_val, 3),
                     round(t_pl, 3), round(t_pct, 2)])
        rows.append(["∑ الربح المحقق (المباع)", "", "", "", "", "", "", "", "",
                     round(realized, 3), ""])
        rows.append(["∑ الإجمالي (محقق + غير محقق)", "", "", "", "", "", "", "", "",
                     round(t_pl + realized, 3), ""])
        try:
            write_xlsx(path, headers, rows)
        except Exception as e:
            messagebox.showerror("تصدير Excel", "تعذّر حفظ الملف:\n{}".format(e))
            return
        messagebox.showinfo("تصدير Excel",
                            "تم تصدير {} سجل بنجاح.".format(len(self.holdings) + len(self.sells)))

    # ============================ جلب البيانات المباشر ============================
    def start_live_refresh(self):
        if not getattr(self, "user", None):
            return
        self._live_on = True
        # Arm the supervisor BEFORE any network or display work can raise.
        self._live_due = time.monotonic()
        self._ensure_result_pump()
        self._supervise_live_refresh()

    def _arm_live_loop(self):
        # Compatibility hook: one supervisor owns scheduling, not two loops.
        job = getattr(self, "_live_job", None)
        if job:
            try:
                self.root.after_cancel(job)
            except tk.TclError:
                pass
        self._live_job = None
        self._live_due = time.monotonic() + self.refresh_interval() / 1000.0
        self._ensure_result_pump()

    def _loop_refresh(self):
        self._live_job = None
        self._supervise_live_refresh()
        self._ensure_result_pump()

    def _supervise_live_refresh(self):
        if not self._live_on or not self.user:
            return
        now = time.monotonic()
        if now < self._live_due:
            return
        # Schedule from a monotonic deadline, independent of fetch completion.
        self._live_due = now + self.refresh_interval() / 1000.0
        try:
            if getattr(self, "_live_chart_views", []):
                self.draw_chart()
            self.refresh_all()
        except Exception as exc:
            self._live_error = str(exc)[:160]

    def refresh_interval(self):
        """تحديث مستمر وثابت كل ٥ ثوانٍ (لا تباطؤ يُشعرك بتوقّف التحديث)."""
        return LIVE_MS

    def _ensure_result_pump(self):
        # Recover a stale timer ID, e.g. after a UI rebuild/cancelled callback.
        job = self._result_job
        if job is not None:
            try:
                self.root.tk.call("after", "info", job)
                return
            except (tk.TclError, AttributeError):
                if not hasattr(self.root, "tk"):
                    return
                self._result_job = None
        self._result_job = self.root.after(100, self._drain_fetch_results)

    def refresh_all(self):
        if not getattr(self, "user", None):
            return
        self._ensure_result_pump()
        now = time.monotonic()
        # Bound worker count: a stuck DNS/network call cannot freeze polling
        # forever, but also cannot create an unlimited chain of threads.
        self._price_workers = {k: t for k, t in self._price_workers.items() if t.is_alive()}
        with self._fetch_lock:
            if self._fetch_busy and now - self._fetch_started < 4.5:
                return
            if len(self._price_workers) >= 2:
                self._set_update_caption()
                return
            self._fetch_busy = True
            self._fetch_started = now
            self._fetch_serial += 1
            serial = self._fetch_serial
        self._live_attempts += 1
        context = (self._fetch_generation, self.user)
        mode = self.source_mode
        self._set_update_caption()
        try:
            worker = threading.Thread(target=self._fetch_once,
                                      args=(context, serial, mode, (self._last_quote_source, self._last_quote_timestamp)), daemon=True)
            self._price_workers[serial] = worker
            worker.start()
        except Exception as exc:
            self._price_workers.pop(serial, None)
            self._fetch_results.put(("price", context, serial, None, str(exc)))
        self._start_fx_fetch(context)

    def _fetch_once(self, context, serial, mode, previous=None):
        # No Tk calls and no account/model mutations from this worker.
        _market_http_local.previous_quote = previous
        try:
            value = fetch_spot_usd(mode)
            result = ("price", context, serial, value, None)
        except Exception as exc:
            result = ("price", context, serial, None, str(exc))
        self._fetch_results.put(result)
        _market_http_local.previous_quote = None
        _close_market_session()

    def _start_fx_fetch(self, context):
        now = time.monotonic()
        if self._fx_busy or now < self._fx_retry_at:
            return
        if self._fx_last and now - self._fx_last <= 300:
            return
        self._fx_busy = True
        try:
            threading.Thread(target=self._fetch_fx_once, args=(context,), daemon=True).start()
        except Exception as exc:
            self._fetch_results.put(("fx", context, None, None, str(exc)))

    def _fetch_fx_once(self, context):
        try:
            data = http_json("https://open.er-api.com/v6/latest/USD", timeout=4)
            if data.get("result") != "success" or data.get("base_code") != "USD":
                raise ValueError("بيانات أسعار الصرف غير صحيحة")
            rates = {k: _valid_price(data["rates"][k])
                     for k in ("SAR", "JOD", "KWD", "QAR", "EGP")}
            result = ("fx", context, None, rates, None)
        except Exception as exc:
            result = ("fx", context, None, None, str(exc))
        self._fetch_results.put(result)
        _close_market_session()

    def _drain_fetch_results(self):
        self._result_job = None
        try:
            # Bounded drain keeps the UI responsive even after being busy.
            for _ in range(20):
                try:
                    kind, context, serial, value, error = self._fetch_results.get_nowait()
                except queue.Empty:
                    break
                if kind == "price":
                    self._price_workers.pop(serial, None)
                    if serial != self._fetch_serial:
                        continue  # Ignore a late response superseded by a newer request.
                    with self._fetch_lock:
                        if serial == self._fetch_serial:
                            self._fetch_busy = False
                else:
                    self._fx_busy = False
                if context != (self._fetch_generation, self.user) or not self.user:
                    continue  # A result from a previous login/restore is never applied.
                if kind == "fx":
                    self._fx_retry_at = time.monotonic() + (60 if error else 0)
                    if not error:
                        changed = any(abs(value[k] - self.rates.get(k, 0)) > 1e-9
                                      for k in value)
                        self.rates.update(value)
                        self._fx_last = time.monotonic()
                        if changed:
                            self.render()
                            self.save()
                else:
                    self._apply_price_result(value, error)
        except Exception as exc:
            self._live_error = str(exc)[:160]
        finally:
            try:
                self._supervise_live_refresh()
            finally:
                self._ensure_result_pump()

    @staticmethod
    def _short_price_source(source):
        raw=str(source or "").strip()
        if not raw:
            return ""
        aliases={"swissquote":"Swissquote XAU/USD", "gold-api":"gold-api", "goldprice.org":"goldprice.org",
                 "yahoo":"Yahoo Finance", "paxg":"PAXG", "gold futures":"Gold Futures"}
        for key,label in aliases.items():
            if key in raw.lower():
                return label
        clean=raw.split("•",1)[0].split("|",1)[0].strip()
        return clean[:40]+("…" if len(clean)>40 else "")

    def _update_caption(self):
        return "آخر تحديث: {}  •  المصدر: {}".format(
            getattr(self,"last_update","") or "—",getattr(self,"last_price_source","") or "—")

    def _set_update_caption(self):
        self.lbl_upd.config(text=self._update_caption())

    def _connection_label(self):
        if getattr(self,"_connection_state","waiting")=="retrying":
            return "السعر غير محدث • جارٍ إعادة الاتصال تلقائياً كل ٥ ثوانٍ • آخر جلب ناجح {}".format(self.last_update or "—")
        if not getattr(self,"_ticks",0):
            return "سعر محفوظ / بانتظار أول جلب ناجح في هذه الجلسة"
        points=getattr(self,"price_hist",[])
        if points and time.time()-points[-1][0]>15:
            return "السعر غير محدث • المصدر لم يرسل عينة حديثة • آخر جلب {}".format(self.last_update or "—")
        return "متصل • آخر جلب ناجح {}".format(self.last_update or "—")

    def _apply_price_result(self, value, error):
        if not error:
            try:
                new_price, source = value
                new_price = _valid_price(new_price)
                if isinstance(value, MarketQuote):
                    if (self._last_quote_timestamp is not None and
                            value.quote_time <= self._last_quote_timestamp):
                        self._set_update_caption()
                        return
                    self._last_quote_source = value.source_id
                    self._last_quote_timestamp = value.quote_time
            except Exception as exc:
                error = str(exc)
        if error:
            self._connection_state = "retrying"
            self._fails += 1
            self._live_error = str(error)[:500]
            self._set_update_caption()
            self.draw_chart()
            return
        old_price = self.spot_oz_usd
        changed = not old_price or abs(new_price - old_price) > 1e-9
        direction = ("up" if new_price > old_price else "down") if old_price and changed else None
        self._delta_oz = new_price - old_price if old_price else 0.0
        self.spot_oz_usd = new_price
        self._record_price_sample(new_price)
        self._ticks += 1
        self._connection_state = "connected"
        self._fails = 0
        self.last_update = datetime.now().strftime("%H:%M:%S")
        self.last_price_source=self._short_price_source(value.source_id if isinstance(value,MarketQuote) else source)
        self._set_update_caption()
        self.update_price_labels()  # Refresh cards/widget even for an unchanged quote.
        if changed:
            self.render()
            if direction:
                self.flash_update(direction)
        if self._ticks == 1 or self._ticks % 12 == 0:
            self.save()


    @staticmethod
    def _mix_color(c1, c2, a=0.18):
        """يمزج لونين بنسبة a (لتلوين خفيف للخلفية أثناء الوميض)."""
        try:
            p1 = [int(c1[i:i + 2], 16) for i in (1, 3, 5)]
            p2 = [int(c2[i:i + 2], 16) for i in (1, 3, 5)]
            return "#%02x%02x%02x" % tuple(
                max(0, min(255, int(p1[i] + (p2[i] - p1[i]) * a))) for i in range(3))
        except Exception:
            return c2

    def flash_update(self, direction=None):
        """وميض قصير متناسق مع اتجاه تغيّر السعر (أخضر صعود / أحمر هبوط)."""
        t = self.theme
        hi = {"up": t.get("profit", "#2ecc71"),
              "down": t.get("loss", "#ff5c5c"),
              "same": t.get("accent", "#e9c684")}.get(direction, t.get("muted", "#888"))
        rest = t.get("muted", "#888")
        card = t.get("card", "#12161e")
        soft = self._mix_color(card, hi, 0.20)
        # لا نلتقط لوناً وسط وميض سابق
        if getattr(self, "_flash_on", False):
            return
        self._flash_on = True
        saved = []
        for name in ("lbl_gr", "lbl_oz"):
            lbl = getattr(self, name, None)
            if lbl is None:
                continue
            try:
                saved.append((lbl, lbl.cget("bg"), lbl.cget("fg")))
                lbl.config(bg=soft, fg=hi)
            except Exception:
                pass

        def revert():
            for lbl, bg, fg in saved:
                try:
                    lbl.config(bg=bg, fg=fg)
                except Exception:
                    pass
            self._flash_on = False
        try:
            self.root.after(FLASH_HOLD, revert)
        except Exception:
            self._flash_on = False

    # ============================ العرض ============================
    def price_diag(self):
        """يفحص كل مصدر سعر على حدة ويعرض النتيجة بالثواني (لتشخيص أي مشكلة)."""
        t = self.theme
        BG, CARD = t.get("bg", "#0a0c11"), t.get("card", "#151a23")
        TXT, MUT, GOLD = t.get("text", "#fff"), t.get("muted", "#888"), t.get("accent", "#e9c684")
        win = tk.Toplevel(self.root)
        try:
            win.title("تشخيص تحديث السعر")
            win.configure(bg=BG)
            win.transient(self.root)
            w, h = 600, 430
            sw, sh = self._screen_size()
            win.geometry("{}x{}+{}+{}".format(w, h, max(0, (sw - w) // 2),
                                              max(0, (sh - h) // 3)))
        except Exception:
            pass
        tk.Label(win, text="🔍  تشخيص تحديث السعر", bg=BG, fg=GOLD,
                 font=("Segoe UI", 14, "bold")).pack(pady=(12, 2))
        tk.Label(win, text=("السعر الحالي: {}  •  آخر تحديث: {}  •  المصدر المختار: {}"
                            "\nالسوق: {}  •  محاولات فاشلة متتالية: {}").format(
            ("${:,.2f}".format(self.spot_oz_usd) if self.spot_oz_usd else "—"),
            self.last_update or "—", self.source_mode or "auto",
            ("مغلق" if market_closed() else "مفتوح"), self._fails),
                 bg=BG, fg=MUT, justify="right", font=("Segoe UI", 10)).pack(pady=(0, 6))
        box = tk.Frame(win, bg=CARD, highlightthickness=1,
                       highlightbackground=t.get("border", "#252c39"))
        box.pack(fill="both", expand=True, padx=14, pady=6)
        txt = tk.Text(box, height=9, bg=CARD, fg=TXT, relief="flat", bd=0,
                      font=("Segoe UI", 10), wrap="word")
        txt.pack(fill="both", expand=True, padx=8, pady=8)
        txt.insert("end", "{} • الطلبات {} • الجلب الناجح {}\nآخر خطأ داخلي: {}\nجارٍ فحص المصادر…\n".format(
            APP_VERSION, self._live_attempts, self._ticks, self._live_error or "لا يوجد"))

        def add(line):
            def do():
                try:
                    txt.insert("end", line + "\n")
                    txt.see("end")
                except Exception:
                    pass
            try:
                win.after(0, do)
            except Exception:
                pass

        def run():
            names = ("swissquote", "gold-api", "goldprice.org", "yahoo", "paxg", "stooq")
            for name in names:
                t0 = time.time()
                try:
                    price, _src = fetch_spot_usd(name, timeout=6, only=True)
                    add("✔ {}   ${:,.2f}   ({:.2f} ثانية)".format(
                        name, price, time.time() - t0))
                except Exception as e:
                    reason = str(e).split(":", 1)[-1].strip()[:70] or type(e).__name__
                    add("✘ {}   {}   ({:.2f} ثانية)".format(
                        name, reason, time.time() - t0))
            add("")
            add("النتيجة: التحديث التلقائي يستخدم مصادر الذهب الفوري فقط؛ العقود والرموز الرقمية للتشخيص."
                if self.spot_oz_usd else
                "النتيجة: لا يوجد سعر بعد — إن فشلت كل المصادر فالمشكلة في الاتصال/الجدار الناري.")

        threading.Thread(target=run, daemon=True).start()
        self._mkbtn(win, "إغلاق", win.destroy, kind="neutral",
                    font_size=10).pack(pady=(0, 10))

    def toggle_sells(self):
        """زر واحد: إخفاء/إظهار عمليات البيع التي تمّت من جدول العمليات."""
        self.show_sells = not bool(getattr(self, "show_sells", True))
        self._paint_sells_btn()
        try:
            self.save()
        except Exception:
            pass
        self.render()

    def _paint_sells_btn(self):
        btn = getattr(self, "btn_sells", None)
        if btn is None:
            return
        n = len(getattr(self, "sells", []) or [])
        try:
            btn.config(text=("👁 إخفاء البيع ({})".format(n) if
                             bool(getattr(self, "show_sells", True))
                             else "👁 إظهار البيع ({})".format(n)))
        except Exception:
            pass

    def _fit_operations_table(self):
        """Keep only visible rows; the page handles scrolling for long lists."""
        count = len(self.tree.get_children())
        self.tree.configure(height=count)
        empty = getattr(self, "empty_operations", None)
        toolbar = getattr(self, "operations_toolbar", None)
        if empty is not None and toolbar is not None:
            if count:
                empty.pack_forget()
            else:
                empty.pack(fill="x", padx=self._px(18), pady=(0, 2), before=toolbar)

    def render(self):
        selected = self.tree.selection() if hasattr(self.tree, "selection") else ()
        viewport = None
        canvas = getattr(getattr(self, "main_wrap", None), "canvas", None)
        if canvas is not None:
            try: viewport = canvas.yview()[0]
            except (tk.TclError, AttributeError, TypeError): pass
        for i in self.tree.get_children():
            self.tree.delete(i)
        self.row_map = {}
        t_cost = t_val = 0.0
        t_qty_g = 0.0
        realized = 0.0
        sym = CUR[self.cur]["sym"]
        rowi = 0
        # المراكز المفتوحة
        for idx, h in enumerate(self.holdings):
            unit_txt = "أونصة" if h["unit"] == "oz" else "غرام"
            cost_disp = h["qty"] * h["buy"] * self.rate()
            karat = h.get("karat", 24)
            cur_disp = self.current_per_unit_disp(h["unit"], karat)
            value_disp = h["qty"] * cur_disp
            pl = value_disp - cost_disp
            pct = (pl / cost_disp * 100) if cost_disp else 0
            t_cost += cost_disp
            t_val += value_disp
            t_qty_g += h["qty"] * (OZ_TO_G if h["unit"] == "oz" else 1)
            tag = "pos" if pl > 0 else ("neg" if pl < 0 else "")
            stripe = "odd" if rowi % 2 else "even"
            row_tags = (tag, stripe) if tag else (stripe,)
            iid = "h{}".format(idx)
            self.row_map[iid] = ("hold", idx)
            self.tree.insert("", "end", iid=iid, tags=row_tags,
                             values=tuple(self._hide(_v) for _v in (
                h["name"], "● مفتوح ({})".format(karat),
                "{:g} {}".format(h["qty"], unit_txt),
                "{}{:,.2f}".format(sym, h["buy"] * self.rate()),
                "{}{:,.2f}".format(sym, cost_disp),
                "{}{:,.2f}".format(sym, cur_disp),
                "{}{:,.2f}".format(sym, value_disp),
                "-",
                "{}{}{:,.2f}".format("+" if pl > 0 else "", sym, pl),
                "{}{:.2f}%".format("+" if pl > 0 else "", pct))))
            rowi += 1
        # عمليات البيع (خضراء — سعر بيع مثبّت). الربح المحقق يُحسب دائماً
        show_sells = bool(getattr(self, "show_sells", True))
        for sidx, s in enumerate(self.sells):
            unit_txt = "أونصة" if s["unit"] == "oz" else "غرام"
            cost_disp = s["qty"] * s["buy"] * self.rate()
            sell_disp = s["sell"] * self.rate()
            s_karat = s.get("karat", 24)
            value_disp = s["qty"] * sell_disp
            pl = value_disp - cost_disp
            pct = (pl / cost_disp * 100) if cost_disp else 0
            realized += pl
            if not show_sells:
                continue
            iid = "s{}".format(sidx)
            self.row_map[iid] = ("sell", sidx)
            self.tree.insert("", "end", iid=iid, tags=("soldgreen",),
                             values=tuple(self._hide(_v) for _v in (
                "💰 {}".format(s.get("name", "بيع ذهب")),
                "✓ تم البيع ({})".format(s_karat),
                "{:g} {}".format(s["qty"], unit_txt),
                "{}{:,.2f}".format(sym, s["buy"] * self.rate()),
                "{}{:,.2f}".format(sym, cost_disp),
                "-",
                "{}{:,.2f}".format(sym, value_disp),
                "{}{:,.2f}".format(sym, sell_disp),
                "{}{}{:,.2f}".format("+" if pl > 0 else "", sym, pl),
                "{}{:.2f}%".format("+" if pl > 0 else "", pct))))
        self._paint_sells_btn()
        self.update_price_labels()
        self.check_alerts()
        t_pl = t_val - t_cost
        t_pct = (t_pl / t_cost * 100) if t_cost else 0
        total_pl = t_pl + realized
        if self.holdings:
            ttag = "totalpos" if t_pl > 0 else ("totalneg" if t_pl < 0 else "total")
            self.tree.insert("", "end", tags=(ttag,),
                             values=tuple(self._hide(_v) for _v in (
                "∑ الرصيد المتبقي", "-", "{:,.2f} غرام".format(t_qty_g),
                "-", "{}{:,.2f}".format(sym, t_cost), "-",
                "{}{:,.2f}".format(sym, t_val), "-",
                "{}{}{:,.2f}".format("+" if t_pl > 0 else "", sym, t_pl),
                "{}{:.2f}%".format("+" if t_pl > 0 else "", t_pct))))
        self._fit_operations_table()
        if selected and selected[0] in self.tree.get_children():
            self.tree.selection_set(selected[0])
        self._sync_selection_actions()
        self.lbl_sum.config(
            text=self._hide(
                ("الربح المحقق: {sg_r}{s}{r:,.2f}    |    غير المحقق: {sg_u}{s}{u:,.2f}"
                 "    |    إجمالي الربح/الخسارة: {sg_t}{s}{tp:,.2f}").format(
                    s=sym, r=realized, u=t_pl, tp=total_pl,
                    sg_r="+" if realized > 0 else "",
                    sg_u="+" if t_pl > 0 else "",
                    sg_t="+" if total_pl > 0 else "")),
            fg=self.theme["total"])
        self._update_dashboard(t_qty_g, t_cost, t_val, t_pl)
        if viewport is not None:
            try: canvas.yview_moveto(viewport)
            except tk.TclError: pass

    def update_price_labels(self):
        if self.spot_oz_usd:
            sym = CUR[self.cur]["sym"]
            kf = karat_factor(self.karat)
            # الأونصة العالمية تبقى بسعر الذهب الخالص
            self.lbl_oz.config(text=self._hide("{}{:,.3f}".format(
                sym, self.spot_oz_usd * self.rate())))
            self.lbl_gr.config(text=self._hide("{}{:,.3f}".format(sym, self.gram_market_disp() * kf)))
            self.lbl_gr_buy.config(text=self._hide("{}{:,.3f}".format(sym, self.gram_buy_disp() * kf)))
            self.lbl_gr_adj.config(text=self._hide("{}{:,.3f}".format(sym, self.gram_eff_disp() * kf)))
            bases = getattr(self, "_card_base_titles", {})
            for attr in ("lbl_gr", "lbl_gr_buy", "lbl_gr_adj"):
                lbl_t = getattr(self, attr + "_title", None)
                if lbl_t is None or attr not in bases:
                    continue
                base = bases[attr]
                lbl_t.config(text=base if self.karat == 24
                             else "{} — عيار {}".format(base, self.karat))
            self._update_delta_indicator()
        else:
            for lbl in (self.lbl_oz, self.lbl_gr, self.lbl_gr_buy, self.lbl_gr_adj):
                lbl.config(text="--")
        self._refresh_widget()
        self._refresh_chart_if_visible()

    def _update_delta_indicator(self):
        lbl = getattr(self, "lbl_delta", None)
        if lbl is None:
            return
        t = self.theme
        sym = CUR[self.cur]["sym"]
        delta_oz = self._delta_oz or 0.0
        if abs(delta_oz) < 1e-9 or not self.spot_oz_usd:
            lbl.config(text=self._hide("━ ثابت"), fg=t.get("flash_same", "#ffd93b"))
            return
        gram_delta = (delta_oz / OZ_TO_G) * self.rate()
        prev_oz = self.spot_oz_usd - delta_oz
        pct = (delta_oz / prev_oz * 100) if prev_oz else 0.0
        if delta_oz > 0:
            arrow, color = "▲", t.get("profit", "#2ecc71")
        else:
            arrow, color = "▼", t.get("loss", "#ff5c5c")
        lbl.config(text=self._hide("{a} {sym}{d:,.3f}/غ ({p:+.2f}%)".format(
            a=arrow, sym=sym, d=abs(gram_delta), p=pct)), fg=color)


# ============================ تصدير Excel ============================
def write_xlsx(path, headers, rows):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
    except ImportError:
        raise RuntimeError("مكتبة openpyxl غير مثبّتة. نفّذ: pip install openpyxl")
    wb = Workbook()
    ws = wb.active
    ws.title = "محفظة الذهب"
    ws.sheet_view.rightToLeft = True
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E79")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for r in rows:
        ws.append(r)
    for col in ws.columns:
        width = max((len(str(c.value)) if c.value is not None else 0) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(width + 4, 12), 34)
    wb.save(path)


def _enable_dpi_awareness():
    """خطوط حادة على شاشات ويندوز عالية الدقة + معرفة نسبة التكبير."""
    global SYSTEM_DPI
    if os.name != "nt":
        return
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
        try:
            dpi = int(ctypes.windll.user32.GetDpiForSystem())
        except Exception:
            dpi = 0
        if not dpi:
            try:
                hdc = ctypes.windll.user32.GetDC(0)
                dpi = int(ctypes.windll.gdi32.GetDeviceCaps(hdc, 88))  # LOGPIXELSX
                ctypes.windll.user32.ReleaseDC(0, hdc)
            except Exception:
                dpi = 96
        if dpi and 72 <= dpi <= 480:
            SYSTEM_DPI = dpi
    except Exception:
        pass


def main():
    _enable_dpi_awareness()
    root = tk.Tk()
    def report_error(exc_type, exc_value, exc_traceback):
        import traceback
        details = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
        try:
            log_path = os.path.join(HOME_DIR, "GoldPortfolio-startup-error.log")
            with open(log_path, "a", encoding="utf-8") as log:
                log.write("\n" + datetime.now().isoformat() + " " + APP_VERSION + "\n" + details)
        except OSError:
            pass
        print(details, file=sys.stderr)
        try:
            messagebox.showerror("خطأ في تشغيل محفظة الذهب",
                                 "{}\n\n{}\n\nتمت محاولة حفظ التفاصيل في GoldPortfolio-startup-error.log داخل مجلد المستخدم.".format(
                                     APP_VERSION, str(exc_value)[:350]), parent=root)
        except tk.TclError:
            pass
    root.report_callback_exception = report_error
    try:
        GoldApp(root)
    except Exception:
        report_error(*sys.exc_info())
        root.destroy()
        raise
    root.mainloop()


if __name__ == "__main__":
    main()
