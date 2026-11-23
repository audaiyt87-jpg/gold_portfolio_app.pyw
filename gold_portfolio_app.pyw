# -*- coding: utf-8 -*-
"""
محفظة الذهب — Gold Portfolio
تطبيق سطح مكتب (Tkinter) لتتبّع محفظة الذهب: شراء / بيع / ربح وخسارة /
حاسبة زكاة / سجل عمليات / مخطط سعر الغرام. يعمل بمكتبة بايثون القياسية فقط
(التصدير إلى Excel اختياري إن توفّر openpyxl).
"""
import os
import json
import time
import hashlib
import hmac
import shutil
import threading
import math
import re
import sys
import glob
import base64
import subprocess
import tempfile
from datetime import datetime
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
from email.mime.text import MIMEText
from email import message_from_bytes

# ============================ ثوابت عامة ============================
OZ_TO_G = 31.1034768                      # غرام في الأونصة الواحدة
LIVE_MS = 5000                            # تحديث السعر المباشر كل 5 ثوانٍ
FLASH_MS = 7000                           # وميض تغيّر السعر: إيقاع ثابت كل 7 ثوانٍ
FLASH_HOLD = 320                          # مدة بقاء الوميض (مللي ثانية)
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
             "sound_on", "source_mode", "dark_mode", "lock_minutes")


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
CHART_RANGES = (("hour", "ساعة", 3600),
                ("day", "يوم", 86400),
                ("week", "أسبوع", 7 * 86400))


def karat_factor(karat):
    """معامل نقاوة العيار نسبةً إلى الذهب الخالص 24."""
    try:
        k = float(karat)
    except (TypeError, ValueError):
        k = 24.0
    return k / 24.0

DEFAULT_THEME = {
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
LIGHT_THEME = {
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
def http_json(url, timeout=10):
    req = urllib.request.Request(url, headers={"User-Agent": BROWSER_UA,
                                              "Accept": "application/json",
                                              "Cache-Control": "no-cache",
                                              "Pragma": "no-cache"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8", "ignore"))


def _http_text(url, timeout=10):
    req = urllib.request.Request(url, headers={"User-Agent": BROWSER_UA,
                                              "Cache-Control": "no-cache",
                                              "Pragma": "no-cache"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read().decode("utf-8", "ignore")


def _yahoo_last(symbol):
    data = http_json("https://query1.finance.yahoo.com/v8/finance/chart/{}"
                     "?interval=1m&range=1d".format(symbol), timeout=5)
    res = (data.get("chart") or {}).get("result") or []
    meta = (res[0] or {}).get("meta") or {}
    price = meta.get("regularMarketPrice")
    if not price:
        q = ((res[0].get("indicators") or {}).get("quote") or [{}])[0]
        closes = [c for c in (q.get("close") or []) if c]
        price = closes[-1] if closes else None
    return price


def fetch_spot_usd(mode="auto", timeout=4, only=False):
    """جلب سعر أونصة الذهب بالدولار من عدة مصادر مجانية (بمهلة قصيرة لكل مصدر)."""
    sources = [
        ("gold-api", lambda: http_json("https://api.gold-api.com/price/XAU",
                                       timeout=timeout).get("price")),
        ("goldprice.org", lambda: http_json(
            "https://data-asg.goldprice.org/dbXRates/USD",
            timeout=timeout)["items"][0]["xauPrice"]),
        ("yahoo", lambda: _yahoo_last("GC=F")),
        ("paxg", lambda: http_json(
            "https://api.coingecko.com/api/v3/simple/price"
            "?ids=pax-gold&vs_currencies=usd", timeout=timeout)["pax-gold"]["usd"]),
        ("stooq", lambda: float(
            _http_text("https://stooq.com/q/l/?s=xauusd&f=sd2t2ohlcv&h&e=csv",
                       timeout=timeout).strip().splitlines()[-1].split(",")[6])),
    ]
    if mode and mode != "auto":
        if only:
            sources = [s for s in sources if s[0] == mode]
        else:
            # المصدر المختار أولاً، ثم البقية تلقائياً إن تعذّر (لا يتوقف التحديث أبداً)
            first = [s for s in sources if s[0] == mode]
            rest = [s for s in sources if s[0] != mode]
            sources = first + rest
    last_err = None
    started = time.time()
    for name, fn in sources:
        if time.time() - started > 12:
            break
        try:
            val = fn()
            if val and float(val) > 0:
                return float(val), name
        except Exception as e:  # noqa
            last_err = "{}: {}".format(name, e)
            continue
    raise RuntimeError("تعذّر جلب السعر: {}".format(last_err))


def market_closed():
    """سوق الذهب العالمي مغلق (السبت والأحد المبكر)."""
    wd = datetime.utcnow().weekday()   # 0=الاثنين ... 5=السبت 6=الأحد
    return wd == 5 or (wd == 6 and datetime.utcnow().hour < 22)


# ==================== أيقونة التطبيق (مضمّنة داخل الملف) ====================
# صورة ICO لعملة ذهبية، مضمّنة كنص Base64 حتى يعمل الملف وحده دون ملفات مساعدة.
APP_ICON_ICO_B64 = (
    "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"
)


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

    WIDTH = 340

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

        # -------- ثلاثة أعمدة بجانب بعضها: شراء ر.س | شراء د.أ | القيمة --------
        strip = tk.Frame(body, bg=self.card)
        strip.pack(fill="x")
        self.v_sar = self._cell(strip, "شراء (ر.س)")
        self._vsep(strip)
        self.v_jod = self._cell(strip, "شراء (د.أ)")
        self._vsep(strip)
        self.v_val = self._cell(strip, "القيمة الحالية", color=self.accent,
                                big=True)
        self.cols = strip

    def _cell(self, parent, title, color=None, big=False):
        """عمود صغير: العنوان فوق القيمة، بداخل الشريط الأفقي."""
        px = self.app._px
        fr = tk.Frame(parent, bg=self.card)
        fr.pack(side="right", padx=px(4), pady=px(1))
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
            x = max(-20, min(int(x), sw - 120))
            y = max(0, min(int(y), sh - 60))
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
                x = max(-20, min(x, sw - 120))
                y = max(0, min(y, sh - 60))
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
            self.lbl_karat.config(text=str(getattr(self.app, "karat", 24)))
            return
        prev = getattr(self, "_prev", None) or {}
        for key, lbl in (("sar", self.v_sar), ("jod", self.v_jod)):
            cur = d[key]
            color = self.text
            if key in prev:
                if cur["buy"] > prev[key] + 1e-9:
                    color = self.profit
                elif cur["buy"] < prev[key] - 1e-9:
                    color = self.loss
            lbl.config(text=hide("{:,.3f}".format(cur["buy"])), fg=color)
        self._prev = {"sar": d["sar"]["buy"], "jod": d["jod"]["buy"]}
        try:
            sym = CUR[self.app.cur]["sym"]
            val = self.app.portfolio_value_disp()
            self.v_val.config(text=hide("{}{:,.2f}".format(sym, val)))
        except Exception:
            self.v_val.config(text="--")
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
        self.chart_range = "day"    # النطاق الزمني للمخطط: hour / day / week
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
        self._flash_job = None      # مؤقّت الوميض الثابت (7 ثوانٍ)
        self._flash_price = 0.0     # السعر عند آخر وميض
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
            when = "  •  آخر تحديث {}".format(self.last_update) if self.last_update else ""
            try:
                self.lbl_upd.config(
                    text="○ آخر سعر محفوظ • الأونصة ${:.3f}{} • جارٍ التحديث…".format(
                        self.spot_oz_usd, when))
            except Exception:
                pass
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
            cr = d.get("chart_range", "day")
            self.chart_range = cr if cr in ("hour", "day", "week") else "day"
            self.adjust = float(d.get("adjust", 5.22))
            self.buy_margin = float(d.get("buy_margin", 0.0))
            self.alerts = d.get("alerts", [])
            self.dark_mode = bool(d.get("dark_mode", True))
            self.theme = dict(self._base_theme())
            if d.get("theme_ver") == THEME_VER:
                self.theme.update(d.get("theme", {}))
            self.titles.update(d.get("titles", {}))
            self.sound_on = bool(d.get("sound_on", True))
            self.source_mode = d.get("source_mode", "auto")
            self.show_sells = bool(d.get("show_sells", True))
            self.spot_oz_usd = float(d.get("spot_oz_usd", 0.0) or 0.0)
            self.rates.update(d.get("rates", {}))
            self.last_update = d.get("last_update", "")
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
                    "price_hist": self.price_hist[-500:],
                    "cur": self.cur,
                    "karat": self.karat,
                    "chart_range": self.chart_range,
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
        self.chart_range = "day"
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
        palettes = {
            "primary": (t["accent"], t["accent_text"], t.get("hover", "#f3d8a5")),
            "success": (t.get("profit", "#3ddc97"), "#07130d", "#6ae8b4"),
            "danger": (t.get("loss", "#ff6b6b"), "#1a0a0a", "#ff9292"),
            "neutral": (t["card2"], t["text"], t.get("border", "#252c39")),
            "ghost": (t["bg"], t["accent"], t.get("card2", "#1a2029")),
        }
        bg, fg, hv = palettes.get(kind, palettes["primary"])
        btn = tk.Button(parent, text=text, command=cmd, bg=bg, fg=fg,
                        activebackground=hv, activeforeground=fg, relief="flat",
                        bd=0, padx=16, pady=7, cursor="hand2",
                        highlightthickness=1,
                        highlightbackground=t.get("border", "#252c39"),
                        highlightcolor=hv,
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

    def _price_card(self, parent, title, attr, color, col):
        t = self.theme
        card = tk.Frame(parent, bg=t["card"], highlightthickness=1,
                        highlightbackground=t.get("border", "#252c39"))
        card.grid(row=0, column=col, padx=7, pady=(0, 4), sticky="nsew")
        tk.Frame(card, bg=color, height=3).pack(fill="x")
        title_lbl = tk.Label(card, text=title, bg=t["card"], fg=t["muted"],
                             font=("Segoe UI", 10))
        title_lbl.pack(pady=(11, 2))
        lbl = tk.Label(card, text="--", bg=t["card"], fg=color,
                       font=("Segoe UI", 18, "bold"))
        lbl.pack(pady=(0, 13))
        setattr(self, attr, lbl)
        setattr(self, attr + "_title", title_lbl)
        self._card_base_titles = getattr(self, "_card_base_titles", {})
        self._card_base_titles[attr] = title
        parent.grid_columnconfigure(col, weight=1)
        return card

    # ============================ بناء الواجهة ============================
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
        self.dark_mode = not getattr(self, "dark_mode", True)
        self.theme = dict(self._base_theme())
        self.save()
        self.rebuild_ui()

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
        card = tk.Frame(parent, bg=t.get("card2", t["card"]),
                        highlightthickness=1,
                        highlightbackground=t.get("border", "#252c39"))
        card.grid(row=0, column=col, padx=6, sticky="nsew")
        tk.Label(card, text=title, bg=t.get("card2", t["card"]), fg=t["muted"],
                 font=("Segoe UI", 10)).pack(pady=(5, 0))
        lbl = tk.Label(card, text="--", bg=t.get("card2", t["card"]), fg=color,
                       font=("Segoe UI", 13, "bold"))
        lbl.pack(pady=(1, 6))
        setattr(self, attr, lbl)
        parent.grid_columnconfigure(col, weight=1)
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

    def build_ui(self):
        t = self.theme
        BG, CARD, CARD2 = t["bg"], t["card"], t["card2"]
        GOLD, ATXT, TXT, MUT = t["accent"], t["accent_text"], t["text"], t["muted"]
        self.root.configure(bg=BG)
        self.root.title(self.titles.get("window", DEFAULT_TITLES["window"]))
        self._apply_screen_layout()
        p = self._px

        # حاوية الصفحات: الرئيسية + المخطط (تنقّل داخل نفس النافذة)
        self.page_holder = tk.Frame(self.root, bg=BG)
        self.page_holder.pack(fill="both", expand=True)
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
                        padding=(p(9), p(10)))
        style.map("Treeview",
                  background=[("selected", t.get("border", "#2c3a52"))],
                  foreground=[("selected", "#ffffff")])
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

        # العنوان
        hdr = tk.Frame(self.root, bg=BG)
        hdr.pack(fill="x", padx=p(22), pady=(p(16), p(4)),
                 before=self.page_holder)
        compact = (getattr(self, "screen_h", 900) < 800 or
                   getattr(self, "screen_w", 1400) < 1180)
        tk.Label(hdr, text=self.titles.get("title"), bg=BG, fg=GOLD,
                 font=("Segoe UI", 18 if compact else 22, "bold")).pack(side="right")
        if not compact:
            tk.Label(hdr, text=self.titles.get("subtitle"), bg=BG, fg=MUT,
                     font=("Segoe UI", 10)).pack(side="right", padx=p(16))
        if getattr(self, "user", None):
            chip = tk.Frame(hdr, bg=BG)
            chip.pack(side="left")
            self.btn_privacy = self._mkbtn(
                chip, "👁" if getattr(self, "privacy", False) else "🙈",
                self.toggle_privacy, kind="ghost", font_size=10)
            self.btn_privacy.pack(side="left", padx=(8, 0))
            self._mkbtn(chip, "🌓", self.toggle_theme, kind="ghost",
                        font_size=10).pack(side="left", padx=(4, 0))
            self.btn_cloud = self._mkbtn(chip, "☁",
                                         lambda: self.cloud_sync_now(True),
                                         kind="ghost", font_size=10)
            self.btn_cloud.pack(side="left", padx=(4, 0))
            self._mkbtn(chip, "الحساب", self.open_account, kind="ghost",
                        font_size=10).pack(side="left", padx=(4, 0))
            tk.Label(chip, text="👤  " + self.user, bg=BG, fg=TXT,
                     font=("Segoe UI", 11, "bold")).pack(side="left")
        tk.Frame(self.root, bg=t.get("border", "#252c39"),
                 height=1).pack(fill="x", padx=20, pady=(4, 0),
                                before=self.page_holder)

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
        self.rajhi_frame = tk.Frame(ctrl, bg=BG)
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

        self._mkbtn(ctrl, "⚙ الإعدادات", self.open_settings, kind="neutral",
                    font_size=10).pack(side="left", padx=6)
        self._mkbtn(ctrl, "🧮 حاسبة التكلفة", self.open_calc, kind="neutral",
                    font_size=10).pack(side="left", padx=6)

        # بطاقات الأسعار تحت سطر العملة والعيار مباشرة
        prices.pack(fill="x", padx=p(16), pady=(p(8), p(8)))

        # ملاحظة تظهر عند اختيار عملة غير الريال السعودي
        self.rajhi_note = tk.Label(
            ctrl, text="ℹ أسعار الراجحي متاحة بالريال السعودي فقط",
            bg=BG, fg=MUT, font=("Segoe UI", 10))

        # لوحة ملخص سريعة (تُقرأ بنظرة واحدة)
        dash = tk.Frame(self.main_page, bg=BG)
        dash.pack(fill="x", padx=p(16), pady=(p(2), p(2)))
        self._dash_card(dash, "إجمالي الغرامات", "dash_qty", TXT, 0)
        self._dash_card(dash, "التكلفة", "dash_cost", TXT, 1)
        self._dash_card(dash, "القيمة الحالية", "dash_val", GOLD, 2)
        self._dash_card(dash, "الربح غير المحقق", "dash_pl", t["profit"], 3)

        # صف التحديث والمؤشر
        updfr = tk.Frame(self.main_page, bg=BG)
        updfr.pack(fill="x", padx=20)
        self.dot = tk.Label(updfr, text="●", bg=BG, fg=MUT, font=("Segoe UI", 10))
        self.dot.pack(side="right", padx=(6, 0))
        self.lbl_upd = tk.Label(updfr, text="جارٍ التحميل...", bg=BG, fg=MUT,
                                font=("Segoe UI", 10), cursor="hand2")
        self.lbl_upd.pack(side="right")
        try:
            # نقرة على سطر الحالة = تحديث فوري للسعر
            self.lbl_upd.bind("<Button-1>", lambda e: self.refresh_all())
        except Exception:
            pass
        self.lbl_delta = tk.Label(updfr, text="", bg=BG, fg=MUT,
                                  font=("Segoe UI", 11, "bold"))
        self.lbl_delta.pack(side="left", padx=(0, 0))
        self._mkbtn(updfr, "🔍 تشخيص", self.price_diag, kind="ghost",
                    font_size=10).pack(side="left", padx=(8, 0))

        # شريط الإضافة: زر واحد فقط — تفاصيل العملية تُدخَل في صفحة مستقلة
        addbar = tk.Frame(self.main_page, bg=BG)
        addbar.pack(fill="x", padx=p(18), pady=(p(4), p(2)))
        self.btn_add = self._mkbtn(addbar, "＋  إضافة عملية (شراء / بيع)",
                                   self.open_add_chooser, kind="primary",
                                   font_size=12)
        self.btn_add.pack(side="right", pady=2)
        tk.Label(addbar,
                 text="اختر «شراء» أو «بيع» ثم أدخل الوزن والعيار والسعر في صفحة مستقلة",
                 bg=BG, fg=MUT, font=("Segoe UI", 10)).pack(side="right",
                                                            padx=p(10))
        # إظهار/إخفاء صفوف البيع في قائمة العمليات
        self.btn_sells = self._mkbtn(addbar, "👁 إخفاء البيع", self.toggle_sells,
                                     kind="ghost", font_size=10)
        self.btn_sells.pack(side="left", pady=2)
        self._paint_sells_btn()

        # الجدول
        cols = ("name", "status", "qty", "buy", "cost", "cur", "val",
                "sellprice", "pl", "pct")
        heads = ("الاسم", "الحالة", "الكمية", "سعر الشراء", "التكلفة",
                 "السعر الحالي", "القيمة", "سعر البيع", "الربح/الخسارة", "%")
        _sh = getattr(self, "screen_h", 900)
        tree_rows = 9 if _sh <= 768 else (12 if _sh <= 820 else 15)
        tree_wrap = tk.Frame(self.main_page, bg=BG)
        tree_wrap.pack(fill="both", expand=True, padx=p(18), pady=p(6))
        self.tree = ttk.Treeview(tree_wrap, columns=cols, show="headings",
                                 height=tree_rows)
        for c, h in zip(cols, heads):
            self.tree.heading(c, text=h)
            self.tree.column(c, anchor="center", width=110)
        self.tree.column("name", width=130)
        self.tree.column("status", width=110)
        self.tree.column("sellprice", width=125)
        self.tree.column("pct", width=80)
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
        self.tree.pack(side="top", fill="both", expand=True)
        self.tree_x.pack(side="bottom", fill="x", pady=(2, 0))
        # نقر مزدوج على أي عملية شراء => فتح أمر البيع عليها مباشرة
        try:
            self.tree.bind("<Double-1>", self._tree_dblclick)
        except Exception:
            pass

        # الأزرار: كل الخيارات في شطر واحد (بلا «أمر بيع» — البيع من
        # زر «إضافة عملية (شراء / بيع)» أو Ctrl+S)
        btns = tk.Frame(self.main_page, bg=BG)
        btns.pack(fill="x", padx=p(12), pady=(p(6), p(2)))
        self._mkbtn(btns, "🗑 حذف", self.del_item, kind="danger",
                    font_size=10).pack(side="right", padx=2)
        self._mkbtn(btns, "مسح الكل", self.clear_all, kind="neutral",
                    font_size=10).pack(side="right", padx=2)
        self._mkbtn(btns, "📜 السجل", self.open_history, kind="neutral",
                    font_size=10).pack(side="right", padx=2)
        self._mkbtn(btns, "⬇ Excel", self.export_excel, kind="success",
                    font_size=10).pack(side="right", padx=2)
        self._mkbtn(btns, "🔔 تنبيه", self.open_alerts, kind="primary",
                    font_size=10).pack(side="left", padx=2)
        self._mkbtn(btns, "📈 المخطط", self.show_chart, kind="primary",
                    font_size=10).pack(side="left", padx=2)
        self._mkbtn(btns, "📐 الزكاة", self.open_zakat, kind="neutral",
                    font_size=10).pack(side="left", padx=2)
        self.btn_widget = self._mkbtn(btns, "🪟 ويدجيت", self.toggle_widget,
                                      kind="neutral", font_size=10)
        self.btn_widget.pack(side="left", padx=2)
        self._paint_widget_btn()

        # الملخص
        self.lbl_sum = tk.Label(self.main_page, text="", bg=BG, fg=t["total"],
                                font=("Segoe UI", 11, "bold"))
        self.lbl_sum.pack(fill="x", padx=p(20), pady=p(10))
        self._update_rajhi_visibility()

        # صفحة المخطط (مخفية حتى يُطلب عرضها)
        self.chart_page = tk.Frame(self.page_holder, bg=BG)
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
        body = self._open_page("addbuy", "＋ إضافة عملية شراء")
        body.title("＋ إضافة عملية شراء")

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

        grid = tk.Frame(card, bg=CARD)
        grid.pack(fill="x", padx=self._px(10), pady=(self._px(6), self._px(14)))
        for c in (0, 1):
            grid.columnconfigure(c, weight=1, uniform="ab")

        def entry(parent, var, fsize=12):
            return tk.Entry(parent, textvariable=var, justify="center",
                            bg=t.get("entry_bg", "#151a23"), fg=t["text"],
                            insertbackground=t["text"], relief="flat", bd=0,
                            highlightthickness=1,
                            highlightbackground=t.get("border", "#252c39"),
                            highlightcolor=t["accent"],
                            font=("Segoe UI", fsize))

        def cell(row, col, label, builder, hint=""):
            box = tk.Frame(grid, bg=CARD)
            box.grid(row=row, column=col, sticky="ew", padx=self._px(8),
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
        self.ab_qty_e = cell(0, 0, "الوزن", lambda p: entry(p, self.ab_qty),
                             hint="بالوحدة المختارة")
        cell(1, 1, "الوحدة",
             lambda p: ttk.OptionMenu(p, self.ab_unit, self.ab_unit.get(),
                                      "غرام", "أونصة"))
        cell(1, 0, "العيار",
             lambda p: ttk.OptionMenu(p, self.ab_karat, self.ab_karat.get(),
                                      *["{} قيراط".format(k) for k in KARATS]))
        self.ab_price_e = cell(2, 1, "سعر الشراء للوحدة",
                               lambda p: entry(p, self.ab_price),
                               hint="بالعملة الحالية: {}".format(
                                   CUR[self.cur]["name"]))
        tot_box = tk.Frame(grid, bg=CARD)
        tot_box.grid(row=2, column=0, sticky="ew", padx=self._px(8),
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

        bar = tk.Frame(body, bg=t["bg"])
        bar.pack(fill="x", padx=self._px(4), pady=(self._px(2), self._px(8)))
        self._mkbtn(bar, "💾  حفظ العملية", self.add_item, kind="primary",
                    font_size=12).pack(side="right")
        self._mkbtn(bar, "⟲  السعر الحالي", self._addbuy_use_price,
                    kind="neutral", font_size=10).pack(side="right",
                                                       padx=self._px(8))
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

    def del_item(self):
        sel = self.tree.selection()
        if not sel:
            return
        kind_idx = self.row_map.get(sel[0])
        if not kind_idx:
            return
        kind, idx = kind_idx
        if kind == "hold" and 0 <= idx < len(self.holdings):
            del self.holdings[idx]
            self.save()
            self.render()
        elif kind == "sell" and 0 <= idx < len(self.sells):
            if messagebox.askyesno("تراجع عن بيع",
                                   "حذف عملية البيع؟ ستُعاد الكمية إلى الرصيد."):
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

        bar = tk.Frame(win, bg=BG)
        bar.pack(fill="x", padx=px(12), pady=(px(4), px(10)))
        self._mkbtn(bar, "✔  تأكيد البيع", do_sell, kind="success",
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
        """يحفظ عينة سعر: كل ١٥ ثانية كحدّ أقصى لآخر 24 ساعة، وعينة/5 دقائق حتى أسبوع."""
        try:
            now = time.time()
            val = float(oz_usd)
            if self.price_hist:
                try:
                    last_ts = float(self.price_hist[-1][0])
                except Exception:
                    last_ts = 0.0
                if now - last_ts < 15:
                    # نحدّث آخر نقطة بدل إضافة نقطة كل ٥ ثوانٍ (يمنع امتلاء السجل)
                    self.price_hist[-1] = [now, val]
                    return
            self.price_hist.append([now, val])
            keep_from = now - 7 * 86400
            day_ago = now - 86400
            recent, older = [], []
            for ts, oz in self.price_hist:
                if ts < keep_from:
                    continue
                if ts >= day_ago:
                    recent.append([ts, oz])
                else:
                    older.append([ts, oz])
            # تخفيف الأقدم من 24 ساعة: عينة واحدة لكل 5 دقائق
            thinned, bucket = [], None
            for ts, oz in older:
                b = int(ts // 300)
                if b != bucket:
                    thinned.append([ts, oz])
                    bucket = b
                else:
                    thinned[-1] = [ts, oz]
            self.price_hist = (thinned + recent)[-9000:]
        except Exception:
            pass

    # ============================ صفحة المخطط ============================
    def _build_chart_page(self):
        t = self.theme
        BG, CARD = t["bg"], t["card"]
        GOLD, MUT = t["accent"], t["muted"]
        px = self._px
        head = tk.Frame(self.chart_page, bg=BG)
        head.pack(fill="x", padx=px(24), pady=(px(14), px(8)))
        tk.Label(head, text="📈 مخطط سعر الغرام", bg=BG, fg=GOLD,
                 font=("Segoe UI", 16, "bold")).pack(side="right")
        self._mkbtn(head, "🏠 الرئيسية", lambda: self._show_page("main"),
                    kind="ghost", font_size=10).pack(side="left")
        tk.Label(head, text="يتحدّث تلقائياً مع كل تحديث للسعر",
                 bg=BG, fg=MUT, font=("Segoe UI", 10)).pack(side="left", padx=12)
        tk.Frame(self.chart_page, bg=t.get("border", "#252c39"),
                 height=1).pack(fill="x", padx=px(22))

        rng = tk.Frame(self.chart_page, bg=BG)
        rng.pack(fill="x", padx=px(24), pady=(px(10), 0))
        tk.Label(rng, text="النطاق الزمني:", bg=BG, fg=MUT,
                 font=("Segoe UI", 10)).pack(side="right", padx=(8, 6))
        self._range_btns = {}
        for key, label, _secs in CHART_RANGES:
            b = tk.Label(rng, text=label, bg=t.get("card2", CARD), fg=MUT,
                         padx=18, pady=4, cursor="hand2",
                         highlightthickness=1,
                         highlightbackground=t.get("border", "#252c39"),
                         font=("Segoe UI", 11, "bold"))
            b.pack(side="right", padx=4)
            b.bind("<Button-1>", lambda e, k=key: self._set_chart_range(k))
            self._range_btns[key] = b

        self.chart_info = tk.Label(self.chart_page, text="", bg=BG, fg=MUT,
                                   font=("Segoe UI", 11, "bold"))
        self.chart_info.pack(anchor="e", padx=26, pady=(10, 4))
        self._paint_range_buttons()

        wrap = tk.Frame(self.chart_page, bg=CARD, highlightthickness=1,
                        highlightbackground=t.get("border", "#252c39"))
        wrap.pack(fill="both", expand=True, padx=20, pady=(4, 14))
        self.chart_canvas = tk.Canvas(wrap, bg=CARD, highlightthickness=0)
        self.chart_canvas.pack(fill="both", expand=True, padx=8, pady=8)
        self.chart_canvas.bind("<Configure>", lambda e: self.draw_chart())

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
        for page in pages.values():
            if page is not None and page.winfo_manager():
                page.pack_forget()
        pages[name].pack(fill="both", expand=True)
        # إعادة أي صفحة إلى أعلاها عند فتحها (لا يظهر منتصفها)
        info = getattr(self, "_pages", {}).get(name) or {}
        scroller = info.get("scroll")
        if name == "main":
            scroller = getattr(self, "main_wrap", None)
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
                             font=("Segoe UI", 18, "bold"))
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
                            "body": body, "scroll": scroll}
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
        if was_widget and getattr(self, "user", None):
            self.open_widget()

    def show_chart(self):
        self._show_page("chart")

    def _set_chart_range(self, key):
        if key not in ("hour", "day", "week") or key == self.chart_range:
            return
        self.chart_range = key
        self.save()
        self._paint_range_buttons()
        self.draw_chart()

    def _paint_range_buttons(self):
        """تلوين زر النطاق النشط."""
        t = self.theme
        for key, btn in getattr(self, "_range_btns", {}).items():
            active = (key == self.chart_range)
            try:
                btn.config(bg=t["accent"] if active else t.get("card2", t["card"]),
                           fg=t.get("accent_text", "#000") if active else t["muted"])
            except Exception:
                pass

    def _range_seconds(self):
        for key, _label, secs in CHART_RANGES:
            if key == self.chart_range:
                return secs
        return 86400

    def _range_label(self):
        for key, label, _secs in CHART_RANGES:
            if key == self.chart_range:
                return label
        return "يوم"

    def _chart_points(self):
        """نقاط النطاق المختار (وقت، سعر الغرام) بعيار وعملة العرض الحاليين."""
        now = time.time()
        cutoff = now - self._range_seconds()
        rate = self.rate()
        kf = karat_factor(self.karat)
        pts = [(ts, (oz / OZ_TO_G) * rate * kf) for ts, oz in self.price_hist
               if ts >= cutoff]
        # تخفيف نقاط الرسم للنطاقات الطويلة (مع الإبقاء على أول وآخر نقطة)
        if len(pts) > 500:
            step = len(pts) // 500 + 1
            pts = pts[::step] + [pts[-1]]
        return pts

    def _fmt_tick(self, ts):
        d = datetime.fromtimestamp(ts)
        if self.chart_range == "week":
            return d.strftime("%m-%d\n%H:%M")
        return d.strftime("%H:%M")

    def draw_chart(self):
        """يرسم المخطط على صفحة المخطط (يُستدعى عند العرض وعند كل تحديث سعر)."""
        cv = getattr(self, "chart_canvas", None)
        if cv is None:
            return
        t = self.theme
        CARD = t["card"]
        GOLD, MUT = t["accent"], t["muted"]
        sym = CUR[self.cur]["sym"]
        try:
            cv.delete("all")
            W = cv.winfo_width() or 900
            H = cv.winfo_height() or 430
            if W < 260:
                W = 900
            if H < 180:
                H = 430
            pts = self._chart_points()
            if len(pts) < 2:
                cv.create_text(W // 2, H // 2,
                               text=("لا توجد بيانات كافية في نطاق «{}».\n"
                                     "اترك البرنامج مفتوحاً ليجمّع الأسعار تلقائياً.").format(
                                         self._range_label()),
                               fill=MUT, font=("Segoe UI", 12), justify="center")
                pf = getattr(self, "chart_info", None)
                if pf is not None:
                    pf.config(text="عدد العينات في هذا النطاق: {}".format(len(pts)), fg=MUT)
                return
            pl, pr, pt, pb = 82, 26, 26, 46
            xs = [p[0] for p in pts]
            ys = [p[1] for p in pts]
            xmin, xmax = min(xs), max(xs)
            ymin, ymax = min(ys), max(ys)
            if ymax - ymin < 1e-9:
                ymin -= 1
                ymax += 1
            pad = (ymax - ymin) * 0.1
            ymin -= pad
            ymax += pad

            def X(x):
                return (pl + (x - xmin) / (xmax - xmin) * (W - pl - pr)
                        if xmax > xmin else pl)

            def Y(y):
                return pt + (ymax - y) / (ymax - ymin) * (H - pt - pb)

            for i in range(5):
                yy = ymin + (ymax - ymin) * i / 4
                y = Y(yy)
                cv.create_line(pl, y, W - pr, y, fill=t.get("border", "#252c39"))
                cv.create_text(pl - 8, y, text="{:,.1f}".format(yy), fill=MUT,
                               anchor="e", font=("Segoe UI", 10))
            # علامات زمنية على المحور الأفقي
            fracs = [0.0, 0.5, 1.0] if self.chart_range == "hour" \
                else [0.0, 0.25, 0.5, 0.75, 1.0]
            for frac in fracs:
                if frac == 0.0:
                    anchor = "w"
                elif frac == 1.0:
                    anchor = "e"
                else:
                    anchor = "center"
                ts = xmin + (xmax - xmin) * frac
                x = pl + (W - pl - pr) * frac
                cv.create_text(x, H - pb + 16, text=self._fmt_tick(ts), fill=MUT,
                               anchor=anchor, font=("Segoe UI", 11),
                               justify="center")
            cv.create_text(W - pr, pt - 10, text="({})".format(sym), fill=MUT,
                           anchor="e", font=("Segoe UI", 10))

            coords = []
            for x, y in pts:
                coords += [X(x), Y(y)]
            up = ys[-1] >= ys[0]
            line_col = t.get("profit", "#2ecc71") if up else t.get("loss", "#ff5c5c")
            cv.create_line(*coords, fill=line_col, width=2, smooth=True)
            cv.create_oval(X(xs[-1]) - 4, Y(ys[-1]) - 4,
                           X(xs[-1]) + 4, Y(ys[-1]) + 4,
                           fill=line_col, outline="")

            change = ys[-1] - ys[0]
            pct = (change / ys[0] * 100) if ys[0] else 0
            info = ("{rg}  •  عيار {k}  •  الافتتاح: {s}{o:,.2f}  •  الآن: {s}{c:,.2f}  •  "
                    "التغيّر: {s}{sg}{d:,.2f} ({sg}{p:.2f}%)  •  آخر تحديث: {u}").format(
                rg=self._range_label(), k=self.karat, s=sym, o=ys[0], c=ys[-1],
                d=change, p=pct, sg="+" if change > 0 else "",
                u=self.last_update or "--:--:--")
            lbl = getattr(self, "chart_info", None)
            if lbl is not None:
                lbl.config(text=info, fg=line_col)
        except Exception:
            pass

    def _refresh_chart_if_visible(self):
        if getattr(self, "_page", "main") == "chart":
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
                self.lbl_upd.config(text="🔔 تحقق تنبيه: {} {} {:.2f}".format(
                    unit, a["op"], a["price"]), fg=self.theme["accent"])
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
        if getattr(self, "_live_on", False):
            self.refresh_all()
            return
        self._live_on = True
        self.refresh_all()
        self.root.after(LIVE_MS, self._loop_refresh)
        self._start_flash_loop()

    def _loop_refresh(self):
        if not getattr(self, "_live_on", False) or not getattr(self, "user", None):
            return
        self.refresh_all()
        self.root.after(self.refresh_interval(), self._loop_refresh)

    def refresh_interval(self):
        """٥ ثوانٍ في الوضع الطبيعي، وتباطؤ تلقائي عند تكرار فشل الاتصال."""
        f = int(getattr(self, "_fails", 0) or 0)
        if f >= 6:
            return 60000
        if f >= 4:
            return 30000
        if f >= 2:
            return 15000
        return LIVE_MS

    def refresh_all(self):
        # لا نبدأ طلباً جديداً قبل انتهاء السابق (يمنع تراكم الطلبات وبطء الواجهة)
        lock = getattr(self, "_fetch_lock", None) or threading.Lock()
        self._fetch_lock = lock
        with lock:
            if getattr(self, "_fetch_busy", False):
                return
            self._fetch_busy = True
        threading.Thread(target=self._fetch_once, daemon=True).start()

    def _fetch_once(self):
        msg = ""
        direction = None
        changed = False
        try:
            new_price, source = fetch_spot_usd(self.source_mode)
            old_price = self.spot_oz_usd
            changed = (not old_price) or (abs(float(new_price) - float(old_price)) > 1e-9)
            self._delta_oz = (new_price - old_price) if old_price else 0.0
            direction = ("up" if new_price > old_price else
                         ("down" if new_price < old_price else "same")) if old_price else None
            self.spot_oz_usd = new_price
            self._record_price_sample(new_price)
            self._ticks += 1
            self._fails = 0
            now = datetime.now().strftime("%H:%M:%S")
            self.last_update = now
            if self._ticks == 1 or self._ticks % 12 == 0:
                self.save()
            closed = "  •  🔒 السوق مغلق — السعر ثابت حتى الفتح" if market_closed() else ""
            msg = "● مباشر ({src}) • الأونصة ${p:.3f} • {t} • تحديثات {n}{c}".format(
                src=source, p=new_price, t=now, n=self._ticks, c=closed)
        except Exception as e:
            direction = None
            self._fails += 1
            reason = ""
            try:
                reason = str(e).split(":", 1)[-1].strip()[:45]
            except Exception:
                reason = ""
            tag = " ({})".format(reason) if reason and self._fails <= 3 else ""
            if self.spot_oz_usd:
                when = " — آخر سعر {}".format(self.last_update) if self.last_update else ""
                msg = "تعذّر التحديث (محاولة {}){}… عرض آخر سعر معروف{}".format(
                    self._fails, tag, when)
            else:
                msg = ("تعذّر جلب سعر الذهب{} — اضغط «🔍 تشخيص» أو «F5»"
                       .format(tag))

        nowt = time.monotonic()
        fx_changed = False
        if nowt - self._fx_last > 300:
            try:
                fx = http_json("https://open.er-api.com/v6/latest/USD", timeout=6)
                if fx.get("rates"):
                    for k in ("SAR", "JOD", "KWD", "QAR", "EGP"):
                        if k in fx["rates"] and abs(
                                float(fx["rates"][k]) - float(self.rates.get(k, 0) or 0)) > 1e-9:
                            self.rates[k] = fx["rates"][k]
                            fx_changed = True
                self._fx_last = nowt
            except Exception:
                pass

        def paint():
            try:
                self.lbl_upd.config(text=msg)
            except Exception:
                pass
            # إعادة الرسم فقط عند تغيّر السعر/سعر الصرف فعلاً (تمنع رجفة الجدول)
            if changed or fx_changed:
                try:
                    self.render()
                except Exception:
                    pass
            # الوميض له إيقاعه الثابت (كل 7 ثوانٍ) وليس مع كل جلب

        try:
            self.root.after(0, paint)
        except Exception:
            pass
        finally:
            lock = getattr(self, "_fetch_lock", None)
            if lock is not None:
                with lock:
                    self._fetch_busy = False


    # ---------- وميض تغيّر السعر: إيقاع ثابت كل 7 ثوانٍ ----------
    def _start_flash_loop(self):
        """يضبط مرجع الوميض وينطلق بمؤقّت ثابت (7 ثوانٍ)."""
        job = getattr(self, "_flash_job", None)
        if job:
            try:
                self.root.after_cancel(job)
            except Exception:
                pass
        self._flash_job = None
        self._flash_price = float(getattr(self, "spot_oz_usd", 0) or 0)
        try:
            self._flash_job = self.root.after(FLASH_MS, self._flash_tick)
        except Exception:
            self._flash_job = None

    def _flash_tick(self):
        """كل 7 ثوانٍ: يومض فقط إذا تغيّر السعر فعلاً منذ الوميض السابق."""
        self._flash_job = None
        if not getattr(self, "_live_on", False) or not getattr(self, "user", None):
            return
        cur = float(getattr(self, "spot_oz_usd", 0) or 0)
        prev = float(getattr(self, "_flash_price", 0) or 0)
        if cur:
            self._flash_price = cur
        if cur and prev and abs(cur - prev) > 1e-9 and not getattr(self, "locked", False):
            self.flash_update("up" if cur > prev else "down")
        try:
            self._flash_job = self.root.after(FLASH_MS, self._flash_tick)
        except Exception:
            self._flash_job = None

    def flash_update(self, direction=None):
        t = self.theme
        hi = {"up": t.get("profit", "#2ecc71"),
              "down": t.get("loss", "#ff5c5c"),
              "same": t.get("flash_same", "#ffd93b")}.get(direction, t.get("muted", "#888"))
        rest = t.get("muted", "#888")
        try:
            self.dot.config(fg=hi)
        except Exception:
            pass
        originals = []
        for name in ("lbl_gr",):
            lbl = getattr(self, name, None)
            if lbl is None:
                continue
            try:
                originals.append((lbl, lbl.cget("bg")))
                lbl.config(bg=hi)
            except Exception:
                pass

        def revert():
            for lbl, bg in originals:
                try:
                    lbl.config(bg=bg)
                except Exception:
                    pass
            try:
                self.dot.config(fg=rest)
            except Exception:
                pass
        self.root.after(FLASH_HOLD, revert)

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
        txt.insert("end", "جارٍ فحص المصادر…\n")

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
            names = ("gold-api", "goldprice.org", "yahoo", "paxg", "stooq")
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
            add("النتيجة: البرنامج يستخدم أول مصدر ينجح، ويعود للبقية تلقائياً."
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

    def render(self):
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
                "✔ بيع ({})".format(s_karat),
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
    GoldApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
