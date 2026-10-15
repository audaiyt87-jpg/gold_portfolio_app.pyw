# -*- coding: utf-8 -*-
"""
محفظة الذهب — تطبيق ويندوز (Gold Portfolio for Windows)
- يجلب سعر الذهب الحالي مباشرة من الإنترنت
- يحسب الربح/الخسارة بمقارنة سعر الشراء بسعر السوق
- تحويل العملة تلقائياً (USD / SAR / AED)
- يطرح مقداراً ثابتاً (افتراضي 5.22) من سعر الغرام الحالي
يعتمد فقط على المكتبة القياسية لبايثون (لا توجد حزم خارجية).
"""
import json
import os
import threading
import time
import urllib.request
import xml.sax.saxutils as _xml
import zipfile
from datetime import datetime

import tkinter as tk
from tkinter import ttk, messagebox

OZ_TO_G = 31.1034768
# المصدر الأساسي (نفس مصدر livepriceofgold - تحديث لحظي)
GOLDPRICE_URL = "https://data-asg.goldprice.org/dbXRates/USD"
# مصدر فوري موثوق (سعر فوري بدقة عالية، بدون مفتاح)
SWISS_URL = "https://forex-data-feed.swissquote.com/public-quotes/bboquotes/instrument/XAU/USD"
# احتياطي 1
YAHOO_URL = "https://query1.finance.yahoo.com/v8/finance/chart/GC=F?interval=1m&range=1d"
# احتياطي 2
GOLD_URL = "https://api.gold-api.com/price/XAU"
FX_URL = "https://open.er-api.com/v6/latest/USD"
DATA_FILE = os.path.join(os.path.expanduser("~"), "gold_portfolio_data.json")

CUR = {
    "USD": {"sym": "$", "name": "USD", "label": "دولار أمريكي (USD)"},
    "SAR": {"sym": "ر.س ", "name": "SAR", "label": "ريال سعودي (SAR)"},
    "AED": {"sym": "د.إ ", "name": "AED", "label": "درهم إماراتي (AED)"},
}
DEFAULT_RATES = {"USD": 1.0, "SAR": 3.75, "AED": 3.6725}

# نسخة سمة الألوان — عند زيادتها تُطبَّق السمة الحديثة الجديدة
# حتى لو كان لدى المستخدم سمة قديمة محفوظة.
THEME_VER = 3

DEFAULT_THEME = {
    "bg": "#0b0e14",        # الخلفية العامة (أزرق-رمادي داكن حديث)
    "card": "#151a23",      # البطاقات
    "card2": "#1e2530",     # خلفية ثانوية / مرتفعة
    "accent": "#e3b23c",    # اللون المميز (ذهبي راقٍ)
    "accent_text": "#141000",
    "text": "#eef1f6",      # النص الأساسي
    "muted": "#8b94a7",     # نص ثانوي
    "profit": "#2ecc71",    # لون الربح
    "loss": "#ff5c5c",      # لون الخسارة
    "total": "#ffffff",     # لون سطر الإجمالي
    "total_bg": "#232b38",
    "border": "#2a3240",    # حدود خفيفة
    "entry_bg": "#1b2029",  # خلفية حقول الإدخال
    "row_alt": "#171c26",   # تظليل الصفوف المتناوبة
    "hover": "#f0c954",     # لون التمرير على الزر الذهبي
    "flash_same": "#ffd93b", # وميض أصفر عند ثبات السعر
    "sold_bg": "#16351f",   # خلفية سطر عملية البيع (أخضر داكن)
}
DEFAULT_TITLES = {
    "title": "✦ محفظة الذهب ✦",
    "subtitle": "تتبّع الربح والخسارة مع جلب السعر مباشرة من الإنترنت",
    "window": "محفظة الذهب - Gold Portfolio",
}


BROWSER_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/122.0 Safari/537.36"
)


def http_json(url, timeout=10, headers=None):
    hdrs = {"User-Agent": BROWSER_UA, "Accept": "application/json, text/plain, */*",
            "Cache-Control": "no-cache", "Pragma": "no-cache"}
    if headers:
        hdrs.update(headers)
    req = urllib.request.Request(url, headers=hdrs)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))


def _bust(url):
    """يضيف طابعاً زمنياً لمنع التخزين المؤقت (ضمان سعر طازج كل مرة)."""
    sep = "&" if "?" in url else "?"
    return "{}{}_t={}".format(url, sep, int(time.time() * 1000))


def market_closed():
    """يرجع True إذا كان سوق الذهب الفوري مغلقاً (عطلة نهاية الأسبوع).
    السوق يفتح الأحد 22:00 UTC ويغلق الجمعة 22:00 UTC تقريباً.
    خلال الإغلاق تبقى الأسعار ثابتة عند سعر إغلاق الجمعة (وليس خللاً).
    """
    now = datetime.utcnow()
    wd = now.weekday()   # الاثنين=0 ... الأحد=6
    h = now.hour
    if wd == 5:                 # السبت كاملاً
        return True
    if wd == 6 and h < 22:      # الأحد قبل 22:00 UTC
        return True
    if wd == 4 and h >= 22:     # الجمعة بعد 22:00 UTC
        return True
    return False


def _src_goldprice():
    d = http_json(_bust(GOLDPRICE_URL), timeout=7,
                  headers={"Referer": "https://goldprice.org/"})
    items = d.get("items") or []
    if items and items[0].get("xauPrice"):
        return float(items[0]["xauPrice"])
    raise ValueError("no data")


def _src_swissquote():
    d = http_json(_bust(SWISS_URL), timeout=7)
    for block in d:
        for sp in block.get("spreadProfilePrices", []):
            bid, ask = sp.get("bid"), sp.get("ask")
            if bid and ask:
                return (float(bid) + float(ask)) / 2.0
    raise ValueError("no data")


def _src_yahoo():
    d = http_json(_bust(YAHOO_URL), timeout=7)
    meta = d["chart"]["result"][0]["meta"]
    p = meta.get("regularMarketPrice")
    if p:
        return float(p)
    raise ValueError("no data")


def _src_goldapi():
    d = http_json(_bust(GOLD_URL), timeout=7)
    return float(d["price"])


# المصادر المتاحة: المفتاح -> (الاسم العربي للعرض، الدالة)
SOURCES = {
    "goldprice":  ("goldprice — مطابق livepriceofgold", _src_goldprice),
    "swissquote": ("swissquote — فوري دقيق", _src_swissquote),
    "yahoo":      ("yahoo — عقد آجل (أعلى قليلاً)", _src_yahoo),
    "gold-api":   ("gold-api — تقريبي (احتياطي)", _src_goldapi),
}
AUTO_ORDER = ["goldprice", "swissquote", "yahoo", "gold-api"]


def fetch_spot_usd(mode="auto"):
    """يرجع (سعر الأونصة بالدولار، اسم المصدر).
    mode="auto" يجرّب المصادر بالترتيب؛ أو اسم مصدر محدد لإجباره.
    عند فشل المصدر المحدد يسقط تلقائياً لباقي المصادر لمنع تجمّد الشاشة.
    """
    if mode and mode != "auto" and mode in SOURCES:
        try:
            return SOURCES[mode][1](), mode
        except Exception:
            pass  # سقوط تلقائي لباقي المصادر
    for key in AUTO_ORDER:
        try:
            return SOURCES[key][1](), key
        except Exception:
            continue
    raise RuntimeError("تعذّر جلب السعر من جميع المصادر")


def _col_letter(idx):
    """رقم عمود (1-مبدوء) -> حرف Excel (A, B, ... AA)."""
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def write_xlsx(path, headers, rows, sheet_name="المحفظة"):
    """يكتب ملف Excel (.xlsx) صحيحاً باستخدام المكتبة القياسية فقط.
    headers: قائمة نصوص. rows: قائمة صفوف؛ كل خلية نص أو رقم.
    """
    def esc(v):
        return _xml.escape(str(v))

    def cell_xml(ref, value):
        if isinstance(value, bool):
            value = str(value)
        if isinstance(value, (int, float)):
            return '<c r="{r}"><v>{v}</v></c>'.format(r=ref, v=value)
        return ('<c r="{r}" t="inlineStr"><is><t xml:space="preserve">'
                '{v}</t></is></c>').format(r=ref, v=esc(value))

    all_rows = [headers] + list(rows)
    rows_xml = []
    for ri, rowvals in enumerate(all_rows, start=1):
        cells = "".join(
            cell_xml(_col_letter(ci) + str(ri), val)
            for ci, val in enumerate(rowvals, start=1))
        rows_xml.append('<row r="{r}">{c}</row>'.format(r=ri, c=cells))
    sheet = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetData>{}</sheetData></worksheet>'.format("".join(rows_xml)))

    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '</Types>')
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>')
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="{n}" sheetId="1" r:id="rId1"/></sheets>'
        '</workbook>'.format(n=esc(sheet_name)))
    wb_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '</Relationships>')

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/worksheets/sheet1.xml", sheet)


class GoldApp:
    def __init__(self, root):
        self.root = root
        self.spot_oz_usd = 0.0            # سعر الأونصة بالدولار (الأساس)
        self.rates = dict(DEFAULT_RATES)
        self.cur = "USD"
        self.adjust = 5.22               # فرق الراجحي للبيع (يُطرح من سعر الغرام العالمي)
        self.buy_margin = 0.0            # فرق الراجحي للشراء (يُضاف لسعر الغرام العالمي)
        self.holdings = []               # أسعار الشراء تُخزّن داخلياً بالدولار
        self.sells = []                  # عمليات البيع المحققة (سعر بيع مثبّت)
        self.history = []                # سجل كل العمليات (شراء/بيع)
        self.price_hist = []             # عينات سعر الأونصة [epoch, oz_usd] خلال اليوم
        self.alerts = []                 # تنبيهات الأسعار
        self.theme = dict(DEFAULT_THEME)
        self.titles = dict(DEFAULT_TITLES)
        self.sound_on = True             # تفعيل التنبيه الصوتي
        self.source_mode = "auto"        # مصدر السعر: auto أو اسم مصدر محدد
        self.last_update = ""            # آخر وقت تحديث ناجح (لعرضه عند الفتح)
        self.load()

        self.build_ui()
        # عرض آخر سعر محفوظ فوراً عند الفتح (قبل وصول أول تحديث مباشر)
        if self.spot_oz_usd:
            self.render()
            when = "  •  آخر تحديث {}".format(self.last_update) if self.last_update else ""
            self.lbl_upd.config(
                text="○ آخر سعر محفوظ • الأونصة ${:.3f}{} • جارٍ التحديث…".format(
                    self.spot_oz_usd, when))
        self.start_live_refresh()

    # ---------- التخزين ----------
    def load(self):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                d = json.load(f)
            self.holdings = d.get("holdings", [])
            self.sells = d.get("sells", [])
            # ترحيل بيانات قديمة: أي مركز يحمل سعر بيع → تحويله إلى عملية بيع
            migrated = []
            for h in list(self.holdings):
                if h.get("sell") is not None:
                    factor = OZ_TO_G if h.get("unit") == "oz" else 1
                    qty_g = h.get("qty", 0) * factor
                    self.sells.append({
                        "name": h.get("name", "بيع ذهب"),
                        "qty_g": qty_g,
                        "sell": h["sell"] / factor,
                        "cost": h.get("qty", 0) * h.get("buy", 0),
                        "date": h.get("sell_date", ""),
                    })
                else:
                    migrated.append(h)
            self.holdings = migrated
            self.history = d.get("history", [])
            self.price_hist = d.get("price_hist", [])
            self.cur = d.get("cur", "USD")
            self.adjust = float(d.get("adjust", 5.22))
            self.buy_margin = float(d.get("buy_margin", 0.0))
            self.alerts = d.get("alerts", [])
            # طبّق السمة المحفوظة فقط إذا طابقت النسخة الحالية،
            # وإلا أبقِ السمة الحديثة الافتراضية (مع الحفاظ على باقي البيانات).
            if d.get("theme_ver") == THEME_VER:
                self.theme.update(d.get("theme", {}))
            self.titles.update(d.get("titles", {}))
            self.sound_on = bool(d.get("sound_on", True))
            self.source_mode = d.get("source_mode", "auto")
            # استعادة آخر سعر معروف + آخر أسعار صرف لعرضها فور فتح البرنامج
            self.spot_oz_usd = float(d.get("spot_oz_usd", 0.0) or 0.0)
            self.rates.update(d.get("rates", {}))
            self.last_update = d.get("last_update", "")
        except Exception:
            pass

    def save(self):
        try:
            with open(DATA_FILE, "w", encoding="utf-8") as f:
                json.dump({"holdings": self.holdings, "cur": self.cur,
                           "sells": self.sells,
                           "history": self.history,
                           "price_hist": self.price_hist[-500:],
                           "adjust": self.adjust, "buy_margin": self.buy_margin,
                           "alerts": self.alerts,
                           "theme": self.theme, "theme_ver": THEME_VER,
                           "titles": self.titles,
                           "sound_on": self.sound_on,
                           "source_mode": self.source_mode,
                           "spot_oz_usd": self.spot_oz_usd,
                           "rates": self.rates,
                           "last_update": getattr(self, "last_update", "")},
                          f, ensure_ascii=False)
        except Exception:
            pass

    # ---------- حسابات ----------
    def rate(self):
        return self.rates.get(self.cur, 1.0)

    def money(self, usd_amount):
        v = usd_amount * self.rate()
        return "{}{:,.2f}".format(CUR[self.cur]["sym"], v)

    def gram_market_disp(self):
        """سعر الغرام السوقي بالعملة المعروضة (بدون طرح)."""
        return (self.spot_oz_usd / OZ_TO_G) * self.rate()

    def gram_eff_disp(self):
        """سعر الراجحي للبيع (العالمي ناقص فرق البيع) للغرام."""
        return self.gram_market_disp() - self.adjust

    def gram_buy_disp(self):
        """سعر الراجحي للشراء (العالمي زائد فرق الشراء) للغرام."""
        return self.gram_market_disp() + self.buy_margin

    def current_per_unit_disp(self, unit):
        g = self.gram_eff_disp()
        return g * OZ_TO_G if unit == "oz" else g

    # ---------- الواجهة ----------
    # ---------- عناصر واجهة حديثة (مساعدات) ----------
    def _mkbtn(self, parent, text, cmd, kind="primary", font_size=10, **kw):
        """زر مسطّح حديث مع تأثير تمرير (hover)."""
        t = self.theme
        palette = {
            "primary": (t["accent"], t["accent_text"], t.get("hover", "#f0c954")),
            "ghost":   (t["card2"], t["accent"], t.get("border", "#2a3240")),
            "danger":  (t["loss"], "#ffffff", "#ff7a7a"),
            "neutral": (t["card2"], t["text"], t.get("border", "#2e3744")),
            "success": (t.get("profit", "#2ecc71"), "#ffffff", "#3fe088"),
        }
        bg, fg, hover = palette.get(kind, palette["primary"])
        b = tk.Button(parent, text=text, command=cmd, bg=bg, fg=fg,
                      activebackground=hover, activeforeground=fg,
                      relief="flat", bd=0, cursor="hand2",
                      font=("Segoe UI", font_size, "bold"),
                      padx=14, pady=7, **kw)
        b.bind("<Enter>", lambda e: b.config(bg=hover))
        b.bind("<Leave>", lambda e: b.config(bg=bg))
        return b

    def _mkentry(self, parent, textvariable=None, width=16):
        """حقل إدخال مسطّح بحدود خفيفة."""
        t = self.theme
        return tk.Entry(parent, textvariable=textvariable, width=width,
                        bg=t.get("entry_bg", "#1b2029"), fg=t["text"],
                        insertbackground=t["accent"], relief="flat", bd=0,
                        highlightthickness=1,
                        highlightbackground=t.get("border", "#2a3240"),
                        highlightcolor=t["accent"],
                        font=("Segoe UI", 10))

    def _price_card(self, parent, title, value_attr, color, col):
        """بطاقة سعر حديثة (عنوان خافت + قيمة بارزة)."""
        t = self.theme
        card = tk.Frame(parent, bg=t["card"], highlightthickness=1,
                        highlightbackground=t.get("border", "#2a3240"))
        card.grid(row=0, column=col, sticky="nsew", padx=5)
        parent.grid_columnconfigure(col, weight=1, uniform="price")
        tk.Label(card, text=title, bg=t["card"], fg=t["muted"],
                 font=("Segoe UI", 9), anchor="e").pack(fill="x", padx=12, pady=(10, 0))
        lbl = tk.Label(card, text="--", bg=t["card"], fg=color,
                       font=("Segoe UI", 15, "bold"), anchor="e")
        lbl.pack(fill="x", padx=12, pady=(2, 12))
        setattr(self, value_attr, lbl)
        return lbl

    def build_ui(self):
        t = self.theme
        BG, CARD, CARD2 = t["bg"], t["card"], t["card2"]
        GOLD, ATXT, TXT, MUT = t["accent"], t["accent_text"], t["text"], t["muted"]
        self.root.title(self.titles.get("window", "Gold Portfolio"))
        self.root.geometry("960x790")
        self.root.minsize(900, 680)
        self.root.configure(bg=BG)
        BORDER = t.get("border", "#2a3240")

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Treeview", background=CARD, fieldbackground=CARD,
                        foreground=TXT, rowheight=30, borderwidth=0, relief="flat")
        style.configure("Treeview.Heading", background=CARD2, foreground=GOLD,
                        font=("Segoe UI", 9, "bold"), relief="flat",
                        padding=6, borderwidth=0)
        style.map("Treeview.Heading", background=[("active", CARD2)])
        style.map("Treeview", background=[("selected", GOLD)],
                  foreground=[("selected", ATXT)])
        style.configure("TMenubutton", background=CARD2, foreground=TXT,
                        relief="flat", font=("Segoe UI", 10), padding=6,
                        arrowcolor=GOLD)
        style.map("TMenubutton", background=[("active", BORDER)])

        # شريط الرأس + زر الإعدادات
        head = tk.Frame(self.root, bg=BG)
        head.pack(fill="x", padx=20, pady=(14, 0))
        self._mkbtn(head, "⚙  الإعدادات والألوان", self.open_settings,
                    kind="ghost", font_size=9).pack(side="right")
        tk.Label(self.root, text=self.titles.get("title", ""), bg=BG, fg=GOLD,
                 font=("Segoe UI", 22, "bold")).pack(pady=(6, 0))
        tk.Label(self.root, text=self.titles.get("subtitle", ""),
                 bg=BG, fg=MUT, font=("Segoe UI", 9)).pack(pady=(2, 0))
        tk.Frame(self.root, bg=GOLD, height=2).pack(fill="x", padx=20, pady=(10, 0))

        # بطاقات الأسعار (شبكة أفقية حديثة)
        prices = tk.Frame(self.root, bg=BG)
        prices.pack(fill="x", padx=15, pady=(14, 6))
        self._price_card(prices, "الأونصة (عالمي)", "lbl_oz", GOLD, 0)
        self._price_card(prices, "الغرام (عالمي)", "lbl_gr", GOLD, 1)
        self._price_card(prices, "الراجحي — شراء/غرام", "lbl_gr_buy", t["loss"], 2)
        self._price_card(prices, "الراجحي — بيع/غرام", "lbl_gr_adj", t["profit"], 3)

        # شريط التحكم
        ctrl = tk.Frame(self.root, bg=BG)
        ctrl.pack(fill="x", padx=20, pady=(10, 4))
        self._mkbtn(ctrl, "⟳  تحديث السعر", self.refresh_all,
                    kind="primary", font_size=9).pack(side="right", padx=4)
        tk.Label(ctrl, text="العملة:", bg=BG, fg=TXT,
                 font=("Segoe UI", 10)).pack(side="right", padx=(10, 2))
        self.cur_var = tk.StringVar(value=self.cur)
        ttk.OptionMenu(ctrl, self.cur_var, self.cur, *CUR.keys(),
                       command=self.on_cur_change).pack(side="right")
        tk.Label(ctrl, text="فرق البيع/غرام:", bg=BG, fg=TXT,
                 font=("Segoe UI", 10)).pack(side="right", padx=(10, 2))
        self.adj_var = tk.StringVar(value=str(self.adjust))
        adj_e = self._mkentry(ctrl, self.adj_var, width=6)
        adj_e.pack(side="right", ipady=4)
        adj_e.bind("<Return>", lambda e: self.on_adjust_change())
        self.adj_var.trace_add("write", lambda *a: self.on_adjust_change())
        tk.Label(ctrl, text="فرق الشراء/غرام:", bg=BG, fg=TXT,
                 font=("Segoe UI", 10)).pack(side="right", padx=(10, 2))
        self.buy_var = tk.StringVar(value=str(self.buy_margin))
        buy_e = self._mkentry(ctrl, self.buy_var, width=6)
        buy_e.pack(side="right", ipady=4)
        buy_e.bind("<Return>", lambda e: self.on_buy_margin_change())
        self.buy_var.trace_add("write", lambda *a: self.on_buy_margin_change())

        updfr = tk.Frame(self.root, bg=BG)
        updfr.pack()
        self.dot = tk.Label(updfr, text="●", bg=BG, fg=MUT,
                            font=("Segoe UI", 10))
        self.dot.pack(side="right", padx=(6, 0))
        self.lbl_upd = tk.Label(updfr, text="جارٍ التحميل...", bg=BG, fg=MUT,
                                font=("Segoe UI", 8))
        self.lbl_upd.pack(side="right")
        # مؤشر تغيّر السعر (سهم + مقدار التغيّر منذ آخر تحديث)
        self.lbl_delta = tk.Label(updfr, text="", bg=BG, fg=MUT,
                                  font=("Segoe UI", 9, "bold"))
        self.lbl_delta.pack(side="right", padx=(0, 10))

        # نموذج الإضافة
        form = tk.Frame(self.root, bg=CARD, highlightthickness=1,
                        highlightbackground=BORDER)
        form.pack(fill="x", padx=20, pady=10)
        tk.Label(form, text="الاسم", bg=CARD, fg=MUT,
                 font=("Segoe UI", 9)).grid(row=0, column=0, padx=8, pady=(8, 2))
        tk.Label(form, text="الكمية", bg=CARD, fg=MUT,
                 font=("Segoe UI", 9)).grid(row=0, column=1, padx=8)
        tk.Label(form, text="الوحدة", bg=CARD, fg=MUT,
                 font=("Segoe UI", 9)).grid(row=0, column=2, padx=8)
        self.buy_lbl = tk.Label(form, text="سعر الشراء/وحدة (USD)", bg=CARD, fg=MUT,
                                font=("Segoe UI", 9))
        self.buy_lbl.grid(row=0, column=3, padx=8)
        self.e_name = self._mkentry(form, width=18)
        self.e_name.grid(row=1, column=0, padx=8, pady=(2, 10), ipady=4)
        self.e_qty = self._mkentry(form, width=10)
        self.e_qty.grid(row=1, column=1, padx=8, ipady=4)
        self.unit_var = tk.StringVar(value="g")
        ttk.OptionMenu(form, self.unit_var, "g", "g", "oz").grid(row=1, column=2, padx=8)
        self.e_buy = self._mkentry(form, width=12)
        self.e_buy.grid(row=1, column=3, padx=8, ipady=4)
        self._mkbtn(form, "＋ إضافة", self.add_item, kind="primary",
                    width=10).grid(row=1, column=4, padx=12, pady=(2, 10))

        # الجدول
        cols = ("name", "status", "qty", "buy", "cost", "cur", "val", "pl", "pct")
        heads = ("الاسم", "الحالة", "الكمية", "سعر الشراء", "التكلفة",
                 "السعر الحالي/البيع", "القيمة", "الربح/الخسارة", "%")
        self.tree = ttk.Treeview(self.root, columns=cols, show="headings", height=6)
        for c, h in zip(cols, heads):
            self.tree.heading(c, text=h)
            self.tree.column(c, anchor="center", width=100)
        self.tree.column("name", width=120)
        self.tree.column("status", width=74)
        self.tree.column("pct", width=70)
        self.tree.tag_configure("pos", foreground=t["profit"])
        self.tree.tag_configure("neg", foreground=t["loss"])
        self.tree.tag_configure("odd", background=t.get("row_alt", CARD))
        self.tree.tag_configure("even", background=CARD)
        self.tree.tag_configure("total", background=t["total_bg"], foreground=t["total"],
                                font=("Segoe UI", 9, "bold"))
        self.tree.tag_configure("totalpos", background=t["total_bg"], foreground=t["profit"],
                                font=("Segoe UI", 9, "bold"))
        self.tree.tag_configure("totalneg", background=t["total_bg"], foreground=t["loss"],
                                font=("Segoe UI", 9, "bold"))
        self.tree.tag_configure("sold", background=t.get("card2", "#1e2530"))
        self.tree.tag_configure("soldgreen", background=t.get("sold_bg", "#16351f"),
                                foreground=t.get("profit", "#2ecc71"),
                                font=("Segoe UI", 9, "bold"))

        btns = tk.Frame(self.root, bg=BG)
        self._mkbtn(btns, "🗑 حذف المحدد", self.del_item, kind="danger",
                    font_size=9).pack(side="left", pady=4)
        self._mkbtn(btns, "مسح الكل", self.clear_all, kind="neutral",
                    font_size=9).pack(side="left", padx=6)
        self._mkbtn(btns, "💰 أمر بيع", self.open_sell, kind="neutral",
                    font_size=9).pack(side="left", padx=6)
        self._mkbtn(btns, "📜 سجل العمليات", self.open_history, kind="neutral",
                    font_size=9).pack(side="left", padx=6)
        self._mkbtn(btns, "📐 حاسبة الزكاة", self.open_zakat, kind="neutral",
                    font_size=9).pack(side="left", padx=6)
        self._mkbtn(btns, "📈 مخطط اليوم", self.open_chart, kind="neutral",
                    font_size=9).pack(side="left", padx=6)
        self._mkbtn(btns, "🔔 إضافة تنبيه السعر", self.open_alerts, kind="primary",
                    font_size=9).pack(side="left", padx=6)
        self._mkbtn(btns, "⬇ تصدير Excel", self.export_excel, kind="success",
                    font_size=9).pack(side="left", padx=6)

        # مراجع عناصر نافذة التنبيهات (تُنشأ عند فتحها فقط)
        self.al_win = None
        self.al_list = None
        self.al_price = None
        self.al_cur_lbl = None
        self.al_metric = None
        self.al_op = None

        self.lbl_sum = tk.Label(self.root, text="", bg=t["total_bg"], fg=t["total"],
                                font=("Segoe UI", 11, "bold"), pady=10)
        # ترتيب العرض: الأزرار والإجمالي بالأسفل والجدول يتمدد لأكبر مساحة
        self.lbl_sum.pack(side="bottom", fill="x", padx=20, pady=(0, 10))
        btns.pack(side="bottom", fill="x", padx=20, pady=(0, 4))
        self.tree.pack(side="top", fill="both", expand=True, padx=20, pady=6)
        self.update_buy_label()

    def update_buy_label(self):
        self.buy_lbl.config(text="سعر الشراء/وحدة ({})".format(CUR[self.cur]["name"]))
        lbl = getattr(self, "al_cur_lbl", None)
        if lbl is not None:
            try:
                lbl.config(text="(" + CUR[self.cur]["name"] + ")")
            except Exception:
                pass

    def eff_price_metric(self, metric):
        """السعر الفعلي (بعد الطرح) للوحدة بالعملة المعروضة الحالية."""
        g = self.gram_eff_disp()
        return g * OZ_TO_G if metric == "oz" else g

    # ---------- الأحداث ----------
    def on_cur_change(self, *a):
        self.cur = self.cur_var.get()
        self.update_buy_label()
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

    def add_item(self):
        name = self.e_name.get().strip()
        try:
            qty = float(self.e_qty.get())
            buy_disp = float(self.e_buy.get())
        except ValueError:
            messagebox.showwarning("تنبيه", "الرجاء إدخال كمية وسعر شراء صحيحين")
            return
        if qty <= 0 or buy_disp <= 0:
            messagebox.showwarning("تنبيه", "الكمية والسعر يجب أن يكونا أكبر من صفر")
            return
        buy_usd = buy_disp / self.rate()   # تخزين داخلي بالدولار
        unit = self.unit_var.get()
        self.holdings.append({"name": name or "بدون اسم", "qty": qty,
                              "unit": unit, "buy": buy_usd})
        factor = OZ_TO_G if unit == "oz" else 1
        self.history.append({
            "type": "buy",
            "name": name or "بدون اسم",
            "qty_g": qty * factor,
            "price_usd": buy_usd / factor,     # دولار/غرام
            "amount_usd": qty * buy_usd,        # إجمالي التكلفة (دولار)
            "date": datetime.now().strftime("%Y-%m-%d %H:%M"),
        })
        self.save()
        self.e_name.delete(0, "end"); self.e_qty.delete(0, "end"); self.e_buy.delete(0, "end")
        self.render()

    def del_item(self):
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        kind_idx = getattr(self, "row_map", {}).get(iid)
        if not kind_idx:
            return
        kind, idx = kind_idx
        if kind == "hold" and 0 <= idx < len(self.holdings):
            del self.holdings[idx]
            self.save()
            self.render()
        elif kind == "sell" and 0 <= idx < len(self.sells):
            if messagebox.askyesno("تراجع عن بيع",
                                   "حذف عملية البيع؟ ستُعاد الكمية المباعة إلى الرصيد."):
                s = self.sells.pop(idx)
                # إعادة الكمية المباعة كمركز مفتوح بنفس تكلفتها
                qty_g = s["qty_g"]
                buy_per_g = (s["cost"] / qty_g) if qty_g else 0.0
                self.holdings.append({"name": s.get("name", "مُسترد"),
                                      "qty": qty_g, "unit": "g", "buy": buy_per_g})
                self.save()
                self.render()

    def clear_all(self):
        if (self.holdings or self.sells) and messagebox.askyesno(
                "تأكيد", "مسح جميع المقتنيات وعمليات البيع؟"):
            self.holdings = []
            self.sells = []
            self.save()
            self.render()

    # ---------- مساعدات الكمية/التكلفة ----------
    def total_open_grams(self):
        """إجمالي الغرامات المتاحة (الرصيد المفتوح)."""
        return sum(h["qty"] * (OZ_TO_G if h["unit"] == "oz" else 1)
                   for h in self.holdings)

    def avg_buy_per_gram_usd(self):
        """متوسط سعر الشراء للغرام (دولار) عبر كل المراكز المفتوحة."""
        g = self.total_open_grams()
        if not g:
            return 0.0
        cost = sum(h["qty"] * h["buy"] for h in self.holdings)   # إجمالي التكلفة بالدولار
        return cost / g

    def _consume_grams(self, grams):
        """خصم (grams) غراماً من المراكز بأسلوب FIFO، وإرجاع تكلفة المخصوم بالدولار."""
        remaining = grams
        cost = 0.0
        i = 0
        while remaining > 1e-9 and i < len(self.holdings):
            h = self.holdings[i]
            factor = OZ_TO_G if h["unit"] == "oz" else 1
            g_avail = h["qty"] * factor
            buy_per_g = h["buy"] / factor
            take = min(remaining, g_avail)
            cost += take * buy_per_g
            new_g = g_avail - take
            if new_g <= 1e-9:
                del self.holdings[i]
            else:
                h["qty"] = new_g / factor
                i += 1
            remaining -= take
        return cost

    # ---------- أمر بيع ----------
    def open_sell(self):
        total_g = self.total_open_grams()
        if total_g <= 1e-9:
            messagebox.showinfo("أمر بيع", "لا يوجد رصيد مفتوح لبيعه")
            return
        t = self.theme
        BG, CARD, CARD2 = t["bg"], t["card"], t["card2"]
        GOLD, ATXT, TXT, MUT = t["accent"], t["accent_text"], t["text"], t["muted"]
        win = tk.Toplevel(self.root)
        win.title("أمر بيع")
        win.configure(bg=BG)
        win.geometry("520x340")
        win.transient(self.root)

        frm = tk.LabelFrame(win, text="💰 بيع كمية من الذهب", bg=CARD, fg=GOLD,
                            font=("Segoe UI", 10, "bold"))
        frm.pack(fill="both", expand=True, padx=12, pady=12)
        sym = CUR[self.cur]["sym"]
        cur_sell_disp = self.gram_eff_disp()      # سعر الراجحي للبيع (للغرام) بالعملة الحالية
        avg_buy_disp = self.avg_buy_per_gram_usd() * self.rate()

        tk.Label(frm, text="الرصيد المتاح: {:,.3f} غرام  •  متوسط الشراء: {}{:,.2f}/غرام".format(
            total_g, sym, avg_buy_disp), bg=CARD, fg=MUT,
            font=("Segoe UI", 9)).pack(anchor="e", padx=10, pady=(6, 2))

        # الكمية بالغرام
        r2 = tk.Frame(frm, bg=CARD); r2.pack(fill="x", padx=6, pady=8)
        tk.Label(r2, text="الكمية المباعة (غرام):", bg=CARD, fg=TXT).pack(side="right")
        e_qty = tk.Entry(r2, width=12)
        e_qty.insert(0, "{:g}".format(round(total_g, 3)))
        e_qty.pack(side="right", padx=6)
        tk.Label(r2, text="(الأقصى {:,.3f})".format(total_g), bg=CARD, fg=MUT,
                 font=("Segoe UI", 8)).pack(side="right", padx=4)

        # سعر البيع للغرام (افتراضي: سعر الراجحي للبيع الحالي)
        r3 = tk.Frame(frm, bg=CARD); r3.pack(fill="x", padx=6, pady=8)
        tk.Label(r3, text="سعر البيع/غرام ({}):".format(CUR[self.cur]["name"]),
                 bg=CARD, fg=TXT).pack(side="right")
        e_price = tk.Entry(r3, width=12)
        if cur_sell_disp:
            e_price.insert(0, "{:.2f}".format(cur_sell_disp))
        e_price.pack(side="right", padx=6)
        tk.Label(r3, text="(سعر الراجحي للبيع الآن)", bg=CARD, fg=MUT,
                 font=("Segoe UI", 8)).pack(side="right", padx=4)

        lbl_prev = tk.Label(frm, text="", bg=CARD, fg=GOLD, font=("Segoe UI", 9, "bold"))
        lbl_prev.pack(anchor="e", padx=10, pady=4)

        def preview(*_):
            try:
                qty = float(e_qty.get()); price_disp = float(e_price.get())
            except ValueError:
                lbl_prev.config(text=""); return
            if qty <= 0 or price_disp <= 0:
                lbl_prev.config(text=""); return
            qty = min(qty, total_g)
            val = qty * price_disp
            cost = qty * avg_buy_disp
            pl = val - cost
            lbl_prev.config(text="قيمة البيع: {s}{v:,.2f}  •  الربح/الخسارة: {sg}{s}{p:,.2f}".format(
                s=sym, v=val, p=pl, sg="+" if pl > 0 else ""))
        e_qty.bind("<KeyRelease>", preview)
        e_price.bind("<KeyRelease>", preview)
        preview()

        def do_sell():
            try:
                qty = float(e_qty.get())
                price_disp = float(e_price.get())
            except ValueError:
                messagebox.showwarning("تنبيه", "الرجاء إدخال كمية وسعر بيع صحيحين", parent=win)
                return
            if qty <= 0 or price_disp <= 0:
                messagebox.showwarning("تنبيه", "الكمية والسعر يجب أن يكونا أكبر من صفر", parent=win)
                return
            if qty > total_g + 1e-6:
                messagebox.showwarning("تنبيه", "الكمية المباعة أكبر من الرصيد المتاح", parent=win)
                return
            qty = min(qty, total_g)
            sell_usd = price_disp / self.rate()        # تثبيت سعر البيع (داخلياً بالدولار)
            cost_usd = self._consume_grams(qty)        # خصم الكمية + تكلفتها FIFO
            sell_date = datetime.now().strftime("%Y-%m-%d %H:%M")
            self.sells.append({
                "name": "بيع ذهب",
                "qty_g": qty,
                "sell": sell_usd,          # دولار/غرام (مثبّت)
                "cost": cost_usd,          # تكلفة المخصوم (دولار)
                "date": sell_date,
            })
            self.history.append({
                "type": "sell",
                "name": "بيع ذهب",
                "qty_g": qty,
                "price_usd": sell_usd,              # دولار/غرام
                "amount_usd": qty * sell_usd,       # قيمة البيع (دولار)
                "cost_usd": cost_usd,               # تكلفة المباع (دولار)
                "pl_usd": qty * sell_usd - cost_usd, # الربح/الخسارة المحقق (دولار)
                "date": sell_date,
            })
            self.save()
            self.render()
            win.destroy()

        bar = tk.Frame(frm, bg=CARD); bar.pack(fill="x", padx=6, pady=14)
        self._mkbtn(bar, "✔ تأكيد البيع", do_sell, kind="success",
                    font_size=9).pack(side="right", padx=6)
        self._mkbtn(bar, "إلغاء", win.destroy, kind="neutral",
                    font_size=9).pack(side="right", padx=6)
        tk.Label(frm, text="يُخصم المباع من الرصيد (FIFO) ويُثبّت سعر البيع لحساب الربح المحقق.",
                 bg=CARD, fg=MUT, font=("Segoe UI", 8)).pack(anchor="e", padx=10)

    # ---------- سجل العمليات ----------
    def open_history(self):
        t = self.theme
        BG, CARD, CARD2 = t["bg"], t["card"], t["card2"]
        GOLD, TXT, MUT = t["accent"], t["text"], t["muted"]
        sym = CUR[self.cur]["sym"]
        win = tk.Toplevel(self.root)
        win.title("سجل العمليات")
        win.configure(bg=BG)
        win.geometry("760x440")
        win.transient(self.root)
        tk.Label(win, text="📜 سجل كل العمليات", bg=BG, fg=GOLD,
                 font=("Segoe UI", 12, "bold")).pack(anchor="e", padx=14, pady=(12, 4))

        cols = ("date", "type", "name", "qty", "price", "amount", "pl")
        heads = ("التاريخ", "النوع", "الاسم", "الكمية (غ)",
                 "السعر/غ", "القيمة", "الربح/الخسارة")
        tv = ttk.Treeview(win, columns=cols, show="headings", height=14)
        for c, h in zip(cols, heads):
            tv.heading(c, text=h)
            tv.column(c, anchor="center", width=100)
        tv.column("date", width=120)
        tv.column("name", width=120)
        tv.tag_configure("buy", foreground=t.get("accent", "#e3b23c"))
        tv.tag_configure("sellpos", foreground=t.get("profit", "#2ecc71"))
        tv.tag_configure("sellneg", foreground=t.get("loss", "#ff5c5c"))
        tv.pack(fill="both", expand=True, padx=12, pady=8)

        tot_buy = tot_sell = tot_pl = 0.0
        for e in self.history:
            amount = e.get("amount_usd", 0) * self.rate()
            price = e.get("price_usd", 0) * self.rate()
            if e.get("type") == "buy":
                tot_buy += amount
                tv.insert("", "end", tags=("buy",), values=(
                    e.get("date", ""), "▲ شراء", e.get("name", ""),
                    "{:g}".format(round(e.get("qty_g", 0), 3)),
                    "{}{:,.2f}".format(sym, price),
                    "{}{:,.2f}".format(sym, amount), "-"))
            else:
                pl = e.get("pl_usd", 0) * self.rate()
                tot_sell += amount
                tot_pl += pl
                tag = "sellpos" if pl >= 0 else "sellneg"
                tv.insert("", "end", tags=(tag,), values=(
                    e.get("date", ""), "▼ بيع", e.get("name", ""),
                    "{:g}".format(round(e.get("qty_g", 0), 3)),
                    "{}{:,.2f}".format(sym, price),
                    "{}{:,.2f}".format(sym, amount),
                    "{}{}{:,.2f}".format("+" if pl > 0 else "", sym, pl)))

        info = ("إجمالي الشراء: {s}{b:,.2f}    |    إجمالي البيع: {s}{sl:,.2f}    |    الربح المحقق: {sg}{s}{p:,.2f}").format(
            s=sym, b=tot_buy, sl=tot_sell, p=tot_pl, sg="+" if tot_pl > 0 else "")
        tk.Label(win, text=info if self.history else "لا توجد عمليات بعد.",
                 bg=BG, fg=TXT, font=("Segoe UI", 10, "bold")).pack(anchor="e", padx=14, pady=(0, 6))

        def clear_hist():
            if self.history and messagebox.askyesno(
                    "تأكيد", "مسح سجل العمليات؟ (لا يؤثر على الرصيد أو الأرباح)", parent=win):
                self.history = []
                self.save()
                win.destroy()
        bar = tk.Frame(win, bg=BG); bar.pack(fill="x", padx=12, pady=(0, 10))
        self._mkbtn(bar, "مسح السجل", clear_hist, kind="danger",
                    font_size=9).pack(side="left")
        self._mkbtn(bar, "إغلاق", win.destroy, kind="neutral",
                    font_size=9).pack(side="left", padx=6)

    # ---------- حاسبة الزكاة ----------
    def open_zakat(self):
        t = self.theme
        BG, CARD = t["bg"], t["card"]
        GOLD, TXT, MUT = t["accent"], t["text"], t["muted"]
        sym = CUR[self.cur]["sym"]
        NISAB_G = 85.0            # نصاب الذهب ≈ 85 غراماً
        ZAKAT_RATE = 0.025       # 2.5%
        grams = self.total_open_grams()
        gram_val_disp = self.gram_market_disp()   # سعر الغرام السوقي الحالي
        total_val = grams * gram_val_disp
        zakat_g = grams * ZAKAT_RATE
        zakat_val = total_val * ZAKAT_RATE
        reached = grams >= NISAB_G

        win = tk.Toplevel(self.root)
        win.title("حاسبة الزكاة")
        win.configure(bg=BG)
        win.geometry("460x360")
        win.transient(self.root)
        tk.Label(win, text="📐 حاسبة زكاة الذهب", bg=BG, fg=GOLD,
                 font=("Segoe UI", 13, "bold")).pack(anchor="e", padx=16, pady=(14, 8))

        card = tk.Frame(win, bg=CARD); card.pack(fill="both", expand=True, padx=14, pady=8)

        def row(label, value, color=None):
            fr = tk.Frame(card, bg=CARD); fr.pack(fill="x", padx=14, pady=6)
            tk.Label(fr, text=label, bg=CARD, fg=MUT, font=("Segoe UI", 10)).pack(side="right")
            tk.Label(fr, text=value, bg=CARD, fg=color or TXT,
                     font=("Segoe UI", 11, "bold")).pack(side="left")

        row("الرصيد المملوك:", "{:,.3f} غرام".format(grams))
        row("سعر الغرام الحالي:", "{}{:,.2f}".format(sym, gram_val_disp))
        row("القيمة الإجمالية:", "{}{:,.2f}".format(sym, total_val))
        row("النصاب (85غ):",
            "✔ بلغ النصاب" if reached else "✖ دون النصاب",
            t.get("profit") if reached else t.get("loss"))
        tk.Frame(card, bg=t.get("border", "#2a3240"), height=1).pack(fill="x", padx=14, pady=8)
        row("الزكاة الواجبة (2.5%):",
            "{}{:,.2f}".format(sym, zakat_val) if reached else "لا تجب",
            GOLD if reached else MUT)
        if reached:
            row("أو ما يعادل:", "{:,.3f} غرام ذهب".format(zakat_g), GOLD)

        tk.Label(win, text="ملاحظة: تجب الزكاة بعد مرور الحول (سنة هجرية) على بلوغ النصاب.\n"
                      "الحساب تقديري للاستئناس فقط.",
                 bg=BG, fg=MUT, font=("Segoe UI", 8), justify="right").pack(anchor="e", padx=16, pady=(4, 10))

    # ---------- مخطط سعر الغرام خلال اليوم ----------
    def _record_price_sample(self, oz_usd):
        """تسجيل عينة سعر (أونصة دولار) مع حذف ما تجاوز 24 ساعة."""
        try:
            now = time.time()
            self.price_hist.append([now, float(oz_usd)])
            cutoff = now - 24 * 3600
            self.price_hist = [p for p in self.price_hist if p[0] >= cutoff][-500:]
        except Exception:
            pass

    def open_chart(self):
        t = self.theme
        BG, CARD = t["bg"], t["card"]
        GOLD, TXT, MUT = t["accent"], t["text"], t["muted"]
        sym = CUR[self.cur]["sym"]
        win = tk.Toplevel(self.root)
        win.title("مخطط سعر الغرام - اليوم")
        win.configure(bg=BG)
        win.geometry("640x420")
        win.transient(self.root)
        tk.Label(win, text="📈 حركة سعر الغرام (السوقي) خلال اليوم", bg=BG, fg=GOLD,
                 font=("Segoe UI", 12, "bold")).pack(anchor="e", padx=14, pady=(12, 4))
        W, H = 600, 320
        cv = tk.Canvas(win, width=W, height=H, bg=CARD, highlightthickness=0)
        cv.pack(padx=14, pady=8)

        # عينات اليوم الحالي فقط
        today = datetime.now().strftime("%Y-%m-%d")
        rate = self.rate()
        pts = [(ts, (oz / OZ_TO_G) * rate) for ts, oz in self.price_hist
               if datetime.fromtimestamp(ts).strftime("%Y-%m-%d") == today]
        if len(pts) < 2:
            cv.create_text(W // 2, H // 2,
                           text="لا توجد بيانات كافية بعد.\nاترك البرنامج مفتوحاً لتجميع الأسعار.",
                           fill=MUT, font=("Segoe UI", 11), justify="center")
            self._mkbtn(win, "إغلاق", win.destroy, kind="neutral",
                        font_size=9).pack(pady=6)
            return
        pl, pr, pt, pb = 64, 20, 20, 40    # هوامش
        xs = [p[0] for p in pts]; ys = [p[1] for p in pts]
        xmin, xmax = min(xs), max(xs)
        ymin, ymax = min(ys), max(ys)
        if ymax - ymin < 1e-9:
            ymin -= 1; ymax += 1
        pad = (ymax - ymin) * 0.1
        ymin -= pad; ymax += pad

        def X(x):
            return pl + (x - xmin) / (xmax - xmin) * (W - pl - pr) if xmax > xmin else pl

        def Y(y):
            return pt + (ymax - y) / (ymax - ymin) * (H - pt - pb)

        # شبكة + محاور أفقية
        for i in range(5):
            yy = ymin + (ymax - ymin) * i / 4
            y = Y(yy)
            cv.create_line(pl, y, W - pr, y, fill=t.get("border", "#2a3240"))
            cv.create_text(pl - 6, y, text="{:,.1f}".format(yy), fill=MUT,
                           anchor="e", font=("Segoe UI", 8))
        # تسميات زمنية (بداية/نهاية)
        cv.create_text(pl, H - pb + 14,
                       text=datetime.fromtimestamp(xmin).strftime("%H:%M"),
                       fill=MUT, anchor="w", font=("Segoe UI", 8))
        cv.create_text(W - pr, H - pb + 14,
                       text=datetime.fromtimestamp(xmax).strftime("%H:%M"),
                       fill=MUT, anchor="e", font=("Segoe UI", 8))
        # الخط
        coords = []
        for x, y in pts:
            coords += [X(x), Y(y)]
        up = ys[-1] >= ys[0]
        line_col = t.get("profit", "#2ecc71") if up else t.get("loss", "#ff5c5c")
        cv.create_line(*coords, fill=line_col, width=2, smooth=True)
        # نقطة النهاية
        cv.create_oval(X(xs[-1]) - 4, Y(ys[-1]) - 4, X(xs[-1]) + 4, Y(ys[-1]) + 4,
                       fill=line_col, outline="")
        change = ys[-1] - ys[0]
        pct = (change / ys[0] * 100) if ys[0] else 0
        tk.Label(win,
                 text="الافتتاح: {s}{o:,.2f}   •   الآن: {s}{c:,.2f}   •   التغيّر: {sg}{s}{d:,.2f} ({sg}{p:.2f}%)".format(
                     s=sym, o=ys[0], c=ys[-1], d=change, p=pct,
                     sg="+" if change > 0 else ""),
                 bg=BG, fg=line_col, font=("Segoe UI", 10, "bold")).pack(anchor="e", padx=14, pady=4)
        self._mkbtn(win, "إغلاق", win.destroy, kind="neutral",
                    font_size=9).pack(pady=6)

    # ---------- التنبيهات ----------
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
        self.alerts.append({
            "metric": self.al_metric.get(),
            "op": self.al_op.get(),
            "price": price,
            "cur": self.cur,
            "triggered": False,
        })
        self.al_price.delete(0, "end")
        self.save()
        self.refresh_alert_list()

    def del_alert(self):
        if not getattr(self, "al_list", None):
            return
        sel = self.al_list.curselection()
        if not sel:
            return
        idx = sel[0]
        if 0 <= idx < len(self.alerts):
            del self.alerts[idx]
            self.save()
            self.refresh_alert_list()

    def refresh_alert_list(self):
        if not getattr(self, "al_list", None):
            return
        try:
            self.al_list.delete(0, "end")
            for a in self.alerts:
                unit = "أونصة" if a["metric"] == "oz" else "غرام"
                state = " ✔ تم التنبيه" if a.get("triggered") else ""
                self.al_list.insert("end", "عندما يصبح سعر ال{u} {op} {p:,.2f} {c}{s}".format(
                    u=unit, op=a["op"], p=a["price"], c=CUR[a["cur"]]["name"], s=state))
        except Exception:
            pass

    def eff_price_in_cur(self, metric, ccy):
        """السعر الفعلي للوحدة بعملة محددة (بعد طرح المبلغ)."""
        gram = (self.spot_oz_usd / OZ_TO_G) * self.rates.get(ccy, 1.0) - self.adjust
        return gram * OZ_TO_G if metric == "oz" else gram

    def check_alerts(self):
        if not self.spot_oz_usd:
            return
        fired = []
        changed = False
        for a in self.alerts:
            cur_val = self.eff_price_in_cur(a["metric"], a["cur"])
            hit = (cur_val >= a["price"]) if a["op"] == ">=" else (cur_val <= a["price"])
            if hit and not a.get("triggered"):
                a["triggered"] = True
                changed = True
                unit = "الأونصة" if a["metric"] == "oz" else "الغرام"
                fired.append("سعر {u} وصل {op} {p:,.2f} {c}\nالسعر الحالي: {v:,.2f} {c}".format(
                    u=unit, op=a["op"], p=a["price"], v=cur_val, c=CUR[a["cur"]]["name"]))
            elif not hit and a.get("triggered"):
                a["triggered"] = False   # إعادة التسليح عند زوال الشرط
                changed = True
        if fired:
            self._play_sound()
            messagebox.showinfo("🔔 تنبيه سعر الذهب", "\n\n".join(fired))
        if changed:
            self.save()
            self.refresh_alert_list()

    def _play_sound(self):
        if not self.sound_on:
            return
        threading.Thread(target=self._sound_worker, daemon=True).start()

    def _test_sound(self):
        threading.Thread(target=self._sound_worker, daemon=True).start()

    def _sound_worker(self):
        try:
            import winsound
            # نغمة تنبيه واضحة متكررة
            for _ in range(3):
                winsound.Beep(1046, 220)   # دو
                winsound.Beep(1318, 220)   # مي
                winsound.Beep(1568, 320)   # صول
        except Exception:
            try:
                for _ in range(5):
                    self.root.bell()
                    self.root.after(150)
            except Exception:
                pass

    def on_sound_toggle(self):
        self.sound_on = bool(self.sound_var.get())
        self.save()

    def start_live_refresh(self):
        """تحديث مستمر (شبه لحظي) لسعر السوق."""
        self._live = True
        self._ticks = 0          # عدد التحديثات الناجحة
        self._fails = 0          # فشل متتالٍ
        self._fx_last = 0.0      # آخر وقت جلب سعر الصرف
        threading.Thread(target=self._live_worker, daemon=True).start()

    def _live_worker(self):
        while getattr(self, "_live", True):
            try:
                self._fetch_once()
            except Exception:
                pass   # لا نسمح لأي خطأ بإيقاف حلقة التحديث
            # أقصر مدة آمنة لا تسبب حظراً
            time.sleep(5)

    # ---------- الإعدادات والألوان ----------
    def rebuild_ui(self):
        for w in self.root.winfo_children():
            w.destroy()
        self.build_ui()
        self.update_price_labels()
        self.render()

    def open_settings(self):
        from tkinter import colorchooser
        t = self.theme
        win = tk.Toplevel(self.root)
        win.title("الإعدادات والألوان")
        win.configure(bg=t["bg"])
        win.geometry("460x640")
        win.transient(self.root)
        PAD = {"padx": 10, "pady": 5}

        # ===== مصدر السعر =====
        tk.Label(win, text="مصدر السعر", bg=t["bg"], fg=t["accent"],
                 font=("Segoe UI", 12, "bold")).pack(anchor="e", **PAD)
        src_disp = {"auto": "تلقائي (موصى به — يجرّب المصادر بالترتيب)"}
        for _k in AUTO_ORDER:
            src_disp[_k] = SOURCES[_k][0]
        src_rev = {v: k for k, v in src_disp.items()}
        cur_src_disp = src_disp.get(self.source_mode, src_disp["auto"])
        srcfr = tk.Frame(win, bg=t["bg"]); srcfr.pack(fill="x", **PAD)
        tk.Label(srcfr, text="المصدر", bg=t["bg"], fg=t["text"], width=10,
                 anchor="e").pack(side="right")
        src_var = tk.StringVar(value=cur_src_disp)
        ttk.OptionMenu(srcfr, src_var, cur_src_disp,
                       *src_disp.values()).pack(side="right", fill="x",
                                                expand=True, padx=6)
        tk.Label(win, text="اختر مصدراً محدداً للسعر أو دعه تلقائياً. يظهر المصدر الفعّال بسطر الحالة.",
                 bg=t["bg"], fg=t["muted"], font=("Segoe UI", 8),
                 wraplength=420, justify="right").pack(anchor="e", padx=10)

        tk.Label(win, text="العناوين", bg=t["bg"], fg=t["accent"],
                 font=("Segoe UI", 12, "bold")).pack(anchor="e", **PAD)
        tvars = {}
        for key, lbl in (("title", "العنوان الرئيسي"),
                         ("subtitle", "العنوان الفرعي"),
                         ("window", "عنوان النافذة")):
            fr = tk.Frame(win, bg=t["bg"]); fr.pack(fill="x", **PAD)
            tk.Label(fr, text=lbl, bg=t["bg"], fg=t["text"], width=14,
                     anchor="e").pack(side="right")
            v = tk.StringVar(value=self.titles.get(key, ""))
            tk.Entry(fr, textvariable=v).pack(side="right", fill="x", expand=True, padx=6)
            tvars[key] = v

        tk.Label(win, text="الألوان", bg=t["bg"], fg=t["accent"],
                 font=("Segoe UI", 12, "bold")).pack(anchor="e", **PAD)
        color_fields = [
            ("bg", "خلفية التطبيق"), ("card", "البطاقات"),
            ("accent", "اللون المميز"), ("text", "لون النص"),
            ("profit", "لون الربح"), ("loss", "لون الخسارة"),
            ("total", "لون سطر الإجمالي"),
        ]
        swatches = {}
        tmp = dict(self.theme)

        def make_picker(k, swatch):
            def pick():
                c = colorchooser.askcolor(color=tmp.get(k), title="اختر لوناً", parent=win)
                if c and c[1]:
                    tmp[k] = c[1]
                    swatch.config(bg=c[1])
            return pick

        for key, lbl in color_fields:
            fr = tk.Frame(win, bg=t["bg"]); fr.pack(fill="x", **PAD)
            tk.Label(fr, text=lbl, bg=t["bg"], fg=t["text"], width=16,
                     anchor="e").pack(side="right")
            sw = tk.Label(fr, text="      ", bg=tmp.get(key, "#000"),
                          relief="solid", bd=1)
            sw.pack(side="right", padx=6)
            tk.Button(fr, text="اختيار لون", relief="flat",
                      command=make_picker(key, sw)).pack(side="right")
            swatches[key] = sw

        btnfr = tk.Frame(win, bg=t["bg"]); btnfr.pack(fill="x", pady=16)

        def do_save():
            for k, v in tvars.items():
                self.titles[k] = v.get().strip() or DEFAULT_TITLES.get(k, "")
            self.theme.update(tmp)
            self.source_mode = src_rev.get(src_var.get(), "auto")
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

        tk.Button(btnfr, text="حفظ وتطبيق", command=do_save, bg=t["accent"],
                  fg=t["accent_text"], relief="flat",
                  font=("Segoe UI", 10, "bold")).pack(side="right", padx=10)
        tk.Button(btnfr, text="استعادة الافتراضي", command=do_reset, bg="#444",
                  fg="white", relief="flat").pack(side="right")

    # ---------- تصدير Excel ----------
    def export_excel(self):
        from tkinter import filedialog
        if not self.holdings and not self.sells:
            messagebox.showinfo("تصدير Excel", "لا توجد عمليات لتصديرها.")
            return
        default = "محفظة_الذهب_{}.xlsx".format(datetime.now().strftime("%Y%m%d_%H%M"))
        path = filedialog.asksaveasfilename(
            title="حفظ ملف Excel", defaultextension=".xlsx",
            initialfile=default,
            filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        ccy = CUR[self.cur]["name"]
        headers = ["الاسم", "الحالة", "الكمية", "الوحدة",
                   "سعر الشراء/وحدة ({})".format(ccy),
                   "التكلفة ({})".format(ccy),
                   "السعر الحالي/البيع ({})".format(ccy),
                   "القيمة ({})".format(ccy),
                   "الربح/الخسارة ({})".format(ccy), "النسبة %"]
        rows = []
        t_cost = t_val = 0.0
        t_qty_g = 0.0
        realized = 0.0
        # المراكز المفتوحة (الرصيد المتبقي)
        for h in self.holdings:
            unit_txt = "أونصة" if h["unit"] == "oz" else "غرام"
            cost_disp = h["qty"] * h["buy"] * self.rate()
            cur_disp = self.current_per_unit_disp(h["unit"])
            value_disp = h["qty"] * cur_disp
            pl = value_disp - cost_disp
            pct = (pl / cost_disp * 100) if cost_disp else 0
            t_cost += cost_disp
            t_val += value_disp
            t_qty_g += h["qty"] * (OZ_TO_G if h["unit"] == "oz" else 1)
            rows.append([
                h["name"], "مفتوح", round(h["qty"], 6), unit_txt,
                round(h["buy"] * self.rate(), 3), round(cost_disp, 3),
                round(cur_disp, 3), round(value_disp, 3),
                round(pl, 3), round(pct, 2),
            ])
        # عمليات البيع المحققة
        for s in self.sells:
            qty_g = s["qty_g"]
            cost_disp = s["cost"] * self.rate()
            sell_disp = s["sell"] * self.rate()
            value_disp = qty_g * sell_disp
            pl = value_disp - cost_disp
            pct = (pl / cost_disp * 100) if cost_disp else 0
            realized += pl
            buy_per_g_disp = (cost_disp / qty_g) if qty_g else 0.0
            rows.append([
                "{} ({})".format(s.get("name", "بيع ذهب"), s.get("date", "")),
                "مباع", round(qty_g, 6), "غرام",
                round(buy_per_g_disp, 3), round(cost_disp, 3),
                round(sell_disp, 3), round(value_disp, 3),
                round(pl, 3), round(pct, 2),
            ])
        t_pl = t_val - t_cost
        t_pct = (t_pl / t_cost * 100) if t_cost else 0
        rows.append(["∑ الرصيد المتبقي (غير محقق)", "", round(t_qty_g, 3), "غرام",
                     "", round(t_cost, 3), "", round(t_val, 3),
                     round(t_pl, 3), round(t_pct, 2)])
        rows.append(["∑ الربح المحقق (المباع)", "", "", "",
                     "", "", "", "", round(realized, 3), ""])
        rows.append(["∑ الإجمالي (محقق + غير محقق)", "", "", "",
                     "", "", "", "", round(t_pl + realized, 3), ""])
        try:
            write_xlsx(path, headers, rows)
        except Exception as e:
            messagebox.showerror("تصدير Excel", "تعذّر حفظ الملف:\n{}".format(e))
            return
        messagebox.showinfo("تصدير Excel",
                            "تم تصدير {} سجل بنجاح.".format(len(self.holdings) + len(self.sells)))

    # ---------- نافذة تنبيه السعر (مستقلة) ----------
    def open_alerts(self):
        if getattr(self, "al_win", None) is not None:
            try:
                self.al_win.deiconify()
                self.al_win.lift()
                return
            except Exception:
                self.al_win = None
        t = self.theme
        BG, CARD, CARD2 = t["bg"], t["card"], t["card2"]
        GOLD, ATXT, TXT, MUT = t["accent"], t["accent_text"], t["text"], t["muted"]
        win = tk.Toplevel(self.root)
        self.al_win = win
        win.title("إضافة تنبيه السعر")
        win.configure(bg=BG)
        win.geometry("540x360")
        win.transient(self.root)

        al = tk.LabelFrame(win, text="🔔 تنبيهات الأسعار", bg=CARD, fg=GOLD,
                           font=("Segoe UI", 10, "bold"))
        al.pack(fill="both", expand=True, padx=12, pady=12)

        row = tk.Frame(al, bg=CARD)
        row.pack(fill="x", padx=6, pady=6)
        tk.Label(row, text="نبّهني عندما يصبح سعر", bg=CARD, fg=TXT).pack(side="right")
        self.al_metric = tk.StringVar(value="g")
        ttk.OptionMenu(row, self.al_metric, "g", "g", "oz").pack(side="right", padx=4)
        self.al_op = tk.StringVar(value=">=")
        ttk.OptionMenu(row, self.al_op, ">=", ">=", "<=").pack(side="right", padx=4)
        self.al_price = tk.Entry(row, width=10)
        self.al_price.pack(side="right", padx=4)
        self.al_cur_lbl = tk.Label(row, text="(" + CUR[self.cur]["name"] + ")", bg=CARD, fg=MUT)
        self.al_cur_lbl.pack(side="right")
        tk.Button(row, text="إضافة تنبيه", command=self.add_alert, bg=GOLD,
                  fg=ATXT, relief="flat", font=("Segoe UI", 9, "bold")).pack(side="right", padx=8)
        tk.Label(al, text="(السعر = سعر الراجحي للبيع بعد الطرح / بالوحدة المختارة)",
                 bg=CARD, fg=MUT, font=("Segoe UI", 8)).pack(anchor="e", padx=10)

        row2 = tk.Frame(al, bg=CARD)
        row2.pack(fill="both", expand=True, padx=6, pady=6)
        self.al_list = tk.Listbox(row2, height=6, bg=CARD2, fg=TXT,
                                  selectbackground="#33415c", relief="flat",
                                  font=("Segoe UI", 9))
        self.al_list.pack(side="left", fill="both", expand=True)
        tk.Button(row2, text="حذف التنبيه", command=self.del_alert, bg=t["loss"],
                  fg="white", relief="flat").pack(side="left", padx=6)

        row3 = tk.Frame(al, bg=CARD)
        row3.pack(fill="x", padx=6, pady=6)
        self.sound_var = tk.BooleanVar(value=self.sound_on)
        tk.Checkbutton(row3, text="تفعيل التنبيه الصوتي 🔊", variable=self.sound_var,
                       command=self.on_sound_toggle, bg=CARD, fg=TXT,
                       selectcolor=CARD2, activebackground=CARD,
                       activeforeground=TXT).pack(side="right")
        tk.Button(row3, text="تجربة الصوت", command=self._test_sound, bg=CARD2,
                  fg=GOLD, relief="flat").pack(side="right", padx=8)

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

    # ---------- جلب البيانات ----------
    def refresh_all(self):
        threading.Thread(target=self._fetch_once, daemon=True).start()

    def _fetch_once(self):
        msg = ""
        direction = None   # 'up' / 'down' / 'same' / None(خطأ)
        ok = False
        try:
            new_price, source = fetch_spot_usd(self.source_mode)
            old_price = self.spot_oz_usd
            self._delta_oz = (new_price - old_price) if old_price else 0.0
            if old_price and new_price > old_price:
                direction = "up"
            elif old_price and new_price < old_price:
                direction = "down"
            else:
                direction = "same"
            self.spot_oz_usd = new_price
            ok = True
            self._record_price_sample(new_price)
            self._ticks = getattr(self, "_ticks", 0) + 1
            self._fails = 0
            now = datetime.now().strftime("%H:%M:%S")
            self.last_update = now
            # حفظ آخر سعر معروف دورياً لعرضه عند إعادة الفتح
            if self._ticks == 1 or self._ticks % 12 == 0:
                self.save()
            closed = "  •  🔒 السوق مغلق (عطلة نهاية الأسبوع) — السعر ثابت حتى الفتح" if market_closed() else ""
            msg = "● مباشر ({src}) • الأونصة ${p:.3f} • {t} • تحديثات {n}{c}".format(
                src=source, p=new_price, t=now, n=self._ticks, c=closed)
        except Exception:
            direction = None
            self._fails = getattr(self, "_fails", 0) + 1
            if self.spot_oz_usd:
                # نحتفظ بآخر سعر ناجح بدل تجميد الشاشة
                msg = "تعذّر التحديث (محاولة {})… عرض آخر سعر معروف".format(self._fails)
            else:
                msg = "تعذّر جلب سعر الذهب — تحقق من الإنترنت"

        # سعر الصرف نادر التغير — يُجلب مرة كل ~5 دقائق فقط (تجنّب الحظر)
        nowt = time.monotonic()
        if nowt - getattr(self, "_fx_last", 0.0) > 300:
            try:
                fx = http_json(FX_URL, timeout=8)
                if fx.get("rates"):
                    self.rates["SAR"] = fx["rates"].get("SAR", self.rates["SAR"])
                    self.rates["AED"] = fx["rates"].get("AED", self.rates["AED"])
                self._fx_last = nowt
            except Exception:
                pass

        self.root.after(0, lambda: (self.lbl_upd.config(text=msg),
                                    self.render(),
                                    self.flash_update(direction)))

    def flash_update(self, direction=None):
        """وميض بصري لحظي عند وصول تحديث جديد للسعر."""
        t = self.theme
        hi = {
            "up": t.get("profit", "#2ecc71"),      # أخضر عند الزيادة
            "down": t.get("loss", "#ff5c5c"),      # أحمر عند النزول
            "same": t.get("flash_same", "#ffd93b"),  # أصفر عند ثبات السعر
        }.get(direction, t.get("muted", "#888"))
        rest = t.get("muted", "#888")
        try:
            self.dot.config(fg=hi)
        except Exception:
            pass
        originals = []
        # الوميض فقط على خانة السعر الحالي (الغرام - السوق)
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

        self.root.after(260, revert)

    # ---------- العرض ----------
    def render(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        self.row_map = {}
        t_cost = t_val = 0.0
        t_qty_g = 0.0            # إجمالي الكمية المتبقية (غرام)
        realized = 0.0          # إجمالي الربح/الخسارة المحقق (المباع)
        sym = CUR[self.cur]["sym"]
        rowi = 0
        # ===== المراكز المفتوحة (الرصيد المتبقي) =====
        for idx, h in enumerate(self.holdings):
            unit_txt = "أونصة" if h["unit"] == "oz" else "غرام"
            cost_usd = h["qty"] * h["buy"]
            cost_disp = cost_usd * self.rate()
            cur_disp = self.current_per_unit_disp(h["unit"])
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
            self.tree.insert("", "end", iid=iid, tags=row_tags, values=(
                h["name"],
                "● مفتوح",
                "{:g} {}".format(h["qty"], unit_txt),
                "{}{:,.2f}".format(sym, h["buy"] * self.rate()),
                "{}{:,.2f}".format(sym, cost_disp),
                "{}{:,.2f}".format(sym, cur_disp),
                "{}{:,.2f}".format(sym, value_disp),
                "{}{}{:,.2f}".format("+" if pl > 0 else "", sym, pl),
                "{}{:.2f}%".format("+" if pl > 0 else "", pct),
            ))
            rowi += 1
        # ===== عمليات البيع (محققة — سطر أخضر بسعر بيع مثبّت) =====
        for sidx, s in enumerate(self.sells):
            qty_g = s["qty_g"]
            cost_disp = s["cost"] * self.rate()
            sell_disp = s["sell"] * self.rate()          # دولار/غرام × الصرف
            value_disp = qty_g * sell_disp
            pl = value_disp - cost_disp
            pct = (pl / cost_disp * 100) if cost_disp else 0
            realized += pl
            buy_per_g_disp = (cost_disp / qty_g) if qty_g else 0.0
            iid = "s{}".format(sidx)
            self.row_map[iid] = ("sell", sidx)
            self.tree.insert("", "end", iid=iid, tags=("soldgreen",), values=(
                "💰 {}".format(s.get("name", "بيع ذهب")),
                "✔ بيع",
                "{:g} غرام".format(qty_g),
                "{}{:,.2f}".format(sym, buy_per_g_disp),
                "{}{:,.2f}".format(sym, cost_disp),
                "{}{:,.2f}".format(sym, sell_disp),
                "{}{:,.2f}".format(sym, value_disp),
                "{}{}{:,.2f}".format("+" if pl > 0 else "", sym, pl),
                "{}{:.2f}%".format("+" if pl > 0 else "", pct),
            ))
        self.update_price_labels()
        self.check_alerts()
        t_pl = t_val - t_cost
        t_pct = (t_pl / t_cost * 100) if t_cost else 0
        total_pl = t_pl + realized
        # ===== سطر إجمالي الرصيد المتبقي =====
        if self.holdings:
            ttag = "totalpos" if t_pl > 0 else ("totalneg" if t_pl < 0 else "total")
            self.tree.insert("", "end", tags=(ttag,), values=(
                "∑ الرصيد المتبقي",
                "-",
                "{:,.2f} غرام".format(t_qty_g),
                "-",
                "{}{:,.2f}".format(sym, t_cost),
                "-",
                "{}{:,.2f}".format(sym, t_val),
                "{}{}{:,.2f}".format("+" if t_pl > 0 else "", sym, t_pl),
                "{}{:.2f}%".format("+" if t_pl > 0 else "", t_pct),
            ))
        self.lbl_sum.config(
            text=("الربح المحقق: {sg_r}{s}{r:,.2f}    |    غير المحقق: {sg_u}{s}{u:,.2f}    |    إجمالي الربح/الخسارة: {sg_t}{s}{tp:,.2f}").format(
                s=sym, r=realized, u=t_pl, tp=total_pl,
                sg_r="+" if realized > 0 else "",
                sg_u="+" if t_pl > 0 else "",
                sg_t="+" if total_pl > 0 else ""),
            fg=self.theme["total"])

    def update_price_labels(self):
        if self.spot_oz_usd:
            sym = CUR[self.cur]["sym"]
            oz_disp = self.spot_oz_usd * self.rate()
            self.lbl_oz.config(text="{}{:,.3f}".format(sym, oz_disp))
            self.lbl_gr.config(text="{}{:,.3f}".format(sym, self.gram_market_disp()))
            self.lbl_gr_buy.config(text="{}{:,.3f}".format(sym, self.gram_buy_disp()))
            self.lbl_gr_adj.config(text="{}{:,.3f}".format(sym, self.gram_eff_disp()))
            self._update_delta_indicator()
        else:
            for lbl in (self.lbl_oz, self.lbl_gr, self.lbl_gr_buy, self.lbl_gr_adj):
                lbl.config(text="--")

    def _update_delta_indicator(self):
        """مؤشر تغيّر سعر الغرام السوقي منذ آخر تحديث."""
        lbl = getattr(self, "lbl_delta", None)
        if lbl is None:
            return
        t = self.theme
        sym = CUR[self.cur]["sym"]
        delta_oz = getattr(self, "_delta_oz", 0.0) or 0.0
        if abs(delta_oz) < 1e-9 or not self.spot_oz_usd:
            lbl.config(text="━ ثابت", fg=t.get("flash_same", "#ffd93b"))
            return
        gram_delta_disp = (delta_oz / OZ_TO_G) * self.rate()
        prev_oz = self.spot_oz_usd - delta_oz
        pct = (delta_oz / prev_oz * 100) if prev_oz else 0.0
        if delta_oz > 0:
            arrow, color = "▲", t.get("profit", "#2ecc71")
        else:
            arrow, color = "▼", t.get("loss", "#ff5c5c")
        lbl.config(
            text="{a} {sym}{d:,.3f}/غ ({p:+.2f}%)".format(
                a=arrow, sym=sym, d=abs(gram_delta_disp), p=pct),
            fg=color)


def main():
    root = tk.Tk()
    GoldApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
