import re
from io import BytesIO
from urllib.parse import urlparse, parse_qs

import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import ColorScaleRule


# ============================================================
# MAD Orders Dashboard V7 - Operations Control Center
# ============================================================

DEFAULT_GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1Lf7R_G5hZ6KvyE5OyRc78b1dKVjD1bEDeeZnorANrxI/edit?usp=sharing"
APP_VERSION = "V7.3 Readable Charts + Fillings Reports"


# =========================
# Page Config
# =========================
st.set_page_config(
    page_title="MAD Orders Control Center",
    page_icon="🧁",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =========================
# CSS
# =========================
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap');

    html, body, [class*="css"] { font-family: 'Cairo', sans-serif !important; }

    .stApp {
        background:
            radial-gradient(circle at top left, rgba(14,165,233,.16), transparent 32%),
            radial-gradient(circle at top right, rgba(34,197,94,.12), transparent 28%),
            linear-gradient(135deg, #07111f 0%, #0f172a 45%, #111827 100%);
    }

    .block-container { padding-top: 1.2rem; padding-bottom: 3rem; max-width: 1500px; }

    .hero {
        padding: 24px 26px;
        border-radius: 24px;
        background: linear-gradient(135deg, rgba(15,23,42,.88), rgba(30,41,59,.70));
        border: 1px solid rgba(148,163,184,.28);
        box-shadow: 0 18px 70px rgba(0,0,0,.35);
        margin-bottom: 18px;
        direction: rtl;
    }
    .hero h1 { font-size: 34px; margin: 0 0 8px 0; font-weight: 800; color: #f8fafc; }
    .hero p { margin: 0; color: #cbd5e1; line-height: 1.8; font-size: 14px; }

    .section-title {
        direction: rtl;
        font-size: 22px;
        font-weight: 800;
        color: #f8fafc;
        margin: 18px 0 10px 0;
    }

    .mini-title {
        direction: rtl;
        font-size: 17px;
        font-weight: 800;
        color: #e5e7eb;
        margin: 10px 0 8px 0;
    }

    .kpi-card {
        direction: rtl;
        position: relative;
        min-height: 118px;
        padding: 18px 18px;
        border-radius: 20px;
        background: linear-gradient(145deg, rgba(15,23,42,.94), rgba(30,41,59,.72));
        border: 1px solid rgba(148,163,184,.22);
        overflow: hidden;
        box-shadow: 0 14px 45px rgba(0,0,0,.28);
        margin-bottom: 12px;
    }
    .kpi-card:before {
        content: '';
        position: absolute;
        inset: 0 auto 0 0;
        width: 5px;
        background: var(--accent, #22c55e);
        box-shadow: 0 0 26px var(--accent, #22c55e);
    }
    .kpi-label { font-size: 13px; color: #94a3b8; font-weight: 700; margin-bottom: 9px; }
    .kpi-value { font-size: 26px; color: #ffffff; font-weight: 800; line-height: 1.15; }
    .kpi-sub { font-size: 12px; color: #cbd5e1; margin-top: 7px; }

    .alert-box {
        direction: rtl;
        padding: 14px 16px;
        border-radius: 16px;
        background: rgba(251,191,36,.10);
        border: 1px solid rgba(251,191,36,.35);
        color: #fde68a;
        margin-bottom: 8px;
        line-height: 1.7;
        font-weight: 700;
    }
    .good-box {
        direction: rtl;
        padding: 14px 16px;
        border-radius: 16px;
        background: rgba(34,197,94,.10);
        border: 1px solid rgba(34,197,94,.35);
        color: #bbf7d0;
        margin-bottom: 8px;
        line-height: 1.7;
        font-weight: 700;
    }

    .note-box {
        direction: rtl;
        padding: 15px 16px;
        border-radius: 16px;
        background: rgba(59,130,246,.10);
        border: 1px solid rgba(59,130,246,.28);
        color: #bfdbfe;
        margin: 10px 0;
        line-height: 1.8;
    }

    .stTabs [data-baseweb="tab-list"] { gap: 8px; overflow-x: auto; }
    .stTabs [data-baseweb="tab"] {
        background: rgba(15,23,42,.8);
        border: 1px solid rgba(148,163,184,.20);
        border-radius: 14px 14px 0 0;
        padding: 9px 15px;
        color: #cbd5e1;
        min-width: max-content;
    }
    .stTabs [aria-selected="true"] { background: rgba(34,197,94,.22) !important; color: white !important; }

    div[data-testid="stDataFrame"] { direction: rtl; }

    @media (max-width: 768px) {
        .block-container { padding-left: .70rem !important; padding-right: .70rem !important; padding-top: .75rem !important; max-width: 100% !important; }
        .hero { padding: 18px 15px !important; border-radius: 18px !important; }
        .hero h1 { font-size: 22px !important; line-height: 1.35 !important; }
        .hero p { font-size: 12px !important; }
        .kpi-card { min-height: auto !important; padding: 15px 14px !important; border-radius: 16px !important; }
        .kpi-value { font-size: 21px !important; }
        .section-title { font-size: 18px !important; }
        div[data-testid="column"] { width: 100% !important; flex: 1 1 100% !important; min-width: 100% !important; }
        div[data-testid="stHorizontalBlock"] { flex-wrap: wrap !important; gap: .5rem !important; }
        .stTabs [data-baseweb="tab-list"] { flex-wrap: nowrap !important; overflow-x: auto !important; padding-bottom: 6px !important; }
        .stTabs [data-baseweb="tab"] { font-size: 12px !important; padding: 8px 11px !important; }
        .stButton button, .stDownloadButton button { width: 100% !important; min-height: 44px !important; }
        .js-plotly-plot, .plotly, .plot-container, .svg-container { width: 100% !important; max-width: 100% !important; }
        section[data-testid="stSidebar"] { min-width: 285px !important; }
    }
    
    /* V7.3 Readability fixes */
    .js-plotly-plot .plotly .legend text,
    .js-plotly-plot .plotly .xtick text,
    .js-plotly-plot .plotly .ytick text,
    .js-plotly-plot .plotly .gtitle,
    .js-plotly-plot .plotly .xtitle,
    .js-plotly-plot .plotly .ytitle {
        fill: #f8fafc !important;
        opacity: 1 !important;
        font-weight: 600 !important;
    }

    .js-plotly-plot .plotly .legend {
        opacity: 1 !important;
    }

    .readability-note {
        background: rgba(30, 41, 59, .65);
        border: 1px solid rgba(148, 163, 184, .24);
        padding: 12px 14px;
        border-radius: 16px;
        color: #e5e7eb;
        font-size: 14px;
        line-height: 1.8;
        margin: 8px 0 16px 0;
    }

    @media (max-width: 768px) {
        .js-plotly-plot .plotly .legend text,
        .js-plotly-plot .plotly .xtick text,
        .js-plotly-plot .plotly .ytick text {
            font-size: 11px !important;
        }
    }

    </style>
    """,
    unsafe_allow_html=True,
)


# =========================
# Core Helpers
# =========================

def normalize_arabic_digits(value):
    if pd.isna(value):
        return ""
    text = str(value)
    arabic_digits = "٠١٢٣٤٥٦٧٨٩"
    persian_digits = "۰۱۲۳۴۵۶۷۸۹"
    english_digits = "0123456789"
    for a, e in zip(arabic_digits, english_digits):
        text = text.replace(a, e)
    for p, e in zip(persian_digits, english_digits):
        text = text.replace(p, e)
    return text.strip()


def find_col(df, possible_names):
    cols = list(df.columns)
    lowered = {str(c).strip().lower(): c for c in cols}
    for name in possible_names:
        key = str(name).strip().lower()
        if key in lowered:
            return lowered[key]
    for col in cols:
        clean_col = str(col).strip().lower()
        for name in possible_names:
            if str(name).strip().lower() in clean_col:
                return col
    return None


def clean_money(value):
    text = normalize_arabic_digits(value)
    if not text or text.lower() in ["nan", "none", "(not set)", "not set", "-"]:
        return 0.0
    text = text.replace(",", "")
    m = re.search(r"-?\d+(?:\.\d+)?", text)
    return float(m.group(0)) if m else 0.0


def clean_quantity(value):
    text = normalize_arabic_digits(value).replace(",", "").strip()
    if not text or text.lower() in ["nan", "none", "(not set)", "not set", "-"]:
        return 1.0
    num = pd.to_numeric(text, errors="coerce")
    if pd.isna(num) or float(num) <= 0:
        return 1.0
    return float(num)


def extract_branch(chef_name):
    text = str(chef_name).strip()
    low = text.lower()
    branch_keywords = {
        "قرطبة": ["قرطبة", "qurtuba", "qortoba", "qurtobah"],
        "عريجاء": ["عريجاء", "عريجا", "العريجاء", "uraija", "uraijaa", "urejha"],
        "الروضة": ["الروضة", "روضه", "rawdah", "rawda"],
        "العارض": ["العارض", "arid", "al arid"],
        "الورود": ["الورود", "worood", "al worood"],
    }
    for branch, keywords in branch_keywords.items():
        if any(k.lower() in low for k in keywords):
            return branch
    return "العقيق"


def normalize_variety(value):
    text = str(value).strip()
    if not text or text.lower() in ["nan", "none", "(not set)", "not set", "-"]:
        return ""
    low = text.lower()
    if "sprinkles" in low or "سبرينكلز" in text:
        return "Vanilla Sprinkles"
    if ("lemon" in low and ("raspberry" in low or "رازبيري" in text or "رسبيري" in text)) or ("ليمون" in text and ("رازبيري" in text or "رسبيري" in text)):
        return "Lemon Raspberry"
    if "vanilla" in low or "فانيلا" in text:
        return "Vanilla"
    if "chocolate" in low or "شوكولات" in text or "تشوكلت" in text:
        return "Chocolate"
    if "mango" in low or "مانجو" in text:
        return "Mango"
    if "oreo" in low or "أوريو" in text or "اوريو" in text:
        return "Oreo"
    return re.sub(r"\s+", " ", text)


def normalize_product_name(value):
    text = str(value).strip()
    if not text or text.lower() in ["nan", "none", "(not set)", "not set", "-"]:
        return "بدون اسم منتج"
    return re.sub(r"\s+", " ", text)


def parse_datetime_parts(date_value, time_value):
    date_value = normalize_arabic_digits(date_value)
    time_value = normalize_arabic_digits(time_value)
    candidates = []
    if date_value and time_value:
        candidates.append(f"{date_value} {time_value}")
    if time_value:
        candidates.append(time_value)
    if date_value:
        candidates.append(date_value)
    for candidate in candidates:
        candidate = candidate.strip()
        if not candidate:
            continue
        for dayfirst in [True, False]:
            parsed = pd.to_datetime(candidate, errors="coerce", dayfirst=dayfirst)
            if not pd.isna(parsed):
                return parsed
    return pd.NaT


def hour_label(hour):
    if pd.isna(hour):
        return "بدون وقت"
    hour = int(hour)
    def fmt(h):
        period = "ص" if h < 12 else "م"
        h12 = h % 12 or 12
        return f"{h12} {period}"
    end_hour = (hour + 1) % 24
    return f"{fmt(hour).replace(' ص','').replace(' م','')}-{fmt(end_hour)}"


def is_cancelled_status(value):
    low = str(value).strip().lower()
    return any(k in low for k in ["cancel", "cancelled", "canceled", "ملغي", "ملغاة", "رفض"])


def is_addon_product(product_name):
    low = str(product_name).lower()
    keywords = [
        "candle", "candles", "شموع", "شمعة",
        "balloon", "helium", "بالون", "هيليوم",
        "gift card", "كرت", "بطاقة", "card",
        "night stars", "نجوم", "stars",
        "topper", "توبير",
    ]
    return any(k in low for k in keywords)


def addon_category(product_name):
    low = str(product_name).lower()
    if any(k in low for k in ["candle", "candles", "شموع", "شمعة"]):
        return "Candles / شموع"
    if any(k in low for k in ["balloon", "helium", "بالون", "هيليوم"]):
        return "Balloons / بالونات"
    if any(k in low for k in ["gift card", "كرت", "بطاقة", "card"]):
        return "Gift Cards / كروت"
    if any(k in low for k in ["night stars", "نجوم", "stars"]):
        return "Night Stars"
    if any(k in low for k in ["topper", "توبير"]):
        return "Toppers"
    return "Other Add-ons"


def classify_campaign(product_name):
    low = str(product_name).lower()
    if any(k in low for k in ["father", "dad", "عيد الأب", "عيد الاب", "بابا", "super dad"]):
        return "Father's Day"
    if any(k in low for k in ["birthday", "بيرثداي", "ميلاد"]):
        return "Birthday"
    if any(k in low for k in ["graduation", "تخرج", "التخرج"]):
        return "Graduation"
    if any(k in low for k in ["eid", "عيد الفطر", "عيد"]):
        return "Eid"
    if any(k in low for k in ["new year", "نيو يير"]):
        return "New Year"
    if any(k in low for k in ["valentine", "love", "الحب"]):
        return "Valentine"
    if any(k in low for k in ["gender reveal", "تحديد الجنس"]):
        return "Gender Reveal"
    return "General"


def extract_phone_from_text(value):
    text = normalize_arabic_digits(value)
    compressed = re.sub(r"[\s\-()]+", "", text)
    patterns = [r"\+?9665\d{8}", r"05\d{8}", r"5\d{8}"]
    for pat in patterns:
        m = re.search(pat, compressed)
        if m:
            return m.group(0)
    return ""


def need_action_reasons(note, product_name=""):
    text = f"{note} {product_name}".strip()
    low = text.lower()
    reasons = []
    if any(k in low for k in ["photo", "picture", "image", "صورة", "الصورة", "الصوره"]):
        reasons.append("يحتاج صورة")
    if any(k in low for k in ["contact", "call", "whatsapp", "text me", "تواصل", "اتصال", "واتساب", "جوال"]):
        reasons.append("يحتاج تواصل")
    if extract_phone_from_text(text):
        reasons.append("يوجد رقم جوال")
    if any(k in low for k in ["write", "writing", "اكتب", "كتابة", "العبارة", "الكتابة"]):
        reasons.append("كتابة خاصة")
    if any(k in low for k in ["draw", "design", "color", "hearts", "match", "تعديل", "تصميم", "لون", "قلوب", "ارسم"]):
        reasons.append("تعديل تصميم")
    if any(k in low for k in ["problem", "wrong", "mistake", "خطأ", "مشكلة", "مو نفس"]):
        reasons.append("ملاحظة حساسة")
    if len(str(note).strip()) > 90:
        reasons.append("ملاحظة طويلة")
    # Deduplicate preserving order
    seen, out = set(), []
    for r in reasons:
        if r not in seen:
            seen.add(r)
            out.append(r)
    return "، ".join(out)


def detect_columns(df):
    return {
        "order_no": find_col(df, ["رقم الطلب", "Order No", "order_no"]),
        "order_id": find_col(df, ["معرف الطلب (Order Id)", "Order Id", "order_id", "معرف الطلب"]),
        "customer": find_col(df, ["اسم العميل", "Customer", "customer_name"]),
        "chef": find_col(df, ["اسم الشيف", "Chef", "chef_name", "الفرع"]),
        "status": find_col(df, ["حالة الطلب", "Order Status", "order_status", "Status"]),
        "order_total": find_col(df, ["إجمالي الطلب بالكامل", "Order Total", "order_price", "Total"]),
        "delivery_date": find_col(df, ["تاريخ التوصيل (Delivery Date)", "Delivery Date", "delivery_date", "تاريخ التوصيل"]),
        "pickup_time": find_col(df, ["وقت الاستلام (Pickup Time)", "Pickup Time", "pickup_time", "وقت الاستلام"]),
        "item_id": find_col(df, ["معرف العنصر (Item Id)", "Item Id", "item_id"]),
        "dish_id": find_col(df, ["معرف الطبق (Dish Id)", "Dish Id", "dish_id"]),
        "product": find_col(df, ["اسم الطبق / المنتج", "Product", "Dish", "اسم الطبق", "اسم المنتج"]),
        "variety": find_col(df, ["نوع الحشوة (Variety)", "Variety", "نوع الحشوة", "الحشوة"]),
        "variety_price": find_col(df, ["سعر الحشوة", "Variety Price"]),
        "note": find_col(df, ["ملاحظة للشيف", "Chef Note", "Note", "ملاحظة"]),
        "unit_price": find_col(df, ["سعر الحبة", "Unit Price", "سعر"]),
        "discount": find_col(df, ["الخصم", "Discount"]),
        "quantity": find_col(df, ["الكمية", "Quantity", "qty"]),
        "item_total": find_col(df, ["إجمالي المنتج (Item Total)", "Item Total", "item_total", "إجمالي المنتج"]),
    }


def col_or_blank(df, col):
    if col and col in df.columns:
        return df[col].fillna("").astype(str)
    return pd.Series([""] * len(df), index=df.index)


def col_or_default(df, col, default=""):
    if col and col in df.columns:
        return df[col].fillna(default)
    return pd.Series([default] * len(df), index=df.index)


# =========================
# Data Loaders
# =========================

def google_sheet_to_csv_url(sheet_url, gid=""):
    if not sheet_url:
        return ""
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_url)
    if not m:
        return ""
    spreadsheet_id = m.group(1)
    parsed = urlparse(sheet_url)
    query_gid = parse_qs(parsed.query).get("gid", [""])[0]
    fragment_gid = ""
    if "gid=" in parsed.fragment:
        fragment_gid = parse_qs(parsed.fragment).get("gid", [""])[0]
    final_gid = str(gid or query_gid or fragment_gid or "0").strip()
    return f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv&gid={final_gid}"


@st.cache_data(ttl=600, show_spinner=False)
def load_google_sheet(sheet_url, gid=""):
    csv_url = google_sheet_to_csv_url(sheet_url, gid)
    if not csv_url:
        raise ValueError("رابط Google Sheet غير صحيح")
    return pd.read_csv(csv_url, dtype=str, keep_default_na=False).fillna("")


def read_uploaded_file(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded_file, dtype=str).fillna("")
    try:
        df = pd.read_csv(uploaded_file, sep="\t", dtype=str, keep_default_na=False, engine="python")
        if len(df.columns) > 1:
            return df.fillna("")
    except Exception:
        pass
    uploaded_file.seek(0)
    return pd.read_csv(uploaded_file, sep=None, dtype=str, keep_default_na=False, engine="python").fillna("")


# =========================
# Data Prep and Reports
# =========================

def prepare_data(raw_df):
    df = raw_df.copy().fillna("")
    df.columns = [str(c).strip() for c in df.columns]
    cols = detect_columns(df)

    if not cols["order_id"]:
        if cols["order_no"]:
            cols["order_id"] = cols["order_no"]
        else:
            df["OrderId_Auto"] = range(1, len(df) + 1)
            cols["order_id"] = "OrderId_Auto"

    df["رقم الطلب الموحد"] = col_or_blank(df, cols["order_id"]).replace("", pd.NA).fillna(col_or_blank(df, cols["order_no"])).astype(str)
    df["رقم الطلب الظاهر"] = col_or_blank(df, cols["order_no"]).replace("", pd.NA).fillna(df["رقم الطلب الموحد"]).astype(str)
    df["العميل"] = col_or_blank(df, cols["customer"])
    df["الشيف / الفرع الأصلي"] = col_or_blank(df, cols["chef"])
    df["الفرع"] = df["الشيف / الفرع الأصلي"].apply(extract_branch)
    df["الحالة"] = col_or_blank(df, cols["status"]).replace("", "غير محدد")
    df["المنتج"] = col_or_blank(df, cols["product"]).apply(normalize_product_name)
    df["الحشوة"] = col_or_blank(df, cols["variety"]).apply(normalize_variety)
    df["الملاحظة"] = col_or_blank(df, cols["note"])
    df["الكمية رقم"] = col_or_default(df, cols["quantity"], "1").apply(clean_quantity)
    df["قيمة الطلب رقم"] = col_or_default(df, cols["order_total"], "0").apply(clean_money)
    df["إجمالي المنتج رقم"] = col_or_default(df, cols["item_total"], "0").apply(clean_money)
    df["سعر الحبة رقم"] = col_or_default(df, cols["unit_price"], "0").apply(clean_money)
    df["سعر الحشوة رقم"] = col_or_default(df, cols["variety_price"], "0").apply(clean_money)
    df["الخصم"] = col_or_blank(df, cols["discount"])
    df["تاريخ التوصيل الأصلي"] = col_or_blank(df, cols["delivery_date"])
    df["وقت الاستلام الأصلي"] = col_or_blank(df, cols["pickup_time"])
    df["تاريخ ووقت الاستلام"] = [parse_datetime_parts(d, t) for d, t in zip(df["تاريخ التوصيل الأصلي"], df["وقت الاستلام الأصلي"])]
    df["تاريخ التحليل"] = pd.to_datetime(df["تاريخ ووقت الاستلام"], errors="coerce").dt.date
    df["ساعة رقم"] = pd.to_datetime(df["تاريخ ووقت الاستلام"], errors="coerce").dt.hour
    df["الساعة"] = df["ساعة رقم"].apply(hour_label)
    df["ملغي؟"] = df["الحالة"].apply(is_cancelled_status)
    df["إضافة؟"] = df["المنتج"].apply(is_addon_product)
    df["تصنيف الإضافة"] = df["المنتج"].apply(addon_category)
    df["الحملة"] = df["المنتج"].apply(classify_campaign)
    df["رقم الجوال المستخرج"] = df.apply(lambda r: extract_phone_from_text(f"{r['الملاحظة']} {r['العميل']}"), axis=1)
    df["سبب المتابعة"] = df.apply(lambda r: need_action_reasons(r["الملاحظة"], r["المنتج"]), axis=1)
    df["يحتاج متابعة؟"] = df["سبب المتابعة"].astype(str).str.len() > 0

    # Order-level table without duplicated order total.
    group = df.sort_values(["تاريخ ووقت الاستلام", "رقم الطلب الموحد"], na_position="last").groupby("رقم الطلب الموحد", dropna=False)
    order_level = group.agg(
        رقم_الطلب=("رقم الطلب الظاهر", "first"),
        العميل=("العميل", "first"),
        الفرع=("الفرع", "first"),
        الحالة=("الحالة", "first"),
        قيمة_الطلب=("قيمة الطلب رقم", "max"),
        تاريخ_التحليل=("تاريخ التحليل", "first"),
        وقت_الاستلام=("وقت الاستلام الأصلي", "first"),
        تاريخ_ووقت_الاستلام=("تاريخ ووقت الاستلام", "first"),
        ساعة_رقم=("ساعة رقم", "first"),
        الساعة=("الساعة", "first"),
        عدد_الأصناف=("المنتج", "count"),
        عدد_المنتجات=("إضافة؟", lambda s: int((~s).sum())),
        عدد_الإضافات=("إضافة؟", lambda s: int(s.sum())),
        فيه_إضافات=("إضافة؟", "max"),
        يحتاج_متابعة=("يحتاج متابعة؟", "max"),
        ملغي=("ملغي؟", "max"),
    ).reset_index()
    # Cleanup duplicate Arabic/underscore naming from agg.
    order_level = order_level.rename(columns={
        "رقم_الطلب": "رقم الطلب",
        "قيمة_الطلب": "قيمة الطلب",
        "تاريخ_التحليل": "تاريخ التحليل",
        "وقت_الاستلام": "وقت الاستلام",
        "تاريخ_ووقت_الاستلام": "تاريخ ووقت الاستلام",
        "ساعة_رقم": "ساعة رقم",
        "عدد_الأصناف": "عدد الأصناف",
        "عدد_المنتجات": "عدد المنتجات",
        "عدد_الإضافات": "عدد الإضافات",
        "فيه_إضافات": "فيه إضافات",
        "يحتاج_متابعة": "يحتاج متابعة",
    })
    if "ساعة رقم" not in order_level.columns:
        order_level["ساعة رقم"] = pd.NA
    order_level["ساعة رقم"] = pd.to_numeric(order_level["ساعة رقم"], errors="coerce")

    return df, order_level, cols


def safe_group_count(df, group_cols, value_col="رقم الطلب الموحد"):
    if df.empty:
        return pd.DataFrame()
    return df.groupby(group_cols, dropna=False)[value_col].nunique().reset_index(name="عدد الطلبات")


def build_reports(items, orders):
    active_items = items[~items["ملغي؟"]].copy()
    active_orders = orders[~orders["ملغي"].astype(bool)].copy() if "ملغي" in orders.columns else orders.copy()

    reports = {}
    reports["raw_filtered"] = items
    reports["items_active"] = active_items
    reports["orders_active"] = active_orders

    if not active_orders.empty:
        branch_sales = active_orders.groupby("الفرع", dropna=False).agg(
            عدد_الطلبات=("رقم الطلب الموحد", "nunique"),
            المبيعات=("قيمة الطلب", "sum"),
            متوسط_الطلب=("قيمة الطلب", "mean"),
            طلبات_بإضافات=("فيه إضافات", "sum"),
            تحتاج_متابعة=("يحتاج متابعة", "sum"),
        ).reset_index()
        branch_sales["نسبة Upsell %"] = (branch_sales["طلبات_بإضافات"] / branch_sales["عدد_الطلبات"].replace(0, pd.NA) * 100).fillna(0).round(1)
        reports["sales_by_branch"] = branch_sales.sort_values("المبيعات", ascending=False)

        hour_sales = active_orders.groupby("الساعة", dropna=False).agg(
            عدد_الطلبات=("رقم الطلب الموحد", "nunique"),
            المبيعات=("قيمة الطلب", "sum"),
            متوسط_الطلب=("قيمة الطلب", "mean"),
        ).reset_index()
        hour_order = active_orders.groupby("الساعة", dropna=False)["ساعة رقم"].min().reset_index(name="ساعة رقم")
        hour_sales = hour_sales.merge(hour_order, on="الساعة", how="left").sort_values("ساعة رقم", na_position="last")
        reports["sales_by_hour"] = hour_sales.drop(columns=["ساعة رقم"], errors="ignore")

        reports["status_report"] = active_orders.groupby(["الفرع", "الحالة"], dropna=False)["رقم الطلب الموحد"].nunique().reset_index(name="عدد الطلبات")

        heat = active_orders.pivot_table(index="الفرع", columns="الساعة", values="رقم الطلب الموحد", aggfunc="nunique", fill_value=0)
        hour_map = active_orders.drop_duplicates("الساعة").set_index("الساعة")["ساعة رقم"].to_dict()
        heat = heat.reindex(sorted(heat.columns, key=lambda x: hour_map.get(x, 999)), axis=1)
        reports["branch_hour_heatmap"] = heat
    else:
        reports["sales_by_branch"] = pd.DataFrame()
        reports["sales_by_hour"] = pd.DataFrame()
        reports["status_report"] = pd.DataFrame()
        reports["branch_hour_heatmap"] = pd.DataFrame()

    non_addon = active_items[~active_items["إضافة؟"]].copy()
    if not non_addon.empty:
        product_perf = non_addon.groupby("المنتج", dropna=False).agg(
            عدد_الطلبات=("رقم الطلب الموحد", "nunique"),
            الكمية=("الكمية رقم", "sum"),
            المبيعات_منتجات=("إجمالي المنتج رقم", "sum"),
            متوسط_سعر_الحبة=("سعر الحبة رقم", "mean"),
            تحتاج_متابعة=("يحتاج متابعة؟", "sum"),
        ).reset_index().sort_values(["المبيعات_منتجات", "الكمية"], ascending=False)
        reports["product_performance"] = product_perf
        reports["product_by_branch"] = non_addon.groupby(["الفرع", "المنتج"], dropna=False).agg(
            عدد_الطلبات=("رقم الطلب الموحد", "nunique"),
            الكمية=("الكمية رقم", "sum"),
            المبيعات=("إجمالي المنتج رقم", "sum"),
        ).reset_index().sort_values(["الفرع", "المبيعات"], ascending=[True, False])
        variety_source = non_addon[non_addon["الحشوة"].astype(str).str.strip().str.len() > 0].copy()

        reports["variety_report"] = variety_source.groupby("الحشوة", dropna=False).agg(
            عدد_الطلبات=("رقم الطلب الموحد", "nunique"),
            الكمية=("الكمية رقم", "sum"),
            المبيعات=("إجمالي المنتج رقم", "sum"),
            متوسط_سعر_الحبة=("سعر الحبة رقم", "mean"),
            تحتاج_متابعة=("يحتاج متابعة؟", "sum"),
        ).reset_index().sort_values(["الكمية", "المبيعات"], ascending=False)

        reports["variety_by_branch"] = variety_source.groupby(["الفرع", "الحشوة"], dropna=False).agg(
            عدد_الطلبات=("رقم الطلب الموحد", "nunique"),
            الكمية=("الكمية رقم", "sum"),
            المبيعات=("إجمالي المنتج رقم", "sum"),
        ).reset_index().sort_values(["الفرع", "الكمية"], ascending=[True, False])

        reports["variety_by_product"] = variety_source.groupby(["المنتج", "الحشوة"], dropna=False).agg(
            عدد_الطلبات=("رقم الطلب الموحد", "nunique"),
            الكمية=("الكمية رقم", "sum"),
            المبيعات=("إجمالي المنتج رقم", "sum"),
        ).reset_index().sort_values(["المنتج", "الكمية"], ascending=[True, False])

        reports["variety_by_hour"] = variety_source.groupby(["الساعة", "الحشوة"], dropna=False).agg(
            عدد_الطلبات=("رقم الطلب الموحد", "nunique"),
            الكمية=("الكمية رقم", "sum"),
            المبيعات=("إجمالي المنتج رقم", "sum"),
        ).reset_index()

        if not reports["variety_by_hour"].empty:
            hour_order_for_variety = variety_source.groupby("الساعة", dropna=False)["ساعة رقم"].min().reset_index(name="ساعة رقم")
            reports["variety_by_hour"] = reports["variety_by_hour"].merge(hour_order_for_variety, on="الساعة", how="left").sort_values(["ساعة رقم", "الكمية"], ascending=[True, False]).drop(columns=["ساعة رقم"], errors="ignore")

        reports["variety_by_campaign"] = variety_source.groupby(["الحملة", "الحشوة"], dropna=False).agg(
            عدد_الطلبات=("رقم الطلب الموحد", "nunique"),
            الكمية=("الكمية رقم", "sum"),
            المبيعات=("إجمالي المنتج رقم", "sum"),
        ).reset_index().sort_values(["الحملة", "الكمية"], ascending=[True, False])

        if not variety_source.empty:
            reports["branch_variety_heatmap"] = variety_source.pivot_table(
                index="الفرع",
                columns="الحشوة",
                values="الكمية رقم",
                aggfunc="sum",
                fill_value=0,
            )
            reports["product_variety_heatmap"] = variety_source.pivot_table(
                index="المنتج",
                columns="الحشوة",
                values="الكمية رقم",
                aggfunc="sum",
                fill_value=0,
            )
            reports["variety_order_details"] = variety_source[[c for c in [
                "رقم الطلب الظاهر", "رقم الطلب الموحد", "الفرع", "الحالة", "تاريخ التحليل",
                "وقت الاستلام الأصلي", "الساعة", "العميل", "المنتج", "الحشوة",
                "الكمية رقم", "إجمالي المنتج رقم", "سبب المتابعة", "الملاحظة"
            ] if c in variety_source.columns]].sort_values(["الحشوة", "الفرع", "وقت الاستلام الأصلي"], na_position="last")
        else:
            reports["branch_variety_heatmap"] = pd.DataFrame()
            reports["product_variety_heatmap"] = pd.DataFrame()
            reports["variety_order_details"] = pd.DataFrame()
    else:
        reports["product_performance"] = pd.DataFrame()
        reports["product_by_branch"] = pd.DataFrame()
        reports["variety_report"] = pd.DataFrame()
        reports["variety_by_branch"] = pd.DataFrame()
        reports["variety_by_product"] = pd.DataFrame()
        reports["variety_by_hour"] = pd.DataFrame()
        reports["variety_by_campaign"] = pd.DataFrame()
        reports["branch_variety_heatmap"] = pd.DataFrame()
        reports["product_variety_heatmap"] = pd.DataFrame()
        reports["variety_order_details"] = pd.DataFrame()

    addons = active_items[active_items["إضافة؟"]].copy()
    reports["addon_items"] = addons
    if not addons.empty:
        reports["addons_summary"] = addons.groupby("تصنيف الإضافة", dropna=False).agg(
            عدد_الطلبات=("رقم الطلب الموحد", "nunique"),
            الكمية=("الكمية رقم", "sum"),
            المبيعات=("إجمالي المنتج رقم", "sum"),
        ).reset_index().sort_values("الكمية", ascending=False)
        reports["addons_by_branch"] = addons.groupby(["الفرع", "تصنيف الإضافة"], dropna=False).agg(
            عدد_الطلبات=("رقم الطلب الموحد", "nunique"),
            الكمية=("الكمية رقم", "sum"),
            المبيعات=("إجمالي المنتج رقم", "sum"),
        ).reset_index().sort_values(["الفرع", "الكمية"], ascending=[True, False])
    else:
        reports["addons_summary"] = pd.DataFrame()
        reports["addons_by_branch"] = pd.DataFrame()

    need_action = active_items[active_items["يحتاج متابعة؟"]].copy()
    if not need_action.empty:
        cols = ["رقم الطلب الظاهر", "رقم الطلب الموحد", "الفرع", "الحالة", "تاريخ التوصيل الأصلي", "وقت الاستلام الأصلي", "العميل", "المنتج", "الحشوة", "الكمية رقم", "رقم الجوال المستخرج", "سبب المتابعة", "الملاحظة"]
        reports["need_action"] = need_action[[c for c in cols if c in need_action.columns]].drop_duplicates().sort_values(["الفرع", "وقت الاستلام الأصلي"])
        reason_rows = []
        for reasons in need_action["سبب المتابعة"].dropna().astype(str):
            for r in [x.strip() for x in reasons.split("،") if x.strip()]:
                reason_rows.append(r)
        if reason_rows:
            reports["need_action_reasons"] = (
                pd.Series(reason_rows)
                .value_counts()
                .rename("عدد الحالات")
                .reset_index()
                .rename(columns={"index": "سبب المتابعة"})
            )
        else:
            reports["need_action_reasons"] = pd.DataFrame(columns=["سبب المتابعة", "عدد الحالات"])
    else:
        reports["need_action"] = pd.DataFrame()
        reports["need_action_reasons"] = pd.DataFrame()

    quality_rows = []
    def add_issue(name, mask, severity):
        count = int(mask.sum()) if len(mask) else 0
        quality_rows.append({"المشكلة": name, "عدد الصفوف": count, "الأهمية": severity})

    add_issue("تاريخ توصيل ناقص", items["تاريخ التوصيل الأصلي"].astype(str).str.strip().eq(""), "عالية")
    add_issue("وقت استلام ناقص", items["وقت الاستلام الأصلي"].astype(str).str.strip().eq(""), "عالية")
    add_issue("فرع غير محدد", items["الفرع"].eq("العقيق"), "متوسطة")
    add_issue("حالة طلب ناقصة", items["الحالة"].astype(str).str.strip().isin(["", "غير محدد"]), "متوسطة")
    add_issue("منتج بدون اسم", items["المنتج"].eq("بدون اسم منتج"), "عالية")
    add_issue("حشوة ناقصة للمنتجات", (~items["إضافة؟"]) & items["الحشوة"].astype(str).str.strip().eq(""), "منخفضة")
    add_issue("قيمة طلب صفرية", items["قيمة الطلب رقم"].fillna(0).eq(0), "متوسطة")
    reports["data_quality_summary"] = pd.DataFrame(quality_rows)

    quality_detail = items[
        items["تاريخ التوصيل الأصلي"].astype(str).str.strip().eq("") |
        items["وقت الاستلام الأصلي"].astype(str).str.strip().eq("") |
        items["الفرع"].eq("العقيق") |
        items["المنتج"].eq("بدون اسم منتج") |
        items["قيمة الطلب رقم"].fillna(0).eq(0)
    ].copy()
    reports["data_quality_details"] = quality_detail[[c for c in ["رقم الطلب الظاهر", "رقم الطلب الموحد", "الفرع", "الحالة", "تاريخ التوصيل الأصلي", "وقت الاستلام الأصلي", "العميل", "المنتج", "قيمة الطلب رقم"] if c in quality_detail.columns]]

    if not active_items.empty:
        camp = active_items.groupby("الحملة", dropna=False).agg(
            عدد_الطلبات=("رقم الطلب الموحد", "nunique"),
            الكمية=("الكمية رقم", "sum"),
            مبيعات_الأصناف=("إجمالي المنتج رقم", "sum"),
            تحتاج_متابعة=("يحتاج متابعة؟", "sum"),
        ).reset_index().sort_values("عدد_الطلبات", ascending=False)
        reports["campaign_summary"] = camp
        reports["campaign_products"] = active_items.groupby(["الحملة", "المنتج"], dropna=False).agg(
            عدد_الطلبات=("رقم الطلب الموحد", "nunique"),
            الكمية=("الكمية رقم", "sum"),
            مبيعات_الأصناف=("إجمالي المنتج رقم", "sum"),
        ).reset_index().sort_values(["الحملة", "عدد_الطلبات"], ascending=[True, False])
    else:
        reports["campaign_summary"] = pd.DataFrame()
        reports["campaign_products"] = pd.DataFrame()

    if not active_orders.empty:
        reports["multi_item_orders"] = active_orders[active_orders["عدد الأصناف"] > 1].sort_values(["عدد الأصناف", "قيمة الطلب"], ascending=False)
    else:
        reports["multi_item_orders"] = pd.DataFrame()

    return reports


# =========================
# Display Helpers
# =========================

def format_int(value):
    try:
        return f"{int(round(float(value))):,}"
    except Exception:
        return "0"


def format_money(value):
    try:
        return f"{float(value):,.0f} SAR"
    except Exception:
        return "0 SAR"


def short_label(value, max_len=28):
    """Short readable label for charts while keeping full value in tables."""
    text = "" if pd.isna(value) else str(value)
    text = re.sub(r"\\s+", " ", text).strip()
    if len(text) <= max_len:
        return text
    return text[:max_len - 1].rstrip() + "…"


def wrap_label(value, width=18):
    """Wrap long Arabic/English labels for Plotly axes."""
    text = "" if pd.isna(value) else str(value)
    text = re.sub(r"\\s+", " ", text).strip()
    if len(text) <= width:
        return text
    words = text.split(" ")
    lines = []
    current = ""
    for word in words:
        if len(current) + len(word) + 1 <= width:
            current = (current + " " + word).strip()
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return "<br>".join(lines[:3]) + ("…" if len(lines) > 3 else "")


def make_readable_fig(fig, height=480, showlegend=True, legend_orientation="h"):
    """Global readability styling for dark dashboard charts."""
    fig.update_layout(
        height=height,
        font=dict(size=15, color="#f8fafc", family="Arial"),
        title=dict(font=dict(size=20, color="#f8fafc"), x=0.02, xanchor="left"),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(15,23,42,0.18)",
        margin=dict(l=70, r=35, t=90, b=105),
        showlegend=showlegend,
    )
    fig.update_xaxes(
        tickfont=dict(size=13, color="#e5e7eb"),
        title_font=dict(size=15, color="#e5e7eb"),
        gridcolor="rgba(148,163,184,0.22)",
        automargin=True,
    )
    fig.update_yaxes(
        tickfont=dict(size=13, color="#e5e7eb"),
        title_font=dict(size=15, color="#e5e7eb"),
        gridcolor="rgba(148,163,184,0.22)",
        automargin=True,
    )
    if showlegend:
        if legend_orientation == "h":
            fig.update_layout(
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=-0.33,
                    xanchor="center",
                    x=0.5,
                    font=dict(size=13, color="#f8fafc"),
                    bgcolor="rgba(15,23,42,0.65)",
                    bordercolor="rgba(148,163,184,0.25)",
                    borderwidth=1,
                    itemwidth=30,
                )
            )
        else:
            fig.update_layout(
                legend=dict(
                    orientation="v",
                    yanchor="top",
                    y=1,
                    xanchor="left",
                    x=1.02,
                    font=dict(size=13, color="#f8fafc"),
                    bgcolor="rgba(15,23,42,0.65)",
                    bordercolor="rgba(148,163,184,0.25)",
                    borderwidth=1,
                )
            )
    return fig


def render_kpi(label, value, sub="", color="#22c55e"):
    st.markdown(
        f"""
        <div class="kpi-card" style="--accent:{color};">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-sub">{sub}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def fig_layout(fig, height=420):
    fig.update_layout(
        height=height,
        template="plotly_dark",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(15,23,42,.20)",
        font=dict(family="Cairo, Arial", color="#e5e7eb", size=12),
        margin=dict(l=20, r=20, t=62, b=35),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    fig.update_xaxes(gridcolor="rgba(255,255,255,.08)", zerolinecolor="rgba(255,255,255,.08)")
    fig.update_yaxes(gridcolor="rgba(255,255,255,.08)", zerolinecolor="rgba(255,255,255,.08)")
    return fig


def chart_config():
    return {"responsive": True, "displayModeBar": False, "scrollZoom": False}


def display_df(df, height=420):
    if df is None or df.empty:
        st.info("لا توجد بيانات لهذا التقرير ضمن الفلاتر الحالية.")
        return
    st.dataframe(df, use_container_width=True, height=height)


def write_excel_sheet(writer, df, sheet_name):
    if df is None or df.empty:
        df = pd.DataFrame({"ملاحظة": ["لا توجد بيانات"]})
    safe_name = str(sheet_name)[:31]
    clean_df = df.copy()
    for col in clean_df.columns:
        if pd.api.types.is_datetime64_any_dtype(clean_df[col]):
            clean_df[col] = clean_df[col].astype(str)
        else:
            clean_df[col] = clean_df[col].apply(lambda x: str(x) if isinstance(x, (list, dict, tuple)) else x)
    clean_df.to_excel(writer, sheet_name=safe_name, index=False)
    ws = writer.book[safe_name]
    ws.freeze_panes = "A2"
    header_fill = PatternFill("solid", fgColor="111827")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="CBD5E1")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for idx, column_cells in enumerate(ws.columns, 1):
        max_len = 12
        for cell in column_cells[:200]:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)) + 2)
        ws.column_dimensions[get_column_letter(idx)].width = min(max_len, 42)
    if ws.max_row >= 2 and ws.max_column >= 1:
        ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
        tab = Table(displayName=re.sub(r"\W+", "", safe_name)[:20] + "Tbl", ref=ref)
        style = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        tab.tableStyleInfo = style
        try:
            ws.add_table(tab)
        except Exception:
            pass


def build_excel_export(reports, filters_summary):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        write_excel_sheet(writer, filters_summary, "Executive Summary")
        sheet_map = {
            "orders_active": "Daily Operations",
            "sales_by_branch": "Sales by Branch",
            "sales_by_hour": "Sales by Hour",
            "status_report": "Status Report",
            "product_performance": "Product Performance",
            "product_by_branch": "Products by Branch",
            "variety_report": "Varieties Summary",
            "variety_by_branch": "Varieties by Branch",
            "variety_by_product": "Varieties by Product",
            "variety_by_hour": "Varieties by Hour",
            "variety_by_campaign": "Varieties by Campaign",
            "variety_order_details": "Variety Order Details",
            "addons_summary": "Addons Summary",
            "addons_by_branch": "Addons by Branch",
            "need_action": "Need Action",
            "need_action_reasons": "Action Reasons",
            "data_quality_summary": "Data Quality",
            "data_quality_details": "Quality Details",
            "campaign_summary": "Campaigns",
            "campaign_products": "Campaign Products",
            "multi_item_orders": "Multi Item Orders",
            "raw_filtered": "Filtered Raw Data",
        }
        for key, sheet in sheet_map.items():
            write_excel_sheet(writer, reports.get(key, pd.DataFrame()), sheet)
    output.seek(0)
    return output


def branch_prep_excel(branch_df, branch_name):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        write_excel_sheet(writer, branch_df, f"Prep {branch_name}"[:31])
    output.seek(0)
    return output


# =========================
# Header
# =========================
st.markdown(
    f"""
    <div class="hero">
        <h1>🧁 MAD Orders Control Center</h1>
        <p>{APP_VERSION} — لوحة تشغيل يومية للمبيعات، الفروع، تجهيز الطلبات، الإضافات، الحملات، جودة البيانات، والطلبات التي تحتاج متابعة.</p>
    </div>
    """,
    unsafe_allow_html=True,
)


# =========================
# Sidebar - Data Source
# =========================
st.sidebar.markdown("## ⚙️ مصدر البيانات")
source_type = st.sidebar.radio("اختر المصدر", ["Google Sheet", "رفع ملف"], index=0, horizontal=True)

raw_df = None
load_error = None

if source_type == "Google Sheet":
    sheet_url = st.sidebar.text_input("رابط Google Sheet", value=DEFAULT_GOOGLE_SHEET_URL, placeholder="https://docs.google.com/spreadsheets/d/...")
    gid = st.sidebar.text_input("Sheet GID", value="0", help="لو الشيت تبويب مختلف، انسخ رقم gid من الرابط. غالباً أول تبويب = 0")
    c_refresh, c_status = st.sidebar.columns([1, 1])
    with c_refresh:
        if st.button("🔄 تحديث البيانات"):
            st.cache_data.clear()
    try:
        with st.spinner("جاري تحميل Google Sheet..."):
            raw_df = load_google_sheet(sheet_url, gid)
        with c_status:
            st.success("تم")
    except Exception as e:
        load_error = str(e)
else:
    uploaded = st.sidebar.file_uploader("ارفع ملف TXT / CSV / Excel", type=["txt", "csv", "xlsx", "xls"])
    if uploaded is not None:
        try:
            raw_df = read_uploaded_file(uploaded)
        except Exception as e:
            load_error = str(e)

if load_error:
    st.error("لم أستطع تحميل البيانات. تأكد أن Google Sheet متاح لأي شخص لديه الرابط Viewer أو ارفع ملف مباشرة.")
    with st.expander("تفاصيل الخطأ"):
        st.code(load_error)
    st.stop()

if raw_df is None or raw_df.empty:
    st.markdown('<div class="note-box">اربط Google Sheet أو ارفع ملف الطلبات لبدء التحليل.</div>', unsafe_allow_html=True)
    st.stop()


# =========================
# Prepare Data
# =========================
try:
    items_all, orders_all, cols = prepare_data(raw_df)
except Exception as e:
    st.error("حدث خطأ أثناء تجهيز البيانات. تأكد من أسماء الأعمدة وترتيب الملف.")
    with st.expander("تفاصيل الخطأ"):
        st.exception(e)
    st.stop()


# =========================
# Sidebar - Filters
# =========================
st.sidebar.markdown("## 🔎 فلاتر التحليل")

available_dates = sorted([d for d in items_all["تاريخ التحليل"].dropna().unique()])
if available_dates:
    min_date, max_date = min(available_dates), max(available_dates)
    date_range = st.sidebar.date_input("فترة تاريخ التوصيل", value=(min_date, max_date), min_value=min_date, max_value=max_date)
else:
    date_range = None

branches = sorted(items_all["الفرع"].dropna().unique().tolist())
selected_branches = st.sidebar.multiselect("الفروع", branches, default=branches)

statuses = sorted(items_all["الحالة"].dropna().unique().tolist())
selected_statuses = st.sidebar.multiselect("الحالات", statuses, default=statuses)

varieties = sorted([v for v in items_all["الحشوة"].dropna().unique().tolist() if str(v).strip()])
selected_varieties = st.sidebar.multiselect("الحشوات", varieties, default=[])

campaigns = sorted(items_all["الحملة"].dropna().unique().tolist())
selected_campaigns = st.sidebar.multiselect("الحملات", campaigns, default=[])

search_text = st.sidebar.text_input("بحث سريع", placeholder="رقم طلب / عميل / منتج / جوال")
include_cancelled = st.sidebar.checkbox("إظهار الملغي ضمن التقارير", value=False)

filtered = items_all.copy()
if available_dates and date_range:
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range
    else:
        start_date = end_date = date_range
    filtered = filtered[(filtered["تاريخ التحليل"].isna()) | ((filtered["تاريخ التحليل"] >= start_date) & (filtered["تاريخ التحليل"] <= end_date))]

if selected_branches:
    filtered = filtered[filtered["الفرع"].isin(selected_branches)]
if selected_statuses:
    filtered = filtered[filtered["الحالة"].isin(selected_statuses)]
if selected_varieties:
    filtered = filtered[filtered["الحشوة"].isin(selected_varieties)]
if selected_campaigns:
    filtered = filtered[filtered["الحملة"].isin(selected_campaigns)]
if not include_cancelled:
    filtered = filtered[~filtered["ملغي؟"]]

if search_text.strip():
    q = search_text.strip().lower()
    search_blob = (
        filtered["رقم الطلب الموحد"].astype(str) + " " +
        filtered["رقم الطلب الظاهر"].astype(str) + " " +
        filtered["العميل"].astype(str) + " " +
        filtered["المنتج"].astype(str) + " " +
        filtered["رقم الجوال المستخرج"].astype(str) + " " +
        filtered["الملاحظة"].astype(str)
    ).str.lower()
    filtered = filtered[search_blob.str.contains(re.escape(q), na=False)]

_, orders_filtered, _ = prepare_data(filtered.drop(columns=[c for c in []], errors="ignore"))
reports = build_reports(filtered, orders_filtered)
active_items = reports["items_active"]
active_orders = reports["orders_active"]

# Core KPIs
total_rows = len(filtered)
total_orders = int(active_orders["رقم الطلب الموحد"].nunique()) if not active_orders.empty else 0
total_sales = float(active_orders["قيمة الطلب"].fillna(0).sum()) if not active_orders.empty else 0.0
avg_order = float(active_orders["قيمة الطلب"].fillna(0).mean()) if not active_orders.empty else 0.0
need_action_count = int(active_orders["يحتاج متابعة"].sum()) if not active_orders.empty and "يحتاج متابعة" in active_orders.columns else 0
missing_date_count = int(filtered["تاريخ التوصيل الأصلي"].astype(str).str.strip().eq("").sum())
missing_time_count = int(filtered["وقت الاستلام الأصلي"].astype(str).str.strip().eq("").sum())
addon_orders = int(active_orders["فيه إضافات"].sum()) if not active_orders.empty and "فيه إضافات" in active_orders.columns else 0
upsell_rate = (addon_orders / total_orders * 100) if total_orders else 0

# Smart insights
top_branch = "-"
top_branch_count = 0
if not active_orders.empty:
    tb = active_orders.groupby("الفرع")["رقم الطلب الموحد"].nunique().sort_values(ascending=False)
    if len(tb):
        top_branch, top_branch_count = tb.index[0], int(tb.iloc[0])

top_hour = "-"
top_hour_count = 0
if not active_orders.empty:
    th = active_orders.groupby("الساعة")["رقم الطلب الموحد"].nunique().sort_values(ascending=False)
    if len(th):
        top_hour, top_hour_count = th.index[0], int(th.iloc[0])

st.sidebar.success("تم تحليل البيانات")
st.sidebar.write(f"الصفوف: **{format_int(total_rows)}**")
st.sidebar.write(f"الطلبات: **{format_int(total_orders)}**")
st.sidebar.write(f"المبيعات: **{format_money(total_sales)}**")
st.sidebar.write(f"تحتاج متابعة: **{format_int(need_action_count)}**")


# =========================
# Main KPIs
# =========================
k1, k2, k3, k4, k5 = st.columns(5)
with k1:
    render_kpi("عدد الطلبات", format_int(total_orders), "Unique Orders", "#2563eb")
with k2:
    render_kpi("إجمالي المبيعات", format_money(total_sales), "بدون تكرار قيمة الطلب", "#16a34a")
with k3:
    render_kpi("متوسط الطلب", format_money(avg_order), "AOV", "#0891b2")
with k4:
    render_kpi("تحتاج متابعة", format_int(need_action_count), "صورة / تواصل / كتابة", "#dc2626")
with k5:
    render_kpi("Upsell", f"{upsell_rate:.1f}%", f"{format_int(addon_orders)} طلب بإضافات", "#f59e0b")


# =========================
# Smart Alerts
# =========================
st.markdown('<div class="section-title">🚦 تنبيهات ذكية</div>', unsafe_allow_html=True)
alerts = []
if need_action_count > 0:
    alerts.append(f"⚠️ يوجد {format_int(need_action_count)} طلب يحتاج متابعة مع العميل.")
if missing_date_count > 0:
    alerts.append(f"⚠️ يوجد {format_int(missing_date_count)} صف بدون تاريخ توصيل.")
if missing_time_count > 0:
    alerts.append(f"⚠️ يوجد {format_int(missing_time_count)} صف بدون وقت استلام.")
if top_hour != "-":
    alerts.append(f"🔥 أعلى ساعة ضغط: {top_hour} بعدد {format_int(top_hour_count)} طلب.")
if top_branch != "-":
    alerts.append(f"🏬 أعلى فرع طلبات: {top_branch} بعدد {format_int(top_branch_count)} طلب.")
if not alerts:
    st.markdown('<div class="good-box">✅ لا توجد تنبيهات حرجة ضمن الفلاتر الحالية.</div>', unsafe_allow_html=True)
else:
    for alert in alerts[:6]:
        st.markdown(f'<div class="alert-box">{alert}</div>', unsafe_allow_html=True)


# =========================
# Tabs
# =========================
(
    tab_daily,
    tab_prep,
    tab_sales,
    tab_products,
    tab_varieties,
    tab_addons,
    tab_actions,
    tab_campaigns,
    tab_branch,
    tab_product,
    tab_quality,
    tab_export,
) = st.tabs([
    "📍 Daily Ops",
    "🧾 Branch Prep",
    "💰 Sales",
    "🧁 Products",
    "🍰 Fillings",
    "🎈 Add-ons",
    "🚨 Need Action",
    "🎯 Campaigns",
    "🏬 Branch Deep Dive",
    "🔍 Product Deep Dive",
    "🧹 Data Quality",
    "⬇️ Export",
])


with tab_daily:
    st.markdown('<div class="section-title">📍 Daily Operations</div>', unsafe_allow_html=True)
    c1, c2 = st.columns([1.1, 1])
    with c1:
        sales_branch = reports.get("sales_by_branch", pd.DataFrame())
        if not sales_branch.empty:
            fig = px.bar(sales_branch.sort_values("عدد_الطلبات"), x="عدد_الطلبات", y="الفرع", orientation="h", text="عدد_الطلبات", title="ضغط الطلبات حسب الفرع", color="عدد_الطلبات", color_continuous_scale="Viridis")
            fig.update_traces(textposition="outside")
            st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=chart_config())
        else:
            st.info("لا يوجد بيانات فروع.")
    with c2:
        sales_hour = reports.get("sales_by_hour", pd.DataFrame())
        if not sales_hour.empty:
            fig = px.line(sales_hour, x="الساعة", y="عدد_الطلبات", markers=True, title="ضغط الطلبات حسب الساعة")
            fig.update_traces(line=dict(width=4), marker=dict(size=10))
            st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=chart_config())
        else:
            st.info("لا يوجد بيانات ساعات.")

    heat = reports.get("branch_hour_heatmap", pd.DataFrame())
    if not heat.empty:
        fig = px.imshow(heat, text_auto=True, aspect="auto", title="Heatmap الفرع × الساعة", color_continuous_scale="YlGnBu")
        fig.update_xaxes(side="top")
        st.plotly_chart(fig_layout(fig, 520), use_container_width=True, config=chart_config())

    st.markdown('<div class="mini-title">أقرب جدول تشغيل ضمن الفلاتر</div>', unsafe_allow_html=True)
    ops_cols = ["رقم الطلب", "العميل", "الفرع", "الحالة", "تاريخ التحليل", "وقت الاستلام", "الساعة", "قيمة الطلب", "عدد الأصناف", "عدد الإضافات", "يحتاج متابعة"]
    ops_view_cols = [c for c in ops_cols if c in active_orders.columns]
    ops_sort_cols = [c for c in ["تاريخ ووقت الاستلام", "الفرع"] if c in active_orders.columns]

    if not active_orders.empty:
        ops_view = active_orders.copy()
        if ops_sort_cols:
            ops_view = ops_view.sort_values(ops_sort_cols, na_position="last")
        ops_view = ops_view[ops_view_cols] if ops_view_cols else ops_view
    else:
        ops_view = pd.DataFrame(columns=ops_view_cols)

    display_df(ops_view, 420)


with tab_prep:
    st.markdown('<div class="section-title">🧾 Branch Prep Sheet</div>', unsafe_allow_html=True)
    prep_branches = sorted(active_items["الفرع"].dropna().unique().tolist()) if not active_items.empty else []
    selected_prep_branch = st.selectbox("اختار الفرع", prep_branches if prep_branches else ["-"])
    prep = active_items[active_items["الفرع"].eq(selected_prep_branch)].copy() if selected_prep_branch != "-" else pd.DataFrame()
    prep_cols = ["وقت الاستلام الأصلي", "رقم الطلب الظاهر", "الحالة", "العميل", "المنتج", "الحشوة", "الكمية رقم", "إجمالي المنتج رقم", "سبب المتابعة", "الملاحظة"]
    prep_view = prep[[c for c in prep_cols if c in prep.columns]].sort_values(["وقت الاستلام الأصلي", "رقم الطلب الظاهر"], na_position="last") if not prep.empty else pd.DataFrame()
    display_df(prep_view, 560)
    if not prep_view.empty:
        st.download_button(
            "⬇️ تحميل تقرير تجهيز الفرع Excel",
            data=branch_prep_excel(prep_view, selected_prep_branch),
            file_name=f"branch_prep_{selected_prep_branch}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


with tab_sales:
    st.markdown('<div class="section-title">💰 Sales Analytics</div>', unsafe_allow_html=True)
    b1, b2 = st.columns(2)
    with b1:
        dfb = reports.get("sales_by_branch", pd.DataFrame())
        if not dfb.empty:
            fig = px.bar(dfb, x="الفرع", y="المبيعات", text="المبيعات", title="المبيعات حسب الفرع", color="المبيعات", color_continuous_scale="Tealgrn")
            st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=chart_config())
        display_df(dfb, 330)
    with b2:
        dfh = reports.get("sales_by_hour", pd.DataFrame())
        if not dfh.empty:
            fig = px.bar(dfh, x="الساعة", y="المبيعات", text="المبيعات", title="المبيعات حسب الساعة", color="المبيعات", color_continuous_scale="Blues")
            st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=chart_config())
        display_df(dfh, 330)

    st.markdown('<div class="mini-title">الحالات حسب الفرع</div>', unsafe_allow_html=True)
    status_df = reports.get("status_report", pd.DataFrame())
    if not status_df.empty:
        fig = px.bar(status_df, x="الفرع", y="عدد الطلبات", color="الحالة", barmode="group", title="حالات الطلبات حسب الفرع")
        st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=chart_config())
    display_df(status_df, 330)


with tab_products:
    st.markdown('<div class="section-title">🧁 Product Performance</div>', unsafe_allow_html=True)
    prod = reports.get("product_performance", pd.DataFrame())
    if not prod.empty:
        top_prod = prod.head(15)
        fig = px.bar(top_prod.sort_values("الكمية"), x="الكمية", y="المنتج", orientation="h", text="الكمية", title="أعلى المنتجات حسب الكمية", color="الكمية", color_continuous_scale="Agsunset")
        st.plotly_chart(fig_layout(fig, 560), use_container_width=True, config=chart_config())
    display_df(prod, 500)

    st.markdown('<div class="mini-title">المنتجات حسب الفرع</div>', unsafe_allow_html=True)
    display_df(reports.get("product_by_branch", pd.DataFrame()), 420)

    st.markdown('<div class="note-box">تم نقل تقارير الحشوات إلى تبويب مستقل باسم 🍰 Fillings حتى تكون واضحة ومفصلة.</div>', unsafe_allow_html=True)


with tab_varieties:
    st.markdown('<div class="section-title">🍰 Fillings / الحشوات</div>', unsafe_allow_html=True)

    variety = reports.get("variety_report", pd.DataFrame())
    variety_by_branch = reports.get("variety_by_branch", pd.DataFrame())
    variety_by_product = reports.get("variety_by_product", pd.DataFrame())
    variety_by_hour = reports.get("variety_by_hour", pd.DataFrame())
    variety_by_campaign = reports.get("variety_by_campaign", pd.DataFrame())
    branch_variety_heatmap = reports.get("branch_variety_heatmap", pd.DataFrame())
    product_variety_heatmap = reports.get("product_variety_heatmap", pd.DataFrame())
    variety_order_details = reports.get("variety_order_details", pd.DataFrame())

    fv1, fv2, fv3, fv4 = st.columns(4)
    if not variety.empty:
        top_variety_name = str(variety.iloc[0]["الحشوة"])
        top_variety_qty = float(variety.iloc[0]["الكمية"])
        total_variety_qty = float(variety["الكمية"].fillna(0).sum())
        total_variety_sales = float(variety["المبيعات"].fillna(0).sum())
        unique_varieties = int(variety["الحشوة"].nunique())
    else:
        top_variety_name = "-"
        top_variety_qty = 0
        total_variety_qty = 0
        total_variety_sales = 0
        unique_varieties = 0

    with fv1:
        render_kpi("عدد الحشوات", format_int(unique_varieties), "ضمن الفلاتر", "#7c3aed")
    with fv2:
        render_kpi("إجمالي كمية الحشوات", format_int(total_variety_qty), "Quantity", "#16a34a")
    with fv3:
        render_kpi("مبيعات منتجات لها حشوة", format_money(total_variety_sales), "Item Total", "#0891b2")
    with fv4:
        render_kpi("أعلى حشوة", top_variety_name, f"كمية: {format_int(top_variety_qty)}", "#f59e0b")

    cv1, cv2 = st.columns([1, 1])
    with cv1:
        if not variety.empty:
            top_v = variety.head(12).copy()
            top_v["الحشوة المختصرة"] = top_v["الحشوة"].apply(lambda x: short_label(x, 34))
            fig = px.bar(
                top_v.sort_values("الكمية"),
                x="الكمية",
                y="الحشوة المختصرة",
                orientation="h",
                text="الكمية",
                title="أكثر الحشوات حسب الكمية",
                color="الكمية",
                color_continuous_scale="Magma",
                hover_data={"الحشوة": True, "الكمية": ":,.0f", "المبيعات": ":,.0f", "الحشوة المختصرة": False},
            )
            fig.update_layout(yaxis_title="الحشوة", xaxis_title="الكمية")
            fig.update_traces(textposition="outside", textfont_size=13, cliponaxis=False)
            st.plotly_chart(make_readable_fig(fig, 500, showlegend=False), use_container_width=True, config=chart_config())
        else:
            st.info("لا توجد بيانات حشوات ضمن الفلاتر الحالية.")

    with cv2:
        if not variety.empty:
            pie_v = variety.copy()
            top_pie_names = pie_v.sort_values("الكمية", ascending=False).head(7)["الحشوة"].tolist()
            pie_v["الحشوة للعرض"] = pie_v["الحشوة"].apply(lambda x: short_label(x, 22) if x in top_pie_names else "Other")
            pie_v = pie_v.groupby("الحشوة للعرض", dropna=False)["الكمية"].sum().reset_index()
            fig = px.pie(pie_v, names="الحشوة للعرض", values="الكمية", hole=.55, title="Mix الحشوات")
            fig.update_traces(textposition="inside", textinfo="percent+label", textfont_size=14)
            st.plotly_chart(make_readable_fig(fig, 500, showlegend=True, legend_orientation="h"), use_container_width=True, config=chart_config())
        else:
            st.info("لا توجد بيانات حشوات.")

    st.markdown('<div class="mini-title">ملخص الحشوات</div>', unsafe_allow_html=True)
    display_df(variety, 340)

    st.markdown('<div class="mini-title">الحشوات حسب الفرع</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="readability-note">تم تحسين هذا الرسم للقراءة: نعرض أعلى 6 حشوات فقط في الرسم، وباقي التفاصيل كاملة موجودة في الجدول وملف Excel.</div>',
        unsafe_allow_html=True,
    )

    if not variety_by_branch.empty:
        top_fillings_for_branch_chart = (
            variety_by_branch.groupby("الحشوة", dropna=False)["الكمية"]
            .sum()
            .sort_values(ascending=False)
            .head(6)
            .index
            .tolist()
        )
        vbb_chart = variety_by_branch[variety_by_branch["الحشوة"].isin(top_fillings_for_branch_chart)].copy()
        vbb_chart["الحشوة المختصرة"] = vbb_chart["الحشوة"].apply(lambda x: short_label(x, 22))
        vbb_chart["الفرع المختصر"] = vbb_chart["الفرع"].apply(lambda x: wrap_label(x, 18))

        fig = px.bar(
            vbb_chart,
            x="الفرع المختصر",
            y="الكمية",
            color="الحشوة المختصرة",
            title="توزيع أعلى الحشوات على الفروع",
            barmode="stack",
            text="الكمية",
            hover_data={
                "الفرع": True,
                "الحشوة": True,
                "الكمية": ":,.0f",
                "المبيعات": ":,.0f",
                "الفرع المختصر": False,
                "الحشوة المختصرة": False,
            },
        )
        fig.update_traces(textposition="inside", textfont_size=13, cliponaxis=False)
        fig.update_layout(xaxis_title="الفرع", yaxis_title="الكمية")
        st.plotly_chart(make_readable_fig(fig, 560, showlegend=True, legend_orientation="h"), use_container_width=True, config=chart_config())
    display_df(variety_by_branch, 420)

    if not branch_variety_heatmap.empty:
        st.markdown('<div class="mini-title">Heatmap الفرع × الحشوة</div>', unsafe_allow_html=True)
        heat = branch_variety_heatmap.copy()
        top_cols = heat.sum(axis=0).sort_values(ascending=False).head(8).index
        heat = heat[top_cols]
        heat.index = [wrap_label(x, 18) for x in heat.index]
        heat.columns = [short_label(x, 18) for x in heat.columns]

        fig = px.imshow(
            heat,
            text_auto=True,
            aspect="auto",
            title="كمية أعلى الحشوات حسب الفروع",
            color_continuous_scale="YlGnBu",
        )
        fig.update_xaxes(side="top", tickangle=0)
        st.plotly_chart(make_readable_fig(fig, 560, showlegend=False), use_container_width=True, config=chart_config())

    st.markdown('<div class="mini-title">الحشوات حسب المنتج</div>', unsafe_allow_html=True)
    display_df(variety_by_product, 520)

    if not product_variety_heatmap.empty:
        top_products_for_heatmap = product_variety_heatmap.sum(axis=1).sort_values(ascending=False).head(15).index
        pv_heat = product_variety_heatmap.loc[top_products_for_heatmap]
        top_variety_cols = pv_heat.sum(axis=0).sort_values(ascending=False).head(8).index
        pv_heat = pv_heat[top_variety_cols]
        pv_heat.index = [wrap_label(x, 26) for x in pv_heat.index]
        pv_heat.columns = [short_label(x, 18) for x in pv_heat.columns]
        st.markdown('<div class="mini-title">Heatmap المنتج × الحشوة — أعلى المنتجات</div>', unsafe_allow_html=True)
        fig = px.imshow(pv_heat, text_auto=True, aspect="auto", title="توزيع الحشوات حسب المنتجات", color_continuous_scale="Teal")
        fig.update_xaxes(side="top", tickangle=0)
        st.plotly_chart(make_readable_fig(fig, 700, showlegend=False), use_container_width=True, config=chart_config())

    st.markdown('<div class="mini-title">الحشوات حسب الساعة</div>', unsafe_allow_html=True)
    if not variety_by_hour.empty:
        top_hour_fillings = (
            variety_by_hour.groupby("الحشوة", dropna=False)["الكمية"]
            .sum()
            .sort_values(ascending=False)
            .head(6)
            .index
            .tolist()
        )
        vbh_chart = variety_by_hour[variety_by_hour["الحشوة"].isin(top_hour_fillings)].copy()
        vbh_chart["الحشوة المختصرة"] = vbh_chart["الحشوة"].apply(lambda x: short_label(x, 22))
        fig = px.line(
            vbh_chart,
            x="الساعة",
            y="الكمية",
            color="الحشوة المختصرة",
            markers=True,
            title="طلب أعلى الحشوات حسب وقت الاستلام",
            hover_data={"الحشوة": True, "الكمية": ":,.0f", "الحشوة المختصرة": False},
        )
        fig.update_traces(line_width=3, marker_size=8)
        fig.update_layout(xaxis_title="الساعة", yaxis_title="الكمية")
        st.plotly_chart(make_readable_fig(fig, 500, showlegend=True, legend_orientation="h"), use_container_width=True, config=chart_config())
    display_df(variety_by_hour, 420)

    st.markdown('<div class="mini-title">الحشوات حسب الحملة</div>', unsafe_allow_html=True)
    display_df(variety_by_campaign, 420)

    st.markdown('<div class="mini-title">تفاصيل طلبات الحشوات</div>', unsafe_allow_html=True)
    display_df(variety_order_details, 560)


with tab_addons:
    st.markdown('<div class="section-title">🎈 Add-ons & Upsell</div>', unsafe_allow_html=True)
    a1, a2, a3 = st.columns(3)
    with a1:
        render_kpi("طلبات بإضافات", format_int(addon_orders), "Orders with add-ons", "#f97316")
    with a2:
        render_kpi("نسبة Upsell", f"{upsell_rate:.1f}%", "من إجمالي الطلبات", "#f59e0b")
    with a3:
        addons_sales = float(reports.get("addon_items", pd.DataFrame()).get("إجمالي المنتج رقم", pd.Series(dtype=float)).sum()) if not reports.get("addon_items", pd.DataFrame()).empty else 0
        render_kpi("مبيعات الإضافات", format_money(addons_sales), "حسب Item Total", "#22c55e")

    addons_summary = reports.get("addons_summary", pd.DataFrame())
    if not addons_summary.empty:
        fig = px.bar(addons_summary, x="تصنيف الإضافة", y="الكمية", text="الكمية", title="أكثر الإضافات مبيعًا", color="الكمية", color_continuous_scale="Oranges")
        st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=chart_config())
    display_df(addons_summary, 320)
    st.markdown('<div class="mini-title">الإضافات حسب الفرع</div>', unsafe_allow_html=True)
    display_df(reports.get("addons_by_branch", pd.DataFrame()), 420)


with tab_actions:
    st.markdown('<div class="section-title">🚨 Orders Need Action</div>', unsafe_allow_html=True)
    reasons = reports.get("need_action_reasons", pd.DataFrame())
    if not reasons.empty:
        fig = px.bar(reasons, x="سبب المتابعة", y="عدد الحالات", text="عدد الحالات", title="أسباب المتابعة", color="عدد الحالات", color_continuous_scale="Reds")
        st.plotly_chart(fig_layout(fig, 390), use_container_width=True, config=chart_config())
    display_df(reports.get("need_action", pd.DataFrame()), 620)


with tab_campaigns:
    st.markdown('<div class="section-title">🎯 Campaign Analyzer</div>', unsafe_allow_html=True)
    camp = reports.get("campaign_summary", pd.DataFrame())
    if not camp.empty:
        fig = px.bar(camp, x="الحملة", y="عدد_الطلبات", text="عدد_الطلبات", title="أداء الحملات حسب عدد الطلبات", color="عدد_الطلبات", color_continuous_scale="Purples")
        st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=chart_config())
    display_df(camp, 330)
    st.markdown('<div class="mini-title">منتجات كل حملة</div>', unsafe_allow_html=True)
    display_df(reports.get("campaign_products", pd.DataFrame()), 500)


with tab_branch:
    st.markdown('<div class="section-title">🏬 Branch Deep Dive</div>', unsafe_allow_html=True)
    bd_branches = sorted(active_items["الفرع"].dropna().unique().tolist()) if not active_items.empty else []
    bd_branch = st.selectbox("اختار فرع للتحليل العميق", bd_branches if bd_branches else ["-"], key="bd_branch")
    b_items = active_items[active_items["الفرع"].eq(bd_branch)].copy() if bd_branch != "-" else pd.DataFrame()
    b_orders = active_orders[active_orders["الفرع"].eq(bd_branch)].copy() if bd_branch != "-" else pd.DataFrame()
    if not b_orders.empty:
        c1, c2, c3, c4 = st.columns(4)
        with c1: render_kpi("طلبات الفرع", format_int(b_orders["رقم الطلب الموحد"].nunique()), bd_branch, "#2563eb")
        with c2: render_kpi("مبيعات الفرع", format_money(b_orders["قيمة الطلب"].sum()), "", "#16a34a")
        with c3: render_kpi("متوسط الطلب", format_money(b_orders["قيمة الطلب"].mean()), "", "#0891b2")
        with c4: render_kpi("تحتاج متابعة", format_int(b_orders["يحتاج متابعة"].sum()), "", "#dc2626")
        col1, col2 = st.columns(2)
        with col1:
            bp = b_items[~b_items["إضافة؟"]].groupby("المنتج").agg(الكمية=("الكمية رقم", "sum"), الطلبات=("رقم الطلب الموحد", "nunique")).reset_index().sort_values("الكمية", ascending=False).head(12)
            if not bp.empty:
                fig = px.bar(bp.sort_values("الكمية"), x="الكمية", y="المنتج", orientation="h", title="أفضل منتجات الفرع")
                st.plotly_chart(fig_layout(fig, 460), use_container_width=True, config=chart_config())
        with col2:
            bh = b_orders.groupby("الساعة")["رقم الطلب الموحد"].nunique().reset_index(name="عدد الطلبات")
            if not bh.empty:
                fig = px.line(bh, x="الساعة", y="عدد الطلبات", markers=True, title="ضغط الفرع حسب الساعة")
                st.plotly_chart(fig_layout(fig, 460), use_container_width=True, config=chart_config())
        display_df(b_items[[c for c in ["رقم الطلب الظاهر", "الحالة", "وقت الاستلام الأصلي", "العميل", "المنتج", "الحشوة", "الكمية رقم", "سبب المتابعة", "الملاحظة"] if c in b_items.columns]], 500)
    else:
        st.info("لا توجد بيانات لهذا الفرع ضمن الفلاتر.")


with tab_product:
    st.markdown('<div class="section-title">🔍 Product Deep Dive</div>', unsafe_allow_html=True)
    product_list = sorted(active_items[~active_items["إضافة؟"]]["المنتج"].dropna().unique().tolist()) if not active_items.empty else []
    selected_product = st.selectbox("اختار المنتج", product_list if product_list else ["-"])
    p_items = active_items[active_items["المنتج"].eq(selected_product)].copy() if selected_product != "-" else pd.DataFrame()
    if not p_items.empty:
        p_order_ids = p_items["رقم الطلب الموحد"].unique().tolist()
        p_orders = active_orders[active_orders["رقم الطلب الموحد"].isin(p_order_ids)].copy()
        c1, c2, c3, c4 = st.columns(4)
        with c1: render_kpi("طلبات المنتج", format_int(p_items["رقم الطلب الموحد"].nunique()), "", "#2563eb")
        with c2: render_kpi("كمية المنتج", format_int(p_items["الكمية رقم"].sum()), "", "#16a34a")
        with c3: render_kpi("مبيعات المنتج", format_money(p_items["إجمالي المنتج رقم"].sum()), "Item Total", "#0891b2")
        with c4: render_kpi("طلبات بإضافات", format_int(p_orders["فيه إضافات"].sum()) if not p_orders.empty else "0", "", "#f59e0b")
        col1, col2 = st.columns(2)
        with col1:
            pb = p_items.groupby("الفرع")["رقم الطلب الموحد"].nunique().reset_index(name="عدد الطلبات").sort_values("عدد الطلبات", ascending=False)
            if not pb.empty:
                fig = px.bar(pb, x="الفرع", y="عدد الطلبات", text="عدد الطلبات", title="المنتج حسب الفرع")
                st.plotly_chart(fig_layout(fig, 420), use_container_width=True, config=chart_config())
        with col2:
            pv = p_items[p_items["الحشوة"].astype(str).str.len() > 0].groupby("الحشوة")["الكمية رقم"].sum().reset_index(name="الكمية").sort_values("الكمية", ascending=False)
            if not pv.empty:
                fig = px.pie(pv, names="الحشوة", values="الكمية", hole=.52, title="حشوات المنتج")
                st.plotly_chart(fig_layout(fig, 420), use_container_width=True, config=chart_config())
        st.markdown('<div class="mini-title">كل طلبات المنتج</div>', unsafe_allow_html=True)
        display_df(p_items[[c for c in ["رقم الطلب الظاهر", "الفرع", "الحالة", "وقت الاستلام الأصلي", "العميل", "الحشوة", "الكمية رقم", "إجمالي المنتج رقم", "سبب المتابعة", "الملاحظة"] if c in p_items.columns]], 520)
    else:
        st.info("لا توجد بيانات لهذا المنتج ضمن الفلاتر.")


with tab_quality:
    st.markdown('<div class="section-title">🧹 Data Quality</div>', unsafe_allow_html=True)
    qsum = reports.get("data_quality_summary", pd.DataFrame())
    if not qsum.empty:
        fig = px.bar(qsum, x="المشكلة", y="عدد الصفوف", color="الأهمية", text="عدد الصفوف", title="مشاكل جودة البيانات")
        st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=chart_config())
    display_df(qsum, 300)
    st.markdown('<div class="mini-title">تفاصيل الصفوف التي تحتاج تنظيف</div>', unsafe_allow_html=True)
    display_df(reports.get("data_quality_details", pd.DataFrame()), 520)


with tab_export:
    st.markdown('<div class="section-title">⬇️ Export Center</div>', unsafe_allow_html=True)
    summary_rows = [
        {"البند": "الإصدار", "القيمة": APP_VERSION},
        {"البند": "عدد الصفوف بعد الفلاتر", "القيمة": total_rows},
        {"البند": "عدد الطلبات", "القيمة": total_orders},
        {"البند": "إجمالي المبيعات", "القيمة": total_sales},
        {"البند": "متوسط الطلب", "القيمة": avg_order},
        {"البند": "طلبات تحتاج متابعة", "القيمة": need_action_count},
        {"البند": "طلبات بإضافات", "القيمة": addon_orders},
        {"البند": "نسبة Upsell", "القيمة": round(upsell_rate, 1)},
        {"البند": "أعلى فرع", "القيمة": top_branch},
        {"البند": "أعلى ساعة", "القيمة": top_hour},
    ]
    filters_summary = pd.DataFrame(summary_rows)
    display_df(filters_summary, 280)
    excel_file = build_excel_export(reports, filters_summary)
    st.download_button(
        "⬇️ تحميل Excel شامل كل التقارير V7.3",
        data=excel_file,
        file_name="MAD_Orders_Control_Center_V7_2.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.markdown(
        '<div class="note-box">الملف يحتوي على Executive Summary، Daily Operations، Sales، Products، Fillings، Add-ons، Need Action، Campaigns، Data Quality، Multi Item Orders، والبيانات المفلترة.</div>',
        unsafe_allow_html=True,
    )
