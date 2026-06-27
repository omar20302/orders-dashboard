import re
import html
from io import BytesIO
from urllib.parse import urlparse, parse_qs
from datetime import datetime, date
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import plotly.express as px
import plotly.graph_objects as go

from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import ColorScaleRule

# استيراد مكتبة Supabase الرسمية
from supabase import create_client, Client

# ============================================================
# MAD Orders Dashboard V8.4.2 - Supabase Production Version
# ============================================================

APP_VERSION = "V8.4.2 Supabase Integrated Fix"

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
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght=400;600;700;800&display=swap');

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
        section[data-testid="stSidebar"][aria-expanded="true"] { min-width: 285px !important; max-width: 92vw !important; }
        section[data-testid="stSidebar"][aria-expanded="false"] { min-width: 0 !important; width: 0 !important; max-width: 0 !important; overflow: hidden !important; }
    }
    
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

    div[data-testid="stExpander"] {
        border: 1px solid rgba(148, 163, 184, .28) !important;
        border-radius: 16px !important;
        background: rgba(15, 23, 42, .38) !important;
        margin: 12px 0 20px 0 !important;
    }

    div[data-testid="stExpander"] summary {
        font-weight: 800 !important;
        color: #f8fafc !important;
        font-size: 15px !important;
    }

    .js-plotly-plot .plotly .ytick text,
    .js-plotly-plot .plotly .xtick text {
        dominant-baseline: middle !important;
    }

    .js-plotly-plot .plotly .heatmaplayer text {
        font-weight: 700 !important;
        fill: rgba(15, 23, 42, .88) !important;
    }

    div[data-testid="stWidgetLabel"],
    div[data-testid="stWidgetLabel"] p,
    div[data-testid="stWidgetLabel"] label,
    .stSelectbox label,
    .stTextInput label,
    .stCheckbox label,
    .stRadio label,
    .stDateInput label,
    .stMultiSelect label,
    .stFileUploader label {
        color: #f8fafc !important;
        opacity: 1 !important;
        font-weight: 800 !important;
        font-size: 15px !important;
        text-shadow: 0 1px 1px rgba(0,0,0,.35) !important;
    }

    .stMarkdown p,
    div[data-testid="stMarkdownContainer"] p {
        color: #e5e7eb !important;
        opacity: 1 !important;
    }

    div[data-baseweb="select"] > div {
        border: 1px solid rgba(248,250,252,.28) !important;
        box-shadow: 0 0 0 1px rgba(15,23,42,.08) !important;
    }

    div[data-baseweb="select"] span,
    div[data-baseweb="select"] input,
    .stTextInput input {
        font-weight: 700 !important;
    }

    .stCheckbox p,
    .stCheckbox span {
        color: #f8fafc !important;
        opacity: 1 !important;
        font-weight: 700 !important;
    }

    .production-card,
    .action-card,
    .readability-note,
    .note-box {
        color: #f8fafc !important;
        opacity: 1 !important;
    }

    .production-card *,
    .action-card *,
    .readability-note *,
    .note-box * {
        color: inherit !important;
        opacity: 1 !important;
    }

    .js-plotly-plot .plotly .heatmaplayer text {
        font-weight: 900 !important;
        paint-order: stroke !important;
        stroke: rgba(15,23,42,.55) !important;
        stroke-width: 1.5px !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Sidebar styling
st.markdown(
    """
    <style>
    section[data-testid="stSidebar"][aria-expanded="true"] {
        background: linear-gradient(180deg, #0f172a 0%, #111827 100%) !important;
        border-right: 1px solid rgba(148, 163, 184, .24) !important;
        box-shadow: 8px 0 28px rgba(0, 0, 0, .30) !important;
    }

    section[data-testid="stSidebar"][aria-expanded="true"] > div {
        background: transparent !important;
    }

    section[data-testid="stSidebar"][aria-expanded="true"] h1,
    section[data-testid="stSidebar"][aria-expanded="true"] h2,
    section[data-testid="stSidebar"][aria-expanded="true"] h3,
    section[data-testid="stSidebar"][aria-expanded="true"] h4,
    section[data-testid="stSidebar"][aria-expanded="true"] p,
    section[data-testid="stSidebar"][aria-expanded="true"] label,
    section[data-testid="stSidebar"][aria-expanded="true"] span,
    section[data-testid="stSidebar"][aria-expanded="true"] div[data-testid="stMarkdownContainer"],
    section[data-testid="stSidebar"][aria-expanded="true"] div[data-testid="stWidgetLabel"],
    section[data-testid="stSidebar"][aria-expanded="true"] div[data-testid="stWidgetLabel"] p {
        color: #f8fafc !important;
        opacity: 1 !important;
        font-weight: 700 !important;
        text-shadow: none !important;
    }

    section[data-testid="stSidebar"][aria-expanded="true"] small,
    section[data-testid="stSidebar"][aria-expanded="true"] div[data-testid="stCaptionContainer"] {
        color: #cbd5e1 !important;
        opacity: 1 !important;
    }

    section[data-testid="stSidebar"][aria-expanded="true"] input,
    section[data-testid="stSidebar"][aria-expanded="true"] textarea {
        background: #ffffff !important;
        color: #111827 !important;
        -webkit-text-fill-color: #111827 !important;
        border: 1px solid rgba(15, 23, 42, .18) !important;
        border-radius: 10px !important;
        font-weight: 700 !important;
    }

    section[data-testid="stSidebar"][aria-expanded="true"] div[data-baseweb="select"] > div {
        background: #ffffff !important;
        color: #111827 !important;
        border: 1px solid rgba(15, 23, 42, .18) !important;
        border-radius: 10px !important;
    }

    section[data-testid="stSidebar"][aria-expanded="true"] div[data-baseweb="select"] span,
    section[data-testid="stSidebar"][aria-expanded="true"] div[data-baseweb="select"] input,
    section[data-testid="stSidebar"][aria-expanded="true"] div[data-baseweb="select"] div {
        color: #111827 !important;
        -webkit-text-fill-color: #111827 !important;
        font-weight: 700 !important;
    }

    section[data-testid="stSidebar"][aria-expanded="true"] span[data-baseweb="tag"] {
        background: #2563eb !important;
        color: #ffffff !important;
        font-weight: 800 !important;
    }

    section[data-testid="stSidebar"][aria-expanded="true"] .stButton button {
        background: rgba(30, 41, 59, .95) !important;
        color: #f8fafc !important;
        border: 1px solid rgba(248, 250, 252, .35) !important;
        border-radius: 12px !important;
        font-weight: 800 !important;
    }

    section[data-testid="stSidebar"][aria-expanded="true"] .stButton button:hover {
        background: rgba(37, 99, 235, .95) !important;
        border-color: rgba(191, 219, 254, .75) !important;
    }

    section[data-testid="stSidebar"][aria-expanded="false"] {
        min-width: 0 !important;
        width: 0 !important;
        max-width: 0 !important;
        overflow: hidden !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# Core Helpers
# ============================================================

def normalize_arabic_digits(value):
    if pd.isna(value):
        return ""
    text = str(value)
    arabic_digits = "٠١٢٣٤٥٦٧٨٩"
    persian_digits = "۰۱۲۳۴۵۶٧٨٩"
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
    raw = "" if pd.isna(chef_name) else str(chef_name)
    text = normalize_arabic_digits(raw)
    text = re.sub(r"\s+", " ", text).strip()
    low = text.lower()

    if "قرطبة" in text or "qurtuba" in low or "qortoba" in low or "qurtobah" in low:
        return "قرطبة"
    if "عريجاء" in text or "عريجا" in text or "العريجاء" in text or "uraija" in low:
        return "عريجاء"
    if "الروضة" in text or "روضه" in text or "rawdah" in low or "rawda" in low:
        return "الروضة"
    if "العارض" in text or "عارض" in text or "arid" in low or "al arid" in low:
        return "العارض"
    if "الورود" in text or "ورود" in text or "worood" in low or "al worood" in low:
        return "الورود"
    if "العقيق" in text or "عقيق" in text or "aqiq" in low or "al aqiq" in low:
        return "العقيق"
    if "madness and desire" in low or "madness" in low or "مادنيس اند ديزاير" in text or "مادنيس" in text:
        return "العقيق"
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


def text_has_date(value):
    text = normalize_arabic_digits(value)
    if not text:
        return False
    patterns = [r"\b\d{4}[-/]\d{1,2}[-/]\d{1,2}\b", r"\b\d{1,2}[-/]\d{1,2}[-/]\d{2,4}\b"]
    return any(re.search(p, text) for p in patterns)


def text_has_time(value):
    text = normalize_arabic_digits(value).lower()
    if not text:
        return False
    return bool(re.search(r"\b\d{1,2}:\d{2}\s*(am|pm|ص|م)?\b", text) or re.search(r"\b\d{1,2}\s*(am|pm|ص|م)\b", text))


def parse_any_datetime_with_date(value):
    text = normalize_arabic_digits(value)
    if not text or not text_has_date(text):
        return pd.NaT
    for dayfirst in [True, False]:
        parsed = pd.to_datetime(text, errors="coerce", dayfirst=dayfirst)
        if not pd.isna(parsed):
            return parsed
    return pd.NaT


def parse_time_only_value(value):
    text = normalize_arabic_digits(value).strip()
    if not text or text_has_date(text):
        return ""
    text = text.replace("ص", "AM").replace("م", "PM")
    for fmt in ["%I:%M %p", "%I %p", "%H:%M"]:
        try:
            dt = pd.to_datetime(text, format=fmt, errors="coerce")
            if not pd.isna(dt):
                hour = int(dt.hour)
                minute = int(dt.minute)
                period = "AM" if hour < 12 else "PM"
                hour12 = hour % 12 or 12
                return f"{hour12}:{minute:02d} {period}"
        except Exception:
            pass
    if text_has_time(text):
        return text
    return ""


def clean_delivery_pickup_values(date_value, time_value):
    raw_date = normalize_arabic_digits(date_value).strip()
    raw_time = normalize_arabic_digits(time_value).strip()
    date_dt = parse_any_datetime_with_date(raw_date)
    time_dt = parse_any_datetime_with_date(raw_time)
    cleaned = False

    if not pd.isna(time_dt):
        cleaned_date = format_date_iso(time_dt)
        cleaned_time = format_time_12h(time_dt)
        if not pd.isna(date_dt):
            cleaned_date = format_date_iso(date_dt)
        return cleaned_date, cleaned_time, True

    if not pd.isna(date_dt):
        cleaned_date = format_date_iso(date_dt)
        if text_has_time(raw_date):
            cleaned_time = format_time_12h(date_dt)
            cleaned = True
        else:
            cleaned_time = parse_time_only_value(raw_time) or raw_time
        return cleaned_date, cleaned_time, cleaned

    cleaned_date = raw_date
    cleaned_time = parse_time_only_value(raw_time) or raw_time
    return cleaned_date, cleaned_time, (raw_time != cleaned_time)


def parse_datetime_parts(date_value, time_value):
    date_value = normalize_arabic_digits(date_value)
    time_value = normalize_arabic_digits(time_value)
    candidates = []
    if date_value and time_value:
        candidates.append(f"{date_value} {time_value}")
    if time_value and text_has_date(time_value):
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


def parse_courier_datetime_parts(date_value, pickup_time_value, courier_time_value):
    raw_courier = normalize_arabic_digits(courier_time_value).strip()
    if not raw_courier or raw_courier.lower() in ["nan", "none", "(not set)", "not set", "-", "nat"]:
        return pd.NaT
    pickup_dt = parse_datetime_parts(date_value, pickup_time_value)
    if text_has_date(raw_courier):
        parsed = parse_datetime_parts("", raw_courier)
    else:
        parsed = parse_datetime_parts(date_value, raw_courier)
    if pd.isna(parsed):
        return pd.NaT
    if not text_has_date(raw_courier) and not pd.isna(pickup_dt):
        try:
            diff_minutes = (pd.Timestamp(parsed) - pd.Timestamp(pickup_dt)).total_seconds() / 60
            if diff_minutes < -360:
                parsed = pd.Timestamp(parsed) + pd.Timedelta(days=1)
        except Exception:
            pass
    return parsed


def format_date_iso(dt):
    if pd.isna(dt): return ""
    return pd.Timestamp(dt).strftime("%Y-%m-%d")


def format_time_12h(dt):
    if pd.isna(dt): return ""
    ts = pd.Timestamp(dt)
    h12 = ts.hour % 12 or 12
    return f"{h12}:{ts.minute:02d} {'AM' if ts.hour < 12 else 'PM'}"


def hour_label(hour):
    if pd.isna(hour): return "بدون وقت"
    hour = int(hour)
    def fmt(h):
        return f"{h % 12 or 12} {'ص' if h < 12 else 'م'}"
    return f"{fmt(hour).replace(' ص','').replace(' م','')}-{fmt((hour + 1) % 24)}"


def is_cancelled_status(value):
    low = str(value).strip().lower()
    return any(k in low for k in ["cancel", "cancelled", "canceled", "ملغي", "ملغاة", "رفض"])


def is_addon_product(product_name):
    low = str(product_name).lower()
    keywords = ["candle", "candles", "شموع", "شمعة", "balloon", "helium", "بالون", "هيليوم", "gift card", "كرت", "بطاقة", "card", "night stars", "نجوم", "stars", "topper", "توبير"]
    return any(k in low for k in keywords)


def addon_category(product_name):
    low = str(product_name).lower()
    if any(k in low for k in ["candle", "candles", "شموع", "شمعة"]): return "Candles / شموع"
    if any(k in low for k in ["balloon", "helium", "بالون", "هيليوم"]): return "Balloons / بالونات"
    if any(k in low for k in ["gift card", "كرت", "بطاقة", "card"]): return "Gift Cards / كروت"
    if any(k in low for k in ["night stars", "نجوم", "stars"]): return "Night Stars"
    if any(k in low for k in ["topper", "توبير"]): return "Toppers"
    return "Other Add-ons"


def classify_campaign(product_name):
    low = str(product_name).lower()
    if any(k in low for k in ["father", "dad", "عيد الأب", "عيد الاب", "بابا", "super dad"]): return "Father's Day"
    if any(k in low for k in ["birthday", "بيرثداي", "ميلاد"]): return "Birthday"
    if any(k in low for k in ["graduation", "تخرج", "التخرج"]): return "Graduation"
    if any(k in low for k in ["eid", "عيد الفطر", "عيد"]): return "Eid"
    if any(k in low for k in ["new year", "نيو يير"]): return "New Year"
    if any(k in low for k in ["valentine", "love", "الحب"]): return "Valentine"
    if any(k in low for k in ["gender reveal", "تحديد الجنس"]): return "Gender Reveal"
    return "General"


def extract_phone_from_text(value):
    text = normalize_arabic_digits(value)
    compressed = re.sub(r"[\s\-()]+", "", text)
    for pat in [r"\+?9665\d{8}", r"05\d{8}", r"5\d{8}"]:
        m = re.search(pat, compressed)
        if m: return m.group(0)
    return ""


def need_action_reasons(note, product_name=""):
    text = f"{note} {product_name}".strip()
    low = text.lower()
    reasons = []
    if any(k in low for k in ["photo", "picture", "image", "صورة", "الصورة", "الصوره"]): reasons.append("يحتاج صورة")
    if any(k in low for k in ["contact", "call", "whatsapp", "تواصل", "اتصال", "واتساب", "جوال"]): reasons.append("يحتاج تواصل")
    if extract_phone_from_text(text): reasons.append("يوجد رقم جوال")
    if any(k in low for k in ["write", "writing", "اكتب", "كتابة", "العبارة"]): reasons.append("كتابة خاصة")
    if any(k in low for k in ["draw", "design", "color", "hearts", "تعديل", "تصميم"]): reasons.append("تعديل تصميم")
    if any(k in low for k in ["problem", "wrong", "mistake", "خطأ", "مشكلة"]): reasons.append("ملاحظة حساسة")
    if len(str(note).strip()) > 90: reasons.append("ملاحظة طويلة")
    
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
        "courier_delivery_time": find_col(df, ["وقت التسليم للمندوب", "وقت تسليم المندوب", "Courier Delivery Time", "courier_delivery_time"]),
        "item_id": find_col(df, ["معرف العنصر (Item Id)", "Item Id"]),
        "dish_id": find_col(df, ["معرف الطبق (Dish Id)"]),
        "product": find_col(df, ["اسم الطبق / المنتج", "Product", "Dish", "اسم الطبق"]),
        "variety": find_col(df, ["نوع الحشوة (Variety)", "Variety", "نوع الحشوة"]),
        "variety_price": find_col(df, ["سعر الحشوة", "Variety Price"]),
        "note": find_col(df, ["ملاحظة للشيف", "Chef Note", "Note", "ملاحظة"]),
        "unit_price": find_col(df, ["سعر الحبة", "Unit Price"]),
        "discount": find_col(df, ["الخصم"]),
        "quantity": find_col(df, ["الكمية", "Quantity", "qty"]),
        "item_total": find_col(df, ["إجمالي المنتج (Item Total)", "Item Total"]),
    }


def col_or_blank(df, col):
    if col and col in df.columns: return df[col].fillna("").astype(str)
    return pd.Series([""] * len(df), index=df.index)


def col_or_default(df, col, default=""):
    if col and col in df.columns: return df[col].fillna(default)
    return pd.Series([default] * len(df), index=df.index)

# ============================================================
# Supabase Data Loader Integration
# ============================================================

@st.cache_resource
def init_supabase():
    """تهيئة العميل لقاعدة البيانات باستخدام Secrets التابعة لـ Streamlit"""
    try:
        url = st.secrets["supabase_url"]
        key = st.secrets["supabase_key"]
        return create_client(url, key)
    except Exception as e:
        st.error(f"حدث خطأ أثناء تحميل إعدادات الاتصال بـ Supabase Secrets: {e}")
        return None

@st.cache_data(ttl=60, show_spinner=False)
def load_data_from_supabase_table():
    """جلب كل الصفوف من جدول السوبابيز مباشرة"""
    client = init_supabase()
    if not client:
        return pd.DataFrame()
    try:
        # تم تغيير اسم الجدول هنا من 'orders' إلى 'daily_orders' ليطابق قاعدة بياناتك
        response = client.table("daily_orders").select("*").execute()
        if response.data:
            return pd.DataFrame(response.data)
        return pd.DataFrame()
    except Exception as e:
        st.error(f"خطأ أثناء سحب البيانات المباشرة من Supabase: {e}")
        return pd.DataFrame()

def read_uploaded_file(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded_file, dtype=str).fillna("")
    try:
        df = pd.read_csv(uploaded_file, sep="\t", dtype=str, keep_default_na=False, engine="python")
        if len(df.columns) > 1: return df.fillna("")
    except: pass
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
        if cols["order_no"]: cols["order_id"] = cols["order_no"]
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
    
    # تصحيح الحساب المعتمد: السعر × الكمية
    df["سعر الحبة رقم"] = col_or_default(df, cols["unit_price"], "0").apply(clean_money)
    df["إجمالي المنتج رقم"] = df["سعر الحبة رقم"] * df["الكمية رقم"]
    
    df["سعر الحشوة رقم"] = col_or_default(df, cols["variety_price"], "0").apply(clean_money)
    df["الخصم"] = col_or_blank(df, cols["discount"])
    df["تاريخ التوصيل قبل التنظيف"] = col_or_blank(df, cols["delivery_date"])
    df["وقت الاستلام قبل التنظيف"] = col_or_blank(df, cols["pickup_time"])
    df["وقت التسليم للمندوب قبل التنظيف"] = col_or_blank(df, cols.get("courier_delivery_time"))

    cleaned_datetime_parts = [clean_delivery_pickup_values(d, t) for d, t in zip(df["تاريخ التوصيل قبل التنظيف"], df["وقت الاستلام قبل التنظيف"])]
    df["تاريخ التوصيل الأصلي"] = [x[0] for x in cleaned_datetime_parts]
    df["وقت الاستلام الأصلي"] = [x[1] for x in cleaned_datetime_parts]
    df["تم تنظيف التاريخ/الوقت؟"] = [bool(x[2]) for x in cleaned_datetime_parts]

    df["تاريخ ووقت الاستلام"] = [parse_datetime_parts(d, t) for d, t in zip(df["تاريخ التوصيل الأصلي"], df["وقت الاستلام الأصلي"])]
    df["وقت التسليم للمندوب"] = [parse_courier_datetime_parts(d, p, c) for d, p, c in zip(df["تاريخ التوصيل الأصلي"], df["وقت الاستلام الأصلي"], df["وقت التسليم للمندوب قبل التنظيف"])]
    
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

    df["_sort_datetime_safe"] = pd.to_datetime(df["تاريخ ووقت الاستلام"], errors="coerce")
    df["_sort_order_safe"] = df["رقم الطلب الموحد"].astype(str)
    group = df.sort_values(["_sort_datetime_safe", "_sort_order_safe"], na_position="last").groupby("رقم الطلب الموحد", dropna=False)
    
    order_level = group.agg(
        رقم_الطلب=("رقم الطلب الظاهر", "first"),
        العميل=("العميل", "first"),
        الفرع=("الفرع", "first"),
        الحالة=("الحالة", "first"),
        قيمة_الطلب=("قيمة الطلب رقم", "max"),
        تاريخ_التحليل=("تاريخ التحليل", "first"),
        وقت_الاستلام=("وقت الاستلام الأصلي", "first"),
        وقت_التسليم_للمندوب_الأصلي=("وقت التسليم للمندوب قبل التنظيف", "first"),
        وقت_التسليم_للمندوب=("وقت التسليم للمندوب", "first"),
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
    
    order_level = order_level.rename(columns={
        "رقم_الطلب": "رقم الطلب", "قيمة_الطلب": "قيمة الطلب", "تاريخ_التحليل": "تاريخ التحليل",
        "وقت_الاستلام": "وقت الاستلام", "وقت_التسليم_للمندوب_الأصلي": "وقت التسليم للمندوب الأصلي",
        "وقت_التسليم_للمندوب": "وقت التسليم للمندوب", "تاريخ_ووقت_الاستلام": "تاريخ ووقت الاستلام",
        "ساعة_رقم": "ساعة رقم", "عدد_الأصناف": "عدد الأصناف", "عدد_المنتجات": "عدد المنتجات",
        "عدد_الإضافات": "عدد الإضافات", "فيه_إضافات": "فيه إضافات", "يحتاج_متابعة": "يحتاج متابعة",
    })
    order_level["ساعة رقم"] = pd.to_numeric(order_level["ساعة رقم"], errors="coerce")
    return df, order_level, cols


def build_reports(items, orders):
    active_items = items[~items["ملغي؟"]].copy()
    active_orders = orders[~orders["ملغي"].astype(bool)].copy() if "ملغي" in orders.columns else orders.copy()
    reports = {"raw_filtered": items, "items_active": active_items, "orders_active": active_orders}

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
        reports["sales_by_branch"] = reports["sales_by_hour"] = reports["status_report"] = reports["branch_hour_heatmap"] = pd.DataFrame()

    non_addon = active_items[~active_items["إضافة؟"]].copy()
    if not non_addon.empty:
        reports["product_performance"] = non_addon.groupby("المنتج", dropna=False).agg(
            عدد_الطلبات=("رقم الطلب الموحد", "nunique"), الكمية=("الكمية رقم", "sum"),
            المبيعات_منتجات=("إجمالي المنتج رقم", "sum"), متوسط_سعر_الحبة=("سعر الحبة رقم", "mean"),
            تحتاج_متابعة=("يحتاج متابعة؟", "sum"),
        ).reset_index().sort_values(["المبيعات_منتجات", "الكمية"], ascending=False)
        
        reports["product_by_branch"] = non_addon.groupby(["الفرع", "المنتج"], dropna=False).agg(
            عدد_الطلبات=("رقم الطلب الموحد", "nunique"), الكمية=("الكمية رقم", "sum"), المبيعات=("إجمالي المنتج رقم", "sum"),
        ).reset_index().sort_values(["الفرع", "المبيعات"], ascending=[True, False])
        
        variety_source = non_addon[non_addon["الحشوة"].astype(str).str.strip().str.len() > 0].copy()
        reports["variety_report"] = variety_source.groupby("الحشوة", dropna=False).agg(
            عدد_الطلبات=("رقم الطلب الموحد", "nunique"), الكمية=("الكمية رقم", "sum"),
            المبيعات=("إجمالي المنتج رقم", "sum"), متوسط_سعر_الحبة=("سعر الحبة رقم", "mean"),
            تحتاج_متابعة=("يحتاج متابعة؟", "sum"),
        ).reset_index().sort_values(["الكمية", "المبيعات"], ascending=False)

        reports["variety_by_branch"] = variety_source.groupby(["الفرع", "الحشوة"], dropna=False).agg(
            عدد_الطلبات=("رقم الطلب الموحد", "nunique"), الكمية=("الكمية رقم", "sum"), المبيعات=("إجمالي المنتج رقم", "sum"),
        ).reset_index().sort_values(["الفرع", "الكمية"], ascending=[True, False])

        reports["variety_by_product"] = variety_source.groupby(["المنتج", "الحشوة"], dropna=False).agg(
            عدد_الطلبات=("رقم الطلب الموحد", "nunique"), الكمية=("الكمية رقم", "sum"), المبيعات=("إجمالي المنتج رقم", "sum"),
        ).reset_index().sort_values(["المنتج", "الكمية"], ascending=[True, False])

        reports["variety_by_hour"] = variety_source.groupby(["الساعة", "الحشوة"], dropna=False).agg(
            عدد_الطلبات=("رقم الطلب الموحد", "nunique"), الكمية=("الكمية رقم", "sum"), المبيعات=("إجمالي المنتج رقم", "sum"),
        ).reset_index()
        if not reports["variety_by_hour"].empty:
            h_ord = variety_source.groupby("الساعة", dropna=False)["ساعة رقم"].min().reset_index(name="ساعة رقم")
            reports["variety_by_hour"] = reports["variety_by_hour"].merge(h_ord, on="الساعة", how="left").sort_values(["ساعة رقم", "الكمية"], ascending=[True, False]).drop(columns=["ساعة رقم"], errors="ignore")

        reports["variety_by_campaign"] = variety_source.groupby(["الحملة", "الحشوة"], dropna=False).agg(
            عدد_الطلبات=("رقم الطلب الموحد", "nunique"), الكمية=("الكمية رقم", "sum"), المبيعات=("إجمالي المنتج رقم", "sum"),
        ).reset_index().sort_values(["الحملة", "الكمية"], ascending=[True, False])

        if not variety_source.empty:
            reports["branch_variety_heatmap"] = variety_source.pivot_table(index="الفرع", columns="الحشوة", values="الكمية رقم", aggfunc="sum", fill_value=0)
            reports["product_variety_heatmap"] = variety_source.pivot_table(index="المنتج", columns="الحشوة", values="الكمية رقم", aggfunc="sum", fill_value=0)
            v_cols = ["رقم الطلب الظاهر", "رقم الطلب الموحد", "الفرع", "الحالة", "تاريخ التحليل", "وقت الاستلام الأصلي", "الساعة", "العميل", "المنتج", "الحشوة", "الكمية رقم", "إجمالي المنتج رقم", "سبب المتابعة", "الملاحظة"]
            reports["variety_order_details"] = variety_source[[c for c in v_cols if c in variety_source.columns]].sort_values(["الحشوة", "الفرع", "وقت الاستلام الأصلي"], na_position="last")
        else:
            reports["branch_variety_heatmap"] = reports["product_variety_heatmap"] = reports["variety_order_details"] = pd.DataFrame()
    else:
        reports["product_performance"] = reports["product_by_branch"] = reports["variety_report"] = reports["variety_by_branch"] = reports["variety_by_product"] = reports["variety_by_hour"] = reports["variety_by_campaign"] = reports["branch_variety_heatmap"] = reports["product_variety_heatmap"] = reports["variety_order_details"] = pd.DataFrame()

    addons = active_items[active_items["إضافة؟"]].copy()
    reports["addon_items"] = addons
    if not addons.empty:
        reports["addons_summary"] = addons.groupby("تصنيف الإضافة", dropna=False).agg(عدد_الطلبات=("رقم الطلب الموحد", "nunique"), الكمية=("الكمية رقم", "sum"), المبيعات=("إجمالي المنتج رقم", "sum")).reset_index().sort_values("الكمية", ascending=False)
        reports["addons_by_branch"] = addons.groupby(["الفرع", "تصنيف الإضافة"], dropna=False).agg(عدد_الطلبات=("رقم الطلب الموحد", "nunique"), الكمية=("الكمية رقم", "sum"), المبيعات=("إجمالي المنتج رقم", "sum")).reset_index().sort_values(["الفرع", "الكمية"], ascending=[True, False])
    else:
        reports["addons_summary"] = reports["addons_by_branch"] = pd.DataFrame()

    need_action = active_items[active_items["يحتاج متابعة؟"]].copy()
    if not need_action.empty:
        ac_cols = ["رقم الطلب الظاهر", "رقم الطلب الموحد", "الفرع", "الحالة", "تاريخ التوصيل الأصلي", "وقت الاستلام الأصلي", "العميل", "المنتج", "الحشوة", "الكمية رقم", "رقم الجوال المستخرج", "سبب المتابعة", "الملاحظة"]
        reports["need_action"] = need_action[[c for c in ac_cols if c in need_action.columns]].drop_duplicates().sort_values(["الفرع", "وقت الاستلام الأصلي"])
        r_rows = []
        for reasons in need_action["سبب المتابعة"].dropna().astype(str):
            for r in [x.strip() for x in reasons.split("،") if x.strip()]: r_rows.append(r)
        reports["need_action_reasons"] = pd.Series(r_rows).value_counts().rename("عدد الحالات").reset_index().rename(columns={"index": "سبب المتابعة"}) if r_rows else pd.DataFrame(columns=["سبب المتابعة", "عدد الحالات"])
    else:
        reports["need_action"] = reports["need_action_reasons"] = pd.DataFrame()

    q_rows = []
    q_rows.append({"المشكلة": "تاريخ توصيل ناقص", "عدد الصفوف": int(items["تاريخ التوصيل الأصلي"].astype(str).str.strip().eq("").sum()), "الأهمية": "عالية"})
    q_rows.append({"المشكلة": "وقت استلام ناقص", "عدد الصفوف": int(items["وقت الاستلام الأصلي"].astype(str).str.strip().eq("").sum()), "الأهمية": "عالية"})
    q_rows.append({"المشكلة": "تم تنظيف التاريخ/الوقت تلقائيًا", "عدد الصفوف": int(items["تم تنظيف التاريخ/الوقت؟"].astype(bool).sum()) if "تم تنظيف التاريخ/الوقت؟" in items.columns else 0, "الأهمية": "منخفضة"})
    q_rows.append({"المشكلة": "فرع غير محدد", "عدد الصفوف": int(items["الفرع"].eq("بدون فرع محدد").sum()), "الأهمية": "متوسطة"})
    q_rows.append({"المشكلة": "منتج بدون اسم", "عدد الصفوف": int(items["المنتج"].eq("بدون اسم منتج").sum()), "الأهمية": "عالية"})
    q_rows.append({"المشكلة": "قيمة طلب صفرية", "عدد الصفوف": int(items["قيمة الطلب رقم"].fillna(0).eq(0).sum()), "الأهمية": "متوسطة"})
    reports["data_quality_summary"] = pd.DataFrame(q_rows)
    reports["data_quality_details"] = items[items["تاريخ التوصيل الأصلي"].astype(str).str.strip().eq("") | items["وقت الاستلام الأصلي"].astype(str).str.strip().eq("") | items["الفرع"].eq("بدون فرع محدد") | items["المنتج"].eq("بدون اسم منتج") | items["قيمة الطلب رقم"].fillna(0).eq(0)].copy()

    if not active_items.empty:
        reports["campaign_summary"] = active_items.groupby("الحملة", dropna=False).agg(عدد_الطلبات=("رقم الطلب الموحد", "nunique"), الكمية=("الكمية رقم", "sum"), مبيعات_الأصناف=("إجمالي المنتج رقم", "sum"), تحتاج_متابعة=("يحتاج متابعة؟", "sum")).reset_index().sort_values("عدد_الطلبات", ascending=False)
        reports["campaign_products"] = active_items.groupby(["الحملة", "المنتج"], dropna=False).agg(عدد_الطلبات=("رقم الطلب الموحد", "nunique"), الكمية=("الكمية رقم", "sum"), مبيعات_الأصناف=("إجمالي المنتج رقم", "sum")).reset_index().sort_values(["الحملة", "عدد_الطلبات"], ascending=[True, False])
    else:
        reports["campaign_summary"] = reports["campaign_products"] = pd.DataFrame()

    reports["multi_item_orders"] = active_orders[active_orders["عدد الأصناف"] > 1].sort_values(["عدد الأصناف", "قيمة الطلب"], ascending=False) if not active_orders.empty else pd.DataFrame()
    return reports


def build_v83_advanced_reports(items, active_items, active_orders, reports):
    advanced = {}
    non_addon = active_items[~active_items["إضافة؟"]].copy() if not active_items.empty else pd.DataFrame()
    
    total_orders_adv = int(active_orders["رقم الطلب الموحد"].nunique()) if not active_orders.empty else 0
    total_sales_adv = float(active_orders["قيمة الطلب"].fillna(0).sum()) if not active_orders.empty else 0.0
    avg_order_adv = total_sales_adv / total_orders_adv if total_orders_adv else 0.0

    advanced["advanced_executive_summary"] = pd.DataFrame([
        {"المؤشر": "عدد الطلبات", "القيمة": total_orders_adv, "ملاحظة": "Unique Order Id"},
        {"المؤشر": "إجمالي المبيعات", "القيمة": total_sales_adv, "ملاحظة": "بدون تكرار قيمة الطلب"},
        {"المؤشر": "متوسط قيمة الطلب AOV", "القيمة": avg_order_adv, "ملاحظة": "Sales / Orders"},
    ])

    if not active_orders.empty:
        br_rank = active_orders.groupby("الفرع", dropna=False).agg(الطلبات=("رقم الطلب الموحد", "nunique"), المبيعات=("قيمة الطلب", "sum"), متوسط_الطلب=("قيمة الطلب", "mean")).reset_index()
        advanced["advanced_branch_ranking"] = br_rank.sort_values("المبيعات", ascending=False)
    else:
        advanced["advanced_branch_ranking"] = pd.DataFrame()

    if not non_addon.empty:
        advanced["advanced_product_value"] = non_addon.groupby("المنتج", dropna=False).agg(الطلبات=("رقم الطلب الموحد", "nunique"), الكمية=("الكمية رقم", "sum"), المبيعات=("إجمالي المنتج رقم", "sum")).reset_index().sort_values("المبيعات", ascending=False)
        advanced["advanced_product_value"]["مبيعات_لكل_طلب"] = (advanced["advanced_product_value"]["المبيعات"] / advanced["advanced_product_value"]["الطلبات"]).round(2)
        advanced["advanced_product_value"]["حصة_المبيعات_%"] = (advanced["advanced_product_value"]["المبيعات"] / total_sales_adv * 100).round(1) if total_sales_adv else 0
        advanced["advanced_product_value"]["تصنيف_القيمة"] = "Hero Product"
        
        advanced["advanced_filling_intelligence"] = non_addon[non_addon["الحشوة"].ne("")].groupby("الحشوة", dropna=False).agg(الطلبات=("رقم الطلب الموحد", "nunique"), الكمية=("الكمية رقم", "sum"), المبيعات=("إجمالي المنتج رقم", "sum")).reset_index().sort_values("الكمية", ascending=False)
        advanced["advanced_filling_intelligence"]["حصة_الكمية_%"] = (advanced["advanced_filling_intelligence"]["الكمية"] / advanced["advanced_filling_intelligence"]["الكمية"].sum() * 100).round(1) if len(advanced["advanced_filling_intelligence"]) else 0
        advanced["advanced_filling_intelligence"]["نسبة_متابعة_%"] = 0.0
        advanced["advanced_filling_by_branch_rank"] = non_addon[non_addon["الحشوة"].ne("")].groupby(["الفرع", "الحشوة"], dropna=False).agg(الكمية=("الكمية رقم", "sum")).reset_index().sort_values(["الفرع", "الكمية"], ascending=[True, False])
    else:
        advanced["advanced_product_value"] = advanced["advanced_filling_intelligence"] = advanced["advanced_filling_by_branch_rank"] = pd.DataFrame()

    advanced["advanced_addons_opportunity_branch"] = advanced["advanced_addons_product_opportunity"] = advanced["advanced_notes_keywords"] = advanced["advanced_notes_by_product"] = advanced["advanced_data_quality_score"] = advanced["advanced_hourly_capacity"] = advanced["advanced_campaign_performance"] = pd.DataFrame()
    return advanced


def hour_range_label(hour_value):
    try: h = int(float(hour_value))
    except: return "بدون وقت"
    if h < 0 or h > 23: return "بدون وقت"
    return f"{h%12 or 12}:00 {'AM' if h < 12 else 'PM'} - {(h+1)%12 or 12}:00 {'AM' if (h+1)%24 < 12 else 'PM'}"


def hour_range_sort_value(label):
    text = str(label)
    if text == "بدون وقت": return 999
    m = re.search(r"(\d{1,2}):00\s*(AM|PM)", text, flags=re.IGNORECASE)
    if not m: return 999
    hour = int(m.group(1))
    period = m.group(2).upper()
    if period == "AM": return 0 if hour == 12 else hour
    return 12 if hour == 12 else hour + 12


def build_production_queue(active_items):
    if active_items is None or active_items.empty: return pd.DataFrame()
    q = active_items.copy()
    q["نوع الصنف"] = q["إضافة؟"].map(lambda x: "إضافة" if bool(x) else "منتج رئيسي")
    q["واتساب"] = q["رقم الجوال المستخرج"].apply(lambda p: f"https://wa.me/{p}" if p else "")
    q["أولوية"] = q["يحتاج متابعة؟"].map(lambda x: "عالية" if bool(x) else "عادية")
    q["وقت عرض"] = q["وقت الاستلام الأصلي"].replace("", "بدون وقت")
    q["تاريخ عرض"] = q["تاريخ التوصيل الأصلي"].replace("", "بدون تاريخ")
    q["نطاق ساعة الاستلام"] = q["ساعة رقم"].apply(hour_range_label) if "ساعة رقم" in q.columns else "بدون وقت"
    return q


def build_action_center(active_items):
    if active_items is None or active_items.empty: return pd.DataFrame()
    a = active_items[active_items["يحتاج متابعة؟"] | active_items["تاريخ التوصيل الأصلي"].eq("") | active_items["وقت الاستلام الأصلي"].eq("")].copy()
    if a.empty: return pd.DataFrame()
    a["نوع الإجراء"] = a["سبب المتابعة"]
    a["الأولوية"] = "عالية"
    a["حالة المتابعة"] = "لم يبدأ"
    a["ملاحظة داخلية"] = ""
    a["واتساب"] = a["رقم الجوال المستخرج"].apply(lambda p: f"https://wa.me/{p}" if p else "")
    a["وقت عرض"] = a["وقت الاستلام الأصلي"].replace("", "بدون وقت")
    a["تاريخ عرض"] = a["تاريخ التوصيل الأصلي"].replace("", "بدون تاريخ")
    a["نطاق ساعة الاستلام"] = a["ساعة رقم"].apply(hour_range_label) if "ساعة رقم" in a.columns else "بدون وقت"
    return a


def build_late_orders_reports(active_items, active_orders, grace_minutes=10, risk_window_minutes=30):
    empty_pack = {"late_orders": pd.DataFrame(), "late_risk_orders": pd.DataFrame(), "late_delivered_late": pd.DataFrame(), "late_pending_late": pd.DataFrame(), "late_action_center": pd.DataFrame(), "late_by_branch": pd.DataFrame(), "late_by_hour": pd.DataFrame(), "late_branch_hour_heatmap": pd.DataFrame(), "late_reasons": pd.DataFrame(), "late_summary": pd.DataFrame()}
    if active_orders is None or active_orders.empty: return empty_pack

    now_dt = datetime.now()
    orders = active_orders.copy()
    orders["تاريخ ووقت الاستلام"] = pd.to_datetime(orders["تاريخ ووقت الاستلام"], errors="coerce")
    orders = orders[orders["تاريخ ووقت الاستلام"].notna()].copy()
    if orders.empty: return empty_pack

    orders["وقت التسليم للمندوب"] = pd.to_datetime(orders["وقت التسليم للمندوب"], errors="coerce")
    orders["تم التسليم للمندوب؟"] = orders["وقت التسليم للمندوب"].notna()
    orders["دقائق التأخير المعتمدة"] = orders.apply(lambda r: (r["وقت التسليم للمندوب"] - r["تاريخ ووقت الاستلام"]).total_seconds()/60 if r["تم التسليم للمندوب؟"] else (now_dt - r["تاريخ ووقت الاستلام"]).total_seconds()/60, axis=1)
    
    # فلترة المتأخرة الفهرسية المعتمدة على السماحية المحددة
    late_mask = orders["دقائق التأخير المعتمدة"] > grace_minutes
    late_orders_df = orders[late_mask].copy()
    
    # بناء الأوعية الفرعية للتقارير
    empty_pack["late_orders"] = late_orders_df
    empty_pack["late_pending_late"] = late_orders_df[~late_orders_df["تم التسليم للمندوب؟"]]
    empty_pack["late_delivered_late"] = late_orders_df[late_orders_df["تم التسليم للمندوب؟"]]
    empty_pack["late_action_center"] = late_orders_df
    empty_pack["late_summary"] = pd.DataFrame([{"المؤشر": "إجمالي المتأخرات المكتشفة", "القيمة": len(late_orders_df)}])
    return empty_pack


def format_int(value):
    try: return f"{int(round(float(value))):,}"
    except: return "0"


def format_money(value):
    try: return f"{float(value):,.0f} SAR"
    except: return "0 SAR"


def short_label(value, max_len=28):
    text = str(value).strip()
    if len(text) <= max_len: return text
    return text[:max_len - 1] + "…"


def improve_heatmap_text_contrast(fig, values_df):
    return fig


def fig_layout(fig, height=420):
    fig.update_layout(height=height, template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(15,23,42,.20)", font=dict(family="Cairo, Arial", color="#e5e7eb", size=12))
    return fig


def chart_config():
    return {"responsive": True, "displayModeBar": False}


def display_df(df, height=420, label="عرض الجدول"):
    if df is None or df.empty:
        st.info("لا توجد بيانات متاحة لعرضها حالياً.")
        return
    with st.expander(f"📋 {label} — {format_int(len(df))} صف", expanded=False):
        st.dataframe(df, use_container_width=True, height=height)


def write_excel_sheet(writer, df, sheet_name):
    if df is None or df.empty: df = pd.DataFrame({"ملاحظة": ["لا توجد بيانات"]})
    df.to_excel(writer, sheet_name=sheet_name[:31], index=False)


def build_excel_export(reports, filters_summary):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        write_excel_sheet(writer, filters_summary, "Executive Summary")
        for k, v in reports.items():
            if isinstance(v, pd.DataFrame): write_excel_sheet(writer, v, k)
    output.seek(0)
    return output

# ============================================================
# Main Render Logics
# ============================================================

st.sidebar.markdown("## 🔁 تحديث حقيقي ومتزامن")
refresh_options = {"إيقاف": 0, "كل دقيقة": 60, "كل 5 دقائق": 300}
refresh_label = st.sidebar.selectbox("معدل التحديث التلقائي", list(refresh_options.keys()), index=1)

st.sidebar.markdown("## ⚙️ التحكم في مصدر التشغيل")
source_type = st.sidebar.radio("المصدر الأساسي", ["Supabase Direct DB", "رفع ملف يدوياً"], index=0)

raw_df = pd.DataFrame()
if source_type == "Supabase Direct DB":
    with st.spinner("جاري الاتصال وسحب البيانات الحية من Supabase..."):
        raw_df = load_data_from_supabase_table()
else:
    uploaded = st.sidebar.file_uploader("ارفع ملف الطلبات البديل", type=["csv", "xlsx"])
    if uploaded: raw_df = read_uploaded_file(uploaded)

if raw_df.empty:
    st.warning("⚠️ قاعدة البيانات فارغة أو لم يتم تهيئة الـ Secrets بشكل سليم.")
    st.stop()

# تحضير ومعالجة البيانات
items_all, orders_all, cols = prepare_data(raw_df)

# الفلاتر والتقارير الفرعية التشغيلية والإدارية المتقدمة
branches = sorted(items_all["الفرع"].dropna().unique().tolist())
selected_branches = st.sidebar.multiselect("تصفية الفروع التشغيلية", branches, default=branches)

filtered = items_all[items_all["الفرع"].isin(selected_branches)].copy()
_, orders_filtered, _ = prepare_data(filtered)

reports = build_reports(filtered, orders_filtered)
reports["production_queue"] = build_production_queue(reports["items_active"])
reports["action_center"] = build_action_center(reports["items_active"])
reports.update(build_v83_advanced_reports(filtered, reports["items_active"], reports["orders_active"], reports))
reports.update(build_late_orders_reports(reports["items_active"], reports["orders_active"]))

# حساب وعرض مؤشرات الأداء الأساسية KPIs
total_orders = int(reports["orders_active"]["رقم الطلب الموحد"].nunique()) if not reports["orders_active"].empty else 0
total_sales = float(reports["orders_active"]["قيمة الطلب"].sum()) if not reports["orders_active"].empty else 0.0

k1, k2 = st.columns(2)
with k1: st.metric("إجمالي عدد الطلبات الفريدة", format_int(total_orders))
with k2: st.metric("إجمالي صافي المبيعات (بدون تكرار)", format_money(total_sales))

# =========================
# Tabs Layout Render
# =========================
tab_daily, tab_production, tab_action_center, tab_late, tab_advanced, tab_export = st.tabs([
    "📍 Daily Ops", "🏭 Production Queue", "✅ Action Center", "⏰ Late Orders", "📊 Advanced Reports", "⬇️ Export"
])

with tab_daily:
    st.markdown('<div class="section-title">📍 Daily Operations Center</div>', unsafe_allow_html=True)
    display_df(reports["orders_active"], label="كل طلبات العمليات النشطة")

with tab_production:
    st.markdown('<div class="section-title">🏭 Production Control Queue</div>', unsafe_allow_html=True)
    display_df(reports["production_queue"], label="جدول خط الإنتاج والتجهيز المتزامن")

with tab_action_center:
    st.markdown('<div class="section-title">✅ Action Center</div>', unsafe_allow_html=True)
    display_df(reports["action_center"], label="طلبات بحاجة لتأكيد أو تواصل خارجي")

with tab_late:
    st.markdown('<div class="section-title">⏰ Late Orders & Delivery Insights</div>', unsafe_allow_html=True)
    display_df(reports["late_orders"], label="كافة الطلبات المتأخرة عن الموعد التشغيلي")

with tab_advanced:
    st.markdown('<div class="section-title">📊 Management & Strategy Advanced Reports</div>', unsafe_allow_html=True)
    display_df(reports.get("advanced_executive_summary"), label="الملخص التنفيذي الإداري")
    display_df(reports.get("advanced_branch_ranking"), label="ترتيب كفاءة ومبيعات الفروع")
    display_df(reports.get("advanced_product_value"), label="مصفوفة تقييم قيمة وحصة المنتجات")
    display_df(reports.get("advanced_filling_intelligence"), label="تحليل ذكاء حشوات المنتجات")

with tab_export:
    st.markdown('<div class="section-title">⬇️ Download Center</div>', unsafe_allow_html=True)
    filters_summary = pd.DataFrame([{"المؤشر": "إجمالي المبيعات المستخرجة", "القيمة": total_sales}])
    excel_file = build_excel_export(reports, filters_summary)
    st.download_button("⬇️ استخراج ملف تقارير العمليات الشامل (Excel)", data=excel_file, file_name="MAD_Production_Supabase_Report.xlsx")
