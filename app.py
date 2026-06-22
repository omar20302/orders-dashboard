import re
import html
from io import BytesIO
from urllib.parse import urlparse, parse_qs
from datetime import datetime

import pandas as pd
import streamlit as st
try:
    from streamlit_autorefresh import st_autorefresh
except Exception:
    st_autorefresh = None
import plotly.express as px
import plotly.graph_objects as go

from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import ColorScaleRule


# ============================================================
# MAD Orders Dashboard V8.3.4 - Fixed Branch Mapping
# ============================================================

DEFAULT_GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1Lf7R_G5hZ6KvyE5OyRc78b1dKVjD1bEDeeZnorANrxI/edit?usp=sharing"
APP_VERSION = "V8.5.4 Global Chart Guard"


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
        section[data-testid="stSidebar"][aria-expanded="true"] { min-width: 285px !important; max-width: 92vw !important; }
        section[data-testid="stSidebar"][aria-expanded="false"] { min-width: 0 !important; width: 0 !important; max-width: 0 !important; overflow: hidden !important; }
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

    
    /* V8.1 Clean chart and hidden tables */
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

    
    /* =========================
       V8.2 Readable filters
       ========================= */

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

    </style>
    """,
    unsafe_allow_html=True,
)



# =========================
# V8.4.1 Sidebar Collapse Fix
# يمنع ظهور نص القائمة الجانبية بشكل عمودي عند إخفائها
# =========================
st.markdown(
    """
    <style>
    section[data-testid="stSidebar"][aria-expanded="false"] {
        min-width: 0 !important;
        width: 0 !important;
        max-width: 0 !important;
        overflow: hidden !important;
        padding: 0 !important;
        margin: 0 !important;
        border: 0 !important;
    }

    section[data-testid="stSidebar"][aria-expanded="false"] > div {
        display: none !important;
        visibility: hidden !important;
        width: 0 !important;
        min-width: 0 !important;
        max-width: 0 !important;
        overflow: hidden !important;
    }

    section[data-testid="stSidebar"][aria-expanded="true"] {
        overflow: visible !important;
    }

    @media (max-width: 768px) {
        section[data-testid="stSidebar"][aria-expanded="true"] {
            min-width: 285px !important;
            max-width: 92vw !important;
        }

        section[data-testid="stSidebar"][aria-expanded="false"] {
            min-width: 0 !important;
            width: 0 !important;
            max-width: 0 !important;
            overflow: hidden !important;
        }

        section[data-testid="stSidebar"][aria-expanded="false"] * {
            display: none !important;
            visibility: hidden !important;
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
    """
    Fixed branch mapping based on the Chef/Branch name coming from the source sheet.
    Generic Madness And Desire branch is treated as العقيق.
    """
    raw = "" if pd.isna(chef_name) else str(chef_name)
    text = normalize_arabic_digits(raw)
    text = re.sub(r"\s+", " ", text).strip()
    low = text.lower()

    if "قرطبة" in text or "qurtuba" in low or "qortoba" in low or "qurtobah" in low:
        return "قرطبة"

    if "عريجاء" in text or "عريجا" in text or "العريجاء" in text or "uraija" in low or "uraijaa" in low or "urejha" in low:
        return "عريجاء"

    if "الروضة" in text or "روضه" in text or "روضة" in text or "rawdah" in low or "rawda" in low:
        return "الروضة"

    if "العارض" in text or "عارض" in text or "arid" in low or "al arid" in low or "alarid" in low:
        return "العارض"

    if "الورود" in text or "ورود" in text or "worood" in low or "al worood" in low or "alworood" in low:
        return "الورود"

    if "العقيق" in text or "عقيق" in text or "aqiq" in low or "al aqiq" in low or "alaqiq" in low:
        return "العقيق"

    # Generic Madness And Desire branch = Al Aqiq.
    if (
        "madness and desire" in low
        or "madness" in low
        or "مادنيس اند ديزاير" in text
        or "مادنس اند ديزاير" in text
        or "مادنيس" in text
        or "مادنس" in text
    ):
        return "العقيق"

    # Final fallback: route unknown/blank rows to العقيق so no undefined branch appears.
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
    patterns = [
        r"\b\d{4}[-/]\d{1,2}[-/]\d{1,2}\b",
        r"\b\d{1,2}[-/]\d{1,2}[-/]\d{2,4}\b",
    ]
    return any(re.search(p, text) for p in patterns)


def text_has_time(value):
    text = normalize_arabic_digits(value).lower()
    if not text:
        return False
    return bool(re.search(r"\b\d{1,2}:\d{2}\s*(am|pm|ص|م)?\b", text) or re.search(r"\b\d{1,2}\s*(am|pm|ص|م)\b", text))


def parse_any_datetime_with_date(value):
    """Parse only values that actually contain a date; avoids converting '5:00 PM' to today's date."""
    text = normalize_arabic_digits(value)
    if not text or not text_has_date(text):
        return pd.NaT
    for dayfirst in [True, False]:
        parsed = pd.to_datetime(text, errors="coerce", dayfirst=dayfirst)
        if not pd.isna(parsed):
            return parsed
    return pd.NaT


def parse_time_only_value(value):
    """Return time string from time-only values, without creating a fake date."""
    text = normalize_arabic_digits(value).strip()
    if not text or text_has_date(text):
        return ""
    # Normalize Arabic AM/PM if present
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
    # fallback for values already like 5:00 PM
    if text_has_time(text):
        return text
    return ""


def format_date_iso(dt):
    if pd.isna(dt):
        return ""
    return pd.Timestamp(dt).strftime("%Y-%m-%d")


def format_time_12h(dt):
    if pd.isna(dt):
        return ""
    ts = pd.Timestamp(dt)
    hour = int(ts.hour)
    minute = int(ts.minute)
    period = "AM" if hour < 12 else "PM"
    hour12 = hour % 12 or 12
    return f"{hour12}:{minute:02d} {period}"


def clean_delivery_pickup_values(date_value, time_value):
    """
    Cleans cases where date+time is placed in the wrong column.
    Example:
    Delivery Date = ''
    Pickup Time = '2026-06-22 5:00 PM'
    returns:
    ('2026-06-22', '5:00 PM', True)
    """
    raw_date = normalize_arabic_digits(date_value).strip()
    raw_time = normalize_arabic_digits(time_value).strip()

    date_dt = parse_any_datetime_with_date(raw_date)
    time_dt = parse_any_datetime_with_date(raw_time)

    cleaned = False

    # Case 1: Pickup Time contains full datetime.
    if not pd.isna(time_dt):
        cleaned_date = format_date_iso(time_dt)
        cleaned_time = format_time_12h(time_dt)

        # If delivery date is valid and different, keep the valid delivery date but still clean the time.
        if not pd.isna(date_dt):
            cleaned_date = format_date_iso(date_dt)

        if raw_date != cleaned_date or raw_time != cleaned_time:
            cleaned = True
        return cleaned_date, cleaned_time, cleaned

    # Case 2: Delivery Date contains full datetime.
    if not pd.isna(date_dt):
        cleaned_date = format_date_iso(date_dt)

        if text_has_time(raw_date):
            cleaned_time = format_time_12h(date_dt)
            cleaned = True
        else:
            cleaned_time = parse_time_only_value(raw_time) or raw_time

        if raw_date != cleaned_date or (raw_time and raw_time != cleaned_time):
            cleaned = True
        return cleaned_date, cleaned_time, cleaned

    # Case 3: normal date/time or missing date.
    cleaned_date = raw_date
    cleaned_time = parse_time_only_value(raw_time) or raw_time
    if raw_time != cleaned_time:
        cleaned = True
    return cleaned_date, cleaned_time, cleaned


def parse_datetime_parts(date_value, time_value):
    date_value = normalize_arabic_digits(date_value)
    time_value = normalize_arabic_digits(time_value)

    # Do not parse time-only values by themselves because Pandas may attach today's date.
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


@st.cache_data(ttl=60, show_spinner=False)
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
    df["تاريخ التوصيل قبل التنظيف"] = col_or_blank(df, cols["delivery_date"])
    df["وقت الاستلام قبل التنظيف"] = col_or_blank(df, cols["pickup_time"])

    cleaned_datetime_parts = [
        clean_delivery_pickup_values(d, t)
        for d, t in zip(df["تاريخ التوصيل قبل التنظيف"], df["وقت الاستلام قبل التنظيف"])
    ]

    df["تاريخ التوصيل الأصلي"] = [x[0] for x in cleaned_datetime_parts]
    df["وقت الاستلام الأصلي"] = [x[1] for x in cleaned_datetime_parts]
    df["تم تنظيف التاريخ/الوقت؟"] = [bool(x[2]) for x in cleaned_datetime_parts]

    df["تاريخ ووقت الاستلام"] = [
        parse_datetime_parts(d, t)
        for d, t in zip(df["تاريخ التوصيل الأصلي"], df["وقت الاستلام الأصلي"])
    ]
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
    # Safe sorting for Streamlit Cloud / newer Pandas versions:
    # avoid sorting directly on mixed/object/category-like columns.
    df["_sort_datetime_safe"] = pd.to_datetime(df["تاريخ ووقت الاستلام"], errors="coerce")
    df["_sort_order_safe"] = df["رقم الطلب الموحد"].astype(str)
    group = (
        df.sort_values(["_sort_datetime_safe", "_sort_order_safe"], na_position="last")
        .groupby("رقم الطلب الموحد", dropna=False)
    )
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

    df = df.drop(columns=["_sort_datetime_safe", "_sort_order_safe"], errors="ignore")
    order_level = order_level.drop(columns=["_sort_datetime_safe", "_sort_order_safe"], errors="ignore")
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
    if "تم تنظيف التاريخ/الوقت؟" in items.columns:
        add_issue("تم تنظيف التاريخ/الوقت تلقائيًا", items["تم تنظيف التاريخ/الوقت؟"].astype(bool), "منخفضة")
    add_issue("فرع غير محدد", items["الفرع"].eq("بدون فرع محدد"), "متوسطة")
    add_issue("حالة طلب ناقصة", items["الحالة"].astype(str).str.strip().isin(["", "غير محدد"]), "متوسطة")
    add_issue("منتج بدون اسم", items["المنتج"].eq("بدون اسم منتج"), "عالية")
    add_issue("حشوة ناقصة للمنتجات", (~items["إضافة؟"]) & items["الحشوة"].astype(str).str.strip().eq(""), "منخفضة")
    add_issue("قيمة طلب صفرية", items["قيمة الطلب رقم"].fillna(0).eq(0), "متوسطة")
    reports["data_quality_summary"] = pd.DataFrame(quality_rows)

    quality_detail = items[
        items["تاريخ التوصيل الأصلي"].astype(str).str.strip().eq("") |
        items["وقت الاستلام الأصلي"].astype(str).str.strip().eq("") |
        items["الفرع"].eq("بدون فرع محدد") |
        items["المنتج"].eq("بدون اسم منتج") |
        items["قيمة الطلب رقم"].fillna(0).eq(0)
    ].copy()
    reports["data_quality_details"] = quality_detail[[c for c in [
        "رقم الطلب الظاهر", "رقم الطلب الموحد", "الفرع", "الحالة",
        "تاريخ التوصيل قبل التنظيف", "وقت الاستلام قبل التنظيف",
        "تاريخ التوصيل الأصلي", "وقت الاستلام الأصلي", "تم تنظيف التاريخ/الوقت؟",
        "العميل", "المنتج", "قيمة الطلب رقم"
    ] if c in quality_detail.columns]]

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



def first_or_dash(series):
    try:
        s = series.dropna()
        if len(s):
            return str(s.iloc[0])
    except Exception:
        pass
    return "-"


def score_label(score):
    try:
        score = float(score)
    except Exception:
        return "Unknown"
    if score >= 90:
        return "Excellent"
    if score >= 75:
        return "Good"
    if score >= 55:
        return "Needs Attention"
    return "Critical"


def build_v83_advanced_reports(items, active_items, active_orders, reports):
    """Advanced management reports pack for V8.3."""
    advanced = {}

    if items is None:
        items = pd.DataFrame()
    if active_items is None:
        active_items = pd.DataFrame()
    if active_orders is None:
        active_orders = pd.DataFrame()

    non_addon = active_items[~active_items["إضافة؟"]].copy() if not active_items.empty and "إضافة؟" in active_items.columns else pd.DataFrame()
    addons = active_items[active_items["إضافة؟"]].copy() if not active_items.empty and "إضافة؟" in active_items.columns else pd.DataFrame()

    # -------------------------
    # Executive Summary
    # -------------------------
    total_orders_adv = int(active_orders["رقم الطلب الموحد"].nunique()) if not active_orders.empty and "رقم الطلب الموحد" in active_orders.columns else 0
    total_sales_adv = float(active_orders["قيمة الطلب"].fillna(0).sum()) if not active_orders.empty and "قيمة الطلب" in active_orders.columns else 0.0
    avg_order_adv = float(active_orders["قيمة الطلب"].fillna(0).mean()) if not active_orders.empty and "قيمة الطلب" in active_orders.columns else 0.0
    addon_orders_adv = int(active_orders["فيه إضافات"].sum()) if not active_orders.empty and "فيه إضافات" in active_orders.columns else 0
    action_orders_adv = int(active_orders["يحتاج متابعة"].sum()) if not active_orders.empty and "يحتاج متابعة" in active_orders.columns else 0
    upsell_rate_adv = round(addon_orders_adv / total_orders_adv * 100, 1) if total_orders_adv else 0.0
    action_rate_adv = round(action_orders_adv / total_orders_adv * 100, 1) if total_orders_adv else 0.0

    top_branch_adv = "-"
    top_branch_sales_adv = 0.0
    if not active_orders.empty:
        branch_sales_tmp = active_orders.groupby("الفرع", dropna=False)["قيمة الطلب"].sum().sort_values(ascending=False)
        if len(branch_sales_tmp):
            top_branch_adv = str(branch_sales_tmp.index[0])
            top_branch_sales_adv = float(branch_sales_tmp.iloc[0])

    top_product_adv = "-"
    if not non_addon.empty:
        prod_qty_tmp = non_addon.groupby("المنتج", dropna=False)["الكمية رقم"].sum().sort_values(ascending=False)
        if len(prod_qty_tmp):
            top_product_adv = str(prod_qty_tmp.index[0])

    top_filling_adv = "-"
    if not non_addon.empty and "الحشوة" in non_addon.columns:
        fill_qty_tmp = non_addon[non_addon["الحشوة"].astype(str).str.strip().ne("")].groupby("الحشوة", dropna=False)["الكمية رقم"].sum().sort_values(ascending=False)
        if len(fill_qty_tmp):
            top_filling_adv = str(fill_qty_tmp.index[0])

    top_hour_range_adv = "-"
    if not active_orders.empty and "ساعة رقم" in active_orders.columns:
        hr = active_orders.copy()
        hr["نطاق ساعة الاستلام"] = hr["ساعة رقم"].apply(hour_range_label)
        hour_tmp = hr.groupby("نطاق ساعة الاستلام", dropna=False)["رقم الطلب الموحد"].nunique().sort_values(ascending=False)
        if len(hour_tmp):
            top_hour_range_adv = str(hour_tmp.index[0])

    advanced["advanced_executive_summary"] = pd.DataFrame([
        {"المؤشر": "عدد الطلبات", "القيمة": total_orders_adv, "ملاحظة": "Unique Order Id"},
        {"المؤشر": "إجمالي المبيعات", "القيمة": total_sales_adv, "ملاحظة": "بدون تكرار قيمة الطلب"},
        {"المؤشر": "متوسط قيمة الطلب AOV", "القيمة": avg_order_adv, "ملاحظة": "Sales / Orders"},
        {"المؤشر": "نسبة الطلبات بإضافات", "القيمة": upsell_rate_adv, "ملاحظة": "%"},
        {"المؤشر": "نسبة الطلبات التي تحتاج متابعة", "القيمة": action_rate_adv, "ملاحظة": "%"},
        {"المؤشر": "أعلى فرع بالمبيعات", "القيمة": top_branch_adv, "ملاحظة": format_money(top_branch_sales_adv) if "format_money" in globals() else str(top_branch_sales_adv)},
        {"المؤشر": "أعلى منتج بالكمية", "القيمة": top_product_adv, "ملاحظة": ""},
        {"المؤشر": "أعلى حشوة بالكمية", "القيمة": top_filling_adv, "ملاحظة": ""},
        {"المؤشر": "أعلى نطاق ساعة ضغط", "القيمة": top_hour_range_adv, "ملاحظة": ""},
    ])

    # -------------------------
    # Branch Ranking
    # -------------------------
    if not active_orders.empty:
        branch_ranking = active_orders.groupby("الفرع", dropna=False).agg(
            الطلبات=("رقم الطلب الموحد", "nunique"),
            المبيعات=("قيمة الطلب", "sum"),
            متوسط_الطلب=("قيمة الطلب", "mean"),
            طلبات_بإضافات=("فيه إضافات", "sum"),
            طلبات_تحتاج_متابعة=("يحتاج متابعة", "sum"),
            عدد_الأصناف=("عدد الأصناف", "sum"),
        ).reset_index()

        if not active_items.empty:
            qty_branch = active_items.groupby("الفرع", dropna=False)["الكمية رقم"].sum().reset_index(name="إجمالي_الكمية")
            branch_ranking = branch_ranking.merge(qty_branch, on="الفرع", how="left")
        else:
            branch_ranking["إجمالي_الكمية"] = 0

        branch_ranking["نسبة_Upsell_%"] = (branch_ranking["طلبات_بإضافات"] / branch_ranking["الطلبات"].replace(0, pd.NA) * 100).fillna(0).round(1)
        branch_ranking["نسبة_متابعة_%"] = (branch_ranking["طلبات_تحتاج_متابعة"] / branch_ranking["الطلبات"].replace(0, pd.NA) * 100).fillna(0).round(1)
        branch_ranking["رتبة_المبيعات"] = branch_ranking["المبيعات"].rank(method="dense", ascending=False).astype(int)
        branch_ranking["رتبة_الطلبات"] = branch_ranking["الطلبات"].rank(method="dense", ascending=False).astype(int)
        branch_ranking = branch_ranking.sort_values(["رتبة_المبيعات", "رتبة_الطلبات"])
    else:
        branch_ranking = pd.DataFrame()
    advanced["advanced_branch_ranking"] = branch_ranking

    # -------------------------
    # Product Value Report
    # -------------------------
    if not non_addon.empty:
        product_value = non_addon.groupby("المنتج", dropna=False).agg(
            الطلبات=("رقم الطلب الموحد", "nunique"),
            الكمية=("الكمية رقم", "sum"),
            المبيعات=("إجمالي المنتج رقم", "sum"),
            متوسط_السعر=("سعر الحبة رقم", "mean"),
            المتابعة=("يحتاج متابعة؟", "sum"),
        ).reset_index()
        product_value["مبيعات_لكل_طلب"] = (product_value["المبيعات"] / product_value["الطلبات"].replace(0, pd.NA)).fillna(0).round(2)
        product_value["نسبة_متابعة_%"] = (product_value["المتابعة"] / product_value["الطلبات"].replace(0, pd.NA) * 100).fillna(0).round(1)
        product_value["حصة_المبيعات_%"] = (product_value["المبيعات"] / product_value["المبيعات"].sum() * 100).fillna(0).round(1)

        q_qty = product_value["الكمية"].quantile(0.65) if len(product_value) else 0
        q_val = product_value["مبيعات_لكل_طلب"].quantile(0.65) if len(product_value) else 0
        def product_segment(row):
            if row["الكمية"] >= q_qty and row["مبيعات_لكل_طلب"] >= q_val:
                return "Hero Product"
            if row["الكمية"] >= q_qty and row["مبيعات_لكل_طلب"] < q_val:
                return "Volume Driver"
            if row["الكمية"] < q_qty and row["مبيعات_لكل_طلب"] >= q_val:
                return "High Value Niche"
            return "Low Priority"
        product_value["تصنيف_القيمة"] = product_value.apply(product_segment, axis=1)
        product_value = product_value.sort_values(["المبيعات", "الكمية"], ascending=False)
    else:
        product_value = pd.DataFrame()
    advanced["advanced_product_value"] = product_value

    # -------------------------
    # Filling Intelligence
    # -------------------------
    if not non_addon.empty and "الحشوة" in non_addon.columns:
        filling_src = non_addon[non_addon["الحشوة"].astype(str).str.strip().ne("")].copy()
        if not filling_src.empty:
            filling_int = filling_src.groupby("الحشوة", dropna=False).agg(
                الطلبات=("رقم الطلب الموحد", "nunique"),
                الكمية=("الكمية رقم", "sum"),
                المبيعات=("إجمالي المنتج رقم", "sum"),
                متوسط_السعر=("سعر الحبة رقم", "mean"),
                المتابعة=("يحتاج متابعة؟", "sum"),
            ).reset_index()
            filling_int["نسبة_متابعة_%"] = (filling_int["المتابعة"] / filling_int["الطلبات"].replace(0, pd.NA) * 100).fillna(0).round(1)
            filling_int["حصة_الكمية_%"] = (filling_int["الكمية"] / filling_int["الكمية"].sum() * 100).fillna(0).round(1)
            filling_int["حصة_المبيعات_%"] = (filling_int["المبيعات"] / filling_int["المبيعات"].sum() * 100).fillna(0).round(1)
            filling_int["أولوية_الإنتاج"] = pd.cut(
                filling_int["الكمية"].rank(method="first"),
                bins=[0, max(1, len(filling_int)*0.33), max(2, len(filling_int)*0.66), max(3, len(filling_int))],
                labels=["منخفضة", "متوسطة", "عالية"],
                include_lowest=True,
            ).astype(str)
            filling_int = filling_int.sort_values(["الكمية", "المبيعات"], ascending=False)

            filling_by_branch_top = filling_src.groupby(["الفرع", "الحشوة"], dropna=False).agg(
                الطلبات=("رقم الطلب الموحد", "nunique"),
                الكمية=("الكمية رقم", "sum"),
                المبيعات=("إجمالي المنتج رقم", "sum"),
            ).reset_index()
            filling_by_branch_top["رتبة_داخل_الفرع"] = filling_by_branch_top.groupby("الفرع")["الكمية"].rank(method="dense", ascending=False).astype(int)
            filling_by_branch_top = filling_by_branch_top.sort_values(["الفرع", "رتبة_داخل_الفرع"])
        else:
            filling_int = pd.DataFrame()
            filling_by_branch_top = pd.DataFrame()
    else:
        filling_int = pd.DataFrame()
        filling_by_branch_top = pd.DataFrame()
    advanced["advanced_filling_intelligence"] = filling_int
    advanced["advanced_filling_by_branch_rank"] = filling_by_branch_top

    # -------------------------
    # Add-ons Opportunity
    # -------------------------
    if not active_orders.empty:
        addon_branch = active_orders.groupby("الفرع", dropna=False).agg(
            الطلبات=("رقم الطلب الموحد", "nunique"),
            طلبات_بإضافات=("فيه إضافات", "sum"),
            متوسط_الطلب=("قيمة الطلب", "mean"),
        ).reset_index()
        addon_branch["طلبات_بدون_إضافات"] = addon_branch["الطلبات"] - addon_branch["طلبات_بإضافات"]
        addon_branch["نسبة_Upsell_%"] = (addon_branch["طلبات_بإضافات"] / addon_branch["الطلبات"].replace(0, pd.NA) * 100).fillna(0).round(1)

        avg_with = active_orders[active_orders["فيه إضافات"].astype(bool)].groupby("الفرع", dropna=False)["قيمة الطلب"].mean().reset_index(name="متوسط_مع_إضافات")
        avg_without = active_orders[~active_orders["فيه إضافات"].astype(bool)].groupby("الفرع", dropna=False)["قيمة الطلب"].mean().reset_index(name="متوسط_بدون_إضافات")
        addon_branch = addon_branch.merge(avg_with, on="الفرع", how="left").merge(avg_without, on="الفرع", how="left")
        addon_branch["فرق_المتوسط"] = (addon_branch["متوسط_مع_إضافات"].fillna(0) - addon_branch["متوسط_بدون_إضافات"].fillna(0)).round(2)
        addon_branch["فرصة"] = addon_branch["نسبة_Upsell_%"].apply(lambda x: "فرصة عالية" if x < 25 else ("متوسطة" if x < 45 else "جيد"))
        addon_branch = addon_branch.sort_values(["فرصة", "نسبة_Upsell_%", "طلبات_بدون_إضافات"], ascending=[True, True, False])
    else:
        addon_branch = pd.DataFrame()
    advanced["advanced_addons_opportunity_branch"] = addon_branch

    if not non_addon.empty and not active_orders.empty:
        order_addon_flag = active_orders[["رقم الطلب الموحد", "فيه إضافات"]].drop_duplicates()
        prod_orders = non_addon[["رقم الطلب الموحد", "المنتج"]].drop_duplicates().merge(order_addon_flag, on="رقم الطلب الموحد", how="left")
        prod_addon = prod_orders.groupby("المنتج", dropna=False).agg(
            الطلبات=("رقم الطلب الموحد", "nunique"),
            طلبات_بإضافات=("فيه إضافات", "sum"),
        ).reset_index()
        prod_addon["طلبات_بدون_إضافات"] = prod_addon["الطلبات"] - prod_addon["طلبات_بإضافات"]
        prod_addon["نسبة_Upsell_%"] = (prod_addon["طلبات_بإضافات"] / prod_addon["الطلبات"].replace(0, pd.NA) * 100).fillna(0).round(1)
        prod_addon["اقتراح Bundle"] = prod_addon["المنتج"].apply(lambda x: f"{short_label(x, 28)} + Candles/Gift Card")
        prod_addon = prod_addon.sort_values(["طلبات_بدون_إضافات", "الطلبات"], ascending=False)
    else:
        prod_addon = pd.DataFrame()
    advanced["advanced_addons_product_opportunity"] = prod_addon

    # -------------------------
    # Notes Intelligence
    # -------------------------
    if not active_items.empty:
        note_src = active_items.copy()
        note_src["طول_الملاحظة"] = note_src["الملاحظة"].astype(str).str.len()
        notes_product = note_src.groupby("المنتج", dropna=False).agg(
            عدد_الطلبات=("رقم الطلب الموحد", "nunique"),
            متوسط_طول_الملاحظة=("طول_الملاحظة", "mean"),
            تحتاج_متابعة=("يحتاج متابعة؟", "sum"),
        ).reset_index()
        notes_product["نسبة_متابعة_%"] = (notes_product["تحتاج_متابعة"] / notes_product["عدد_الطلبات"].replace(0, pd.NA) * 100).fillna(0).round(1)
        notes_product = notes_product.sort_values(["تحتاج_متابعة", "متوسط_طول_الملاحظة"], ascending=False)

        keyword_map = {
            "صورة / Photo": ["صورة", "الصورة", "الصوره", "photo", "picture", "image"],
            "كتابة / Writing": ["كتابة", "اكتب", "العبارة", "write", "writing"],
            "تواصل / Contact": ["تواصل", "اتصال", "واتساب", "contact", "call", "whatsapp", "text"],
            "تعديل تصميم": ["تعديل", "تصميم", "لون", "قلوب", "draw", "design", "color", "hearts"],
            "مشكلة / Problem": ["مشكلة", "خطأ", "غلط", "problem", "wrong", "mistake"],
        }
        notes_text = note_src["الملاحظة"].fillna("").astype(str).str.lower()
        keyword_rows = []
        for label, keys in keyword_map.items():
            mask = notes_text.apply(lambda x: any(k.lower() in x for k in keys))
            keyword_rows.append({"الكلمة/التصنيف": label, "عدد الصفوف": int(mask.sum()), "نسبة من الصفوف %": round(mask.mean()*100, 1) if len(mask) else 0})
        notes_keywords = pd.DataFrame(keyword_rows).sort_values("عدد الصفوف", ascending=False)
    else:
        notes_product = pd.DataFrame()
        notes_keywords = pd.DataFrame()
    advanced["advanced_notes_by_product"] = notes_product
    advanced["advanced_notes_keywords"] = notes_keywords

    # -------------------------
    # Data Quality Score
    # -------------------------
    if not items.empty:
        quality = items.copy()
        issue_masks = pd.DataFrame({
            "تاريخ_ناقص": quality["تاريخ التوصيل الأصلي"].astype(str).str.strip().eq(""),
            "وقت_ناقص": quality["وقت الاستلام الأصلي"].astype(str).str.strip().eq(""),
            "فرع_غير_محدد": quality["الفرع"].eq("بدون فرع محدد"),
            "منتج_بدون_اسم": quality["المنتج"].eq("بدون اسم منتج"),
            "قيمة_صفرية": quality["قيمة الطلب رقم"].fillna(0).eq(0),
            "حشوة_ناقصة": (~quality["إضافة؟"]) & quality["الحشوة"].astype(str).str.strip().eq(""),
        })
        quality["عدد_مشاكل_الجودة"] = issue_masks.sum(axis=1)
        quality_score = quality.groupby("الفرع", dropna=False).agg(
            الصفوف=("رقم الطلب الموحد", "count"),
            الطلبات=("رقم الطلب الموحد", "nunique"),
            إجمالي_مشاكل_الجودة=("عدد_مشاكل_الجودة", "sum"),
        ).reset_index()
        quality_score["مشاكل_لكل_100_صف"] = (quality_score["إجمالي_مشاكل_الجودة"] / quality_score["الصفوف"].replace(0, pd.NA) * 100).fillna(0).round(1)
        quality_score["Quality Score"] = (100 - quality_score["مشاكل_لكل_100_صف"]).clip(lower=0, upper=100).round(1)
        quality_score["التقييم"] = quality_score["Quality Score"].apply(score_label)
        quality_score = quality_score.sort_values(["Quality Score", "إجمالي_مشاكل_الجودة"], ascending=[True, False])
    else:
        quality_score = pd.DataFrame()
    advanced["advanced_data_quality_score"] = quality_score

    # -------------------------
    # Hourly Capacity
    # -------------------------
    if not active_orders.empty:
        cap_orders = active_orders.copy()
        cap_orders["نطاق ساعة الاستلام"] = cap_orders["ساعة رقم"].apply(hour_range_label)
        hourly_capacity = cap_orders.groupby("نطاق ساعة الاستلام", dropna=False).agg(
            الطلبات=("رقم الطلب الموحد", "nunique"),
            المبيعات=("قيمة الطلب", "sum"),
            متوسط_الطلب=("قيمة الطلب", "mean"),
            تحتاج_متابعة=("يحتاج متابعة", "sum"),
        ).reset_index()
        if not active_items.empty:
            cap_items = active_items.copy()
            cap_items["نطاق ساعة الاستلام"] = cap_items["ساعة رقم"].apply(hour_range_label)
            qty_by_hour = cap_items.groupby("نطاق ساعة الاستلام", dropna=False).agg(
                الكمية=("الكمية رقم", "sum"),
                صفوف_الأصناف=("رقم الطلب الموحد", "count"),
            ).reset_index()
            hourly_capacity = hourly_capacity.merge(qty_by_hour, on="نطاق ساعة الاستلام", how="left")
        else:
            hourly_capacity["الكمية"] = 0
            hourly_capacity["صفوف_الأصناف"] = 0
        hourly_capacity["Action Rate %"] = (hourly_capacity["تحتاج_متابعة"] / hourly_capacity["الطلبات"].replace(0, pd.NA) * 100).fillna(0).round(1)
        hourly_capacity["_sort"] = hourly_capacity["نطاق ساعة الاستلام"].apply(hour_range_sort_value)
        hourly_capacity = hourly_capacity.sort_values("_sort").drop(columns=["_sort"])
    else:
        hourly_capacity = pd.DataFrame()
    advanced["advanced_hourly_capacity"] = hourly_capacity

    # -------------------------
    # Campaign Performance
    # -------------------------
    if not active_items.empty:
        camp_base = active_items.copy()
        camp_perf = camp_base.groupby("الحملة", dropna=False).agg(
            الطلبات=("رقم الطلب الموحد", "nunique"),
            الكمية=("الكمية رقم", "sum"),
            مبيعات_الأصناف=("إجمالي المنتج رقم", "sum"),
            تحتاج_متابعة=("يحتاج متابعة؟", "sum"),
        ).reset_index()
        camp_perf["نسبة_متابعة_%"] = (camp_perf["تحتاج_متابعة"] / camp_perf["الطلبات"].replace(0, pd.NA) * 100).fillna(0).round(1)
        camp_perf = camp_perf.sort_values(["مبيعات_الأصناف", "الطلبات"], ascending=False)

        def top_value_for_campaign(campaign, col):
            sub = camp_base[camp_base["الحملة"].eq(campaign)]
            if sub.empty or col not in sub.columns:
                return "-"
            vc = sub.groupby(col, dropna=False)["الكمية رقم"].sum().sort_values(ascending=False)
            return str(vc.index[0]) if len(vc) else "-"

        if not camp_perf.empty:
            camp_perf["أفضل_فرع"] = camp_perf["الحملة"].apply(lambda c: top_value_for_campaign(c, "الفرع"))
            camp_perf["أفضل_منتج"] = camp_perf["الحملة"].apply(lambda c: top_value_for_campaign(c, "المنتج"))
            camp_perf["أفضل_حشوة"] = camp_perf["الحملة"].apply(lambda c: top_value_for_campaign(c, "الحشوة"))
    else:
        camp_perf = pd.DataFrame()
    advanced["advanced_campaign_performance"] = camp_perf

    return advanced


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
    fig = translate_fig_for_language(fig)
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




def heatmap_height(row_count, min_height=520, max_height=900, row_px=44):
    """Dynamic heatmap height to avoid label overlap."""
    try:
        rows = int(row_count)
    except Exception:
        rows = 8
    return max(min_height, min(max_height, 170 + rows * row_px))




def hour_range_label(hour_value):
    """Convert numeric pickup hour into readable 1-hour range."""
    try:
        h = int(float(hour_value))
    except Exception:
        return "بدون وقت"

    if h < 0 or h > 23:
        return "بدون وقت"

    def fmt(hour_24):
        period = "AM" if hour_24 < 12 else "PM"
        hour_12 = hour_24 % 12
        if hour_12 == 0:
            hour_12 = 12
        return f"{hour_12}:00 {period}"

    return f"{fmt(h)} - {fmt((h + 1) % 24)}"


def hour_range_sort_value(label):
    """Sort hour range labels by the starting hour."""
    text = str(label)
    if text == "بدون وقت":
        return 999

    m = re.search(r"(\d{1,2}):00\s*(AM|PM)", text, flags=re.IGNORECASE)
    if not m:
        return 999

    hour = int(m.group(1))
    period = m.group(2).upper()

    if period == "AM":
        return 0 if hour == 12 else hour
    return 12 if hour == 12 else hour + 12




# =========================
# V8.5 Language Display Helpers
# يحول واجهة العرض والرسوم والجداول بدون تغيير أسماء الأعمدة الأصلية أو منطق التقارير
# =========================

UI_EN = {
    # General
    "عرض الجدول": "Show table",
    "لا توجد بيانات لهذا التقرير ضمن الفلاتر الحالية.": "No data for this report with the current filters.",
    "صف": "rows",
    "عمود": "columns",
    "الكل": "All",
    "عالية": "High",
    "عادية": "Normal",
    "متوسطة": "Medium",
    "منخفضة": "Low",
    "منخفضة": "Low",
    "عالية": "High",
    "منخفض": "Low",
    "جيد": "Good",
    "فرصة عالية": "High opportunity",
    "بدون وقت": "No time",
    "بدون فرع محدد": "Unknown branch",

    # Common columns
    "رقم الطلب": "Order No.",
    "رقم الطلب الظاهر": "Order No.",
    "رقم الطلب الموحد": "Order ID",
    "العميل": "Customer",
    "الفرع": "Branch",
    "الحالة": "Status",
    "تاريخ التحليل": "Analysis date",
    "تاريخ التوصيل الأصلي": "Delivery date",
    "وقت الاستلام": "Pickup time",
    "وقت الاستلام الأصلي": "Pickup time",
    "الساعة": "Hour",
    "ساعة رقم": "Hour no.",
    "قيمة الطلب": "Order value",
    "عدد الأصناف": "Items count",
    "عدد المنتجات": "Products count",
    "عدد الإضافات": "Add-ons count",
    "يحتاج متابعة": "Need action",
    "يحتاج متابعة؟": "Need action?",
    "إضافة؟": "Add-on?",
    "فيه إضافات": "Has add-ons",
    "ملغي": "Cancelled",
    "ملغي؟": "Cancelled?",
    "المنتج": "Product",
    "الحشوة": "Filling",
    "الكمية": "Quantity",
    "الكمية رقم": "Quantity",
    "المبيعات": "Sales",
    "المبيعات_منتجات": "Product sales",
    "متوسط_الطلب": "Average order",
    "عدد_الطلبات": "Orders",
    "طلبات": "Orders",
    "الطلبات": "Orders",
    "عدد الحالات": "Cases",
    "سبب المتابعة": "Action reason",
    "الملاحظة": "Note",
    "رقم الجوال المستخرج": "Extracted mobile",
    "الحملة": "Campaign",
    "تصنيف الإضافة": "Add-on category",
    "نوع الصنف": "Item type",
    "أولوية": "Priority",
    "الأولوية": "Priority",
    "حالة المتابعة": "Follow-up status",
    "نوع الإجراء": "Action type",
    "نطاق ساعة الاستلام": "Pickup hour range",
    "وقت عرض": "Display time",
    "تاريخ عرض": "Display date",
    "إجمالي المنتج رقم": "Item total",
    "سعر الحبة رقم": "Unit price",
    "سعر الحشوة رقم": "Filling price",
    "الخصم": "Discount",
    "البند": "Item",
    "القيمة": "Value",
    "المؤشر": "Metric",
    "ملاحظة": "Note",
    "المشكلة": "Issue",
    "عدد الصفوف": "Rows",
    "الأهمية": "Severity",

    # Chart / section titles
    "ضغط الطلبات حسب الفرع": "Orders pressure by branch",
    "ضغط الطلبات حسب الساعة": "Orders pressure by hour",
    "Heatmap الفرع × الساعة": "Branch × Hour Heatmap",
    "أقرب جدول تشغيل ضمن الفلاتر": "Nearest operating schedule within filters",
    "ضغط التجهيز حسب نطاق الساعة": "Preparation pressure by hour range",
    "طلبات التجهيز حسب الفرع": "Preparation orders by branch",
    "أسباب الإجراءات": "Action reasons",
    "المبيعات حسب الفرع": "Sales by branch",
    "المبيعات حسب الساعة": "Sales by hour",
    "الحالات حسب الفرع": "Statuses by branch",
    "حالات الطلبات حسب الفرع": "Order statuses by branch",
    "أعلى المنتجات حسب الكمية": "Top products by quantity",
    "المنتجات حسب الفرع": "Products by branch",
    "أعلى الحشوات حسب الكمية": "Top fillings by quantity",
    "الحشوات حسب الفرع": "Fillings by branch",
    "الحشوات حسب المنتج": "Fillings by product",
    "الحشوات حسب الساعة": "Fillings by hour",
    "الإضافات حسب التصنيف": "Add-ons by category",
    "الإضافات حسب الفرع": "Add-ons by branch",
    "المنتجات حسب الحملة": "Products by campaign",
    "أفضل منتجات الفرع": "Top branch products",
    "ضغط الفرع حسب الساعة": "Branch pressure by hour",
    "المنتج حسب الفرع": "Product by branch",
    "حشوات المنتج": "Product fillings",
    "كل طلبات المنتج": "All product orders",
    "مشاكل جودة البيانات": "Data quality issues",
    "تفاصيل الصفوف التي تحتاج تنظيف": "Rows that need cleanup",
    "ترتيب الفروع حسب المبيعات": "Branch ranking by sales",
    "خريطة قيمة المنتجات: كمية × مبيعات لكل طلب": "Product value map: quantity × sales per order",
    "أعلى المنتجات بالقيمة": "Top products by value",
    "عرض Executive Summary": "Show Executive Summary",
    "عرض جدول Branch Ranking": "Show Branch Ranking table",
    "عرض جدول Product Value": "Show Product Value table",

    # KPI / labels
    "عدد الطلبات": "Orders",
    "إجمالي المبيعات": "Total sales",
    "متوسط الطلب": "Average order",
    "تحتاج متابعة": "Need action",
    "بدون تكرار قيمة الطلب": "Order value without duplication",
    "صورة / تواصل / كتابة": "Photo / Contact / Writing",
    "طلب بإضافات": "orders with add-ons",
    "طلبات الفرع": "Branch orders",
    "مبيعات الفرع": "Branch sales",
    "طلبات المنتج": "Product orders",
    "كمية المنتج": "Product quantity",
    "مبيعات المنتج": "Product sales",
    "طلبات بإضافات": "Orders with add-ons",
    "أفضل فرع بالمبيعات": "Top branch by sales",
    "أعلى نطاق ساعة": "Peak hour range",

    # Filters
    "فرع التجهيز": "Preparation branch",
    "نطاق ساعة الاستلام": "Pickup hour range",
    "الحالة": "Status",
    "الأولوية": "Priority",
    "إظهار الإضافات": "Show add-ons",
    "فقط ما يحتاج متابعة": "Only need action",
    "بحث داخل التجهيز": "Search preparation queue",
    "رقم طلب / منتج / عميل / جوال": "Order / product / customer / mobile",
    "اختار الفرع": "Choose branch",
    "اختار فرع للتحليل العميق": "Choose branch for deep analysis",
    "اختار المنتج": "Choose product",
}

VALUE_EN = {
    # Branches
    "العقيق": "Al Aqiq",
    "العارض": "Al Arid",
    "عريجاء": "Uraija",
    "الروضة": "Rawdah",
    "قرطبة": "Qurtuba",
    "الورود": "Al Worood",
    "بدون فرع محدد": "Unknown branch",

    # Common values
    "الكل": "All",
    "بدون وقت": "No time",
    "ملغي": "Cancelled",
    "ملغاة": "Cancelled",
    "غير محدد": "Undefined",
    "منتج": "Product",
    "إضافة": "Add-on",
    "عالية": "High",
    "عادية": "Normal",
    "متوسطة": "Medium",
    "منخفضة": "Low",
    "صورة / Photo": "Photo",
    "كتابة / Writing": "Writing",
    "تواصل / Contact": "Contact",
    "مشكلة / Problem": "Problem",
    "تعديل تصميم": "Design change",
}

def is_english_ui():
    return globals().get("IS_AR", True) is False

def ui(text):
    text = "" if text is None else str(text)
    if not is_english_ui():
        return text
    return UI_EN.get(text, VALUE_EN.get(text, text))

def translate_value(value):
    if not is_english_ui():
        return value
    if pd.isna(value):
        return value
    text = str(value)
    if text in VALUE_EN:
        return VALUE_EN[text]
    if text in UI_EN:
        return UI_EN[text]

    # Convert Arabic hour labels مثل 5-6 م / 4-5 ص
    text2 = text
    text2 = re.sub(r"\s*ص\b", " AM", text2)
    text2 = re.sub(r"\s*م\b", " PM", text2)
    text2 = text2.replace("بدون وقت", "No time")
    return text2

def translate_df_for_display(df):
    if df is None or df.empty or not is_english_ui():
        return df
    out = df.copy()
    # Translate visible values only; report logic remains untouched.
    for col in out.columns:
        if out[col].dtype == object or str(out[col].dtype).startswith("string"):
            out[col] = out[col].map(translate_value)
    out = out.rename(columns={c: ui(c) for c in out.columns})
    return out


def px_labels(mapping=None):
    """Return Plotly labels without changing dataframe column names."""
    base_mapping = {
        "عدد الطلبات": ui("عدد الطلبات"),
        "الفرع": ui("الفرع"),
        "الحالة": ui("الحالة"),
        "الساعة": ui("الساعة"),
        "الحشوة": ui("الحشوة"),
        "المنتج": ui("المنتج"),
        "المبيعات": ui("المبيعات"),
        "الكمية": ui("الكمية"),
        "الكمية رقم": ui("الكمية رقم"),
        "قيمة الطلب": ui("قيمة الطلب"),
        "إجمالي المنتج رقم": ui("إجمالي المنتج رقم"),
        "عدد_الطلبات": ui("عدد_الطلبات"),
        "المبيعات_منتجات": ui("المبيعات_منتجات"),
        "متوسط_الطلب": ui("متوسط_الطلب"),
        "تصنيف الإضافة": ui("تصنيف الإضافة"),
        "الحملة": ui("الحملة"),
        "سبب المتابعة": ui("سبب المتابعة"),
        "نطاق ساعة الاستلام": ui("نطاق ساعة الاستلام"),
    }
    if mapping:
        base_mapping.update(mapping)
    return base_mapping



def ensure_chart_column(df, preferred, fallbacks=None):
    """Return an existing dataframe column. Never use translated labels as data bindings."""
    fallbacks = fallbacks or []
    if df is None or getattr(df, "empty", True):
        return None

    reverse_map = {}
    try:
        reverse_map.update({v: k for k, v in UI_EN.items()})
        reverse_map.update({v: k for k, v in VALUE_EN.items()})
    except Exception:
        pass

    candidates = []
    for c in [preferred] + list(fallbacks):
        if c is None:
            continue
        candidates.append(c)
        candidates.append(str(c))
        if str(c) in reverse_map:
            candidates.append(reverse_map[str(c)])

    for c in candidates:
        if c in df.columns:
            return c

    return None


def make_safe_bar(df, x_col, y_col, *, title="", text_col=None, color_col=None, fallback_x=None, fallback_y=None, **kwargs):
    """Safe bar chart wrapper used only for fragile dynamic charts."""
    if df is None or df.empty:
        return go.Figure()

    x_real = ensure_chart_column(df, x_col, fallback_x or [])
    y_real = ensure_chart_column(df, y_col, fallback_y or [])
    text_real = ensure_chart_column(df, text_col, []) if text_col is not None else None
    color_real = ensure_chart_column(df, color_col, []) if color_col is not None else None

    if x_real is None:
        x_real = df.columns[0]
    if y_real is None:
        numeric_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
        y_real = numeric_cols[0] if numeric_cols else (df.columns[1] if len(df.columns) > 1 else df.columns[0])

    chart_kwargs = dict(kwargs)
    if text_real is not None:
        chart_kwargs["text"] = text_real
    if color_real is not None:
        chart_kwargs["color"] = color_real

    return px.bar(
        df,
        x=x_real,
        y=y_real,
        title=ui(title) if title else None,
        labels=px_labels(),
        **chart_kwargs
    )



# =========================
# V8.5.4 Global Plotly Guard
# يمنع انهيار التطبيق لو أي رسم Plotly استخدم عمود مترجم أو غير موجود
# بدون تغيير منطق التقارير أو حذف أي وظيفة
# =========================

def _chart_reverse_label_map():
    reverse = {}
    try:
        reverse.update({v: k for k, v in UI_EN.items()})
        reverse.update({v: k for k, v in VALUE_EN.items()})
    except Exception:
        pass
    return reverse


def _resolve_plotly_column(data_frame, value, prefer_numeric=False):
    """Resolve a Plotly column argument to a real dataframe column."""
    if data_frame is None or not hasattr(data_frame, "columns"):
        return value

    cols = list(data_frame.columns)
    if value is None:
        return None

    # Lists/dicts used by Plotly should stay as-is unless they are a list of column names.
    if isinstance(value, (list, tuple)):
        resolved = []
        for item in value:
            resolved_item = _resolve_plotly_column(data_frame, item, prefer_numeric=prefer_numeric)
            if resolved_item is not None:
                resolved.append(resolved_item)
        return resolved or value

    if not isinstance(value, str):
        return value

    if value in cols:
        return value

    reverse = _chart_reverse_label_map()
    if value in reverse and reverse[value] in cols:
        return reverse[value]

    # Normalize very common English translated labels back to Arabic source columns
    common_reverse = {
        "Orders": "عدد الطلبات",
        "Order count": "عدد الطلبات",
        "Branch": "الفرع",
        "Hour": "الساعة",
        "Status": "الحالة",
        "Sales": "المبيعات",
        "Quantity": "الكمية",
        "Product": "المنتج",
        "Filling": "الحشوة",
        "Campaign": "الحملة",
        "Need action": "يحتاج متابعة",
        "Pickup hour range": "نطاق ساعة الاستلام",
    }
    if value in common_reverse and common_reverse[value] in cols:
        return common_reverse[value]

    # Fallbacks by intent
    numeric_cols = []
    non_numeric_cols = []
    try:
        numeric_cols = [c for c in cols if pd.api.types.is_numeric_dtype(data_frame[c])]
        non_numeric_cols = [c for c in cols if c not in numeric_cols]
    except Exception:
        pass

    if prefer_numeric and numeric_cols:
        return numeric_cols[0]
    if (not prefer_numeric) and non_numeric_cols:
        return non_numeric_cols[0]
    if cols:
        return cols[0]
    return value


def _clean_plotly_kwargs(data_frame, kwargs):
    """Resolve fragile dataframe column bindings before Plotly Express sees them."""
    if data_frame is None or not hasattr(data_frame, "columns"):
        return kwargs

    out = dict(kwargs)

    # Positional data bindings
    for key in ["x", "y", "color", "text", "names", "values", "size", "hover_name", "facet_row", "facet_col", "animation_frame", "line_group", "symbol"]:
        if key in out:
            out[key] = _resolve_plotly_column(
                data_frame,
                out[key],
                prefer_numeric=key in ["y", "values", "size"]
            )

    # hover_data can be list/dict
    if "hover_data" in out:
        hd = out["hover_data"]
        if isinstance(hd, dict):
            new_hd = {}
            for k, v in hd.items():
                rk = _resolve_plotly_column(data_frame, k)
                if rk in data_frame.columns:
                    new_hd[rk] = v
            out["hover_data"] = new_hd if new_hd else None
        elif isinstance(hd, (list, tuple)):
            new_hd = []
            for k in hd:
                rk = _resolve_plotly_column(data_frame, k)
                if rk in data_frame.columns:
                    new_hd.append(rk)
            out["hover_data"] = new_hd if new_hd else None

    # Labels must keep original dataframe column keys, values can be translated.
    if "labels" not in out or out["labels"] is None:
        try:
            out["labels"] = px_labels()
        except Exception:
            pass

    return {k: v for k, v in out.items() if v is not None}


def _install_plotly_guard():
    if getattr(px, "_mad_guard_installed", False):
        return

    px._mad_original_bar = px.bar
    px._mad_original_line = px.line
    px._mad_original_scatter = px.scatter
    px._mad_original_pie = px.pie

    def guarded_bar(data_frame=None, *args, **kwargs):
        kwargs = _clean_plotly_kwargs(data_frame, kwargs)
        try:
            return px._mad_original_bar(data_frame, *args, **kwargs)
        except ValueError:
            # Last-resort fallback: pick safe x/y columns and keep the chart alive.
            if data_frame is not None and hasattr(data_frame, "columns") and len(data_frame.columns) > 0:
                safe_x = _resolve_plotly_column(data_frame, kwargs.get("x"), prefer_numeric=False)
                safe_y = _resolve_plotly_column(data_frame, kwargs.get("y"), prefer_numeric=True)
                basic_kwargs = {
                    "x": safe_x,
                    "y": safe_y,
                    "title": kwargs.get("title"),
                    "labels": kwargs.get("labels", px_labels() if "px_labels" in globals() else None),
                }
                return px._mad_original_bar(data_frame, **{k:v for k,v in basic_kwargs.items() if v is not None})
            return go.Figure()

    def guarded_line(data_frame=None, *args, **kwargs):
        kwargs = _clean_plotly_kwargs(data_frame, kwargs)
        try:
            return px._mad_original_line(data_frame, *args, **kwargs)
        except ValueError:
            if data_frame is not None and hasattr(data_frame, "columns") and len(data_frame.columns) > 0:
                safe_x = _resolve_plotly_column(data_frame, kwargs.get("x"), prefer_numeric=False)
                safe_y = _resolve_plotly_column(data_frame, kwargs.get("y"), prefer_numeric=True)
                basic_kwargs = {
                    "x": safe_x,
                    "y": safe_y,
                    "markers": kwargs.get("markers", True),
                    "title": kwargs.get("title"),
                    "labels": kwargs.get("labels", px_labels() if "px_labels" in globals() else None),
                }
                return px._mad_original_line(data_frame, **{k:v for k,v in basic_kwargs.items() if v is not None})
            return go.Figure()

    def guarded_scatter(data_frame=None, *args, **kwargs):
        kwargs = _clean_plotly_kwargs(data_frame, kwargs)
        try:
            return px._mad_original_scatter(data_frame, *args, **kwargs)
        except ValueError:
            if data_frame is not None and hasattr(data_frame, "columns") and len(data_frame.columns) > 0:
                safe_x = _resolve_plotly_column(data_frame, kwargs.get("x"), prefer_numeric=True)
                safe_y = _resolve_plotly_column(data_frame, kwargs.get("y"), prefer_numeric=True)
                basic_kwargs = {
                    "x": safe_x,
                    "y": safe_y,
                    "title": kwargs.get("title"),
                    "labels": kwargs.get("labels", px_labels() if "px_labels" in globals() else None),
                }
                return px._mad_original_scatter(data_frame, **{k:v for k,v in basic_kwargs.items() if v is not None})
            return go.Figure()

    def guarded_pie(data_frame=None, *args, **kwargs):
        kwargs = _clean_plotly_kwargs(data_frame, kwargs)
        try:
            return px._mad_original_pie(data_frame, *args, **kwargs)
        except ValueError:
            return go.Figure()

    px.bar = guarded_bar
    px.line = guarded_line
    px.scatter = guarded_scatter
    px.pie = guarded_pie
    px._mad_guard_installed = True


_install_plotly_guard()


def translate_fig_for_language(fig):
    if not is_english_ui():
        return fig

    def _map_array(arr):
        if arr is None:
            return arr
        try:
            return [translate_value(v) for v in list(arr)]
        except Exception:
            return arr

    for trace in fig.data:
        try:
            if hasattr(trace, "x"):
                trace.x = _map_array(trace.x)
        except Exception:
            pass
        try:
            if hasattr(trace, "y"):
                trace.y = _map_array(trace.y)
        except Exception:
            pass
        try:
            if hasattr(trace, "labels"):
                trace.labels = _map_array(trace.labels)
        except Exception:
            pass
        try:
            if hasattr(trace, "names"):
                trace.names = _map_array(trace.names)
        except Exception:
            pass
        try:
            if getattr(trace, "name", None):
                trace.name = ui(trace.name)
        except Exception:
            pass

    try:
        if fig.layout.title and fig.layout.title.text:
            fig.update_layout(title_text=ui(fig.layout.title.text))
    except Exception:
        pass

    for ax in ["xaxis", "yaxis", "coloraxis"]:
        try:
            axis_obj = getattr(fig.layout, ax, None)
            if axis_obj and getattr(axis_obj, "title", None) and axis_obj.title.text:
                axis_obj.title.text = ui(axis_obj.title.text)
        except Exception:
            pass

    # Some colorbar titles live here
    try:
        if fig.layout.coloraxis and fig.layout.coloraxis.colorbar and fig.layout.coloraxis.colorbar.title.text:
            fig.layout.coloraxis.colorbar.title.text = ui(fig.layout.coloraxis.colorbar.title.text)
    except Exception:
        pass

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
    fig = translate_fig_for_language(fig)
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


def display_df(df, height=420, label="عرض الجدول"):
    """Keep dashboards clean: tables are hidden by default behind an expander."""
    if df is None or df.empty:
        st.info(ui("لا توجد بيانات لهذا التقرير ضمن الفلاتر الحالية."))
        return

    display_data = translate_df_for_display(df)
    rows_count = len(display_data)
    cols_count = len(display_data.columns)
    expander_label = f"📋 {ui(label)} — {format_int(rows_count)} {ui('صف')} / {format_int(cols_count)} {ui('عمود')}"
    with st.expander(expander_label, expanded=False):
        st.dataframe(display_data, use_container_width=True, height=height)


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
            "production_queue": "Production Queue",
            "action_center": "Action Center",
            "advanced_executive_summary": "V83 Executive Summary",
            "advanced_branch_ranking": "V83 Branch Ranking",
            "advanced_product_value": "V83 Product Value",
            "advanced_filling_intelligence": "V83 Filling Intel",
            "advanced_filling_by_branch_rank": "V83 Filling by Branch",
            "advanced_addons_opportunity_branch": "V83 Addons Branch",
            "advanced_addons_product_opportunity": "V83 Addons Product",
            "advanced_notes_keywords": "V83 Notes Keywords",
            "advanced_notes_by_product": "V83 Notes by Product",
            "advanced_data_quality_score": "V83 Quality Score",
            "advanced_hourly_capacity": "V83 Hourly Capacity",
            "advanced_campaign_performance": "V83 Campaign Perf",
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


def normalize_phone_for_whatsapp(phone):
    """Convert local Saudi mobile numbers into WhatsApp-friendly international digits."""
    digits = re.sub(r"\D+", "", normalize_arabic_digits(phone))
    if not digits:
        return ""
    if digits.startswith("00"):
        digits = digits[2:]
    if digits.startswith("9665") and len(digits) >= 12:
        return digits[:12]
    if digits.startswith("05") and len(digits) >= 10:
        return "966" + digits[1:10]
    if digits.startswith("5") and len(digits) >= 9:
        return "966" + digits[:9]
    return digits


def whatsapp_url(phone):
    normalized = normalize_phone_for_whatsapp(phone)
    return f"https://wa.me/{normalized}" if normalized else ""


def html_escape(value):
    return html.escape("" if pd.isna(value) else str(value))


def build_production_queue(active_items):
    """Production-friendly item list ordered by pickup time."""
    if active_items is None or active_items.empty:
        return pd.DataFrame()

    q = active_items.copy()
    q["نوع الصنف"] = q["إضافة؟"].map(lambda x: "إضافة" if bool(x) else "منتج رئيسي")
    q["واتساب"] = q["رقم الجوال المستخرج"].apply(whatsapp_url)
    q["أولوية"] = q["يحتاج متابعة؟"].map(lambda x: "عالية" if bool(x) else "عادية")
    q["وقت عرض"] = q["وقت الاستلام الأصلي"].replace("", "بدون وقت")
    q["تاريخ عرض"] = q["تاريخ التوصيل الأصلي"].replace("", "بدون تاريخ")
    if "ساعة رقم" in q.columns:
        q["نطاق ساعة الاستلام"] = q["ساعة رقم"].apply(hour_range_label)
    else:
        q["نطاق ساعة الاستلام"] = "بدون وقت"

    cols = [
        "نطاق ساعة الاستلام", "وقت عرض", "تاريخ عرض", "الفرع", "رقم الطلب الظاهر", "رقم الطلب الموحد",
        "الحالة", "العميل", "نوع الصنف", "المنتج", "الحشوة", "الكمية رقم",
        "أولوية", "سبب المتابعة", "رقم الجوال المستخرج", "واتساب", "الملاحظة",
        "تاريخ ووقت الاستلام", "ساعة رقم"
    ]
    q = q[[c for c in cols if c in q.columns]].copy()

    sort_cols = [c for c in ["تاريخ ووقت الاستلام", "الفرع", "رقم الطلب الموحد"] if c in q.columns]
    if sort_cols:
        q = q.sort_values(sort_cols, na_position="last")

    return q


def build_action_center(active_items):
    """Action-focused table: photos, contact, phone, writing, design, data issues."""
    if active_items is None or active_items.empty:
        return pd.DataFrame()

    a = active_items.copy()

    missing_date = a["تاريخ التوصيل الأصلي"].astype(str).str.strip().eq("") if "تاريخ التوصيل الأصلي" in a.columns else False
    missing_time = a["وقت الاستلام الأصلي"].astype(str).str.strip().eq("") if "وقت الاستلام الأصلي" in a.columns else False
    missing_branch = a["الفرع"].astype(str).eq("بدون فرع محدد") if "الفرع" in a.columns else False
    action_mask = a["يحتاج متابعة؟"].astype(bool) | missing_date | missing_time | missing_branch
    a = a[action_mask].copy()

    if a.empty:
        return pd.DataFrame()

    def enrich_reasons(row):
        reasons = str(row.get("سبب المتابعة", "")).strip()
        parts = [p.strip() for p in reasons.split("،") if p.strip()]
        if not str(row.get("تاريخ التوصيل الأصلي", "")).strip():
            parts.append("تاريخ ناقص")
        if not str(row.get("وقت الاستلام الأصلي", "")).strip():
            parts.append("وقت ناقص")
        if str(row.get("الفرع", "")) == "بدون فرع محدد":
            parts.append("فرع غير محدد")
        seen = []
        for p in parts:
            if p not in seen:
                seen.append(p)
        return "، ".join(seen)

    def priority(reasons):
        txt = str(reasons)
        if any(k in txt for k in ["ملاحظة حساسة", "تاريخ ناقص", "وقت ناقص", "فرع غير محدد"]):
            return "حرجة"
        if any(k in txt for k in ["يحتاج صورة", "يحتاج تواصل", "يوجد رقم جوال"]):
            return "عالية"
        return "متوسطة"

    a["نوع الإجراء"] = a.apply(enrich_reasons, axis=1)
    a["الأولوية"] = a["نوع الإجراء"].apply(priority)
    a["حالة المتابعة"] = "لم يبدأ"
    a["ملاحظة داخلية"] = ""
    a["واتساب"] = a["رقم الجوال المستخرج"].apply(whatsapp_url)
    a["وقت عرض"] = a["وقت الاستلام الأصلي"].replace("", "بدون وقت")
    a["تاريخ عرض"] = a["تاريخ التوصيل الأصلي"].replace("", "بدون تاريخ")
    if "ساعة رقم" in a.columns:
        a["نطاق ساعة الاستلام"] = a["ساعة رقم"].apply(hour_range_label)
    else:
        a["نطاق ساعة الاستلام"] = "بدون وقت"

    cols = [
        "الأولوية", "حالة المتابعة", "نوع الإجراء", "نطاق ساعة الاستلام", "وقت عرض", "تاريخ عرض",
        "الفرع", "رقم الطلب الظاهر", "رقم الطلب الموحد", "الحالة", "العميل",
        "المنتج", "الحشوة", "الكمية رقم", "رقم الجوال المستخرج", "واتساب",
        "ملاحظة داخلية", "الملاحظة", "تاريخ ووقت الاستلام", "ساعة رقم"
    ]
    a = a[[c for c in cols if c in a.columns]].drop_duplicates().copy()

    sort_priority = {"حرجة": 0, "عالية": 1, "متوسطة": 2}
    a["_priority_sort"] = a["الأولوية"].map(sort_priority).fillna(9)
    sort_cols = [c for c in ["_priority_sort", "تاريخ ووقت الاستلام", "الفرع"] if c in a.columns]
    if sort_cols:
        a = a.sort_values(sort_cols, na_position="last")
    return a.drop(columns=["_priority_sort"], errors="ignore")


def build_branch_prep_detail(queue_df, branch):
    if queue_df is None or queue_df.empty or not branch:
        return pd.DataFrame()
    b = queue_df[queue_df["الفرع"].eq(branch)].copy() if "الفرع" in queue_df.columns else pd.DataFrame()
    cols = [
        "نطاق ساعة الاستلام", "وقت عرض", "رقم الطلب الظاهر", "الحالة", "العميل", "نوع الصنف",
        "المنتج", "الحشوة", "الكمية رقم", "أولوية", "سبب المتابعة",
        "رقم الجوال المستخرج", "الملاحظة"
    ]
    return b[[c for c in cols if c in b.columns]]


def build_multi_sheet_excel(sheet_map):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheet_map.items():
            write_excel_sheet(writer, df if isinstance(df, pd.DataFrame) else pd.DataFrame(), sheet_name)
    output.seek(0)
    return output


def build_branch_workbook(branch, queue_df, action_df, active_items):
    branch_queue = queue_df[queue_df["الفرع"].eq(branch)].copy() if queue_df is not None and not queue_df.empty and "الفرع" in queue_df.columns else pd.DataFrame()
    branch_action = action_df[action_df["الفرع"].eq(branch)].copy() if action_df is not None and not action_df.empty and "الفرع" in action_df.columns else pd.DataFrame()
    branch_items = active_items[active_items["الفرع"].eq(branch)].copy() if active_items is not None and not active_items.empty and "الفرع" in active_items.columns else pd.DataFrame()

    if not branch_items.empty:
        products = branch_items[~branch_items["إضافة؟"]].groupby(["المنتج", "الحشوة"], dropna=False).agg(
            الكمية=("الكمية رقم", "sum"),
            عدد_الطلبات=("رقم الطلب الموحد", "nunique"),
        ).reset_index().sort_values(["الكمية", "عدد_الطلبات"], ascending=False)

        fillings = branch_items[branch_items["الحشوة"].astype(str).str.strip().ne("")].groupby("الحشوة", dropna=False).agg(
            الكمية=("الكمية رقم", "sum"),
            عدد_الطلبات=("رقم الطلب الموحد", "nunique"),
        ).reset_index().sort_values("الكمية", ascending=False)

        addons = branch_items[branch_items["إضافة؟"]].groupby("تصنيف الإضافة", dropna=False).agg(
            الكمية=("الكمية رقم", "sum"),
            عدد_الطلبات=("رقم الطلب الموحد", "nunique"),
        ).reset_index().sort_values("الكمية", ascending=False)

        by_hour = branch_items.drop_duplicates("رقم الطلب الموحد").groupby("الساعة", dropna=False).agg(
            عدد_الطلبات=("رقم الطلب الموحد", "nunique"),
        ).reset_index()
    else:
        products = fillings = addons = by_hour = pd.DataFrame()

    summary = pd.DataFrame([
        {"البند": "الفرع", "القيمة": branch},
        {"البند": "عدد صفوف التجهيز", "القيمة": len(branch_queue)},
        {"البند": "طلبات تحتاج متابعة", "القيمة": len(branch_action)},
        {"البند": "عدد المنتجات", "القيمة": len(products)},
        {"البند": "عدد الحشوات", "القيمة": len(fillings)},
    ])

    return build_multi_sheet_excel({
        "Branch Summary": summary,
        "Production Queue": branch_queue,
        "Need Action": branch_action,
        "Products Prep": products,
        "Fillings Prep": fillings,
        "Add-ons Prep": addons,
        "Orders by Hour": by_hour,
    })


def render_print_cards(print_df, limit=120):
    if print_df is None or print_df.empty:
        st.info("لا توجد بيانات للطباعة ضمن الفلاتر الحالية.")
        return
    shown = print_df.head(limit)
    for _, row in shown.iterrows():
        order_no = html_escape(row.get("رقم الطلب الظاهر", ""))
        branch = html_escape(row.get("الفرع", ""))
        time_v = html_escape(row.get("نطاق ساعة الاستلام", row.get("وقت عرض", row.get("وقت الاستلام الأصلي", ""))))
        status = html_escape(row.get("الحالة", ""))
        customer = html_escape(row.get("العميل", ""))
        product = html_escape(row.get("المنتج", ""))
        variety = html_escape(row.get("الحشوة", ""))
        qty = html_escape(row.get("الكمية رقم", ""))
        action = html_escape(row.get("سبب المتابعة", row.get("نوع الإجراء", "")))
        phone = html_escape(row.get("رقم الجوال المستخرج", ""))
        note = html_escape(row.get("الملاحظة", ""))

        st.markdown(
            f"""
            <div class="print-card">
                <b>وقت:</b> {time_v} &nbsp; | &nbsp; <b>فرع:</b> {branch} &nbsp; | &nbsp; <b>طلب:</b> {order_no} &nbsp; | &nbsp; <b>الحالة:</b> {status}<br>
                <b>العميل:</b> {customer} &nbsp; | &nbsp; <b>جوال:</b> {phone}<br>
                <b>المنتج:</b> {product}<br>
                <b>الحشوة:</b> {variety} &nbsp; | &nbsp; <b>الكمية:</b> {qty}<br>
                <b>متابعة:</b> {action}
                <div class="print-note">{note}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    if len(print_df) > limit:
        st.warning(f"تم عرض أول {limit} صف فقط للطباعة. استخدم Excel للعدد الكامل.")




# =========================
# V8.4 Add-ons - Language + Auto Refresh
# هذه الإضافات لا تحذف أي تقرير أو وظيفة من V8.3.4
# =========================
st.sidebar.markdown("## 🌐 Language / اللغة")
LANGUAGE = st.sidebar.radio(
    "اختر اللغة / Choose language",
    ["العربية", "English"],
    index=0,
    horizontal=True,
    key="ui_language_selector",
)
IS_AR = LANGUAGE == "العربية"

def tr(ar_text, en_text):
    return ar_text if IS_AR else en_text

APP_DIRECTION = "rtl" if IS_AR else "ltr"
APP_ALIGN = "right" if IS_AR else "left"

st.markdown(
    f"""
    <style>
        html, body, [data-testid="stAppViewContainer"] {{
            direction: {APP_DIRECTION};
        }}
        .block-container, .hero, .section-title, .mini-title, .kpi-card,
        .alert-box, .good-box, .note-box, .readability-note,
        div[data-testid="stMarkdownContainer"], label, p, h1, h2, h3, h4 {{
            direction: {APP_DIRECTION} !important;
            text-align: {APP_ALIGN} !important;
        }}
        .js-plotly-plot, .plot-container, div[data-testid="stDataFrame"] {{
            direction: ltr !important;
        }}
    </style>
    """,
    unsafe_allow_html=True,
)

st.sidebar.markdown("---")
st.sidebar.markdown(f"## 🔁 {tr('التحديث التلقائي', 'Auto refresh')}")
_refresh_options = {
    tr("إيقاف", "Off"): 0,
    tr("كل دقيقة", "Every 1 minute"): 60,
    tr("كل 5 دقائق", "Every 5 minutes"): 300,
    tr("كل 10 دقائق", "Every 10 minutes"): 600,
    tr("كل 15 دقيقة", "Every 15 minutes"): 900,
}
_auto_refresh_label = st.sidebar.selectbox(
    tr("معدل التحديث", "Refresh interval"),
    list(_refresh_options.keys()),
    index=2,
    key="auto_refresh_interval",
)
_auto_refresh_seconds = _refresh_options[_auto_refresh_label]

if _auto_refresh_seconds > 0:
    if st_autorefresh is not None:
        st_autorefresh(
            interval=_auto_refresh_seconds * 1000,
            key="orders_dashboard_v84_auto_refresh",
        )
        st.sidebar.success(
            tr(
                f"يتم التحديث تلقائيًا: {_auto_refresh_label}",
                f"Auto refresh enabled: {_auto_refresh_label}",
            )
        )
    else:
        st.sidebar.warning(
            tr(
                "التحديث التلقائي يحتاج إضافة streamlit-autorefresh في requirements.txt",
                "Auto refresh requires streamlit-autorefresh in requirements.txt",
            )
        )
else:
    st.sidebar.info(tr("التحديث التلقائي متوقف", "Auto refresh is off"))

st.sidebar.caption(
    tr(
        f"آخر تشغيل للصفحة: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"Last page run: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
    )
)
st.sidebar.markdown("---")



# =========================
# Header
# =========================
st.markdown(
    f"""
    <div class="hero">
        <h1>{tr('🧁 مركز تحكم طلبات MAD', '🧁 MAD Orders Control Center')}</h1>
        <p>{APP_VERSION} — {tr('نفس التقارير التشغيلية والإدارية بالكامل بدون حذف: ملخص تنفيذي، ترتيب الفروع، قيمة المنتجات، ذكاء الحشوات، فرص الإضافات، الطاقة التشغيلية، وجودة البيانات.', 'All operational and management reports are preserved: executive summary, branch ranking, product value, filling intelligence, add-ons opportunities, operational capacity, and data quality.')}</p>
    </div>
    """,
    unsafe_allow_html=True,
)


# =========================
# Sidebar - Data Source
# =========================
st.sidebar.markdown(f"## ⚙️ {tr('مصدر البيانات', 'Data source')}")
_source_display = st.sidebar.radio(
    tr("اختر المصدر", "Choose source"),
    ["Google Sheet", tr("رفع ملف", "Upload file")],
    index=0,
    horizontal=True,
)
source_type = "Google Sheet" if _source_display == "Google Sheet" else "رفع ملف"

raw_df = None
load_error = None

if source_type == "Google Sheet":
    sheet_url = st.sidebar.text_input(
        tr("رابط Google Sheet", "Google Sheet URL"),
        value=DEFAULT_GOOGLE_SHEET_URL,
        placeholder="https://docs.google.com/spreadsheets/d/...",
    )
    gid = st.sidebar.text_input(
        "Sheet GID",
        value="0",
        help=tr(
            "لو الشيت تبويب مختلف، انسخ رقم gid من الرابط. غالباً أول تبويب = 0",
            "If the data is in another sheet tab, copy the gid from the link. Usually first tab = 0",
        ),
    )
    c_refresh, c_status = st.sidebar.columns([1, 1])
    with c_refresh:
        if st.button(tr("🔄 تحديث الآن", "🔄 Update now")):
            st.cache_data.clear()
            st.rerun()
    try:
        with st.spinner(tr("جاري تحميل Google Sheet...", "Loading Google Sheet...")):
            raw_df = load_google_sheet(sheet_url, gid)
        with c_status:
            st.success(tr("تم", "Done"))
    except Exception as e:
        load_error = str(e)
else:
    uploaded = st.sidebar.file_uploader(
        tr("ارفع ملف TXT / CSV / Excel", "Upload TXT / CSV / Excel file"),
        type=["txt", "csv", "xlsx", "xls"],
    )
    if uploaded is not None:
        try:
            raw_df = read_uploaded_file(uploaded)
        except Exception as e:
            load_error = str(e)

if load_error:
    st.error(
        tr(
            "لم أستطع تحميل البيانات. تأكد أن Google Sheet متاح لأي شخص لديه الرابط Viewer أو ارفع ملف مباشرة.",
            "Could not load the data. Make sure the Google Sheet is shared as Viewer with anyone who has the link, or upload a file directly.",
        )
    )
    with st.expander(tr("تفاصيل الخطأ", "Error details")):
        st.code(load_error)
    st.stop()

if raw_df is None or raw_df.empty:
    st.markdown(
        f'<div class="note-box">{tr("اربط Google Sheet أو ارفع ملف الطلبات لبدء التحليل.", "Connect Google Sheet or upload the orders file to start analysis.")}</div>',
        unsafe_allow_html=True,
    )
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
st.sidebar.markdown(f"## 🔎 {tr('فلاتر التحليل', 'Analysis filters')}")

available_dates = sorted([d for d in items_all["تاريخ التحليل"].dropna().unique()])
if available_dates:
    min_date, max_date = min(available_dates), max(available_dates)
    date_range = st.sidebar.date_input(tr("فترة تاريخ التوصيل", "Delivery date range"), value=(min_date, max_date), min_value=min_date, max_value=max_date)
else:
    date_range = None

branches = sorted(items_all["الفرع"].dropna().unique().tolist())
selected_branches = st.sidebar.multiselect(tr("الفروع", "Branches"), branches, default=branches)

statuses = sorted(items_all["الحالة"].dropna().unique().tolist())
selected_statuses = st.sidebar.multiselect(tr("الحالات", "Statuses"), statuses, default=statuses)

varieties = sorted([v for v in items_all["الحشوة"].dropna().unique().tolist() if str(v).strip()])
selected_varieties = st.sidebar.multiselect(tr("الحشوات", "Fillings"), varieties, default=[])

campaigns = sorted(items_all["الحملة"].dropna().unique().tolist())
selected_campaigns = st.sidebar.multiselect(tr("الحملات", "Campaigns"), campaigns, default=[])

search_text = st.sidebar.text_input(tr("بحث سريع", "Quick search"), placeholder=tr("رقم طلب / عميل / منتج / جوال", "Order / customer / product / mobile"))
include_cancelled = st.sidebar.checkbox(tr("إظهار الملغي ضمن التقارير", "Show cancelled in reports"), value=False)

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

# V8 operational reports
reports["production_queue"] = build_production_queue(active_items)
reports["action_center"] = build_action_center(active_items)

# V8.3 advanced management reports
reports.update(build_v83_advanced_reports(filtered, active_items, active_orders, reports))

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

st.sidebar.success(tr("تم تحليل البيانات", "Data analyzed"))
st.sidebar.write(f"{tr('الصفوف', 'Rows')}: **{format_int(total_rows)}**")
st.sidebar.write(f"{tr('الطلبات', 'Orders')}: **{format_int(total_orders)}**")
st.sidebar.write(f"{tr('المبيعات', 'Sales')}: **{format_money(total_sales)}**")
st.sidebar.write(f"{tr('تحتاج متابعة', 'Need action')}: **{format_int(need_action_count)}**")


# =========================
# Main KPIs
# =========================
k1, k2, k3, k4, k5 = st.columns(5)
with k1:
    render_kpi(tr(ui("عدد الطلبات"), "Orders"), format_int(total_orders), "Unique Orders", "#2563eb")
with k2:
    render_kpi(tr("إجمالي المبيعات", "Total sales"), format_money(total_sales), tr("بدون تكرار قيمة الطلب", "Order value without duplication"), "#16a34a")
with k3:
    render_kpi(tr("متوسط الطلب", "Average order"), format_money(avg_order), "AOV", "#0891b2")
with k4:
    render_kpi(tr("تحتاج متابعة", "Need action"), format_int(need_action_count), tr("صورة / تواصل / كتابة", "Photo / Contact / Writing"), "#dc2626")
with k5:
    render_kpi("Upsell", f"{upsell_rate:.1f}%", f"{format_int(addon_orders)} {tr('طلب بإضافات', 'orders with add-ons')}", "#f59e0b")


# =========================
# Smart Alerts
# =========================
st.markdown(f'<div class="section-title">{tr("🚦 تنبيهات ذكية", "🚦 Smart alerts")}</div>', unsafe_allow_html=True)
alerts = []
action_center_count = len(reports.get("action_center", pd.DataFrame()))
if action_center_count > 0:
    alerts.append(
        tr(
            f"🚨 Action Center: يوجد {format_int(action_center_count)} صف يحتاج إجراء تشغيلي.",
            f"🚨 Action Center: {format_int(action_center_count)} rows need operational action."
        )
    )
elif need_action_count > 0:
    alerts.append(
        tr(
            f"⚠️ يوجد {format_int(need_action_count)} طلب يحتاج متابعة مع العميل.",
            f"⚠️ {format_int(need_action_count)} orders need customer follow-up."
        )
    )
if missing_date_count > 0:
    alerts.append(
        tr(
            f"⚠️ يوجد {format_int(missing_date_count)} صف بدون تاريخ توصيل.",
            f"⚠️ {format_int(missing_date_count)} rows have no delivery date."
        )
    )
if missing_time_count > 0:
    alerts.append(
        tr(
            f"⚠️ يوجد {format_int(missing_time_count)} صف بدون وقت استلام.",
            f"⚠️ {format_int(missing_time_count)} rows have no pickup time."
        )
    )
if top_hour != "-":
    alerts.append(
        tr(
            f"🔥 أعلى ساعة ضغط: {top_hour} بعدد {format_int(top_hour_count)} طلب.",
            f"🔥 Peak hour: {translate_value(top_hour)} with {format_int(top_hour_count)} orders."
        )
    )
if top_branch != "-":
    alerts.append(
        tr(
            f"🏬 أعلى فرع طلبات: {top_branch} بعدد {format_int(top_branch_count)} طلب.",
            f"🏬 Top branch by orders: {translate_value(top_branch)} with {format_int(top_branch_count)} orders."
        )
    )
if not alerts:
    st.markdown(f'<div class="good-box">{tr("✅ لا توجد تنبيهات حرجة ضمن الفلاتر الحالية.", "✅ No critical alerts in the current filters.")}</div>', unsafe_allow_html=True)
else:
    for alert in alerts[:6]:
        st.markdown(f'<div class="alert-box">{alert}</div>', unsafe_allow_html=True)


# =========================
# Tabs
# =========================
(
    tab_daily,
    tab_production,
    tab_action_center,
    tab_print,
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
    tab_advanced,
    tab_export,
) = st.tabs([
    tr("📍 التشغيل اليومي", "📍 Daily Ops"),
    tr("🏭 قائمة الإنتاج", "🏭 Production Queue"),
    tr("✅ مركز الإجراءات", "✅ Action Center"),
    tr("🖨️ عرض الطباعة", "🖨️ Print View"),
    tr("🧾 تجهيز الفروع", "🧾 Branch Prep"),
    tr("💰 المبيعات", "💰 Sales"),
    tr("🧁 المنتجات", "🧁 Products"),
    tr("🍰 الحشوات", "🍰 Fillings"),
    tr("🎈 الإضافات", "🎈 Add-ons"),
    tr("🚨 تحتاج متابعة", "🚨 Need Action"),
    tr("🎯 الحملات", "🎯 Campaigns"),
    tr("🏬 تحليل الفرع", "🏬 Branch Deep Dive"),
    tr("🔍 تحليل المنتج", "🔍 Product Deep Dive"),
    tr("🧹 جودة البيانات", "🧹 Data Quality"),
    tr("📊 التقارير المتقدمة", "📊 Advanced Reports"),
    tr("⬇️ التصدير", "⬇️ Export"),
])


with tab_daily:
    st.markdown(f'<div class="section-title">{tr("📍 التشغيل اليومي", "📍 Daily Operations")}</div>', unsafe_allow_html=True)
    c1, c2 = st.columns([1.1, 1])
    with c1:
        sales_branch = reports.get("sales_by_branch", pd.DataFrame())
        if not sales_branch.empty:
            fig = px.bar(sales_branch.sort_values("عدد_الطلبات"), x="عدد_الطلبات", y="الفرع", orientation="h", text="عدد_الطلبات", title="ضغط الطلبات حسب الفرع", color="عدد_الطلبات", color_continuous_scale="Viridis")
            fig.update_traces(textposition="outside")
            st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=chart_config())
        else:
            st.info(tr("لا يوجد بيانات فروع.", "No branch data."))
    with c2:
        sales_hour = reports.get("sales_by_hour", pd.DataFrame())
        if not sales_hour.empty:
            fig = px.line(sales_hour, x="الساعة", y="عدد_الطلبات", markers=True, title="ضغط الطلبات حسب الساعة")
            fig.update_traces(line=dict(width=4), marker=dict(size=10))
            st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=chart_config())
        else:
            st.info(tr("لا يوجد بيانات ساعات.", "No hour data."))

    heat = reports.get("branch_hour_heatmap", pd.DataFrame())
    if not heat.empty:
        fig = px.imshow(heat, text_auto=True, aspect="auto", title="Heatmap الفرع × الساعة", color_continuous_scale="YlGnBu")
        fig.update_xaxes(side="top")
        st.plotly_chart(fig_layout(fig, 520), use_container_width=True, config=chart_config())

    st.markdown(f'<div class="mini-title">{ui("أقرب جدول تشغيل ضمن الفلاتر")}</div>', unsafe_allow_html=True)
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



with tab_production:
    st.markdown(f'<div class="section-title">{tr("🏭 قائمة الإنتاج", "🏭 Production Queue")}</div>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="production-card">{tr("صفحة تجهيز يومية مرتبة حسب وقت الاستلام. استخدمها للفروع والإنتاج لمعرفة المطلوب الآن، وما يحتاج متابعة قبل التجهيز.", "Daily preparation queue sorted by pickup time. Use it for branches and production to know what is needed now and what needs follow-up before preparation.")}</div>',
        unsafe_allow_html=True,
    )

    queue = reports.get("production_queue", pd.DataFrame()).copy()

    if queue.empty:
        st.info("لا توجد بيانات تجهيز ضمن الفلاتر الحالية.")
    else:
        qf1, qf2, qf3, qf4 = st.columns(4)
        with qf1:
            q_branches = ["الكل"] + sorted(queue["الفرع"].dropna().unique().tolist()) if "الفرع" in queue.columns else ["الكل"]
            q_branch = st.selectbox(ui("فرع التجهيز"), q_branches, key="v8_queue_branch")
        with qf2:
            if "نطاق ساعة الاستلام" in queue.columns:
                unique_ranges = queue["نطاق ساعة الاستلام"].dropna().unique().tolist()
                q_hours = ["الكل"] + sorted(unique_ranges, key=hour_range_sort_value)
            else:
                q_hours = ["الكل"]
            q_hour = st.selectbox(ui("نطاق ساعة الاستلام"), q_hours, key="v8_queue_hour")
        with qf3:
            q_statuses = ["الكل"] + sorted(queue["الحالة"].dropna().unique().tolist()) if "الحالة" in queue.columns else ["الكل"]
            q_status = st.selectbox(ui("الحالة"), q_statuses, key="v8_queue_status")
        with qf4:
            q_priority = st.selectbox(ui("الأولوية"), ["الكل", "عالية", "عادية"], key="v8_queue_priority", format_func=translate_value)

        qf5, qf6, qf7 = st.columns(3)
        with qf5:
            show_addons_q = st.checkbox(ui("إظهار الإضافات"), value=True, key="v8_queue_addons")
        with qf6:
            need_action_q = st.checkbox(ui("فقط ما يحتاج متابعة"), value=False, key="v8_queue_need_action")
        with qf7:
            q_search = st.text_input(ui("بحث داخل التجهيز"), placeholder=ui("رقم طلب / منتج / عميل / جوال"), key="v8_queue_search")

        queue_view = queue.copy()
        if q_branch != "الكل" and "الفرع" in queue_view.columns:
            queue_view = queue_view[queue_view["الفرع"].eq(q_branch)]
        if q_hour != "الكل" and "نطاق ساعة الاستلام" in queue_view.columns:
            queue_view = queue_view[queue_view["نطاق ساعة الاستلام"].eq(q_hour)]
        if q_status != "الكل" and "الحالة" in queue_view.columns:
            queue_view = queue_view[queue_view["الحالة"].eq(q_status)]
        if q_priority != "الكل" and "أولوية" in queue_view.columns:
            queue_view = queue_view[queue_view["أولوية"].eq(q_priority)]
        if not show_addons_q and "نوع الصنف" in queue_view.columns:
            queue_view = queue_view[queue_view["نوع الصنف"].ne("إضافة")]
        if need_action_q and "أولوية" in queue_view.columns:
            queue_view = queue_view[queue_view["أولوية"].eq("عالية")]
        if q_search.strip():
            q = q_search.strip().lower()
            blob_cols = [c for c in ["رقم الطلب الظاهر", "رقم الطلب الموحد", "العميل", "المنتج", "الحشوة", "رقم الجوال المستخرج", "الملاحظة"] if c in queue_view.columns]
            blob = queue_view[blob_cols].astype(str).agg(" ".join, axis=1).str.lower() if blob_cols else pd.Series([], dtype=str)
            queue_view = queue_view[blob.str.contains(re.escape(q), na=False)]

        pq1, pq2, pq3, pq4 = st.columns(4)
        with pq1:
            render_kpi("صفوف التجهيز", format_int(len(queue_view)), "Items", "#0ea5e9")
        with pq2:
            order_count_q = queue_view["رقم الطلب الموحد"].nunique() if "رقم الطلب الموحد" in queue_view.columns else 0
            render_kpi(ui("عدد الطلبات"), format_int(order_count_q), "Unique Orders", "#2563eb")
        with pq3:
            qty_q = queue_view["الكمية رقم"].sum() if "الكمية رقم" in queue_view.columns else 0
            render_kpi("إجمالي الكمية", format_int(qty_q), "Qty", "#16a34a")
        with pq4:
            act_q = queue_view["أولوية"].eq("عالية").sum() if "أولوية" in queue_view.columns else 0
            render_kpi("تحتاج متابعة", format_int(act_q), "High priority", "#dc2626")

        if not queue_view.empty:
            qc1, qc2 = st.columns([1, 1])
            with qc1:
                by_hour_col = "نطاق ساعة الاستلام" if "نطاق ساعة الاستلام" in queue_view.columns else "وقت عرض"
                by_hour_q = queue_view.groupby(by_hour_col, dropna=False)["رقم الطلب الموحد"].nunique().reset_index(name=ui("عدد الطلبات"))
                by_hour_q["_sort"] = by_hour_q[by_hour_col].apply(hour_range_sort_value)
                by_hour_q = by_hour_q.sort_values("_sort").drop(columns=["_sort"])
                fig = make_safe_bar(by_hour_q, by_hour_col, "عدد الطلبات", text_col="عدد الطلبات", title="ضغط التجهيز حسب نطاق الساعة", fallback_x=["نطاق ساعة الاستلام", "الساعة"], fallback_y=["عدد الطلبات"])
                st.plotly_chart(make_readable_fig(fig, 430, showlegend=False), use_container_width=True, config=chart_config())
            with qc2:
                by_branch_q = queue_view.groupby("الفرع", dropna=False)["رقم الطلب الموحد"].nunique().reset_index(name=ui("عدد الطلبات"))
                fig = make_safe_bar(by_branch_q, "الفرع", "عدد الطلبات", text_col="عدد الطلبات", title="طلبات التجهيز حسب الفرع", fallback_x=["الفرع"], fallback_y=["عدد الطلبات"])
                st.plotly_chart(make_readable_fig(fig, 430, showlegend=False), use_container_width=True, config=chart_config())

        st.markdown('<div class="mini-title">جدول التجهيز</div>', unsafe_allow_html=True)
        view_cols = [c for c in [
            "نطاق ساعة الاستلام", "وقت عرض", "تاريخ عرض", "الفرع", "رقم الطلب الظاهر", "الحالة", "العميل",
            "نوع الصنف", "المنتج", "الحشوة", "الكمية رقم", "أولوية",
            "سبب المتابعة", "رقم الجوال المستخرج", "الملاحظة"
        ] if c in queue_view.columns]
        display_df(queue_view[view_cols] if view_cols else queue_view, 620)

        st.download_button(
            "⬇️ تحميل Production Queue Excel",
            data=build_multi_sheet_excel({"Production Queue": queue_view}),
            file_name="Production_Queue_V8_2.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_production_queue_v8",
        )


with tab_action_center:
    st.markdown(f'<div class="section-title">{tr("✅ مركز الإجراءات", "✅ Action Center")}</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="action-card">هذه الصفحة تجمع الطلبات التي تحتاج إجراء: صورة، تواصل، رقم جوال، كتابة خاصة، تعديل تصميم، مشكلة، أو بيانات ناقصة. التعديل هنا مؤقت داخل الجلسة فقط، والحفظ الدائم يكون في V9 عبر Supabase أو Google Sheet Tracking.</div>',
        unsafe_allow_html=True,
    )

    action_df = reports.get("action_center", pd.DataFrame()).copy()

    if action_df.empty:
        st.success("✅ لا توجد طلبات تحتاج إجراء ضمن الفلاتر الحالية.")
    else:
        af1, af2, af3, af4 = st.columns(4)
        with af1:
            a_branches = ["الكل"] + sorted(action_df["الفرع"].dropna().unique().tolist()) if "الفرع" in action_df.columns else ["الكل"]
            a_branch = st.selectbox("فرع", a_branches, key="v8_action_branch")
        with af2:
            a_priorities = ["الكل"] + sorted(action_df["الأولوية"].dropna().unique().tolist()) if "الأولوية" in action_df.columns else ["الكل"]
            a_priority = st.selectbox("الأولوية", a_priorities, key="v8_action_priority")
        with af3:
            reason_options = ["الكل"]
            if "نوع الإجراء" in action_df.columns:
                reason_set = []
                for reasons in action_df["نوع الإجراء"].astype(str):
                    for r in [x.strip() for x in reasons.split("،") if x.strip()]:
                        reason_set.append(r)
                reason_options += sorted(set(reason_set))
            a_reason = st.selectbox("نوع الإجراء", reason_options, key="v8_action_reason")
        with af4:
            a_search = st.text_input("بحث", placeholder="رقم طلب / جوال / عميل / منتج", key="v8_action_search")

        action_view = action_df.copy()
        if a_branch != "الكل" and "الفرع" in action_view.columns:
            action_view = action_view[action_view["الفرع"].eq(a_branch)]
        if a_priority != "الكل" and "الأولوية" in action_view.columns:
            action_view = action_view[action_view["الأولوية"].eq(a_priority)]
        if a_reason != "الكل" and "نوع الإجراء" in action_view.columns:
            action_view = action_view[action_view["نوع الإجراء"].astype(str).str.contains(re.escape(a_reason), na=False)]
        if a_search.strip():
            q = a_search.strip().lower()
            blob_cols = [c for c in ["رقم الطلب الظاهر", "رقم الطلب الموحد", "العميل", "المنتج", "رقم الجوال المستخرج", "الملاحظة"] if c in action_view.columns]
            blob = action_view[blob_cols].astype(str).agg(" ".join, axis=1).str.lower() if blob_cols else pd.Series([], dtype=str)
            action_view = action_view[blob.str.contains(re.escape(q), na=False)]

        ac1, ac2, ac3, ac4 = st.columns(4)
        with ac1:
            render_kpi("إجمالي الإجراءات", format_int(len(action_view)), "Rows", "#dc2626")
        with ac2:
            high_count = action_view["الأولوية"].isin(["حرجة", "عالية"]).sum() if "الأولوية" in action_view.columns else 0
            render_kpi("حرجة / عالية", format_int(high_count), "Priority", "#f97316")
        with ac3:
            phone_count = action_view["رقم الجوال المستخرج"].astype(str).str.len().gt(0).sum() if "رقم الجوال المستخرج" in action_view.columns else 0
            render_kpi("فيها جوال", format_int(phone_count), "WhatsApp ready", "#16a34a")
        with ac4:
            orders_count = action_view["رقم الطلب الموحد"].nunique() if "رقم الطلب الموحد" in action_view.columns else 0
            render_kpi(ui("عدد الطلبات"), format_int(orders_count), "Unique Orders", "#2563eb")

        if not action_view.empty:
            reasons_expanded = []
            for reasons in action_view["نوع الإجراء"].astype(str):
                for r in [x.strip() for x in reasons.split("،") if x.strip()]:
                    reasons_expanded.append(r)
            if reasons_expanded:
                reasons_df = pd.Series(reasons_expanded).value_counts().reset_index()
                reasons_df.columns = ["نوع الإجراء", "عدد الحالات"]
                fig = px.bar(reasons_df, x="نوع الإجراء", y="عدد الحالات", text="عدد الحالات", title="أسباب الإجراءات")
                st.plotly_chart(make_readable_fig(fig, 430, showlegend=False), use_container_width=True, config=chart_config())

        st.markdown('<div class="mini-title">جدول المتابعة التفاعلي</div>', unsafe_allow_html=True)
        editor_cols = [c for c in [
            "الأولوية", "حالة المتابعة", "نوع الإجراء", "نطاق ساعة الاستلام", "وقت عرض", "الفرع",
            "رقم الطلب الظاهر", "الحالة", "العميل", "المنتج", "الحشوة",
            "رقم الجوال المستخرج", "واتساب", "ملاحظة داخلية", "الملاحظة"
        ] if c in action_view.columns]
        action_editor = action_view[editor_cols].copy() if editor_cols else action_view.copy()

        edited_action = action_editor
        with st.expander(f"✅ عرض جدول المتابعة التفاعلي — {format_int(len(action_editor))} صف", expanded=False):
            try:
                edited_action = st.data_editor(
                    action_editor,
                    use_container_width=True,
                    height=620,
                    column_config={
                        "واتساب": st.column_config.LinkColumn("واتساب", display_text="فتح واتساب"),
                        "حالة المتابعة": st.column_config.SelectboxColumn(
                            "حالة المتابعة",
                            options=[
                                "لم يبدأ", "تم التواصل", "بانتظار الصورة", "تم استلام الصورة",
                                "تم إرسالها للإنتاج", "تم التجهيز", "جاهز للتسليم",
                                "مشكلة / يحتاج تدخل"
                            ],
                        ),
                        "ملاحظة داخلية": st.column_config.TextColumn("ملاحظة داخلية"),
                    },
                    disabled=[c for c in action_editor.columns if c not in ["حالة المتابعة", "ملاحظة داخلية"]],
                    key="action_center_editor_v8",
                )
            except Exception:
                edited_action = action_editor
                st.dataframe(action_editor, use_container_width=True, height=620)

        st.download_button(
            "⬇️ تحميل Action Center Excel",
            data=build_multi_sheet_excel({"Action Center": edited_action}),
            file_name="Action_Center_V8_2.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_action_center_v8",
        )


with tab_print:
    st.markdown(f'<div class="section-title">{tr("🖨️ عرض الطباعة", "🖨️ Print View")}</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="readability-note">هذه الصفحة مصممة للطباعة أو Screenshot من الجوال. من المتصفح استخدم Ctrl+P أو Print. الجداول هنا مختصرة وواضحة بدون شارتات.</div>',
        unsafe_allow_html=True,
    )

    print_source = reports.get("production_queue", pd.DataFrame()).copy()
    if print_source.empty:
        st.info("لا توجد بيانات للطباعة ضمن الفلاتر الحالية.")
    else:
        pf1, pf2, pf3 = st.columns(3)
        with pf1:
            p_branches = ["الكل"] + sorted(print_source["الفرع"].dropna().unique().tolist()) if "الفرع" in print_source.columns else ["الكل"]
            p_branch = st.selectbox("فرع للطباعة", p_branches, key="v8_print_branch")
        with pf2:
            p_need_action_only = st.checkbox("فقط ما يحتاج متابعة", value=False, key="v8_print_action_only")
        with pf3:
            p_hide_addons = st.checkbox("إخفاء الإضافات", value=False, key="v8_print_hide_addons")

        print_view = print_source.copy()
        if p_branch != "الكل" and "الفرع" in print_view.columns:
            print_view = print_view[print_view["الفرع"].eq(p_branch)]
        if p_need_action_only and "أولوية" in print_view.columns:
            print_view = print_view[print_view["أولوية"].eq("عالية")]
        if p_hide_addons and "نوع الصنف" in print_view.columns:
            print_view = print_view[print_view["نوع الصنف"].ne("إضافة")]

        pp1, pp2, pp3 = st.columns(3)
        with pp1:
            render_kpi("صفوف الطباعة", format_int(len(print_view)), "", "#0ea5e9")
        with pp2:
            render_kpi("طلبات", format_int(print_view["رقم الطلب الموحد"].nunique() if "رقم الطلب الموحد" in print_view.columns else 0), "", "#2563eb")
        with pp3:
            render_kpi("تحتاج متابعة", format_int(print_view["أولوية"].eq("عالية").sum() if "أولوية" in print_view.columns else 0), "", "#dc2626")

        render_print_cards(print_view, limit=140)

        st.download_button(
            "⬇️ تحميل Print View Excel",
            data=build_multi_sheet_excel({"Print View": print_view}),
            file_name="Print_View_V8_2.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_print_view_v8",
        )


with tab_prep:
    st.markdown(f'<div class="section-title">{tr("🧾 تجهيز الفروع", "🧾 Branch Prep Sheet")}</div>', unsafe_allow_html=True)
    prep_branches = sorted(active_items["الفرع"].dropna().unique().tolist()) if not active_items.empty else []
    selected_prep_branch = st.selectbox(ui("اختار الفرع"), prep_branches if prep_branches else ["-"])
    prep = active_items[active_items["الفرع"].eq(selected_prep_branch)].copy() if selected_prep_branch != "-" else pd.DataFrame()
    prep_cols = ["وقت الاستلام الأصلي", "رقم الطلب الظاهر", "الحالة", "العميل", "المنتج", "الحشوة", "الكمية رقم", "إجمالي المنتج رقم", "سبب المتابعة", "الملاحظة"]
    prep_view = prep[[c for c in prep_cols if c in prep.columns]].sort_values(["وقت الاستلام الأصلي", "رقم الطلب الظاهر"], na_position="last") if not prep.empty else pd.DataFrame()
    display_df(prep_view, 560)
    if not prep_view.empty:
        st.download_button(
            tr("⬇️ تحميل تقرير تجهيز الفرع Excel", "⬇️ Download branch prep Excel"),
            data=branch_prep_excel(prep_view, selected_prep_branch),
            file_name=f"branch_prep_{selected_prep_branch}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


with tab_sales:
    st.markdown(f'<div class="section-title">{tr("💰 تحليل المبيعات", "💰 Sales Analytics")}</div>', unsafe_allow_html=True)
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

    st.markdown(f'<div class="mini-title">{ui("الحالات حسب الفرع")}</div>', unsafe_allow_html=True)
    status_df = reports.get("status_report", pd.DataFrame())
    if not status_df.empty:
        fig = px.bar(status_df, x="الفرع", y="عدد الطلبات", color="الحالة", barmode="group", title=ui("حالات الطلبات حسب الفرع"), labels=px_labels())
        st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=chart_config())
    display_df(status_df, 330)


with tab_products:
    st.markdown(f'<div class="section-title">{tr("🧁 أداء المنتجات", "🧁 Product Performance")}</div>', unsafe_allow_html=True)
    prod = reports.get("product_performance", pd.DataFrame())
    if not prod.empty:
        top_prod = prod.head(15)
        fig = px.bar(top_prod.sort_values("الكمية"), x="الكمية", y="المنتج", orientation="h", text="الكمية", title="أعلى المنتجات حسب الكمية", color="الكمية", color_continuous_scale="Agsunset")
        st.plotly_chart(fig_layout(fig, 560), use_container_width=True, config=chart_config())
    display_df(prod, 500)

    st.markdown(f'<div class="mini-title">{ui("المنتجات حسب الفرع")}</div>', unsafe_allow_html=True)
    display_df(reports.get("product_by_branch", pd.DataFrame()), 420)

    st.markdown('<div class="note-box">تم نقل تقارير الحشوات إلى تبويب مستقل باسم 🍰 Fillings حتى تكون واضحة ومفصلة.</div>', unsafe_allow_html=True)


with tab_varieties:
    st.markdown(f'<div class="section-title">{tr("🍰 الحشوات", "🍰 Fillings")}</div>', unsafe_allow_html=True)

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
        st.markdown(
            '<div class="readability-note">لمنع تداخل النصوص، يعرض الرسم أعلى 6 حشوات فقط بأسماء مختصرة. الاسم الكامل موجود في الجدول وExcel.</div>',
            unsafe_allow_html=True,
        )

        heat = branch_variety_heatmap.copy()
        top_cols = heat.sum(axis=0).sort_values(ascending=False).head(6).index
        heat = heat[top_cols]
        heat.index = [short_label(x, 24) for x in heat.index]
        heat.columns = [short_label(x, 18) for x in heat.columns]

        fig = px.imshow(
            heat,
            text_auto=True,
            aspect="auto",
            title="كمية أعلى الحشوات حسب الفروع",
            color_continuous_scale="YlGnBu",
        )
        fig.update_xaxes(side="top", tickangle=-20, tickfont=dict(size=12))
        fig.update_yaxes(tickfont=dict(size=12))
        fig.update_layout(margin=dict(l=170, r=35, t=105, b=80))
        st.plotly_chart(
            make_readable_fig(fig, heatmap_height(len(heat), min_height=520, max_height=760, row_px=54), showlegend=False),
            use_container_width=True,
            config=chart_config(),
        )

    st.markdown('<div class="mini-title">الحشوات حسب المنتج</div>', unsafe_allow_html=True)
    display_df(variety_by_product, 520)

    if not product_variety_heatmap.empty:
        st.markdown('<div class="mini-title">Heatmap المنتج × الحشوة — أعلى المنتجات</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="readability-note">هذا الرسم كان متداخل بسبب أسماء المنتجات الطويلة. الآن نعرض أعلى 10 منتجات × أعلى 6 حشوات بأسماء مختصرة، والتفاصيل الكاملة موجودة في الجدول وExcel.</div>',
            unsafe_allow_html=True,
        )

        top_products_for_heatmap = product_variety_heatmap.sum(axis=1).sort_values(ascending=False).head(10).index
        pv_heat = product_variety_heatmap.loc[top_products_for_heatmap]
        top_variety_cols = pv_heat.sum(axis=0).sort_values(ascending=False).head(6).index
        pv_heat = pv_heat[top_variety_cols]

        pv_heat.index = [short_label(x, 32) for x in pv_heat.index]
        pv_heat.columns = [short_label(x, 18) for x in pv_heat.columns]

        fig = px.imshow(
            pv_heat,
            text_auto=True,
            aspect="auto",
            title="توزيع الحشوات حسب المنتجات",
            color_continuous_scale="Teal",
        )
        fig.update_xaxes(side="top", tickangle=-20, tickfont=dict(size=12))
        fig.update_yaxes(tickfont=dict(size=12))
        fig.update_layout(margin=dict(l=260, r=35, t=110, b=95))
        st.plotly_chart(
            make_readable_fig(fig, heatmap_height(len(pv_heat), min_height=620, max_height=880, row_px=58), showlegend=False),
            use_container_width=True,
            config=chart_config(),
        )

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
    st.markdown(f'<div class="section-title">{tr("🎈 الإضافات والبيع الإضافي", "🎈 Add-ons & Upsell")}</div>', unsafe_allow_html=True)
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
    st.markdown(f'<div class="section-title">{tr("🚨 طلبات تحتاج متابعة", "🚨 Orders Need Action")}</div>', unsafe_allow_html=True)
    reasons = reports.get("need_action_reasons", pd.DataFrame())
    if not reasons.empty:
        fig = px.bar(reasons, x="سبب المتابعة", y="عدد الحالات", text="عدد الحالات", title="أسباب المتابعة", color="عدد الحالات", color_continuous_scale="Reds")
        st.plotly_chart(fig_layout(fig, 390), use_container_width=True, config=chart_config())
    display_df(reports.get("need_action", pd.DataFrame()), 620)


with tab_campaigns:
    st.markdown(f'<div class="section-title">{tr("🎯 تحليل الحملات", "🎯 Campaign Analyzer")}</div>', unsafe_allow_html=True)
    camp = reports.get("campaign_summary", pd.DataFrame())
    if not camp.empty:
        fig = px.bar(camp, x="الحملة", y="عدد_الطلبات", text="عدد_الطلبات", title="أداء الحملات حسب عدد الطلبات", color="عدد_الطلبات", color_continuous_scale="Purples")
        st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=chart_config())
    display_df(camp, 330)
    st.markdown('<div class="mini-title">منتجات كل حملة</div>', unsafe_allow_html=True)
    display_df(reports.get("campaign_products", pd.DataFrame()), 500)


with tab_branch:
    st.markdown(f'<div class="section-title">{tr("🏬 تحليل الفرع", "🏬 Branch Deep Dive")}</div>', unsafe_allow_html=True)
    bd_branches = sorted(active_items["الفرع"].dropna().unique().tolist()) if not active_items.empty else []
    bd_branch = st.selectbox(ui("اختار فرع للتحليل العميق"), bd_branches if bd_branches else ["-"], key="bd_branch")
    b_items = active_items[active_items["الفرع"].eq(bd_branch)].copy() if bd_branch != "-" else pd.DataFrame()
    b_orders = active_orders[active_orders["الفرع"].eq(bd_branch)].copy() if bd_branch != "-" else pd.DataFrame()
    if not b_orders.empty:
        c1, c2, c3, c4 = st.columns(4)
        with c1: render_kpi(ui("طلبات الفرع"), format_int(b_orders["رقم الطلب الموحد"].nunique()), translate_value(bd_branch), "#2563eb")
        with c2: render_kpi(ui("مبيعات الفرع"), format_money(b_orders["قيمة الطلب"].sum()), "", "#16a34a")
        with c3: render_kpi(ui("متوسط الطلب"), format_money(b_orders["قيمة الطلب"].mean()), "", "#0891b2")
        with c4: render_kpi(ui("تحتاج متابعة"), format_int(b_orders["يحتاج متابعة"].sum()), "", "#dc2626")
        col1, col2 = st.columns(2)
        with col1:
            bp = b_items[~b_items["إضافة؟"]].groupby("المنتج").agg(الكمية=("الكمية رقم", "sum"), الطلبات=("رقم الطلب الموحد", "nunique")).reset_index().sort_values("الكمية", ascending=False).head(12)
            if not bp.empty:
                fig = px.bar(bp.sort_values("الكمية"), x="الكمية", y="المنتج", orientation="h", title="أفضل منتجات الفرع")
                st.plotly_chart(fig_layout(fig, 460), use_container_width=True, config=chart_config())
        with col2:
            bh = b_orders.groupby("الساعة")["رقم الطلب الموحد"].nunique().reset_index(name=ui("عدد الطلبات"))
            if not bh.empty:
                fig = px.line(bh, x="الساعة", y="عدد الطلبات", markers=True, title=ui("ضغط الفرع حسب الساعة"), labels=px_labels())
                st.plotly_chart(fig_layout(fig, 460), use_container_width=True, config=chart_config())
        display_df(b_items[[c for c in ["رقم الطلب الظاهر", "الحالة", "وقت الاستلام الأصلي", "العميل", "المنتج", "الحشوة", "الكمية رقم", "سبب المتابعة", "الملاحظة"] if c in b_items.columns]], 500)
    else:
        st.info(tr("لا توجد بيانات لهذا الفرع ضمن الفلاتر.", "No data for this branch with the current filters."))


with tab_product:
    st.markdown(f'<div class="section-title">{tr("🔍 تحليل المنتج", "🔍 Product Deep Dive")}</div>', unsafe_allow_html=True)
    product_list = sorted(active_items[~active_items["إضافة؟"]]["المنتج"].dropna().unique().tolist()) if not active_items.empty else []
    selected_product = st.selectbox(ui("اختار المنتج"), product_list if product_list else ["-"])
    p_items = active_items[active_items["المنتج"].eq(selected_product)].copy() if selected_product != "-" else pd.DataFrame()
    if not p_items.empty:
        p_order_ids = p_items["رقم الطلب الموحد"].unique().tolist()
        p_orders = active_orders[active_orders["رقم الطلب الموحد"].isin(p_order_ids)].copy()
        c1, c2, c3, c4 = st.columns(4)
        with c1: render_kpi(ui("طلبات المنتج"), format_int(p_items["رقم الطلب الموحد"].nunique()), "", "#2563eb")
        with c2: render_kpi(ui("كمية المنتج"), format_int(p_items["الكمية رقم"].sum()), "", "#16a34a")
        with c3: render_kpi(ui("مبيعات المنتج"), format_money(p_items["إجمالي المنتج رقم"].sum()), "Item Total", "#0891b2")
        with c4: render_kpi(ui("طلبات بإضافات"), format_int(p_orders["فيه إضافات"].sum()) if not p_orders.empty else "0", "", "#f59e0b")
        col1, col2 = st.columns(2)
        with col1:
            pb = p_items.groupby("الفرع")["رقم الطلب الموحد"].nunique().reset_index(name=ui("عدد الطلبات")).sort_values(ui("عدد الطلبات"), ascending=False)
            if not pb.empty:
                fig = px.bar(pb, x="الفرع", y="عدد الطلبات", text="عدد الطلبات", title="المنتج حسب الفرع")
                st.plotly_chart(fig_layout(fig, 420), use_container_width=True, config=chart_config())
        with col2:
            pv = p_items[p_items["الحشوة"].astype(str).str.len() > 0].groupby("الحشوة")["الكمية رقم"].sum().reset_index(name="الكمية").sort_values("الكمية", ascending=False)
            if not pv.empty:
                fig = px.pie(pv, names="الحشوة", values="الكمية", hole=.52, title="حشوات المنتج")
                st.plotly_chart(fig_layout(fig, 420), use_container_width=True, config=chart_config())
        st.markdown(f'<div class="mini-title">{ui("كل طلبات المنتج")}</div>', unsafe_allow_html=True)
        display_df(p_items[[c for c in ["رقم الطلب الظاهر", "الفرع", "الحالة", "وقت الاستلام الأصلي", "العميل", "الحشوة", "الكمية رقم", "إجمالي المنتج رقم", "سبب المتابعة", "الملاحظة"] if c in p_items.columns]], 520)
    else:
        st.info(tr("لا توجد بيانات لهذا المنتج ضمن الفلاتر.", "No data for this product with the current filters."))


with tab_quality:
    st.markdown(f'<div class="section-title">{tr("🧹 جودة البيانات", "🧹 Data Quality")}</div>', unsafe_allow_html=True)
    qsum = reports.get("data_quality_summary", pd.DataFrame())
    if not qsum.empty:
        fig = px.bar(qsum, x="المشكلة", y="عدد الصفوف", color="الأهمية", text="عدد الصفوف", title="مشاكل جودة البيانات")
        st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=chart_config())
    display_df(qsum, 300)
    st.markdown(f'<div class="mini-title">{ui("تفاصيل الصفوف التي تحتاج تنظيف")}</div>', unsafe_allow_html=True)
    display_df(reports.get("data_quality_details", pd.DataFrame()), 520)



with tab_advanced:
    st.markdown(f'<div class="section-title">{tr("📊 باقة التقارير المتقدمة V8.4", "📊 V8.4 Advanced Reports Pack")}</div>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="report-card">{tr("تقارير إدارية متقدمة لاتخاذ القرار: ترتيب الفروع، قيمة المنتجات، ذكاء الحشوات، فرص الإضافات، ملاحظات العملاء، جودة البيانات، الطاقة التشغيلية، والحملات.", "Advanced management reports for decision-making: branch ranking, product value, filling intelligence, add-on opportunities, customer notes, data quality, operational capacity, and campaigns.")}</div>',
        unsafe_allow_html=True,
    )

    exec_summary = reports.get("advanced_executive_summary", pd.DataFrame())
    branch_rank = reports.get("advanced_branch_ranking", pd.DataFrame())
    product_value = reports.get("advanced_product_value", pd.DataFrame())
    filling_intel = reports.get("advanced_filling_intelligence", pd.DataFrame())
    filling_branch_rank = reports.get("advanced_filling_by_branch_rank", pd.DataFrame())
    addons_branch = reports.get("advanced_addons_opportunity_branch", pd.DataFrame())
    addons_product = reports.get("advanced_addons_product_opportunity", pd.DataFrame())
    notes_keywords = reports.get("advanced_notes_keywords", pd.DataFrame())
    notes_product = reports.get("advanced_notes_by_product", pd.DataFrame())
    quality_score = reports.get("advanced_data_quality_score", pd.DataFrame())
    hourly_capacity = reports.get("advanced_hourly_capacity", pd.DataFrame())
    campaign_perf = reports.get("advanced_campaign_performance", pd.DataFrame())

    # Executive KPIs
    ar1, ar2, ar3, ar4 = st.columns(4)
    with ar1:
        render_kpi(ui("أفضل فرع بالمبيعات"), short_label(translate_value(top_branch), 26), f"{format_int(top_branch_count)} {tr('طلب', 'orders')}", "#f59e0b")
    with ar2:
        render_kpi("Upsell Rate", f"{upsell_rate:.1f}%", f"{format_int(addon_orders)} طلب بإضافات", "#16a34a")
    with ar3:
        render_kpi("Action Rate", f"{(need_action_count / total_orders * 100 if total_orders else 0):.1f}%", f"{format_int(need_action_count)} طلب", "#dc2626")
    with ar4:
        render_kpi(ui("أعلى نطاق ساعة"), translate_value(top_hour), f"{format_int(top_hour_count)} {tr('طلب', 'orders')}", "#0ea5e9")

    display_df(exec_summary, 320, "عرض Executive Summary")

    # Branch Ranking
    st.markdown('<div class="mini-title">🏬 Branch Ranking Report</div>', unsafe_allow_html=True)
    if not branch_rank.empty:
        top_branches_chart = branch_rank.head(12).copy()
        top_branches_chart["الفرع للعرض"] = top_branches_chart["الفرع"].apply(lambda x: short_label(x, 24))
        fig = px.bar(
            top_branches_chart.sort_values("المبيعات"),
            x="المبيعات",
            y="الفرع للعرض",
            orientation="h",
            text="الطلبات",
            title="ترتيب الفروع حسب المبيعات",
            color="نسبة_متابعة_%",
            color_continuous_scale="RdYlGn_r",
            hover_data={"الفرع": True, "المبيعات": ":,.0f", "متوسط_الطلب": ":,.0f", "نسبة_Upsell_%": True, "نسبة_متابعة_%": True, "الفرع للعرض": False},
        )
        fig.update_traces(texttemplate="%{text} " + tr("طلب", "orders"), textposition="outside", cliponaxis=False)
        st.plotly_chart(make_readable_fig(fig, 560, showlegend=False), use_container_width=True, config=chart_config())
    display_df(branch_rank, 520, "عرض جدول Branch Ranking")

    # Product Value
    st.markdown('<div class="mini-title">🧁 Product Value Report</div>', unsafe_allow_html=True)
    if not product_value.empty:
        pv = product_value.head(20).copy()
        pv["المنتج للعرض"] = pv["المنتج"].apply(lambda x: short_label(x, 34))
        fig = px.scatter(
            pv,
            x="الكمية",
            y="مبيعات_لكل_طلب",
            size="المبيعات",
            color="تصنيف_القيمة",
            hover_name="المنتج",
            title="خريطة قيمة المنتجات: كمية × مبيعات لكل طلب",
        )
        st.plotly_chart(make_readable_fig(fig, 560, showlegend=True), use_container_width=True, config=chart_config())

        top_pv = product_value.head(12).copy()
        top_pv["المنتج للعرض"] = top_pv["المنتج"].apply(lambda x: short_label(x, 30))
        fig = px.bar(
            top_pv.sort_values("المبيعات"),
            x="المبيعات",
            y="المنتج للعرض",
            orientation="h",
            text="الكمية",
            title="أعلى المنتجات بالقيمة",
            color="حصة_المبيعات_%",
            color_continuous_scale="Blues",
            hover_data={"المنتج": True, "الطلبات": True, "الكمية": True, "نسبة_متابعة_%": True, "المنتج للعرض": False},
        )
        st.plotly_chart(make_readable_fig(fig, 560, showlegend=False), use_container_width=True, config=chart_config())
    display_df(product_value, 560, "عرض جدول Product Value")

    # Filling Intelligence
    st.markdown('<div class="mini-title">🍰 Filling Intelligence</div>', unsafe_allow_html=True)
    if not filling_intel.empty:
        fv = filling_intel.head(12).copy()
        fv["الحشوة للعرض"] = fv["الحشوة"].apply(lambda x: short_label(x, 24))
        fig = px.bar(
            fv.sort_values("الكمية"),
            x="الكمية",
            y="الحشوة للعرض",
            orientation="h",
            text="حصة_الكمية_%",
            title="أولوية الحشوات للإنتاج",
            color="نسبة_متابعة_%",
            color_continuous_scale="OrRd",
            hover_data={"الحشوة": True, "المبيعات": ":,.0f", "نسبة_متابعة_%": True, "الحشوة للعرض": False},
        )
        fig.update_traces(texttemplate="%{text}%", textposition="outside", cliponaxis=False)
        st.plotly_chart(make_readable_fig(fig, 520, showlegend=False), use_container_width=True, config=chart_config())
    display_df(filling_intel, 440, "عرض جدول Filling Intelligence")
    display_df(filling_branch_rank, 440, "عرض ترتيب الحشوات داخل كل فرع")

    # Add-ons Opportunity
    st.markdown('<div class="mini-title">🎈 Add-ons Opportunity</div>', unsafe_allow_html=True)
    if not addons_branch.empty:
        ab = addons_branch.copy()
        ab["الفرع للعرض"] = ab["الفرع"].apply(lambda x: short_label(x, 24))
        fig = px.bar(
            ab.sort_values("نسبة_Upsell_%"),
            x="الفرع للعرض",
            y="نسبة_Upsell_%",
            text="طلبات_بدون_إضافات",
            color="فرصة",
            title="فرص رفع Upsell حسب الفرع",
            hover_data={"الفرع": True, "الطلبات": True, "طلبات_بإضافات": True, "فرق_المتوسط": ":,.0f", "الفرع للعرض": False},
        )
        fig.update_traces(texttemplate="%{y:.1f}% | بدون إضافات: %{text}", textposition="outside", cliponaxis=False)
        st.plotly_chart(make_readable_fig(fig, 520, showlegend=True), use_container_width=True, config=chart_config())
    display_df(addons_branch, 420, "عرض فرص الإضافات حسب الفرع")
    display_df(addons_product, 500, "عرض فرص الإضافات حسب المنتج")

    # Notes Intelligence
    st.markdown('<div class="mini-title">📝 Customer Notes Intelligence</div>', unsafe_allow_html=True)
    if not notes_keywords.empty:
        fig = px.bar(
            notes_keywords.sort_values("عدد الصفوف"),
            x="عدد الصفوف",
            y="الكلمة/التصنيف",
            orientation="h",
            text="نسبة من الصفوف %",
            title="تصنيف ملاحظات العملاء",
            color="عدد الصفوف",
            color_continuous_scale="Purples",
        )
        fig.update_traces(texttemplate="%{text}%", textposition="outside", cliponaxis=False)
        st.plotly_chart(make_readable_fig(fig, 430, showlegend=False), use_container_width=True, config=chart_config())
    display_df(notes_keywords, 320, "عرض كلمات وملاحظات العملاء")
    display_df(notes_product, 460, "عرض المنتجات ذات الملاحظات الأعلى")

    # Data Quality Score
    st.markdown('<div class="mini-title">🧹 Data Quality Score</div>', unsafe_allow_html=True)
    if not quality_score.empty:
        qs = quality_score.copy()
        qs["الفرع للعرض"] = qs["الفرع"].apply(lambda x: short_label(x, 24))
        fig = px.bar(
            qs.sort_values("Quality Score"),
            x="Quality Score",
            y="الفرع للعرض",
            orientation="h",
            text="التقييم",
            title="تقييم جودة البيانات حسب الفرع",
            color="Quality Score",
            color_continuous_scale="RdYlGn",
            hover_data={"الفرع": True, "إجمالي_مشاكل_الجودة": True, "مشاكل_لكل_100_صف": True, "الفرع للعرض": False},
        )
        st.plotly_chart(make_readable_fig(fig, 520, showlegend=False), use_container_width=True, config=chart_config())
    display_df(quality_score, 420, "عرض Data Quality Score")

    # Hourly Capacity
    st.markdown('<div class="mini-title">⏰ Hourly Capacity Report</div>', unsafe_allow_html=True)
    if not hourly_capacity.empty:
        fig = go.Figure()
        fig.add_trace(go.Bar(x=hourly_capacity["نطاق ساعة الاستلام"], y=hourly_capacity["الطلبات"], name="الطلبات"))
        if "الكمية" in hourly_capacity.columns:
            fig.add_trace(go.Scatter(x=hourly_capacity["نطاق ساعة الاستلام"], y=hourly_capacity["الكمية"], name="الكمية", mode="lines+markers", yaxis="y2"))
        fig.update_layout(
            title="الطاقة التشغيلية حسب نطاق الساعة",
            yaxis=dict(title=ui("عدد الطلبات")),
            yaxis2=dict(title="الكمية", overlaying="y", side="right"),
        )
        st.plotly_chart(make_readable_fig(fig, 560, showlegend=True), use_container_width=True, config=chart_config())
    display_df(hourly_capacity, 420, "عرض Hourly Capacity")

    # Campaign Performance
    st.markdown('<div class="mini-title">🎯 Campaign Performance Advanced</div>', unsafe_allow_html=True)
    if not campaign_perf.empty:
        cp = campaign_perf.copy()
        fig = px.bar(
            cp,
            x="الحملة",
            y="مبيعات_الأصناف",
            text="الطلبات",
            color="نسبة_متابعة_%",
            color_continuous_scale="RdYlGn_r",
            title="أداء الحملات حسب المبيعات والمتابعة",
            hover_data={"أفضل_فرع": True, "أفضل_منتج": True, "أفضل_حشوة": True},
        )
        fig.update_traces(texttemplate="%{text} " + tr("طلب", "orders"), textposition="outside", cliponaxis=False)
        st.plotly_chart(make_readable_fig(fig, 520, showlegend=False), use_container_width=True, config=chart_config())
    display_df(campaign_perf, 440, "عرض Campaign Performance")


with tab_export:
    st.markdown(f'<div class="section-title">{tr("⬇️ مركز التصدير", "⬇️ Export Center")}</div>', unsafe_allow_html=True)
    summary_rows = [
        {"البند": "الإصدار", "القيمة": APP_VERSION},
        {"البند": "عدد الصفوف بعد الفلاتر", "القيمة": total_rows},
        {"البند": ui("عدد الطلبات"), "القيمة": total_orders},
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
        tr("⬇️ تحميل Excel شامل كل التقارير V8.4", "⬇️ Download Excel with all V8.4 reports"),
        data=excel_file,
        file_name="MAD_Orders_Control_Center_V8_4.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.markdown(
        f'<div class="note-box">{tr("الملف يحتوي على Executive Summary، Advanced Reports Pack، Branch Ranking، Product Value، Filling Intelligence، Add-ons Opportunity، Customer Notes، Data Quality Score، Hourly Capacity، Campaign Performance، وكل تقارير التشغيل السابقة.", "The file includes Executive Summary, Advanced Reports Pack, Branch Ranking, Product Value, Filling Intelligence, Add-ons Opportunity, Customer Notes, Data Quality Score, Hourly Capacity, Campaign Performance, and all previous operational reports.")}</div>',
        unsafe_allow_html=True,
    )
