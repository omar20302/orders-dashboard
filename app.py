import re
from io import BytesIO
from datetime import datetime
from urllib.parse import urlparse, parse_qs

import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import ColorScaleRule


# ============================================================
# MAD Orders Analysis App
# تحليل الطلبات حسب الفرع والساعة + الحشوات + Excel Dashboard
# ============================================================


# =========================
# Helpers
# =========================

def normalize_arabic_digits(value):
    """تحويل الأرقام العربية/الفارسية إلى أرقام إنجليزية."""
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
    """البحث عن اسم العمود سواء مطابق أو يحتوي على الاسم."""
    cols = list(df.columns)

    for name in possible_names:
        if name in cols:
            return name

    for col in cols:
        clean_col = str(col).strip().lower()
        for name in possible_names:
            if str(name).strip().lower() in clean_col:
                return col

    return None


def read_uploaded_file(uploaded_file):
    """قراءة ملفات TXT / CSV / Excel."""
    filename = uploaded_file.name.lower()

    if filename.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded_file, dtype=str).fillna("")

    # الملف غالباً Tab-separated
    try:
        df = pd.read_csv(
            uploaded_file,
            sep="\t",
            dtype=str,
            keep_default_na=False,
            engine="python"
        )
        if len(df.columns) > 1:
            return df.fillna("")
    except Exception:
        pass

    uploaded_file.seek(0)

    # محاولة قراءة CSV بشكل تلقائي
    df = pd.read_csv(
        uploaded_file,
        sep=None,
        dtype=str,
        keep_default_na=False,
        engine="python"
    )

    return df.fillna("")


def extract_branch(chef_name):
    """
    استخراج اسم الفرع من خانة اسم الشيف.
    لو مفيش اسم فرع واضح يرجع: العقيق
    """
    text = str(chef_name).strip()
    lower_text = text.lower()

    branch_keywords = {
        "قرطبة": ["قرطبة", "qurtuba", "qortoba"],
        "عريجاء": ["عريجاء", "عريجا", "uraija", "uraijaa", "urejha"],
        "الروضة": ["الروضة", "روضه", "rawdah", "rawda"],
        "العارض": ["العارض", "arid", "al arid"],
        "الورود": ["الورود", "worood", "al worood"],
    }

    for branch, keywords in branch_keywords.items():
        for keyword in keywords:
            if keyword.lower() in lower_text:
                return branch

    return "العقيق"


def parse_datetime(delivery_date, pickup_time):
    """
    قراءة التاريخ والوقت بأكثر من صيغة:
    21-06-2026 + 8:16 PM
    2026-06-21 7:00 PM
    6/21/2026 19:00
    """
    delivery_date = normalize_arabic_digits(delivery_date)
    pickup_time = normalize_arabic_digits(pickup_time)

    candidates = []

    # لو وقت الاستلام يحتوي على تاريخ كامل
    if re.search(r"\d{4}-\d{1,2}-\d{1,2}|\d{1,2}/\d{1,2}/\d{4}", pickup_time):
        candidates.append(pickup_time)

    # التاريخ في عمود والوقت في عمود
    if delivery_date and pickup_time:
        candidates.append(f"{delivery_date} {pickup_time}")

    # آخر محاولة
    if pickup_time:
        candidates.append(pickup_time)

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
    """تحويل رقم الساعة إلى شكل: 5-6 م"""
    def fmt(h):
        period = "ص" if h < 12 else "م"
        h12 = h % 12
        if h12 == 0:
            h12 = 12
        return f"{h12} {period}"

    end_hour = (hour + 1) % 24

    start_txt = fmt(hour).replace(" ص", "").replace(" م", "")
    end_txt = fmt(end_hour)

    return f"{start_txt}-{end_txt}"


def clean_quantity(value):
    """تنظيف الكمية وتحويلها لرقم."""
    value = normalize_arabic_digits(value)
    value = value.replace(",", "").strip()

    if value == "":
        return 1

    number = pd.to_numeric(value, errors="coerce")

    if pd.isna(number):
        return 1

    return float(number)


def normalize_variety(value):
    """
    توحيد أسماء الحشوات.
    لو الحشوة فاضية معناها غالباً إضافة مثل شموع/بالون/كرت، فيتم تجاهلها.
    """
    text = str(value).strip()

    if not text:
        return ""

    if text.lower() in ["nan", "none", "(not set)", "not set", "-"]:
        return ""

    low = text.lower()

    if "vanilla with sprinkles" in low or "وسبرينكلز" in text or "سبرينكلز" in text:
        return "Vanilla Sprinkles"

    if "lemon" in low and ("raspberry" in low or "رازبيري" in text or "رسبيري" in text):
        return "Lemon Raspberry"

    if "ليمون" in text and ("رازبيري" in text or "رسبيري" in text):
        return "Lemon Raspberry"

    if "vanilla" in low or "فانيلا" in text:
        return "Vanilla"

    if "chocolate" in low or "شوكولات" in text or "شوكولاته" in text:
        return "Chocolate"

    if "oreo" in low or "أوريو" in text or "اوريو" in text:
        return "Oreo"

    if "mango" in low or "مانجو" in text:
        return "Mango"

    if "rocher" in low or "ferrero" in low or "روشيه" in text or "فيريرو" in text:
        return "Rocher / Ferrero"

    if "love you" in low or "أحبك" in text or "احبك" in text:
        return "Love You"

    # لو حشوة غير معروفة، سيبها باسمها
    return text


def is_cancelled(status):
    """معرفة هل الطلب ملغي."""
    status = str(status).strip().lower()

    cancelled_words = [
        "cancelled",
        "canceled",
        "cancel",
        "ملغي",
        "ملغى",
        "الغاء",
        "إلغاء",
    ]

    return any(word in status for word in cancelled_words)


def make_numeric_table(df):
    """تحويل الأرقام التي تظهر كـ float إلى int عندما يكون مناسباً."""
    out = df.copy()
    for col in out.columns:
        try:
            if pd.api.types.is_numeric_dtype(out[col]):
                out[col] = out[col].apply(lambda x: int(x) if pd.notna(x) and float(x).is_integer() else x)
        except Exception:
            pass
    return out


def create_excel_file(
    summary_df,
    hourly_df,
    fillings_qty_df,
    fillings_orders_df,
    fillings_by_hour_branch_df,
    branch_hour_details_dict,
    clean_df,
    unique_orders_df
):
    """
    إنشاء ملف Excel تفاعلي ومتوافق:
    - لا يستخدم FILTER / UNIQUE حتى لا يقوم Excel بعمل Repair.
    - شيت Cleaned Data يحتوي على Helper Columns تبدأ بـ Calc_.
    - الجداول والرسومات تعتمد على SUMIFS بسيطة.
    - عند تعديل أو حذف الصفوف من Cleaned Data تتحدث النتائج والرسومات.
    """
    output = BytesIO()

    COLORS = {
        "dark": "1F2937",
        "navy": "0F172A",
        "blue": "1D4ED8",
        "light_blue": "DBEAFE",
        "green": "16A34A",
        "light_green": "DCFCE7",
        "orange": "EA580C",
        "purple": "7C3AED",
        "red": "B91C1C",
        "light_red": "FEE2E2",
        "yellow": "FDE68A",
        "light_yellow": "FEF3C7",
        "gray": "F3F4F6",
        "white": "FFFFFF",
        "border": "D1D5DB",
    }

    thin_border = Border(
        left=Side(style="thin", color=COLORS["border"]),
        right=Side(style="thin", color=COLORS["border"]),
        top=Side(style="thin", color=COLORS["border"]),
        bottom=Side(style="thin", color=COLORS["border"]),
    )

    header_fill = PatternFill("solid", fgColor=COLORS["navy"])
    header_font = Font(color=COLORS["white"], bold=True, size=11)
    title_font = Font(color=COLORS["white"], bold=True, size=20)
    subtitle_font = Font(color=COLORS["white"], size=11)

    # =========================
    # Prepare export data
    # =========================

    export_df = clean_df.copy()

    order_id_col = find_col(export_df, ["معرف الطلب (Order Id)", "Order Id", "order_id"])
    if order_id_col is None:
        order_id_col = "Calc_Source_OrderId"
        export_df[order_id_col] = ""

    export_df["Calc_OrderId"] = export_df[order_id_col].astype(str)
    export_df["Calc_Branch"] = export_df["الفرع"].astype(str) if "الفرع" in export_df.columns else ""
    export_df["Calc_Hour"] = export_df["الساعة"].astype(str) if "الساعة" in export_df.columns else ""
    export_df["Calc_Filling"] = export_df["الحشوة الموحدة"].astype(str) if "الحشوة الموحدة" in export_df.columns else ""
    export_df["Calc_Qty"] = export_df["كمية رقمية"] if "كمية رقمية" in export_df.columns else 1

    if "ملغي؟" in export_df.columns:
        export_df["Calc_Cancelled"] = export_df["ملغي؟"].apply(lambda x: 1 if bool(x) else 0)
    else:
        export_df["Calc_Cancelled"] = 0

    # Helper columns. Values will be replaced by Excel formulas after writing.
    export_df["Calc_UniqueOrderFlag"] = 0
    export_df["Calc_ActiveOrderFlag"] = 0
    export_df["Calc_CancelledOrderFlag"] = 0
    export_df["Calc_ActiveQty"] = 0

    active_rows = export_df[export_df["Calc_Cancelled"] == 0].copy()

    branches = [
        str(x) for x in active_rows["Calc_Branch"].dropna().unique().tolist()
        if str(x).strip() != ""
    ]
    branches = sorted(branches) or ["العقيق"]

    if "رقم الساعة" in active_rows.columns:
        hour_df = (
            active_rows[["Calc_Hour", "رقم الساعة"]]
            .dropna()
            .drop_duplicates()
            .sort_values("رقم الساعة")
        )
        hours = [str(x) for x in hour_df["Calc_Hour"].tolist() if str(x).strip() != ""]
    else:
        hours = [str(x) for x in active_rows["Calc_Hour"].dropna().unique().tolist() if str(x).strip() != ""]
    hours = hours or ["وقت غير واضح"]

    filling_rows_tmp = active_rows[active_rows["Calc_Filling"].astype(str).str.strip() != ""]
    if len(filling_rows_tmp):
        filling_totals = (
            filling_rows_tmp
            .groupby("Calc_Filling")["Calc_Qty"]
            .sum()
            .sort_values(ascending=False)
        )
        fillings = [str(x) for x in filling_totals.index.tolist()]
    else:
        fillings = ["بدون حشوة"]

    def safe_sheet_name(name):
        name = str(name)
        for ch in ["/", "\\", "*", "?", ":", "[", "]"]:
            name = name.replace(ch, "-")
        return name[:31]

    def add_table_style(ws, table_name, ref):
        try:
            tab = Table(displayName=table_name, ref=ref)
            style = TableStyleInfo(
                name="TableStyleMedium2",
                showFirstColumn=False,
                showLastColumn=False,
                showRowStripes=True,
                showColumnStripes=False
            )
            tab.tableStyleInfo = style
            ws.add_table(tab)
        except Exception:
            pass

    def apply_heatmap(ws, data_range):
        try:
            ws.conditional_formatting.add(
                data_range,
                ColorScaleRule(
                    start_type="min",
                    start_color="FEE2E2",
                    mid_type="percentile",
                    mid_value=50,
                    mid_color="FEF3C7",
                    end_type="max",
                    end_color="DCFCE7"
                )
            )
        except Exception:
            pass

    def style_grid(ws, header_row, start_row, end_row, end_col, title_row=1):
        ws.sheet_view.rightToLeft = True
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = f"A{start_row}"

        # title
        if title_row:
            ws.cell(row=title_row, column=1).fill = PatternFill("solid", fgColor=COLORS["navy"])
            ws.cell(row=title_row, column=1).font = Font(color=COLORS["white"], bold=True, size=14)
            ws.cell(row=title_row, column=1).alignment = Alignment(horizontal="center", vertical="center")

        # header
        for cell in ws[header_row]:
            if cell.column <= end_col:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = thin_border

        # body
        for row in ws.iter_rows(min_row=start_row, max_row=end_row, max_col=end_col):
            for cell in row:
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = thin_border
                cell.number_format = "#,##0"

        for c in range(1, end_col + 1):
            ws.column_dimensions[get_column_letter(c)].width = 17

    def sumifs_order(branch_cell=None, hour_cell=None, filling_cell=None):
        formula = "=SUMIFS(tblData[Calc_ActiveOrderFlag]"
        if branch_cell:
            formula += f",tblData[Calc_Branch],{branch_cell}"
        if hour_cell:
            formula += f",tblData[Calc_Hour],{hour_cell}"
        if filling_cell:
            formula += f",tblData[Calc_Filling],{filling_cell}"
        formula += ")"
        return formula

    def sumifs_qty(branch_cell=None, hour_cell=None, filling_cell=None):
        formula = "=SUMIFS(tblData[Calc_ActiveQty]"
        if branch_cell:
            formula += f",tblData[Calc_Branch],{branch_cell}"
        if hour_cell:
            formula += f",tblData[Calc_Hour],{hour_cell}"
        if filling_cell:
            formula += f",tblData[Calc_Filling],{filling_cell}"
        formula += ")"
        return formula

    def draw_card(ws, cell_range, formula_text, fill_color):
        ws.merge_cells(cell_range)
        start_cell = cell_range.split(":")[0]
        cell = ws[start_cell]
        cell.value = formula_text
        cell.fill = PatternFill("solid", fgColor=fill_color)
        cell.font = Font(color=COLORS["white"], bold=True, size=13)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

        for row in ws[cell_range]:
            for c in row:
                c.fill = PatternFill("solid", fgColor=fill_color)
                c.border = thin_border

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_df.to_excel(writer, sheet_name="Cleaned Data", index=False, startrow=3)
        workbook = writer.book

        try:
            workbook.calculation.calcMode = "auto"
            workbook.calculation.fullCalcOnLoad = True
            workbook.calculation.forceFullCalc = True
        except Exception:
            pass

        data_ws = workbook["Cleaned Data"]
        data_ws.sheet_view.rightToLeft = True
        data_ws.sheet_view.showGridLines = False
        data_ws.freeze_panes = "A5"
        data_ws.sheet_properties.tabColor = COLORS["orange"]

        data_ws.merge_cells("A1:H1")
        data_ws["A1"] = "شيت البيانات التفاعلي | عدّل أو احذف الصفوف من الجدول بالأسفل وسيتم تحديث الـ Dashboard والرسومات"
        data_ws["A1"].fill = PatternFill("solid", fgColor=COLORS["navy"])
        data_ws["A1"].font = Font(color=COLORS["white"], bold=True, size=14)
        data_ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

        data_ws.merge_cells("A2:H2")
        data_ws["A2"] = "مهم: أعمدة Calc_ هي أعمدة الحساب. لو عدّلت الفرع/الساعة/الحشوة/الكمية عدّل عمود Calc_ المقابل."
        data_ws["A2"].fill = PatternFill("solid", fgColor=COLORS["light_yellow"])
        data_ws["A2"].font = Font(color=COLORS["dark"], bold=True, size=11)
        data_ws["A2"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        data_header_row = 4
        data_start_row = 5
        data_end_row = data_ws.max_row
        data_end_col = data_ws.max_column
        data_ref = f"A{data_header_row}:{get_column_letter(data_end_col)}{data_end_row}"
        add_table_style(data_ws, "tblData", data_ref)

        headers = {}
        for c in range(1, data_end_col + 1):
            header_value = data_ws.cell(row=data_header_row, column=c).value
            headers[str(header_value)] = c

        order_c = headers.get("Calc_OrderId")
        cancel_c = headers.get("Calc_Cancelled")
        qty_c = headers.get("Calc_Qty")
        unique_c = headers.get("Calc_UniqueOrderFlag")
        active_flag_c = headers.get("Calc_ActiveOrderFlag")
        cancelled_flag_c = headers.get("Calc_CancelledOrderFlag")
        active_qty_c = headers.get("Calc_ActiveQty")

        order_letter = get_column_letter(order_c)
        cancel_letter = get_column_letter(cancel_c)
        qty_letter = get_column_letter(qty_c)

        # Fill helper formulas using classic Excel formulas, no dynamic functions.
        for r in range(data_start_row, data_end_row + 1):
            data_ws.cell(row=r, column=unique_c).value = (
                f'=IF(${order_letter}{r}="",0,IF(COUNTIF(${order_letter}${data_start_row}:${order_letter}{r},${order_letter}{r})=1,1,0))'
            )
            data_ws.cell(row=r, column=active_flag_c).value = (
                f'=IF(OR(${cancel_letter}{r}=1,${order_letter}{r}=""),0,'
                f'IF(COUNTIFS(${order_letter}${data_start_row}:${order_letter}{r},${order_letter}{r},'
                f'${cancel_letter}${data_start_row}:${cancel_letter}{r},0)=1,1,0))'
            )
            data_ws.cell(row=r, column=cancelled_flag_c).value = (
                f'=IF(OR(${cancel_letter}{r}=0,${order_letter}{r}=""),0,'
                f'IF(COUNTIFS(${order_letter}${data_start_row}:${order_letter}{r},${order_letter}{r},'
                f'${cancel_letter}${data_start_row}:${cancel_letter}{r},1)=1,1,0))'
            )
            data_ws.cell(row=r, column=active_qty_c).value = f'=IF(${cancel_letter}{r}=1,0,${qty_letter}{r})'

        # Style data sheet
        for cell in data_ws[data_header_row]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

        for row in data_ws.iter_rows(min_row=data_start_row, max_row=data_end_row, max_col=data_end_col):
            for cell in row:
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = thin_border

        for c in range(1, data_end_col + 1):
            col_letter = get_column_letter(c)
            header_name = data_ws.cell(row=data_header_row, column=c).value
            data_ws.column_dimensions[col_letter].width = 18
            if str(header_name).startswith("Calc_"):
                for cell in data_ws[col_letter]:
                    if cell.row >= data_header_row:
                        cell.fill = PatternFill("solid", fgColor=COLORS["light_yellow"])

        # =========================
        # Summary Data
        # =========================

        summary_ws = workbook.create_sheet("Summary Data")
        summary_ws.sheet_view.rightToLeft = True
        summary_ws.append(["البند", "القيمة"])
        summary_items = [
            ["عدد سطور الملف", '=COUNTA(tblData[Calc_OrderId])'],
            ["عدد الطلبات المختلفة", '=SUM(tblData[Calc_UniqueOrderFlag])'],
            ["عدد الطلبات بعد الاستبعاد", '=SUM(tblData[Calc_ActiveOrderFlag])'],
            ["عدد الطلبات الملغاة", '=SUM(tblData[Calc_CancelledOrderFlag])'],
            ["عدد السطور المكررة تقريباً", '=B2-B3'],
            ["أعلى ساعة ضغط", ""],
            ["طلبات أعلى ساعة", ""],
            ["أعلى فرع", ""],
            ["طلبات أعلى فرع", ""],
            ["أكثر حشوة", ""],
            ["كمية أكثر حشوة", ""],
        ]
        for item in summary_items:
            summary_ws.append(item)

        # =========================
        # Orders by Hour
        # =========================

        orders_ws = workbook.create_sheet("Orders by Hour")
        orders_ws.sheet_view.rightToLeft = True
        orders_ws.cell(row=1, column=1).value = "تحليل الطلبات لكل فرع بالساعة - تفاعلي"
        orders_ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(3, len(hours) + 2))

        header_row = 3
        first_data_row = 4
        orders_ws.cell(row=header_row, column=1).value = "الفرع"
        for j, hour in enumerate(hours, start=2):
            orders_ws.cell(row=header_row, column=j).value = hour
        total_col = len(hours) + 2
        orders_ws.cell(row=header_row, column=total_col).value = "الإجمالي"

        for i, branch in enumerate(branches, start=first_data_row):
            orders_ws.cell(row=i, column=1).value = branch
            for j, hour in enumerate(hours, start=2):
                col_letter = get_column_letter(j)
                orders_ws.cell(row=i, column=j).value = sumifs_order(f"$A{i}", f"{col_letter}${header_row}")
            orders_ws.cell(row=i, column=total_col).value = f"=SUM(B{i}:{get_column_letter(total_col-1)}{i})"

        total_row = first_data_row + len(branches)
        orders_ws.cell(row=total_row, column=1).value = "الإجمالي"
        for j in range(2, total_col + 1):
            col_letter = get_column_letter(j)
            orders_ws.cell(row=total_row, column=j).value = f"=SUM({col_letter}{first_data_row}:{col_letter}{total_row-1})"

        first_hour_col = "B"
        last_hour_col = get_column_letter(total_col - 1)
        total_orders_col = get_column_letter(total_col)
        summary_ws["B7"] = f'=INDEX(\'Orders by Hour\'!${first_hour_col}${header_row}:${last_hour_col}${header_row},1,MATCH(MAX(\'Orders by Hour\'!${first_hour_col}${total_row}:${last_hour_col}${total_row}),\'Orders by Hour\'!${first_hour_col}${total_row}:${last_hour_col}${total_row},0))'
        summary_ws["B8"] = f'=MAX(\'Orders by Hour\'!${first_hour_col}${total_row}:${last_hour_col}${total_row})'
        summary_ws["B9"] = f'=INDEX(\'Orders by Hour\'!$A${first_data_row}:$A${total_row-1},MATCH(MAX(\'Orders by Hour\'!${total_orders_col}${first_data_row}:${total_orders_col}${total_row-1}),\'Orders by Hour\'!${total_orders_col}${first_data_row}:${total_orders_col}${total_row-1},0))'
        summary_ws["B10"] = f'=MAX(\'Orders by Hour\'!${total_orders_col}${first_data_row}:${total_orders_col}${total_row-1})'

        # =========================
        # Fillings Qty
        # =========================

        fill_ws = workbook.create_sheet("Fillings Qty")
        fill_ws.sheet_view.rightToLeft = True
        fill_ws.cell(row=1, column=1).value = "تحليل الحشوات لكل فرع حسب الكمية - تفاعلي"
        fill_ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(3, len(fillings) + 2))

        fill_header_row = 3
        fill_first_data_row = 4
        fill_ws.cell(row=fill_header_row, column=1).value = "الفرع"
        for j, filling in enumerate(fillings, start=2):
            fill_ws.cell(row=fill_header_row, column=j).value = filling
        fill_total_col = len(fillings) + 2
        fill_ws.cell(row=fill_header_row, column=fill_total_col).value = "الإجمالي"

        for i, branch in enumerate(branches, start=fill_first_data_row):
            fill_ws.cell(row=i, column=1).value = branch
            for j, filling in enumerate(fillings, start=2):
                col_letter = get_column_letter(j)
                fill_ws.cell(row=i, column=j).value = sumifs_qty(f"$A{i}", None, f"{col_letter}${fill_header_row}")
            fill_ws.cell(row=i, column=fill_total_col).value = f"=SUM(B{i}:{get_column_letter(fill_total_col-1)}{i})"

        fill_total_row = fill_first_data_row + len(branches)
        fill_ws.cell(row=fill_total_row, column=1).value = "الإجمالي"
        for j in range(2, fill_total_col + 1):
            col_letter = get_column_letter(j)
            fill_ws.cell(row=fill_total_row, column=j).value = f"=SUM({col_letter}{fill_first_data_row}:{col_letter}{fill_total_row-1})"

        first_filling_col = "B"
        last_filling_col = get_column_letter(fill_total_col - 1)
        summary_ws["B11"] = f'=INDEX(\'Fillings Qty\'!${first_filling_col}${fill_header_row}:${last_filling_col}${fill_header_row},1,MATCH(MAX(\'Fillings Qty\'!${first_filling_col}${fill_total_row}:${last_filling_col}${fill_total_row}),\'Fillings Qty\'!${first_filling_col}${fill_total_row}:${last_filling_col}${fill_total_row},0))'
        summary_ws["B12"] = f'=MAX(\'Fillings Qty\'!${first_filling_col}${fill_total_row}:${last_filling_col}${fill_total_row})'

        # =========================
        # Fillings Orders
        # =========================

        fill_orders_ws = workbook.create_sheet("Fillings Orders")
        fill_orders_ws.sheet_view.rightToLeft = True
        fill_orders_ws.cell(row=1, column=1).value = "تحليل الحشوات لكل فرع حسب عدد الطلبات - تفاعلي"
        fill_orders_ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(3, len(fillings) + 2))

        fill_orders_ws.cell(row=fill_header_row, column=1).value = "الفرع"
        for j, filling in enumerate(fillings, start=2):
            fill_orders_ws.cell(row=fill_header_row, column=j).value = filling
        fill_orders_ws.cell(row=fill_header_row, column=fill_total_col).value = "الإجمالي"

        for i, branch in enumerate(branches, start=fill_first_data_row):
            fill_orders_ws.cell(row=i, column=1).value = branch
            for j, filling in enumerate(fillings, start=2):
                col_letter = get_column_letter(j)
                fill_orders_ws.cell(row=i, column=j).value = sumifs_order(f"$A{i}", None, f"{col_letter}${fill_header_row}")
            fill_orders_ws.cell(row=i, column=fill_total_col).value = f"=SUM(B{i}:{get_column_letter(fill_total_col-1)}{i})"

        fill_orders_ws.cell(row=fill_total_row, column=1).value = "الإجمالي"
        for j in range(2, fill_total_col + 1):
            col_letter = get_column_letter(j)
            fill_orders_ws.cell(row=fill_total_row, column=j).value = f"=SUM({col_letter}{fill_first_data_row}:{col_letter}{fill_total_row-1})"

        # =========================
        # Fillings Hour Branch
        # =========================

        fh_ws = workbook.create_sheet("Fillings Hour Branch")
        fh_ws.sheet_view.rightToLeft = True
        fh_ws.cell(row=1, column=1).value = "تحليل الحشوات بالساعة لكل فرع حسب الكمية - تفاعلي"
        fh_ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(4, len(fillings) + 3))

        fh_header_row = 3
        fh_ws.cell(row=fh_header_row, column=1).value = "الفرع"
        fh_ws.cell(row=fh_header_row, column=2).value = "الساعة"
        for j, filling in enumerate(fillings, start=3):
            fh_ws.cell(row=fh_header_row, column=j).value = filling
        fh_total_col = len(fillings) + 3
        fh_ws.cell(row=fh_header_row, column=fh_total_col).value = "الإجمالي"

        fh_row = 4
        for branch in branches:
            for hour in hours:
                fh_ws.cell(row=fh_row, column=1).value = branch
                fh_ws.cell(row=fh_row, column=2).value = hour
                for j, filling in enumerate(fillings, start=3):
                    col_letter = get_column_letter(j)
                    fh_ws.cell(row=fh_row, column=j).value = sumifs_qty(f"$A{fh_row}", f"$B{fh_row}", f"{col_letter}${fh_header_row}")
                fh_ws.cell(row=fh_row, column=fh_total_col).value = f"=SUM(C{fh_row}:{get_column_letter(fh_total_col-1)}{fh_row})"
                fh_row += 1

        # =========================
        # Dashboard
        # =========================

        dash_ws = workbook.create_sheet("Dashboard", 0)
        dash_ws.sheet_view.rightToLeft = True
        dash_ws.sheet_view.showGridLines = False
        dash_ws.sheet_properties.tabColor = COLORS["navy"]

        for col_idx in range(1, 18):
            dash_ws.column_dimensions[get_column_letter(col_idx)].width = 14
        for row_idx in range(1, 46):
            dash_ws.row_dimensions[row_idx].height = 22

        dash_ws.merge_cells("A1:Q2")
        dash_ws["A1"] = "Dashboard | تحليل الطلبات والحشوات التفاعلي"
        dash_ws["A1"].fill = PatternFill("solid", fgColor=COLORS["navy"])
        dash_ws["A1"].font = title_font
        dash_ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        for row in dash_ws["A1:Q2"]:
            for cell in row:
                cell.fill = PatternFill("solid", fgColor=COLORS["navy"])

        dash_ws.merge_cells("A3:Q3")
        dash_ws["A3"] = f"Generated: {datetime.now().strftime('%Y-%m-%d %I:%M %p')} | عدّل بيانات Cleaned Data وسيتم تحديث الجداول والرسومات"
        dash_ws["A3"].fill = PatternFill("solid", fgColor=COLORS["dark"])
        dash_ws["A3"].font = subtitle_font
        dash_ws["A3"].alignment = Alignment(horizontal="center", vertical="center")
        for cell in dash_ws["A3:Q3"][0]:
            cell.fill = PatternFill("solid", fgColor=COLORS["dark"])

        draw_card(dash_ws, "A5:D7", '="عدد الطلبات"&CHAR(10)&TEXT(\'Summary Data\'!B4,"#,##0")', COLORS["blue"])
        draw_card(dash_ws, "E5:H7", '="أعلى ساعة ضغط"&CHAR(10)&\'Summary Data\'!B7', COLORS["orange"])
        draw_card(dash_ws, "I5:L7", '="أعلى فرع"&CHAR(10)&\'Summary Data\'!B9', COLORS["green"])
        draw_card(dash_ws, "M5:Q7", '="أكثر حشوة"&CHAR(10)&\'Summary Data\'!B11', COLORS["purple"])

        dash_ws.merge_cells("A9:Q9")
        dash_ws["A9"] = "ملخص سريع"
        dash_ws["A9"].fill = PatternFill("solid", fgColor=COLORS["gray"])
        dash_ws["A9"].font = Font(color=COLORS["dark"], bold=True, size=13)
        dash_ws["A9"].alignment = Alignment(horizontal="center")

        summary_labels = [
            ("عدد سطور الملف", "B2"),
            ("عدد الطلبات المختلفة", "B3"),
            ("عدد الطلبات الملغاة", "B5"),
            ("عدد السطور المكررة تقريباً", "B6"),
            ("طلبات أعلى ساعة", "B8"),
            ("طلبات أعلى فرع", "B10"),
            ("كمية أكثر حشوة", "B12"),
        ]

        for idx, (label, ref) in enumerate(summary_labels, start=10):
            dash_ws.cell(row=idx, column=1).value = label
            dash_ws.cell(row=idx, column=2).value = f"='Summary Data'!{ref}"
            dash_ws.cell(row=idx, column=1).fill = PatternFill("solid", fgColor=COLORS["light_blue"])
            dash_ws.cell(row=idx, column=2).fill = PatternFill("solid", fgColor=COLORS["white"])
            dash_ws.cell(row=idx, column=1).font = Font(bold=True, color=COLORS["dark"])
            dash_ws.cell(row=idx, column=2).font = Font(bold=True, color=COLORS["blue"])
            dash_ws.cell(row=idx, column=1).alignment = Alignment(horizontal="center")
            dash_ws.cell(row=idx, column=2).alignment = Alignment(horizontal="center")
            dash_ws.cell(row=idx, column=1).border = thin_border
            dash_ws.cell(row=idx, column=2).border = thin_border

        # Charts
        chart1 = BarChart()
        chart1.type = "bar"
        chart1.style = 10
        chart1.title = "عدد الطلبات حسب الفرع"
        chart1.y_axis.title = "الفرع"
        chart1.x_axis.title = "عدد الطلبات"
        chart1.height = 8
        chart1.width = 15
        data = Reference(orders_ws, min_col=total_col, min_row=header_row, max_row=total_row - 1)
        cats = Reference(orders_ws, min_col=1, min_row=first_data_row, max_row=total_row - 1)
        chart1.add_data(data, titles_from_data=True)
        chart1.set_categories(cats)
        chart1.legend = None
        dash_ws.add_chart(chart1, "D10")

        chart2 = BarChart()
        chart2.type = "col"
        chart2.style = 11
        chart2.title = "عدد الطلبات بالساعة"
        chart2.y_axis.title = "عدد الطلبات"
        chart2.x_axis.title = "الساعة"
        chart2.height = 8
        chart2.width = 15
        data = Reference(orders_ws, min_col=2, min_row=total_row, max_col=total_col - 1, max_row=total_row)
        cats = Reference(orders_ws, min_col=2, min_row=header_row, max_col=total_col - 1, max_row=header_row)
        chart2.add_data(data, from_rows=True, titles_from_data=False)
        chart2.set_categories(cats)
        chart2.legend = None
        dash_ws.add_chart(chart2, "L10")

        chart3 = PieChart()
        chart3.title = "توزيع الحشوات"
        chart3.height = 8
        chart3.width = 15
        data = Reference(fill_ws, min_col=2, min_row=fill_total_row, max_col=fill_total_col - 1, max_row=fill_total_row)
        cats = Reference(fill_ws, min_col=2, min_row=fill_header_row, max_col=fill_total_col - 1, max_row=fill_header_row)
        chart3.add_data(data, from_rows=True, titles_from_data=False)
        chart3.set_categories(cats)
        chart3.dataLabels = DataLabelList()
        chart3.dataLabels.showPercent = True
        chart3.dataLabels.showLeaderLines = True
        dash_ws.add_chart(chart3, "D27")

        chart4 = BarChart()
        chart4.type = "col"
        chart4.style = 12
        chart4.title = "إجمالي الحشوات حسب الفرع"
        chart4.y_axis.title = "الكمية"
        chart4.x_axis.title = "الفرع"
        chart4.height = 8
        chart4.width = 15
        data = Reference(fill_ws, min_col=fill_total_col, min_row=fill_header_row, max_row=fill_total_row - 1)
        cats = Reference(fill_ws, min_col=1, min_row=fill_first_data_row, max_row=fill_total_row - 1)
        chart4.add_data(data, titles_from_data=True)
        chart4.set_categories(cats)
        chart4.legend = None
        dash_ws.add_chart(chart4, "L27")

        # Formatting tabs
        summary_ws.sheet_properties.tabColor = COLORS["purple"]
        style_grid(summary_ws, header_row=1, start_row=2, end_row=summary_ws.max_row, end_col=2, title_row=None)

        for ws, tab_color, hrow, srow, erow, ecol in [
            (orders_ws, COLORS["blue"], header_row, first_data_row, total_row, total_col),
            (fill_ws, COLORS["green"], fill_header_row, fill_first_data_row, fill_total_row, fill_total_col),
            (fill_orders_ws, COLORS["green"], fill_header_row, fill_first_data_row, fill_total_row, fill_total_col),
            (fh_ws, COLORS["green"], fh_header_row, 4, fh_row - 1, fh_total_col),
        ]:
            ws.sheet_properties.tabColor = tab_color
            style_grid(ws, header_row=hrow, start_row=srow, end_row=erow, end_col=ecol, title_row=1)
            if ecol >= 3:
                start_col_for_heat = "B" if ws != fh_ws else "C"
                apply_heatmap(ws, f"{start_col_for_heat}{srow}:{get_column_letter(ecol)}{erow}")

    output.seek(0)
    return output





# =========================
# Google Sheets + Date Filter Helpers
# =========================

def google_sheet_to_csv_url(sheet_url, gid="0"):
    """تحويل رابط Google Sheet إلى CSV export URL."""
    sheet_url = str(sheet_url).strip()
    if not sheet_url:
        raise ValueError("رابط Google Sheet فارغ.")
    if "output=csv" in sheet_url or "format=csv" in sheet_url:
        return sheet_url
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_url)
    if not match:
        raise ValueError("رابط Google Sheet غير صحيح. لازم يكون من docs.google.com/spreadsheets")
    spreadsheet_id = match.group(1)
    parsed = urlparse(sheet_url)
    parsed_qs = parse_qs(parsed.query)
    if "gid" in parsed_qs and parsed_qs["gid"]:
        gid = parsed_qs["gid"][0]
    elif "#gid=" in sheet_url:
        gid = sheet_url.split("#gid=")[-1].split("&")[0].strip() or gid
    return f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv&gid={gid}"


def read_google_sheet(sheet_url, gid="0"):
    csv_url = google_sheet_to_csv_url(sheet_url, gid)
    df = pd.read_csv(csv_url, dtype=str, keep_default_na=False, engine="python")
    return df.fillna("")


def find_date_column(df):
    possible = [
        "تاريخ الاستلام",
        "تاريخ التوصيل (Delivery Date)",
        "Delivery Date",
        "delivery_date",
        "Pickup Date",
        "pickup_date",
        "تاريخ الطلب",
        "Order Date",
        "date",
    ]
    return find_col(df, possible)


def parse_date_series(series):
    def parse_one(value):
        value = normalize_arabic_digits(value)
        if not value:
            return pd.NaT
        for dayfirst in [True, False]:
            parsed = pd.to_datetime(value, errors="coerce", dayfirst=dayfirst)
            if not pd.isna(parsed):
                return parsed.normalize()
        return pd.NaT
    return series.apply(parse_one)

# =========================
# Streamlit App - V4 Premium Dashboard
# =========================

st.set_page_config(
    page_title="MAD Orders Dashboard V6",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------
# Premium CSS
# -------------------------

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Cairo', sans-serif !important;
    }

    .stApp {
        background:
            radial-gradient(circle at top left, rgba(22,163,74,0.22), transparent 28%),
            radial-gradient(circle at top right, rgba(37,99,235,0.22), transparent 30%),
            linear-gradient(135deg, #0f172a 0%, #111827 48%, #020617 100%);
        color: #f8fafc;
    }

    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #020617 0%, #0f172a 100%);
        border-left: 1px solid rgba(255,255,255,0.08);
    }

    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 2rem;
        max-width: 1500px;
    }

    .hero {
        padding: 28px 32px;
        border-radius: 28px;
        background:
            linear-gradient(135deg, rgba(15,23,42,0.92), rgba(30,41,59,0.88)),
            linear-gradient(45deg, rgba(22,163,74,0.28), rgba(37,99,235,0.22));
        border: 1px solid rgba(255,255,255,0.12);
        box-shadow: 0 24px 70px rgba(0,0,0,0.35);
        margin-bottom: 20px;
        direction: rtl;
        text-align: right;
    }

    .hero h1 {
        margin: 0;
        font-size: 36px;
        font-weight: 800;
        color: #ffffff;
        letter-spacing: -0.5px;
    }

    .hero p {
        margin: 10px 0 0 0;
        color: #cbd5e1;
        font-size: 15px;
        line-height: 1.8;
    }

    .kpi-card {
        padding: 22px 20px;
        border-radius: 22px;
        background: linear-gradient(180deg, rgba(255,255,255,0.11), rgba(255,255,255,0.06));
        border: 1px solid rgba(255,255,255,0.12);
        box-shadow: 0 18px 45px rgba(0,0,0,0.25);
        min-height: 138px;
        direction: rtl;
        text-align: right;
        position: relative;
        overflow: hidden;
    }

    .kpi-card:before {
        content: "";
        position: absolute;
        width: 110px;
        height: 110px;
        border-radius: 999px;
        right: -36px;
        top: -42px;
        background: var(--accent);
        opacity: 0.22;
        filter: blur(2px);
    }

    .kpi-label {
        color: #cbd5e1;
        font-size: 14px;
        font-weight: 600;
        margin-bottom: 12px;
    }

    .kpi-value {
        color: #ffffff;
        font-size: 32px;
        font-weight: 800;
        line-height: 1.1;
        word-break: break-word;
    }

    .kpi-sub {
        color: #94a3b8;
        font-size: 12px;
        margin-top: 10px;
    }

    .section-title {
        direction: rtl;
        text-align: right;
        color: #ffffff;
        font-size: 22px;
        font-weight: 800;
        margin: 20px 0 10px 0;
    }

    div[data-testid="stDataFrame"] {
        direction: rtl;
        border-radius: 16px;
        overflow: hidden;
        border: 1px solid rgba(255,255,255,0.08);
    }

    .glass-box {
        background: rgba(255,255,255,0.08);
        border: 1px solid rgba(255,255,255,0.12);
        border-radius: 22px;
        padding: 18px;
        box-shadow: 0 18px 45px rgba(0,0,0,0.18);
    }

    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        direction: rtl;
    }

    .stTabs [data-baseweb="tab"] {
        border-radius: 999px;
        padding: 10px 18px;
        background-color: rgba(255,255,255,0.08);
        color: #e2e8f0;
        border: 1px solid rgba(255,255,255,0.08);
    }

    .stTabs [aria-selected="true"] {
        background: linear-gradient(90deg, #16a34a, #2563eb) !important;
        color: white !important;
        border: 0 !important;
    }

    .small-note {
        direction: rtl;
        text-align: right;
        color: #94a3b8;
        font-size: 13px;
        line-height: 1.8;
    }

    .stDownloadButton button {
        background: linear-gradient(90deg, #16a34a, #2563eb);
        color: white;
        border: 0;
        border-radius: 16px;
        padding: 0.75rem 1rem;
        font-weight: 800;
    }

    .stButton button {
        border-radius: 14px;
        font-weight: 700;
    }
    
    /* =========================
       Mobile Responsive Fixes
       ========================= */

    @media (max-width: 768px) {
        .block-container {
            padding-left: 0.75rem !important;
            padding-right: 0.75rem !important;
            padding-top: 0.75rem !important;
            max-width: 100% !important;
        }

        .hero {
            padding: 18px 16px !important;
            border-radius: 18px !important;
            margin-bottom: 12px !important;
        }

        .hero h1 {
            font-size: 22px !important;
            line-height: 1.35 !important;
        }

        .hero p {
            font-size: 12.5px !important;
            line-height: 1.7 !important;
        }

        .kpi-card {
            min-height: auto !important;
            padding: 16px 14px !important;
            border-radius: 16px !important;
            margin-bottom: 10px !important;
        }

        .kpi-label {
            font-size: 12px !important;
            margin-bottom: 8px !important;
        }

        .kpi-value {
            font-size: 22px !important;
            line-height: 1.25 !important;
        }

        .kpi-sub {
            font-size: 11px !important;
            margin-top: 6px !important;
        }

        .section-title {
            font-size: 18px !important;
            margin: 14px 0 8px 0 !important;
        }

        .glass-box {
            padding: 14px !important;
            border-radius: 16px !important;
        }

        .small-note {
            font-size: 12px !important;
        }

        .source-badge,
        .date-badge {
            display: block !important;
            width: 100% !important;
            box-sizing: border-box !important;
            margin: 6px 0 !important;
            text-align: center !important;
            font-size: 12px !important;
        }

        div[data-testid="column"] {
            width: 100% !important;
            flex: 1 1 100% !important;
            min-width: 100% !important;
        }

        div[data-testid="stHorizontalBlock"] {
            flex-wrap: wrap !important;
            gap: 0.5rem !important;
        }

        .stTabs [data-baseweb="tab-list"] {
            overflow-x: auto !important;
            white-space: nowrap !important;
            flex-wrap: nowrap !important;
            padding-bottom: 6px !important;
        }

        .stTabs [data-baseweb="tab"] {
            padding: 8px 12px !important;
            font-size: 12px !important;
            min-width: max-content !important;
        }

        div[data-testid="stDataFrame"] {
            width: 100% !important;
            overflow-x: auto !important;
            border-radius: 12px !important;
        }

        iframe {
            max-width: 100% !important;
        }

        .js-plotly-plot,
        .plotly,
        .plot-container,
        .svg-container {
            width: 100% !important;
            max-width: 100% !important;
        }

        .stTextInput input,
        .stSelectbox,
        .stDateInput,
        .stRadio,
        .stFileUploader {
            font-size: 13px !important;
        }

        .stDownloadButton button,
        .stButton button {
            width: 100% !important;
            min-height: 44px !important;
            font-size: 13px !important;
        }

        section[data-testid="stSidebar"] {
            min-width: 270px !important;
        }
    }

    @media (max-width: 420px) {
        .hero h1 {
            font-size: 20px !important;
        }

        .hero p {
            font-size: 12px !important;
        }

        .kpi-value {
            font-size: 20px !important;
        }

        .section-title {
            font-size: 16px !important;
        }
    }

    </style>
    """,
    unsafe_allow_html=True
)


def format_int(value):
    try:
        return f"{int(float(value)):,}"
    except Exception:
        return str(value)


def render_kpi(label, value, sub="", color="#16a34a"):
    st.markdown(
        f"""
        <div class="kpi-card" style="--accent:{color};">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-sub">{sub}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def fig_layout(fig, height=390):
    fig.update_layout(
        height=height,
        template="plotly_dark",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(15,23,42,0.25)",
        font=dict(family="Cairo, Arial", color="#e5e7eb", size=13),
        margin=dict(l=20, r=20, t=65, b=30),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        ),
    )
    fig.update_xaxes(gridcolor="rgba(255,255,255,0.08)", zerolinecolor="rgba(255,255,255,0.08)")
    fig.update_yaxes(gridcolor="rgba(255,255,255,0.08)", zerolinecolor="rgba(255,255,255,0.08)")
    return fig



def mobile_chart_config():
    return {
        "responsive": True,
        "displayModeBar": False,
        "scrollZoom": False,
    }


def display_pretty_dataframe(df, height=420):
    try:
        numeric_cols = df.select_dtypes(include="number").columns
        styled = df.style.background_gradient(cmap="YlGnBu", subset=numeric_cols)
        st.dataframe(styled, use_container_width=True, height=height)
    except Exception:
        st.dataframe(df, use_container_width=True, height=height)


# =========================
# Advanced Reports Helpers
# =========================

def clean_money(value):
    """تحويل المبالغ مثل 120 SAR أو 1,200.50 إلى رقم."""
    text = normalize_arabic_digits(value)
    if not text or str(text).strip().lower() in ["nan", "none", "(not set)", "not set", "-"]:
        return 0.0
    text = str(text).replace(",", "")
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return 0.0
    try:
        return float(match.group(0))
    except Exception:
        return 0.0


def normalize_product_name(value):
    """تنظيف اسم المنتج مع الإبقاء على الاسم الأساسي."""
    text = str(value).strip()
    if not text or text.lower() in ["nan", "none"]:
        return "بدون اسم منتج"
    return re.sub(r"\s+", " ", text)


def is_addon_product(product_name):
    """تمييز الإضافات مثل الشموع والبالون والكروت."""
    text = str(product_name).lower()
    addon_keywords = [
        "candle", "candles", "شموع", "شمعة",
        "balloon", "helium", "بالون", "هيليوم",
        "gift card", "card", "كرت", "بطاقة",
        "night stars", "نجوم", "stars",
        "topper", "توبير",
    ]
    return any(k.lower() in text for k in addon_keywords)


def classify_campaign(product_name):
    """تصنيف المنتج ضمن حملة/مناسبة من الاسم."""
    text = str(product_name).lower()
    if any(k in text for k in ["father", "dad", "عيد الأب", "عيد الاب", "بابا"]):
        return "Father's Day"
    if any(k in text for k in ["birthday", "بيرثداي", "ميلاد"]):
        return "Birthday"
    if any(k in text for k in ["graduation", "تخرج", "التخرج"]):
        return "Graduation"
    if any(k in text for k in ["eid", "عيد الفطر", "عيد"]):
        return "Eid"
    if any(k in text for k in ["new year", "نيو يير", "2026", "2025"]):
        return "New Year"
    if any(k in text for k in ["valentine", "love", "الحب"]):
        return "Valentine"
    if any(k in text for k in ["gender reveal", "تحديد الجنس"]):
        return "Gender Reveal"
    return "General"


def extract_phone_from_text(value):
    """استخراج أول رقم جوال سعودي محتمل من الملاحظة."""
    text = normalize_arabic_digits(value)
    if not text:
        return ""
    # يسمح بأرقام مثل 05xxxxxxxx أو 9665xxxxxxxx أو +9665xxxxxxxx
    patterns = [
        r"(?:\+?966\s?)?5\d{8}",
        r"0?5\d{8}",
    ]
    for pat in patterns:
        m = re.search(pat, text.replace(" ", "").replace("-", ""))
        if m:
            return m.group(0)
    return ""


def need_action_reason(note, product_name=""):
    """تحديد هل الطلب يحتاج متابعة بسبب صورة/تواصل/واتساب/ملاحظة خاصة."""
    text = f"{note} {product_name}".strip()
    low = text.lower()
    reasons = []
    if any(k in low for k in ["photo", "picture", "image", "صورة", "الصوره", "الصورة"]):
        reasons.append("صورة")
    if any(k in low for k in ["contact", "call", "whatsapp", "text me", "تواصل", "اتصال", "واتساب", "جوال"]):
        reasons.append("تواصل")
    if extract_phone_from_text(text):
        reasons.append("رقم جوال")
    if any(k in low for k in ["write", "writing", "اكتب", "كتابة", "الكتابة", "عبارة"]):
        reasons.append("كتابة خاصة")
    if any(k in low for k in ["draw", "ارسم", "تصميم", "نفس", "match the design"]):
        reasons.append("تعديل/تصميم خاص")
    return " + ".join(dict.fromkeys(reasons))


def first_existing_col(df, names):
    col = find_col(df, names)
    return col if col in df.columns else None


def build_advanced_reports(clean, active, active_items, unique_orders, cols):
    """بناء كل التقارير الإضافية المستخدمة في الواجهة والتصدير."""
    order_id_col = cols.get("order_id_col")
    order_no_col = cols.get("order_no_col")
    status_col = cols.get("status_col")
    pickup_time_col = cols.get("pickup_time_col")
    delivery_date_col = cols.get("delivery_date_col")
    item_name_col = cols.get("item_name_col")
    variety_col = cols.get("variety_col")
    note_col = cols.get("note_col")

    # Order-level table for accurate sales/AOV without duplicating multi-item orders.
    order_level = (
        active.sort_values(by=["تاريخ ووقت الاستلام"])
        .drop_duplicates(subset=[order_id_col])
        .copy()
    ) if order_id_col else active.copy()

    if "قيمة الطلب رقمية" not in order_level.columns:
        order_level["قيمة الطلب رقمية"] = 0.0

    # Branch sales summary
    branch_sales = order_level.groupby("الفرع", dropna=False).agg(
        عدد_الطلبات=(order_id_col, "nunique") if order_id_col else ("الفرع", "size"),
        إجمالي_المبيعات=("قيمة الطلب رقمية", "sum"),
        متوسط_الطلب=("قيمة الطلب رقمية", "mean"),
    ).reset_index()

    if len(active_items):
        qty_by_branch = active_items.groupby("الفرع")["كمية رقمية"].sum().reset_index(name="إجمالي_الكمية")
        branch_sales = branch_sales.merge(qty_by_branch, on="الفرع", how="left")
    else:
        branch_sales["إجمالي_الكمية"] = 0

    addon_items = active_items[active_items.get("إضافة؟", False) == True].copy() if "إضافة؟" in active_items.columns else pd.DataFrame()
    if len(addon_items):
        addon_branch = addon_items.groupby("الفرع").agg(
            عدد_الإضافات=("كمية رقمية", "sum"),
            مبيعات_الإضافات=("إجمالي المنتج رقمي", "sum"),
        ).reset_index()
        branch_sales = branch_sales.merge(addon_branch, on="الفرع", how="left")
    else:
        branch_sales["عدد_الإضافات"] = 0
        branch_sales["مبيعات_الإضافات"] = 0

    for c in ["إجمالي_الكمية", "عدد_الإضافات", "مبيعات_الإضافات"]:
        if c in branch_sales.columns:
            branch_sales[c] = branch_sales[c].fillna(0)
    branch_sales["متوسط_الطلب"] = branch_sales["متوسط_الطلب"].fillna(0)
    branch_sales = branch_sales.sort_values("إجمالي_المبيعات", ascending=False)

    # Status report
    if status_col and status_col in active.columns:
        status_report = pd.pivot_table(
            active.drop_duplicates(subset=[order_id_col]) if order_id_col else active,
            index="الفرع",
            columns=status_col,
            values=order_id_col if order_id_col else "الفرع",
            aggfunc="nunique" if order_id_col else "size",
            fill_value=0,
        )
        status_report["الإجمالي"] = status_report.sum(axis=1)
        status_report.loc["الإجمالي"] = status_report.sum(axis=0)
        status_report = make_numeric_table(status_report)
    else:
        status_report = pd.DataFrame()

    # Product performance
    if item_name_col and item_name_col in active_items.columns:
        product_perf = active_items.groupby("اسم المنتج النظيف", dropna=False).agg(
            عدد_الطلبات=(order_id_col, "nunique") if order_id_col else ("اسم المنتج النظيف", "size"),
            الكمية=("كمية رقمية", "sum"),
            إجمالي_المبيعات=("إجمالي المنتج رقمي", "sum"),
            متوسط_سعر_المنتج=("سعر الحبة رقمي", "mean"),
        ).reset_index().sort_values(["إجمالي_المبيعات", "الكمية"], ascending=False)
    else:
        product_perf = pd.DataFrame()

    # Products by branch
    if item_name_col and item_name_col in active_items.columns:
        products_by_branch = active_items.groupby(["الفرع", "اسم المنتج النظيف"], dropna=False).agg(
            عدد_الطلبات=(order_id_col, "nunique") if order_id_col else ("اسم المنتج النظيف", "size"),
            الكمية=("كمية رقمية", "sum"),
            إجمالي_المبيعات=("إجمالي المنتج رقمي", "sum"),
        ).reset_index().sort_values(["الفرع", "إجمالي_المبيعات"], ascending=[True, False])
    else:
        products_by_branch = pd.DataFrame()

    # Add-ons report
    if len(addon_items) and item_name_col and item_name_col in addon_items.columns:
        addons_summary = addon_items.groupby("اسم المنتج النظيف", dropna=False).agg(
            عدد_الطلبات=(order_id_col, "nunique") if order_id_col else ("اسم المنتج النظيف", "size"),
            الكمية=("كمية رقمية", "sum"),
            إجمالي_المبيعات=("إجمالي المنتج رقمي", "sum"),
        ).reset_index().sort_values("الكمية", ascending=False)
        add_on_order_ids = set(addon_items[order_id_col].astype(str).unique()) if order_id_col else set()
        orders_with_addons = order_level[order_level[order_id_col].astype(str).isin(add_on_order_ids)] if order_id_col else pd.DataFrame()
        orders_without_addons = order_level[~order_level[order_id_col].astype(str).isin(add_on_order_ids)] if order_id_col else pd.DataFrame()
        addon_kpis = pd.DataFrame({
            "البند": ["طلبات بها إضافات", "طلبات بدون إضافات", "نسبة الطلبات بها إضافات", "متوسط الطلب مع إضافات", "متوسط الطلب بدون إضافات"],
            "القيمة": [
                len(orders_with_addons),
                len(orders_without_addons),
                round((len(orders_with_addons) / len(order_level) * 100), 2) if len(order_level) else 0,
                round(orders_with_addons["قيمة الطلب رقمية"].mean(), 2) if len(orders_with_addons) else 0,
                round(orders_without_addons["قيمة الطلب رقمية"].mean(), 2) if len(orders_without_addons) else 0,
            ]
        })
    else:
        addons_summary = pd.DataFrame(columns=["اسم المنتج النظيف", "عدد_الطلبات", "الكمية", "إجمالي_المبيعات"])
        addon_kpis = pd.DataFrame({"البند": ["طلبات بها إضافات"], "القيمة": [0]})

    # Need action report
    need_action_df = clean[clean.get("يحتاج متابعة؟", False) == True].copy() if "يحتاج متابعة؟" in clean.columns else pd.DataFrame()
    action_cols = []
    for col in [order_no_col, order_id_col, "الفرع", delivery_date_col, pickup_time_col, item_name_col, variety_col, "سبب المتابعة", "رقم تواصل مستخرج", note_col]:
        if col and col in need_action_df.columns and col not in action_cols:
            action_cols.append(col)
    need_action_report = need_action_df[action_cols].copy() if len(need_action_df) and action_cols else pd.DataFrame()

    # Data quality report
    quality_rows = []
    data_check = clean.copy()
    for idx, row in data_check.iterrows():
        order_ref = row.get(order_id_col, "") if order_id_col else ""
        branch = row.get("الفرع", "")
        product = row.get(item_name_col, "") if item_name_col else ""
        def add_issue(issue, severity="متوسط"):
            quality_rows.append({
                "Order Id": order_ref,
                "الفرع": branch,
                "المنتج": product,
                "المشكلة": issue,
                "الأهمية": severity,
            })
        if order_id_col and str(row.get(order_id_col, "")).strip() == "":
            add_issue("معرف الطلب مفقود", "عالي")
        if branch == "العقيق":
            add_issue("الفرع غير واضح", "عالي")
        if pd.isna(row.get("تاريخ ووقت الاستلام")):
            add_issue("تاريخ/وقت الاستلام غير مقروء", "عالي")
        if delivery_date_col and str(row.get(delivery_date_col, "")).strip() == "":
            add_issue("تاريخ التوصيل/الاستلام فارغ", "متوسط")
        if pickup_time_col and str(row.get(pickup_time_col, "")).strip() == "":
            add_issue("وقت الاستلام فارغ", "متوسط")
        if item_name_col and str(row.get(item_name_col, "")).strip() == "":
            add_issue("اسم المنتج فارغ", "متوسط")
        if variety_col and (not bool(row.get("إضافة؟", False))) and str(row.get("الحشوة الموحدة", "")).strip() == "":
            add_issue("حشوة غير واضحة لمنتج أساسي", "منخفض")
        if row.get("كمية رقمية", 0) <= 0:
            add_issue("الكمية صفر أو غير صحيحة", "عالي")
        if not bool(row.get("إضافة؟", False)) and row.get("إجمالي المنتج رقمي", 0) <= 0:
            add_issue("إجمالي المنتج صفر أو غير مقروء", "متوسط")
    data_quality_report = pd.DataFrame(quality_rows)
    if len(data_quality_report):
        quality_summary = data_quality_report.groupby(["المشكلة", "الأهمية"]).size().reset_index(name="عدد الحالات").sort_values("عدد الحالات", ascending=False)
    else:
        quality_summary = pd.DataFrame({"المشكلة": ["لا توجد مشاكل واضحة"], "الأهمية": ["-"], "عدد الحالات": [0]})

    # Multi-item orders
    if order_id_col and len(active_items):
        tmp = active_items.copy()
        multi_item_orders = tmp.groupby(order_id_col).agg(
            عدد_الأصناف=("معرف العنصر (Item Id)" if "معرف العنصر (Item Id)" in tmp.columns else "اسم المنتج النظيف", "count"),
            إجمالي_الكمية=("كمية رقمية", "sum"),
            مبيعات_الأصناف=("إجمالي المنتج رقمي", "sum"),
            عدد_الإضافات=("إضافة؟", "sum") if "إضافة؟" in tmp.columns else ("كمية رقمية", "sum"),
            الفرع=("الفرع", "first"),
            الساعة=("الساعة", "first"),
        ).reset_index()
        multi_item_orders = multi_item_orders[multi_item_orders["عدد_الأصناف"] > 1].sort_values("عدد_الأصناف", ascending=False)
    else:
        multi_item_orders = pd.DataFrame()

    # Campaign report
    campaign_items = active_items.copy()
    if len(campaign_items):
        campaign_summary = campaign_items.groupby("الحملة", dropna=False).agg(
            عدد_الطلبات=(order_id_col, "nunique") if order_id_col else ("الحملة", "size"),
            الكمية=("كمية رقمية", "sum"),
            إجمالي_المبيعات=("إجمالي المنتج رقمي", "sum"),
        ).reset_index().sort_values("إجمالي_المبيعات", ascending=False)
        campaign_products = campaign_items.groupby(["الحملة", "اسم المنتج النظيف"], dropna=False).agg(
            عدد_الطلبات=(order_id_col, "nunique") if order_id_col else ("الحملة", "size"),
            الكمية=("كمية رقمية", "sum"),
            إجمالي_المبيعات=("إجمالي المنتج رقمي", "sum"),
        ).reset_index().sort_values(["الحملة", "إجمالي_المبيعات"], ascending=[True, False])
    else:
        campaign_summary = pd.DataFrame()
        campaign_products = pd.DataFrame()

    # Hour x sales
    if len(order_level):
        hour_sales = order_level.groupby("الساعة", dropna=False).agg(
            عدد_الطلبات=(order_id_col, "nunique") if order_id_col else ("الساعة", "size"),
            إجمالي_المبيعات=("قيمة الطلب رقمية", "sum"),
            متوسط_الطلب=("قيمة الطلب رقمية", "mean"),
        ).reset_index()
        hour_order_map = order_level[["الساعة", "رقم الساعة"]].drop_duplicates().dropna().set_index("الساعة")["رقم الساعة"].to_dict()
        hour_sales["ترتيب"] = hour_sales["الساعة"].map(hour_order_map).fillna(99)
        hour_sales = hour_sales.sort_values("ترتيب").drop(columns="ترتيب")
    else:
        hour_sales = pd.DataFrame()

    return {
        "order_level": order_level,
        "branch_sales": make_numeric_table(branch_sales),
        "hour_sales": make_numeric_table(hour_sales),
        "status_report": status_report,
        "product_perf": make_numeric_table(product_perf),
        "products_by_branch": make_numeric_table(products_by_branch),
        "addons_summary": make_numeric_table(addons_summary),
        "addon_kpis": addon_kpis,
        "need_action_report": need_action_report,
        "quality_summary": make_numeric_table(quality_summary),
        "data_quality_report": data_quality_report,
        "multi_item_orders": make_numeric_table(multi_item_orders),
        "campaign_summary": make_numeric_table(campaign_summary),
        "campaign_products": make_numeric_table(campaign_products),
    }


def create_all_reports_excel(reports, clean_df):
    """تصدير ملف Excel شامل كل التقارير الجديدة."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        sheet_map = {
            "Executive Summary": reports.get("branch_sales", pd.DataFrame()),
            "Sales by Branch": reports.get("branch_sales", pd.DataFrame()),
            "Sales by Hour": reports.get("hour_sales", pd.DataFrame()),
            "Status Report": reports.get("status_report", pd.DataFrame()),
            "Product Performance": reports.get("product_perf", pd.DataFrame()),
            "Products by Branch": reports.get("products_by_branch", pd.DataFrame()),
            "Addons Summary": reports.get("addons_summary", pd.DataFrame()),
            "Addons KPIs": reports.get("addon_kpis", pd.DataFrame()),
            "Need Action": reports.get("need_action_report", pd.DataFrame()),
            "Quality Summary": reports.get("quality_summary", pd.DataFrame()),
            "Data Quality Details": reports.get("data_quality_report", pd.DataFrame()),
            "Multi Item Orders": reports.get("multi_item_orders", pd.DataFrame()),
            "Campaign Summary": reports.get("campaign_summary", pd.DataFrame()),
            "Campaign Products": reports.get("campaign_products", pd.DataFrame()),
            "Clean Data": clean_df,
        }
        for sheet_name, data in sheet_map.items():
            if data is None or len(data) == 0:
                data = pd.DataFrame({"ملاحظة": ["لا توجد بيانات لهذا التقرير في الفترة المختارة"]})
            safe_name = str(sheet_name)[:31]
            data.to_excel(writer, sheet_name=safe_name, index=True if isinstance(data.index, pd.MultiIndex) else False)

        wb = writer.book
        header_fill = PatternFill("solid", fgColor="0F172A")
        header_font = Font(color="FFFFFF", bold=True)
        border = Border(
            left=Side(style="thin", color="D1D5DB"),
            right=Side(style="thin", color="D1D5DB"),
            top=Side(style="thin", color="D1D5DB"),
            bottom=Side(style="thin", color="D1D5DB"),
        )
        for ws in wb.worksheets:
            ws.sheet_view.rightToLeft = True
            ws.sheet_view.showGridLines = False
            ws.freeze_panes = "A2"
            max_row = ws.max_row
            max_col = ws.max_column
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = border
            for row in ws.iter_rows(min_row=2, max_row=max_row, max_col=max_col):
                for cell in row:
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    cell.border = border
            for c in range(1, max_col + 1):
                letter = get_column_letter(c)
                max_len = 12
                for cell in ws[letter]:
                    max_len = max(max_len, min(len(str(cell.value)) if cell.value is not None else 0, 45))
                ws.column_dimensions[letter].width = max_len + 3
            try:
                ref = f"A1:{get_column_letter(max_col)}{max_row}"
                table = Table(displayName=("tbl_" + re.sub(r"[^A-Za-z0-9]", "", ws.title))[:250], ref=ref)
                style = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False)
                table.tableStyleInfo = style
                ws.add_table(table)
            except Exception:
                pass
    output.seek(0)
    return output


DEFAULT_GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1Lf7R_G5hZ6KvyE5OyRc78b1dKVjD1bEDeeZnorANrxI/edit?usp=sharing"

st.sidebar.markdown("## ⚙️ إعدادات التحليل")
st.sidebar.caption("اربط Google Sheet أو ارفع ملف، ثم اختر تاريخ الاستلام للتحليل.")

source_type = st.sidebar.radio(
    "مصدر البيانات",
    ["رفع ملف", "Google Sheet"],
    index=1,
    horizontal=True
)

df = None
source_label = ""

if source_type == "رفع ملف":
    uploaded_file = st.sidebar.file_uploader(
        "ارفع ملف الطلبات TXT / CSV / Excel",
        type=["txt", "csv", "xlsx", "xls"]
    )
    if uploaded_file is not None:
        df = read_uploaded_file(uploaded_file)
        source_label = f"ملف مرفوع: {uploaded_file.name}"
else:
    sheet_url = st.sidebar.text_input(
        "رابط Google Sheet",
        value=DEFAULT_GOOGLE_SHEET_URL,
        placeholder="https://docs.google.com/spreadsheets/d/..."
    )
    sheet_gid = st.sidebar.text_input(
        "Sheet GID",
        value="0",
        help="اتركه 0 لو أول تاب. لو الرابط فيه gid هيتم قراءته تلقائيًا."
    )
    refresh = st.sidebar.button("🔄 تحديث من Google Sheet", use_container_width=True)
    if sheet_url:
        try:
            df = read_google_sheet(sheet_url, sheet_gid)
            source_label = "Google Sheet"
            st.sidebar.success("تم قراءة البيانات من Google Sheet")
        except Exception as e:
            st.sidebar.error("لم أستطع قراءة Google Sheet")
            st.sidebar.caption(str(e))
            st.sidebar.info("لو الشيت خاص، اجعله Anyone with the link can view.")

exclude_cancelled = st.sidebar.toggle("استبعاد الطلبات الملغاة", value=True)

st.markdown(
    """
    <div class="hero">
        <h1>📊 MAD Orders Dashboard V6</h1>
        <p>
            لوحة تحليل تشغيلية تفاعلية للطلبات، المبيعات، المنتجات، الإضافات، جودة البيانات، والحملات.
            تدعم Google Sheet + فلترة تاريخ الاستلام ليوم محدد أو فترة متعددة الأيام.
        </p>
    </div>
    """,
    unsafe_allow_html=True
)

if df is not None and len(df) > 0:
    # =========================
    # Date Filter
    # =========================

    st.sidebar.markdown("---")
    st.sidebar.markdown("## 📅 فلترة تاريخ الاستلام")

    auto_date_col = find_date_column(df)
    date_options = list(df.columns)
    default_date_index = date_options.index(auto_date_col) if auto_date_col in date_options else 0

    selected_date_col = st.sidebar.selectbox(
        "عمود تاريخ الاستلام",
        date_options,
        index=default_date_index
    )

    raw_dates = parse_date_series(df[selected_date_col])
    valid_dates = raw_dates.dropna()

    date_filter_mode = st.sidebar.radio(
        "طريقة الفلترة",
        ["كل البيانات", "يوم محدد", "فترة من / إلى"],
        index=0
    )

    date_badge = "كل البيانات"
    filtered_df = df.copy()

    if len(valid_dates) > 0 and date_filter_mode != "كل البيانات":
        min_date = valid_dates.min().date()
        max_date = valid_dates.max().date()

        if date_filter_mode == "يوم محدد":
            selected_day = st.sidebar.date_input(
                "اختر اليوم",
                value=max_date,
                min_value=min_date,
                max_value=max_date
            )
            mask = raw_dates == pd.Timestamp(selected_day)
            filtered_df = df[mask].copy()
            date_badge = f"يوم محدد: {selected_day}"
        else:
            date_range = st.sidebar.date_input(
                "اختر الفترة",
                value=(min_date, max_date),
                min_value=min_date,
                max_value=max_date
            )
            if isinstance(date_range, tuple) and len(date_range) == 2:
                start_date, end_date = date_range
                mask = (raw_dates >= pd.Timestamp(start_date)) & (raw_dates <= pd.Timestamp(end_date))
                filtered_df = df[mask].copy()
                date_badge = f"من {start_date} إلى {end_date}"
    elif date_filter_mode != "كل البيانات":
        st.sidebar.warning("لم أستطع قراءة تواريخ صحيحة من العمود المختار.")

    df = filtered_df.copy()

    st.markdown(
        f"""
        <div class="glass-box small-note">
            <b>المصدر:</b> {source_label} &nbsp; | &nbsp;
            <b>فلترة التاريخ:</b> {date_badge} &nbsp; | &nbsp;
            <b>الصفوف بعد الفلترة:</b> {format_int(len(df))}
        </div>
        """,
        unsafe_allow_html=True
    )

    if len(df) == 0:
        st.warning("لا توجد بيانات في الفترة المختارة.")
        st.stop()

    # =========================
    # Detect Columns
    # =========================

    order_no_col = find_col(df, ["رقم الطلب", "Order No", "order_no"])
    order_id_col = find_col(df, ["معرف الطلب (Order Id)", "Order Id", "order_id"])
    chef_col = find_col(df, ["اسم الشيف", "chef"])
    status_col = find_col(df, ["حالة الطلب", "status"])
    delivery_date_col = find_col(df, ["تاريخ الاستلام", "تاريخ التوصيل (Delivery Date)", "Delivery Date", "Pickup Date", "delivery_date"])
    pickup_time_col = find_col(df, ["وقت الاستلام (Pickup Time)", "Pickup Time"])
    item_id_col = find_col(df, ["معرف العنصر (Item Id)", "Item Id", "item_id"])
    dish_id_col = find_col(df, ["معرف الطبق (Dish Id)", "Dish Id", "dish_id"])
    item_name_col = find_col(df, ["اسم الطبق / المنتج", "product", "item"])
    variety_col = find_col(df, ["نوع الحشوة (Variety)", "Variety", "الحشوة"])
    qty_col = find_col(df, ["الكمية", "Quantity", "qty"])
    total_order_col = find_col(df, ["إجمالي الطلب بالكامل", "Order Total", "order_price", "order total", "total order"])
    item_total_col = find_col(df, ["إجمالي المنتج (Item Total)", "Item Total", "item_total", "line total"])
    unit_price_col = find_col(df, ["سعر الحبة", "Unit Price", "unit_price", "price"])
    discount_col = find_col(df, ["الخصم", "Discount", "discount"])
    note_col = find_col(df, ["ملاحظة للشيف", "Chef Note", "notes", "note", "ملاحظات"])

    required_cols = {
        "معرف الطلب / Order Id": order_id_col,
        "اسم الشيف / الفرع": chef_col,
        "وقت الاستلام": pickup_time_col,
    }

    missing = [name for name, col in required_cols.items() if col is None]

    if missing:
        st.error(f"في أعمدة ناقصة أو غير واضحة: {', '.join(missing)}")
        st.stop()

    if variety_col is None:
        st.warning("عمود نوع الحشوة غير موجود، لذلك تحليل الحشوات لن يظهر بشكل كامل.")

    # =========================
    # Clean Data
    # =========================

    clean = df.copy()

    clean["الفرع"] = clean[chef_col].apply(extract_branch)

    clean["تاريخ ووقت الاستلام"] = clean.apply(
        lambda row: parse_datetime(
            row[delivery_date_col] if delivery_date_col else "",
            row[pickup_time_col] if pickup_time_col else ""
        ),
        axis=1
    )

    clean["رقم الساعة"] = clean["تاريخ ووقت الاستلام"].dt.hour

    clean["الساعة"] = clean["رقم الساعة"].apply(
        lambda x: hour_label(int(x)) if pd.notna(x) else "وقت غير واضح"
    )

    if status_col:
        clean["ملغي؟"] = clean[status_col].apply(is_cancelled)
    else:
        clean["ملغي؟"] = False

    if qty_col:
        clean["كمية رقمية"] = clean[qty_col].apply(clean_quantity)
    else:
        clean["كمية رقمية"] = 1

    if variety_col:
        clean["الحشوة الموحدة"] = clean[variety_col].apply(normalize_variety)
    else:
        clean["الحشوة الموحدة"] = ""

    # =========================
    # Advanced Clean Columns
    # =========================

    if item_name_col:
        clean["اسم المنتج النظيف"] = clean[item_name_col].apply(normalize_product_name)
    else:
        clean["اسم المنتج النظيف"] = "بدون اسم منتج"

    clean["إضافة؟"] = clean["اسم المنتج النظيف"].apply(is_addon_product)
    clean["الحملة"] = clean["اسم المنتج النظيف"].apply(classify_campaign)

    if total_order_col:
        clean["قيمة الطلب رقمية"] = clean[total_order_col].apply(clean_money)
    else:
        clean["قيمة الطلب رقمية"] = 0.0

    if unit_price_col:
        clean["سعر الحبة رقمي"] = clean[unit_price_col].apply(clean_money)
    else:
        clean["سعر الحبة رقمي"] = 0.0

    if item_total_col:
        clean["إجمالي المنتج رقمي"] = clean[item_total_col].apply(clean_money)
    else:
        clean["إجمالي المنتج رقمي"] = clean["سعر الحبة رقمي"] * clean["كمية رقمية"]

    if discount_col:
        clean["خصم رقمي"] = clean[discount_col].apply(clean_money)
    else:
        clean["خصم رقمي"] = 0.0

    if note_col:
        clean["سبب المتابعة"] = clean.apply(lambda r: need_action_reason(r.get(note_col, ""), r.get("اسم المنتج النظيف", "")), axis=1)
        clean["رقم تواصل مستخرج"] = clean.apply(lambda r: extract_phone_from_text(str(r.get(note_col, "")) + " " + str(r.get("اسم المنتج النظيف", ""))), axis=1)
    else:
        clean["سبب المتابعة"] = clean["اسم المنتج النظيف"].apply(lambda x: need_action_reason("", x))
        clean["رقم تواصل مستخرج"] = ""

    clean["يحتاج متابعة؟"] = clean["سبب المتابعة"].astype(str).str.strip() != ""

    # استبعاد الملغي لو مختار
    active = clean.copy()

    if exclude_cancelled:
        active = active[active["ملغي؟"] == False].copy()

    active_items = active.copy()

    if item_id_col:
        active_items = active_items.drop_duplicates(subset=[item_id_col]).copy()

    # =========================
    # Unique Orders for Hourly Analysis
    # =========================

    unique_orders = (
        active
        .sort_values(by=["تاريخ ووقت الاستلام"])
        .drop_duplicates(subset=[order_id_col])
        .copy()
    )

    # =========================
    # Orders by Hour per Branch
    # =========================

    hourly = pd.pivot_table(
        unique_orders,
        index="الفرع",
        columns="الساعة",
        values=order_id_col,
        aggfunc="nunique",
        fill_value=0
    )

    hour_order = (
        unique_orders[["الساعة", "رقم الساعة"]]
        .dropna()
        .drop_duplicates()
        .sort_values("رقم الساعة")
    )

    sorted_hour_cols = [
        h for h in hour_order["الساعة"].tolist()
        if h in hourly.columns
    ]

    hourly = hourly.reindex(columns=sorted_hour_cols, fill_value=0)

    hourly["الإجمالي"] = hourly.sum(axis=1)
    hourly.loc["الإجمالي"] = hourly.sum(axis=0)
    hourly = make_numeric_table(hourly)

    # =========================
    # Fillings Rows
    # =========================

    filling_rows = active_items[
        active_items["الحشوة الموحدة"].astype(str).str.strip() != ""
    ].copy()

    # =========================
    # Fillings by Branch - Qty
    # =========================

    fillings_qty = pd.pivot_table(
        filling_rows,
        index="الفرع",
        columns="الحشوة الموحدة",
        values="كمية رقمية",
        aggfunc="sum",
        fill_value=0
    )

    fillings_qty["الإجمالي"] = fillings_qty.sum(axis=1)
    fillings_qty.loc["الإجمالي"] = fillings_qty.sum(axis=0)
    fillings_qty = make_numeric_table(fillings_qty)

    # =========================
    # Fillings by Branch - Unique Orders
    # =========================

    fillings_orders_source = filling_rows.drop_duplicates(
        subset=[order_id_col, "الحشوة الموحدة"]
    )

    fillings_orders = pd.pivot_table(
        fillings_orders_source,
        index="الفرع",
        columns="الحشوة الموحدة",
        values=order_id_col,
        aggfunc="nunique",
        fill_value=0
    )

    fillings_orders["الإجمالي"] = fillings_orders.sum(axis=1)
    fillings_orders.loc["الإجمالي"] = fillings_orders.sum(axis=0)
    fillings_orders = make_numeric_table(fillings_orders)

    # =========================
    # Fillings by Hour per Branch - Qty
    # =========================

    fillings_by_hour_branch = pd.pivot_table(
        filling_rows,
        index=["الفرع", "الساعة"],
        columns="الحشوة الموحدة",
        values="كمية رقمية",
        aggfunc="sum",
        fill_value=0
    )

    hour_sort_map = (
        active[["الساعة", "رقم الساعة"]]
        .dropna()
        .drop_duplicates()
        .set_index("الساعة")["رقم الساعة"]
        .to_dict()
    )

    fillings_by_hour_branch = fillings_by_hour_branch.reset_index()
    fillings_by_hour_branch["ترتيب الساعة"] = fillings_by_hour_branch["الساعة"].map(hour_sort_map)

    fillings_by_hour_branch = (
        fillings_by_hour_branch
        .sort_values(by=["الفرع", "ترتيب الساعة"])
        .drop(columns=["ترتيب الساعة"])
        .set_index(["الفرع", "الساعة"])
    )

    fillings_by_hour_branch["الإجمالي"] = fillings_by_hour_branch.sum(axis=1)
    fillings_by_hour_branch = make_numeric_table(fillings_by_hour_branch)

    # =========================
    # Separate table per branch
    # =========================

    branch_hour_details_dict = {}
    branches = sorted(filling_rows["الفرع"].dropna().unique().tolist())

    for branch in branches:
        branch_data = filling_rows[filling_rows["الفرع"] == branch].copy()

        branch_table = pd.pivot_table(
            branch_data,
            index="الساعة",
            columns="الحشوة الموحدة",
            values="كمية رقمية",
            aggfunc="sum",
            fill_value=0
        )

        branch_hour_order = (
            branch_data[["الساعة", "رقم الساعة"]]
            .dropna()
            .drop_duplicates()
            .sort_values("رقم الساعة")
        )

        branch_table = branch_table.reindex(
            branch_hour_order["الساعة"].tolist(),
            fill_value=0
        )

        branch_table["الإجمالي"] = branch_table.sum(axis=1)
        branch_table = make_numeric_table(branch_table)

        branch_hour_details_dict[branch] = branch_table

    # =========================
    # Summary
    # =========================

    total_rows = len(clean)
    total_unique_orders = clean[order_id_col].nunique()
    active_unique_orders = active[order_id_col].nunique()
    cancelled_orders = clean[clean["ملغي؟"] == True][order_id_col].nunique()
    duplicate_rows_estimate = total_rows - total_unique_orders
    total_fillings_qty = int(filling_rows["كمية رقمية"].sum()) if len(filling_rows) else 0

    if len(unique_orders) > 0:
        top_hour = (
            unique_orders.groupby("الساعة")[order_id_col]
            .nunique()
            .sort_values(ascending=False)
            .head(1)
        )

        top_branch = (
            unique_orders.groupby("الفرع")[order_id_col]
            .nunique()
            .sort_values(ascending=False)
            .head(1)
        )
    else:
        top_hour = pd.Series(dtype="float64")
        top_branch = pd.Series(dtype="float64")

    if len(filling_rows) > 0:
        top_filling = (
            filling_rows.groupby("الحشوة الموحدة")["كمية رقمية"]
            .sum()
            .sort_values(ascending=False)
            .head(1)
        )
    else:
        top_filling = pd.Series(dtype="float64")

    summary_df = pd.DataFrame({
        "البند": [
            "عدد سطور الملف",
            "عدد الطلبات المختلفة",
            "عدد الطلبات بعد الاستبعاد",
            "عدد الطلبات الملغاة",
            "عدد السطور المكررة تقريباً",
            "أعلى ساعة ضغط",
            "طلبات أعلى ساعة",
            "أعلى فرع",
            "طلبات أعلى فرع",
            "أكثر حشوة",
            "كمية أكثر حشوة",
        ],
        "القيمة": [
            total_rows,
            total_unique_orders,
            active_unique_orders,
            cancelled_orders,
            duplicate_rows_estimate,
            top_hour.index[0] if len(top_hour) else "",
            int(top_hour.iloc[0]) if len(top_hour) else 0,
            top_branch.index[0] if len(top_branch) else "",
            int(top_branch.iloc[0]) if len(top_branch) else 0,
            top_filling.index[0] if len(top_filling) else "",
            int(top_filling.iloc[0]) if len(top_filling) else 0,
        ]
    })

    # =========================
    # Premium Display
    # =========================

    st.sidebar.success("تم تحميل الملف وتحليل البيانات بنجاح")
    st.sidebar.markdown("### 📌 ملخص")
    st.sidebar.write(f"الصفوف بعد فلترة التاريخ: **{format_int(total_rows)}**")
    st.sidebar.write(f"الطلبات الفعلية: **{format_int(active_unique_orders)}**")
    st.sidebar.write(f"الملغاة: **{format_int(cancelled_orders)}**")
    st.sidebar.write(f"المبيعات: **{total_sales:,.0f} SAR**")
    st.sidebar.write(f"تحتاج متابعة: **{format_int(need_action_count)}**")

    k1, k2, k3, k4, k5 = st.columns(5)

    with k1:
        render_kpi("عدد الطلبات", format_int(active_unique_orders), "Unique Order Id", "#2563eb")
    with k2:
        render_kpi("إجمالي الحشوات", format_int(total_fillings_qty), "حسب الكمية", "#16a34a")
    with k3:
        render_kpi("أعلى ساعة ضغط", top_hour.index[0] if len(top_hour) else "-", f"{format_int(top_hour.iloc[0]) if len(top_hour) else 0} طلب", "#ea580c")
    with k4:
        render_kpi("أعلى فرع", top_branch.index[0] if len(top_branch) else "-", f"{format_int(top_branch.iloc[0]) if len(top_branch) else 0} طلب", "#7c3aed")
    with k5:
        render_kpi("أكثر حشوة", top_filling.index[0] if len(top_filling) else "-", f"{format_int(top_filling.iloc[0]) if len(top_filling) else 0} قطعة", "#db2777")

    r1, r2, r3 = st.columns(3)
    with r1:
        render_kpi("إجمالي المبيعات", f"{total_sales:,.0f} SAR", "حسب إجمالي الطلب بدون تكرار", "#0891b2")
    with r2:
        render_kpi("متوسط الطلب", f"{avg_order_value:,.0f} SAR", "Average Order Value", "#059669")
    with r3:
        render_kpi("طلبات تحتاج متابعة", format_int(need_action_count), "صورة / تواصل / كتابة خاصة", "#dc2626")

    tab_overview, tab_hours, tab_sales, tab_products, tab_addons, tab_actions, tab_quality, tab_campaigns, tab_fillings, tab_branch, tab_tables, tab_export = st.tabs(
        ["🏠 النظرة العامة", "⏱️ الساعات", "💰 المبيعات", "🧁 المنتجات", "🎈 الإضافات", "🚨 متابعة", "🧹 جودة البيانات", "🎯 الحملات", "🍰 الحشوات", "🏬 تحليل فرع", "📋 الجداول", "⬇️ التصدير"]
    )

    with tab_overview:
        st.markdown('<div class="section-title">نظرة تشغيلية سريعة</div>', unsafe_allow_html=True)

        branch_df = hourly.drop(index="الإجمالي", errors="ignore").copy()
        branch_totals = branch_df["الإجمالي"].sort_values(ascending=False).reset_index()
        branch_totals.columns = ["الفرع", "عدد الطلبات"]

        hour_totals = hourly.loc["الإجمالي"].drop("الإجمالي", errors="ignore").reset_index()
        hour_totals.columns = ["الساعة", "عدد الطلبات"]

        c1, c2 = st.columns([1.15, 1])

        with c1:
            fig = px.bar(
                branch_totals,
                x="عدد الطلبات",
                y="الفرع",
                orientation="h",
                text="عدد الطلبات",
                title="ترتيب الفروع حسب عدد الطلبات",
                color="عدد الطلبات",
                color_continuous_scale=["#0ea5e9", "#22c55e", "#f59e0b"]
            )
            fig.update_traces(textposition="outside")
            fig.update_yaxes(categoryorder="total ascending")
            st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=mobile_chart_config())

        with c2:
            fig = px.line(
                hour_totals,
                x="الساعة",
                y="عدد الطلبات",
                markers=True,
                title="منحنى ضغط الطلبات بالساعة"
            )
            fig.update_traces(line=dict(width=4, color="#22c55e"), marker=dict(size=10))
            st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=mobile_chart_config())

        c3, c4 = st.columns([1, 1])

        with c3:
            if len(filling_rows):
                fill_mix = (
                    filling_rows.groupby("الحشوة الموحدة")["كمية رقمية"]
                    .sum()
                    .sort_values(ascending=False)
                    .reset_index()
                )
                fill_mix.columns = ["الحشوة", "الكمية"]

                fig = px.pie(
                    fill_mix,
                    names="الحشوة",
                    values="الكمية",
                    hole=0.55,
                    title="Mix الحشوات"
                )
                fig.update_traces(textposition="inside", textinfo="percent+label")
                st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=mobile_chart_config())
            else:
                st.info("لا توجد حشوات واضحة في الملف.")

        with c4:
            heatmap_df = hourly.drop(index="الإجمالي", errors="ignore").drop(columns="الإجمالي", errors="ignore")
            if not heatmap_df.empty:
                fig = px.imshow(
                    heatmap_df,
                    text_auto=True,
                    aspect="auto",
                    title="Heatmap الطلبات: الفرع × الساعة",
                    color_continuous_scale=["#0f172a", "#2563eb", "#22c55e", "#facc15"]
                )
                fig.update_xaxes(side="top")
                st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=mobile_chart_config())

    with tab_hours:
        st.markdown('<div class="section-title">تحليل الطلبات حسب الساعة</div>', unsafe_allow_html=True)

        hour_totals = hourly.loc["الإجمالي"].drop("الإجمالي", errors="ignore").reset_index()
        hour_totals.columns = ["الساعة", "عدد الطلبات"]

        fig = px.bar(
            hour_totals,
            x="الساعة",
            y="عدد الطلبات",
            text="عدد الطلبات",
            title="إجمالي الطلبات في كل ساعة",
            color="عدد الطلبات",
            color_continuous_scale=["#1d4ed8", "#16a34a", "#f59e0b"]
        )
        fig.update_traces(textposition="outside")
        st.plotly_chart(fig_layout(fig, 440), use_container_width=True, config=mobile_chart_config())

        st.markdown('<div class="section-title">جدول الفرع × الساعة</div>', unsafe_allow_html=True)
        display_pretty_dataframe(hourly, height=450)


    with tab_sales:
        st.markdown('<div class="section-title">تقرير المبيعات حسب الفرع والساعة</div>', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            if len(branch_sales):
                fig = px.bar(
                    branch_sales,
                    x="الفرع",
                    y="إجمالي_المبيعات",
                    text="إجمالي_المبيعات",
                    title="إجمالي المبيعات لكل فرع",
                    color="إجمالي_المبيعات",
                    color_continuous_scale=["#1d4ed8", "#16a34a", "#f59e0b"]
                )
                fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
                st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=mobile_chart_config())
        with c2:
            if len(hour_sales):
                fig = px.line(
                    hour_sales,
                    x="الساعة",
                    y="إجمالي_المبيعات",
                    markers=True,
                    title="المبيعات حسب الساعة"
                )
                fig.update_traces(mode="lines+markers+text", text=hour_sales["إجمالي_المبيعات"].round(0), textposition="top center")
                st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=mobile_chart_config())

        st.markdown('<div class="section-title">جدول مبيعات الفروع</div>', unsafe_allow_html=True)
        display_pretty_dataframe(branch_sales, height=430)

        st.markdown('<div class="section-title">المبيعات حسب الساعة</div>', unsafe_allow_html=True)
        display_pretty_dataframe(hour_sales, height=360)

    with tab_products:
        st.markdown('<div class="section-title">تقرير أداء المنتجات</div>', unsafe_allow_html=True)
        if len(product_perf):
            top_products = product_perf.head(20)
            fig = px.bar(
                top_products.sort_values("إجمالي_المبيعات", ascending=True),
                x="إجمالي_المبيعات",
                y="اسم المنتج النظيف",
                orientation="h",
                text="إجمالي_المبيعات",
                title="أعلى المنتجات حسب المبيعات"
            )
            fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
            st.plotly_chart(fig_layout(fig, 620), use_container_width=True, config=mobile_chart_config())

            c1, c2 = st.columns(2)
            with c1:
                qty_top = product_perf.sort_values("الكمية", ascending=False).head(15)
                fig = px.bar(qty_top, x="اسم المنتج النظيف", y="الكمية", text="الكمية", title="أعلى المنتجات حسب الكمية")
                fig.update_traces(textposition="outside")
                st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=mobile_chart_config())
            with c2:
                if len(products_by_branch):
                    branch_choice = st.selectbox("اختر فرع لعرض منتجاته", sorted(products_by_branch["الفرع"].dropna().unique().tolist()), key="product_branch_choice")
                    branch_products = products_by_branch[products_by_branch["الفرع"] == branch_choice].head(15)
                    fig = px.bar(branch_products, x="اسم المنتج النظيف", y="الكمية", text="الكمية", title=f"Top Products - {branch_choice}")
                    fig.update_traces(textposition="outside")
                    st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=mobile_chart_config())

            st.markdown('<div class="section-title">جدول أداء المنتجات</div>', unsafe_allow_html=True)
            display_pretty_dataframe(product_perf, height=520)
        else:
            st.info("لا توجد بيانات منتجات واضحة.")

    with tab_addons:
        st.markdown('<div class="section-title">تقرير الإضافات والـ Add-ons</div>', unsafe_allow_html=True)
        if len(addons_summary):
            a1, a2, a3 = st.columns(3)
            addon_orders_value = addon_kpis.loc[addon_kpis["البند"] == "طلبات بها إضافات", "القيمة"].iloc[0] if len(addon_kpis) else 0
            addon_pct_value = addon_kpis.loc[addon_kpis["البند"] == "نسبة الطلبات بها إضافات", "القيمة"].iloc[0] if "نسبة الطلبات بها إضافات" in addon_kpis["البند"].tolist() else 0
            addon_sales_value = addons_summary["إجمالي_المبيعات"].sum() if "إجمالي_المبيعات" in addons_summary.columns else 0
            with a1:
                render_kpi("طلبات بها إضافات", format_int(addon_orders_value), "Orders with Add-ons", "#8b5cf6")
            with a2:
                render_kpi("نسبة الإضافات", f"{float(addon_pct_value):.1f}%", "من إجمالي الطلبات", "#06b6d4")
            with a3:
                render_kpi("مبيعات الإضافات", f"{addon_sales_value:,.0f} SAR", "Add-ons Revenue", "#f59e0b")

            c1, c2 = st.columns(2)
            with c1:
                fig = px.bar(addons_summary, x="اسم المنتج النظيف", y="الكمية", text="الكمية", title="أكثر الإضافات مبيعاً")
                fig.update_traces(textposition="outside")
                st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=mobile_chart_config())
            with c2:
                fig = px.pie(addons_summary, names="اسم المنتج النظيف", values="الكمية", hole=0.48, title="Mix الإضافات")
                fig.update_traces(textinfo="percent+label")
                st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=mobile_chart_config())

            st.markdown('<div class="section-title">KPIs الإضافات</div>', unsafe_allow_html=True)
            display_pretty_dataframe(addon_kpis, height=250)
            st.markdown('<div class="section-title">جدول الإضافات</div>', unsafe_allow_html=True)
            display_pretty_dataframe(addons_summary, height=420)
        else:
            st.info("لا توجد إضافات واضحة في الفترة المختارة.")

    with tab_actions:
        st.markdown('<div class="section-title">طلبات تحتاج متابعة</div>', unsafe_allow_html=True)
        st.markdown(
            """
            <div class="glass-box small-note">
            هذا التقرير يلتقط الطلبات التي تحتوي على صورة، رقم تواصل، واتساب، Call، كتابة خاصة، أو تعديل تصميم.
            مهم جداً لطلبات الكيك بالصور والطلبات المخصصة.
            </div>
            """,
            unsafe_allow_html=True
        )
        if len(need_action_report):
            reason_counts = need_action_report["سبب المتابعة"].value_counts().reset_index()
            reason_counts.columns = ["سبب المتابعة", "عدد الطلبات"]
            c1, c2 = st.columns([1, 1.2])
            with c1:
                fig = px.bar(reason_counts, x="سبب المتابعة", y="عدد الطلبات", text="عدد الطلبات", title="أسباب المتابعة")
                fig.update_traces(textposition="outside")
                st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=mobile_chart_config())
            with c2:
                branch_action = need_action_report.groupby("الفرع").size().reset_index(name="عدد الطلبات") if "الفرع" in need_action_report.columns else pd.DataFrame()
                if len(branch_action):
                    fig = px.bar(branch_action, x="الفرع", y="عدد الطلبات", text="عدد الطلبات", title="طلبات المتابعة حسب الفرع")
                    fig.update_traces(textposition="outside")
                    st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=mobile_chart_config())
            display_pretty_dataframe(need_action_report, height=560)
        else:
            st.success("لا توجد طلبات تحتاج متابعة واضحة في الفترة المختارة.")

    with tab_quality:
        st.markdown('<div class="section-title">تقرير جودة البيانات</div>', unsafe_allow_html=True)
        q1, q2 = st.columns(2)
        with q1:
            if len(quality_summary):
                fig = px.bar(quality_summary, x="المشكلة", y="عدد الحالات", color="الأهمية", text="عدد الحالات", title="ملخص مشاكل البيانات")
                fig.update_traces(textposition="outside")
                st.plotly_chart(fig_layout(fig, 460), use_container_width=True, config=mobile_chart_config())
        with q2:
            if len(status_report):
                status_for_chart = status_report.drop(index="الإجمالي", errors="ignore").reset_index().melt(id_vars="الفرع", var_name="الحالة", value_name="عدد الطلبات")
                fig = px.bar(status_for_chart, x="الفرع", y="عدد الطلبات", color="الحالة", title="حالات الطلبات حسب الفرع", barmode="stack")
                st.plotly_chart(fig_layout(fig, 460), use_container_width=True, config=mobile_chart_config())
        st.markdown('<div class="section-title">ملخص جودة البيانات</div>', unsafe_allow_html=True)
        display_pretty_dataframe(quality_summary, height=320)
        st.markdown('<div class="section-title">تفاصيل مشاكل البيانات</div>', unsafe_allow_html=True)
        display_pretty_dataframe(data_quality_report, height=480)
        if len(status_report):
            st.markdown('<div class="section-title">جدول حالات الطلبات</div>', unsafe_allow_html=True)
            display_pretty_dataframe(status_report, height=360)

    with tab_campaigns:
        st.markdown('<div class="section-title">تقرير الحملات والمواسم</div>', unsafe_allow_html=True)
        if len(campaign_summary):
            c1, c2 = st.columns(2)
            with c1:
                fig = px.bar(campaign_summary, x="الحملة", y="إجمالي_المبيعات", text="إجمالي_المبيعات", title="مبيعات كل حملة")
                fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
                st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=mobile_chart_config())
            with c2:
                fig = px.pie(campaign_summary, names="الحملة", values="عدد_الطلبات", hole=0.48, title="نسبة الطلبات حسب الحملة")
                fig.update_traces(textinfo="percent+label")
                st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=mobile_chart_config())

            selected_campaign = st.selectbox("اختر حملة لعرض منتجاتها", campaign_summary["الحملة"].tolist(), key="campaign_choice")
            selected_campaign_products = campaign_products[campaign_products["الحملة"] == selected_campaign].head(20)
            if len(selected_campaign_products):
                fig = px.bar(selected_campaign_products, x="اسم المنتج النظيف", y="الكمية", text="الكمية", title=f"منتجات حملة {selected_campaign}")
                fig.update_traces(textposition="outside")
                st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=mobile_chart_config())

            st.markdown('<div class="section-title">ملخص الحملات</div>', unsafe_allow_html=True)
            display_pretty_dataframe(campaign_summary, height=320)
            st.markdown('<div class="section-title">منتجات كل حملة</div>', unsafe_allow_html=True)
            display_pretty_dataframe(campaign_products, height=520)
        else:
            st.info("لا توجد حملات واضحة في الفترة المختارة.")

    with tab_fillings:
        st.markdown('<div class="section-title">تحليل الحشوات حسب الفرع والكمية</div>', unsafe_allow_html=True)

        if len(filling_rows):
            fill_series = fillings_qty.loc["الإجمالي"].drop("الإجمالي", errors="ignore").sort_values(ascending=False).reset_index()
            fill_series.columns = ["الحشوة", "الكمية"]

            c1, c2 = st.columns([1.1, 1])

            with c1:
                fig = px.bar(
                    fill_series,
                    x="الحشوة",
                    y="الكمية",
                    text="الكمية",
                    title="أكثر الحشوات طلباً",
                    color="الكمية",
                    color_continuous_scale=["#2563eb", "#16a34a", "#f59e0b"]
                )
                fig.update_traces(textposition="outside")
                st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=mobile_chart_config())

            with c2:
                fig = px.pie(
                    fill_series,
                    names="الحشوة",
                    values="الكمية",
                    hole=0.48,
                    title="نسبة كل حشوة من الإجمالي"
                )
                fig.update_traces(textposition="inside", textinfo="percent+label")
                st.plotly_chart(fig_layout(fig, 430), use_container_width=True, config=mobile_chart_config())

            st.markdown('<div class="section-title">الحشوات لكل فرع - حسب الكمية</div>', unsafe_allow_html=True)
            display_pretty_dataframe(fillings_qty, height=420)

            st.markdown('<div class="section-title">الحشوات لكل فرع - حسب عدد الطلبات</div>', unsafe_allow_html=True)
            display_pretty_dataframe(fillings_orders, height=420)
        else:
            st.info("لا توجد حشوات واضحة في الملف.")

    with tab_branch:
        st.markdown('<div class="section-title">تحليل تفصيلي حسب الفرع</div>', unsafe_allow_html=True)

        available_branches = sorted(unique_orders["الفرع"].dropna().unique().tolist())

        if available_branches:
            selected_branch = st.selectbox(
                "اختر الفرع",
                available_branches,
                index=0
            )

            branch_orders = unique_orders[unique_orders["الفرع"] == selected_branch]
            branch_items = filling_rows[filling_rows["الفرع"] == selected_branch] if len(filling_rows) else pd.DataFrame()

            b1, b2, b3 = st.columns(3)
            with b1:
                render_kpi("طلبات الفرع", format_int(branch_orders[order_id_col].nunique()), selected_branch, "#2563eb")
            with b2:
                render_kpi("حشوات الفرع", format_int(branch_items["كمية رقمية"].sum() if len(branch_items) else 0), "إجمالي الكمية", "#16a34a")
            with b3:
                if len(branch_items):
                    bf = branch_items.groupby("الحشوة الموحدة")["كمية رقمية"].sum().sort_values(ascending=False)
                    render_kpi("أعلى حشوة", bf.index[0], f"{format_int(bf.iloc[0])} قطعة", "#ea580c")
                else:
                    render_kpi("أعلى حشوة", "-", "لا توجد حشوات", "#ea580c")

            c1, c2 = st.columns(2)

            with c1:
                branch_hour = (
                    branch_orders.groupby("الساعة")[order_id_col]
                    .nunique()
                    .reindex(sorted_hour_cols, fill_value=0)
                    .reset_index()
                )
                branch_hour.columns = ["الساعة", "عدد الطلبات"]

                fig = px.bar(
                    branch_hour,
                    x="الساعة",
                    y="عدد الطلبات",
                    text="عدد الطلبات",
                    title=f"طلبات {selected_branch} بالساعة",
                    color="عدد الطلبات",
                    color_continuous_scale=["#1d4ed8", "#16a34a"]
                )
                fig.update_traces(textposition="outside")
                st.plotly_chart(fig_layout(fig, 410), use_container_width=True, config=mobile_chart_config())

            with c2:
                if len(branch_items):
                    branch_fill = (
                        branch_items.groupby("الحشوة الموحدة")["كمية رقمية"]
                        .sum()
                        .sort_values(ascending=False)
                        .reset_index()
                    )
                    branch_fill.columns = ["الحشوة", "الكمية"]

                    fig = px.pie(
                        branch_fill,
                        names="الحشوة",
                        values="الكمية",
                        hole=0.50,
                        title=f"Mix الحشوات في {selected_branch}"
                    )
                    fig.update_traces(textinfo="percent+label")
                    st.plotly_chart(fig_layout(fig, 410), use_container_width=True, config=mobile_chart_config())
                else:
                    st.info("لا توجد حشوات لهذا الفرع.")

            if selected_branch in branch_hour_details_dict:
                st.markdown('<div class="section-title">الحشوات بالساعة للفرع المختار</div>', unsafe_allow_html=True)
                selected_table = branch_hour_details_dict[selected_branch]
                display_pretty_dataframe(selected_table, height=430)

                stack_df = selected_table.drop(columns="الإجمالي", errors="ignore").reset_index().melt(
                    id_vars="الساعة",
                    var_name="الحشوة",
                    value_name="الكمية"
                )
                stack_df = stack_df[stack_df["الكمية"] > 0]

                if len(stack_df):
                    fig = px.bar(
                        stack_df,
                        x="الساعة",
                        y="الكمية",
                        color="الحشوة",
                        title=f"الحشوات بالساعة - {selected_branch}",
                        barmode="stack"
                    )
                    st.plotly_chart(fig_layout(fig, 460), use_container_width=True, config=mobile_chart_config())

    with tab_tables:
        st.markdown('<div class="section-title">الجداول التفصيلية</div>', unsafe_allow_html=True)

        with st.expander("معاينة أول 20 صف من البيانات المنظفة", expanded=False):
            st.dataframe(clean.head(20), use_container_width=True, height=420)

        with st.expander("تحليل الطلبات لكل فرع بالساعة", expanded=True):
            display_pretty_dataframe(hourly, height=430)

        with st.expander("تحليل الحشوات بالساعة لكل فرع", expanded=False):
            display_pretty_dataframe(fillings_by_hour_branch, height=520)

        with st.expander("تقرير المبيعات حسب الفرع", expanded=False):
            display_pretty_dataframe(branch_sales, height=430)

        with st.expander("تقرير أداء المنتجات", expanded=False):
            display_pretty_dataframe(product_perf, height=520)

        with st.expander("تقرير الإضافات", expanded=False):
            display_pretty_dataframe(addons_summary, height=420)

        with st.expander("طلبات تحتاج متابعة", expanded=False):
            display_pretty_dataframe(need_action_report, height=520)

        with st.expander("جودة البيانات", expanded=False):
            display_pretty_dataframe(data_quality_report, height=520)

        with st.expander("الطلبات متعددة الأصناف", expanded=False):
            display_pretty_dataframe(multi_item_orders, height=520)

        with st.expander("تقرير الحملات", expanded=False):
            display_pretty_dataframe(campaign_summary, height=360)

        with st.expander("Summary Data", expanded=False):
            st.dataframe(summary_df, use_container_width=True, hide_index=True)

    with tab_export:
        st.markdown('<div class="section-title">تحميل ملف Excel التفاعلي</div>', unsafe_allow_html=True)

        st.markdown(
            """
            <div class="glass-box small-note">
            الملف الناتج Excel تفاعلي متوافق. عدّل أو احذف من شيت <b>Cleaned Data</b>،
            وخصوصاً أعمدة <b>Calc_</b>، وسيتم تحديث الجداول والرسومات داخل Excel.
            </div>
            """,
            unsafe_allow_html=True
        )

        excel_file = create_excel_file(
            summary_df=summary_df,
            hourly_df=hourly,
            fillings_qty_df=fillings_qty,
            fillings_orders_df=fillings_orders,
            fillings_by_hour_branch_df=fillings_by_hour_branch,
            branch_hour_details_dict=branch_hour_details_dict,
            clean_df=clean,
            unique_orders_df=unique_orders
        )

        st.download_button(
            label="⬇️ تحميل ملف Excel V6 للفترة المختارة",
            data=excel_file,
            file_name="تحليل_الساعات_والحشوات_Cloud_V6.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

else:
    st.markdown(
        """
        <div class="glass-box small-note">
        👈 اختر مصدر البيانات من القائمة الجانبية: رفع ملف أو Google Sheet. بعد تحميل البيانات اختر فلترة تاريخ الاستلام ثم سيظهر Dashboard.
        </div>
        """,
        unsafe_allow_html=True
    )
