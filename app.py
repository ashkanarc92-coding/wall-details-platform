# app.py
import streamlit as st
import pandas as pd
import io
import os
import re

st.set_page_config(page_title="Wall Detail Platform", layout="wide")

EXCEL_PATH = "materials.xlsx"

# ------------------- بارگذاری شیت‌ها -------------------
@st.cache_data
def load_all_sheets(path):
    if not os.path.exists(path):
        raise FileNotFoundError(f"فایل '{path}' پیدا نشد. لطفاً آن را کنار app.py قرار دهید.")
    xls = pd.read_excel(path, sheet_name=None, engine="openpyxl", header=None, dtype=object)
    for k, df in xls.items():
        xls[k] = df.fillna("").astype(str)
    return xls

# ------------------- استخراج استان‌ها -------------------
def detect_provinces(df_sheet0):
    provinces = []
    pattern_p = re.compile(r"(?i)\bP-\d{2}\b")
    rows, cols = df_sheet0.shape
    for i in range(rows):
        for j in range(cols):
            cell = df_sheet0.iat[i, j].strip()
            if pattern_p.search(cell):
                code = pattern_p.search(cell).group(0).upper()
                # فرض: ستون بعدی نام استان است
                name = ""
                if j + 1 < cols:
                    name = df_sheet0.iat[i, j + 1].strip()
                if not name:
                    name = code
                if (code, name) not in provinces:
                    provinces.append((code, name))
    return provinces

# ------------------- استخراج شهرها بر اساس استان -------------------
def detect_cities_for_province(df_sheet1, province_code):
    """
    ساختار Sheet1 معمولاً:
    [کد استان | نام استان | کد شهر | نام شهر | سایر ستون‌ها...]
    """
    df = df_sheet1.copy()
    df = df.replace("", pd.NA).dropna(how="all")
    df = df.reset_index(drop=True)

    # فقط چهار ستون اول را نگه داریم تا از ستون‌های انرژی و ... جلوگیری شود
    df = df.iloc[:, :4] if df.shape[1] > 4 else df

    # نام‌گذاری ایمن ستون‌ها
    cols = df.columns
    prov_code_col = cols[0]
    prov_name_col = cols[1] if len(cols) > 1 else cols[0]
    city_code_col = cols[2] if len(cols) > 2 else cols[-1]
    city_name_col = cols[3] if len(cols) > 3 else cols[-1]

    # فیلتر ردیف‌ها برای استان انتخابی
    filtered = df[df[prov_code_col].astype(str).str.contains(province_code, case=False, na=False)]

    cities = []
    for _, row in filtered.iterrows():
        c_code = str(row[city_code_col]).strip()
        c_name = str(row[city_name_col]).strip()
        if not c_name or c_name.lower() == "nan":
            c_name = c_code
        if c_code or c_name:
            cities.append((c_code, c_name))

    # حذف تکراری‌ها
    unique_cities = []
    seen = set()
    for code, name in cities:
        if name not in seen:
            seen.add(name)
            unique_cities.append((code, name))
    return unique_cities

# ------------------- استخراج دیتیل دیوار از Sheet3 -------------------
def extract_details_sheet3(df_sheet3, selected_province_code, selected_city_identifier):
    df = df_sheet3.copy().astype(str)
    rows, cols = df.shape
    city_is_code = bool(re.match(r"(?i)^C-\d{2}-\d{2}$", str(selected_city_identifier)))
    matched_rows = pd.Series([False]*rows)
    for i in range(rows):
        row_text = " | ".join([df.iat[i, j] for j in range(cols)])
        cond_city = selected_city_identifier.strip().lower() in row_text.lower()
        cond_prov = selected_province_code.strip().lower() in row_text.lower()
        if city_is_code:
            if cond_city:
                matched_rows.iat[i] = True
        else:
            if cond_city and cond_prov:
                matched_rows.iat[i] = True
            elif cond_city and not matched_rows.any():
                matched_rows.iat[i] = True
    if matched_rows.any():
        res = df.loc[matched_rows.values, :].reset_index(drop]()
