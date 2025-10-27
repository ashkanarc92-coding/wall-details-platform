# app.py
import streamlit as st
import pandas as pd
import io
import os

st.set_page_config(page_title="پلتفرم دیتیل دیوار", layout="wide")

EXCEL_PATH = "materials.xlsx"

@st.cache_data
def load_sheets(path):
    try:
        if not os.path.exists(path):
            st.error("فایل materials.xlsx در مسیر برنامه پیدا نشد.")
            return {}
        xls = pd.read_excel(path, sheet_name=None, engine="openpyxl")
        for name, df in xls.items():
            df.columns = df.columns.map(lambda c: str(c).strip() if not pd.isna(c) else c)
        return xls
    except Exception as e:
        st.error("خطا در بارگذاری اکسل: " + str(e))
        return {}

def guess_code_name_columns(df):
    cols = list(df.columns)
    if len(cols) >= 2:
        return cols[0], cols[1]
    elif len(cols) == 1:
        return cols[0], None
    else:
        return None, None

def make_province_options(df_sheet0):
    """تولید لیست استان‌ها حتی اگر نام وجود نداشته باشد"""
    code_col, name_col = guess_code_name_columns(df_sheet0)
    if code_col is None:
        return []
    rows = df_sheet0.dropna(how="all")
    opts = []
    for _, r in rows.iterrows():
        code = str(r[code_col]).strip()
        name = str(r[name_col]).strip() if name_col and not pd.isna(r[name_col]) else ""
        if name:
            label = f"{code} — {name}"
        else:
            label = code
        opts.append((code, name, label))
    return opts

def find_cities_for_province(df_sheet1, province_code=None, province_name=None):
    """جستجوی شهرها بر اساس کد یا نام استان. اگر پیدا نشد، همه شهرها را برمی‌گرداند."""
    df = df_sheet1.copy()
    cols = list(df.columns)
    # تلاش برای یافتن شهرها مرتبط با استان
    city_list = []
    if province_code:
        for col in cols:
            try:
                matches = df[col].astype(str).str.strip().str.lower() == province_code.strip().lower()
                if matches.any():
                    # ستون بعدی یا آخر را به عنوان شهر فرض کن
                    next_col = cols[min(cols.index(col) + 1, len(cols)-1)]
                    city_list = df.loc[matches, next_col].dropna().astype(str).str.strip().unique().tolist()
                    break
            except Exception:
                continue
    if not city_list and province_name:
        for col in cols:
            try:
                matches = df[col].astype(str).str.strip().str.lower() == province_name.strip().lower()
                if matches.any():
                    next_col = cols[min(cols.index(col) + 1, len(cols)-1)]
                    city_list = df.loc[matches, next_col].dropna().astype(str).str.strip().unique().tolist()
                    break
            except Exception:
                continue
    # اگر هیچ چیز پیدا نشد، تمام شهرها را برگردان
    if not city_list:
        city_list = df.iloc[:, -1].dropna().astype(str).str.strip().unique().tolist()
    return city_list

def extract_wall_details(df_sheet3, selected_city):
    """یافتن جزئیات دیوار از Sheet3 برای شهر انتخابی"""
    if not selected_city:
        return pd.DataFrame()
    df = df_sheet3.copy().astype(str)
    mask = df.apply(lambda x: x.str.contains(selected_city, case=False, na=False))
    if mask.any().any():
        return df[mask.any(axis=1)]
    return pd.DataFrame()

# رابط کاربری ا
