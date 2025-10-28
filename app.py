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
                name = ""
                for k in range(1, 4):
                    if j + k < cols:
                        cand = df_sheet0.iat[i, j + k].strip()
                        if cand and not pattern_p.search(cand):
                            name = cand
                            break
                if not name:
                    name = code
                if (code, name) not in provinces:
                    provinces.append((code, name))
    return provinces

# ------------------- استخراج شهرها بر اساس استان -------------------
def detect_cities_for_province(df_sheet1, province_code):
    """
    ساختار Sheet1 معمولاً شامل ستون‌های:
    [کد استان | نام استان | کد شهر | نام شهر]
    """
    df = df_sheet1.copy()
    df = df.replace("", pd.NA).dropna(how="all")
    cols = list(df.columns)

    prov_col, city_code_col, city_name_col = None, None, None
    for i, c in enumerate(cols):
        sample_vals = " ".join(df[c].astype(str).head(15).tolist())
        if re.search(r"P-\d{2}", sample_vals, re.I):
            prov_col = c
        if re.search(r"C-\d{2}-\d{2}", sample_vals, re.I):
            city_code_col = c
        if re.search(r"[\u0600-\u06FF]", sample_vals) and not re.search(r"P-|C-", sample_vals, re.I):
            city_name_col = c

    if prov_col is None:
        prov_col = cols[0]
    if city_code_col is None:
        city_code_col = cols[2] if len(cols) > 2 else cols[0]
    if city_name_col is None:
        city_name_col = cols[3] if len(cols) > 3 else cols[-1]

    filtered = df[df[prov_col].astype(str).str.contains(province_code, case=False, na=False)]

    cities = []
    for _, row in filtered.iterrows():
        c_code = str(row[city_code_col]).strip()
        c_name = str(row[city_name_col]).strip()
        if not c_name or c_name.lower() == "nan":
            c_name = c_code
        if c_code or c_name:
            cities.append((c_code, c_name))

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
        res = df.loc[matched_rows.values, :].reset_index(drop=True)
        res.columns = [f"Column_{i+1}" for i in range(res.shape[1])]
        return res
    else:
        return pd.DataFrame()

# ------------------- رابط کاربری Streamlit -------------------
st.title("🧱 Wall Detail Platform — Tehran-based Material Dataset")
st.write("ابتدا استان را انتخاب کنید، سپس شهر مربوطه را انتخاب کنید و در نهایت جزئیات دیوار نمایش داده می‌شود.")

# بارگذاری فایل اکسل
try:
    sheets = load_all_sheets(EXCEL_PATH)
except Exception as e:
    st.error(f"❌ خطا در خواندن فایل اکسل: {e}")
    st.stop()

sheet_names = list(sheets.keys())
sheet0_key = next((k for k in sheet_names if "0" in k.lower()), sheet_names[0])
sheet1_key = next((k for k in sheet_names if "1" in k.lower()), sheet_names[1] if len(sheet_names) > 1 else sheet_names[0])
sheet3_key = next((k for k in sheet_names if "3" in k.lower()), sheet_names[-1])

df0 = sheets[sheet0_key]
df1 = sheets[sheet1_key]
df3 = sheets[sheet3_key]

# --- انتخاب استان ---
provs = detect_provinces(df0)
if not provs:
    st.error("هیچ استانی در Sheet0 یافت نشد.")
    st.stop()

province_labels = [name for code, name in provs]
province_idx = st.selectbox("انتخاب استان:", range(len(provs)), format_func=lambda i: province_labels[i])
selected_province_code, selected_province_name = provs[province_idx][0], provs[province_idx][1]

# --- انتخاب شهر ---
cities = detect_cities_for_province(df1, selected_province_code)
if not cities:
    st.warning("برای این استان، هیچ شهری یافت نشد.")
    st.stop()

city_labels = [name for code, name in cities]
city_idx = st.selectbox("انتخاب شهر:", range(len(cities)), format_func=lambda i: city_labels[i])
selected_city_identifier = cities[city_idx][0]
selected_city_name = cities[city_idx][1]

# --- نمایش دیتیل ---
if st.button("نمایش جزئیات دیوار"):
    with st.spinner("در حال استخراج داده‌ها..."):
        result_df = extract_details_sheet3(df3, selected_province_code, selected_city_identifier)
        if result_df.empty:
            st.warning("هیچ دیتیلی برای این شهر در Sheet3 پیدا نشد.")
        else:
            st.success(f"✅ {len(result_df)} ردیف دیتیل برای شهر '{selected_city_name}' یافت شد.")
            st.dataframe(result_df, use_container_width=True)

            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                result_df.to_excel(writer, index=False, sheet_name="Wall_Details")
            buf.seek(0)
            st.download_button(
                label="📥 دانلود خروجی (Excel)",
                data=buf,
                file_name=f"Wall_Details_{selected_city_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
