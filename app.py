# app.py
import streamlit as st
import pandas as pd
import io
import os

st.set_page_config(page_title="Wall Detail Platform (robust)", layout="wide")
EXCEL_PATH = "materials.xlsx"

# ---------- helpers ----------
@st.cache_data
def load_sheets(path):
    if not os.path.exists(path):
        raise FileNotFoundError(f"فایل '{path}' پیدا نشد. لطفاً آن را کنار app.py قرار دهید.")
    # try reading with header=0 first; also keep header=None copy for raw preview
    sheets_with_header = pd.read_excel(path, sheet_name=None, engine="openpyxl", header=0, dtype=object)
    sheets_raw = pd.read_excel(path, sheet_name=None, engine="openpyxl", header=None, dtype=object)
    # normalize to string and fillna for preview
    for k in sheets_with_header:
        sheets_with_header[k] = sheets_with_header[k].fillna("").astype(str)
    for k in sheets_raw:
        sheets_raw[k] = sheets_raw[k].fillna("").astype(str)
    return sheets_with_header, sheets_raw

def get_column_options(df):
    """Return list of readable column labels for selectbox: use actual names if exist else numeric indices."""
    cols = list(df.columns)
    readable = []
    for c in cols:
        readable.append(str(c))
    return readable

def build_province_list(df_prov, prov_code_col, prov_name_col):
    # take rows where prov_code_col not empty (or prov_name if code missing)
    s_code = df_prov[prov_code_col].astype(str).str.strip()
    s_name = df_prov[prov_name_col].astype(str).str.strip() if prov_name_col in df_prov.columns else s_code
    mask = (s_code != "") | (s_name != "")
    codes = s_code[mask].tolist()
    names = s_name[mask].tolist()
    # unique preserving order based on visible name
    seen = set()
    opts = []
    for c, n in zip(codes, names):
        display = n if n and n.lower() != "nan" else c
        key = (c.strip(), display.strip())
        if key[1] not in seen:
            seen.add(key[1])
            opts.append(key)
    return opts

def build_city_list(df_cities, prov_col_sel, city_code_col, city_name_col, selected_prov_code):
    df = df_cities.copy()
    # restrict to first 4 cols if many (common case)
    if df.shape[1] > 8:
        df = df.iloc[:, :8]
    # filter rows where selected_prov_code appears in prov_col_sel OR anywhere in row
    def row_matches_prov(row):
        try:
            v = str(row[prov_col_sel]).strip()
            if v and selected_prov_code.strip().lower() in v.lower():
                return True
        except Exception:
            pass
        # fallback: check entire row
        joined = " | ".join([str(x).strip() for x in row])
        if selected_prov_code.strip().lower() in joined.lower():
            return True
        return False
    df_filtered = df[df.apply(row_matches_prov, axis=1)]
    # build city tuples
    cities = []
    for _, r in df_filtered.iterrows():
        code = str(r[city_code_col]).strip() if city_code_col in df.columns else ""
        name = str(r[city_name_col]).strip() if city_name_col in df.columns else ""
        if (not name or name.lower() in ["nan","none"]) and code:
            name = code
        if name:
            cities.append((code, name))
    # fallback: if none, try scanning entire sheet for lines containing province code and any C- pattern
    if not cities:
        import re
        pat_c = re.compile(r"(?i)\bC-\d{2}-\d{2}\b")
        for i in range(df.shape[0]):
            row = df.iloc[i]
            joined = " | ".join([str(x).strip() for x in row])
            if selected_prov_code.strip().lower() in joined.lower() and pat_c.search(joined):
                m = pat_c.search(joined).group(0).upper()
                cities.append((m, m))
    # unique preserve
    seen = set()
    uniq = []
    for code, name in cities:
        if name not in seen:
            seen.add(name)
            uniq.append((code, name))
    return uniq

def extract_sheet3_rows(df3, prov_code, city_code_or_name):
    df = df3.copy().astype(str)
    rows = []
    for i in range(df.shape[0]):
        row_text = " | ".join([df.iat[i,j] for j in range(df.shape[1])])
        if prov_code.strip().lower() in row_text.lower() and city_code_or_name.strip().lower() in row_text.lower():
            rows.append(i)
    if not rows:
        # fallback: if city_code_or_name empty, match just province; or match city anywhere
        for i in range(df.shape[0]):
            row_text = " | ".join([df.iat[i,j] for j in range(df.shape[1])])
            if city_code_or_name.strip().lower() in row_text.lower():
                rows.append(i)
    if not rows:
        return pd.DataFrame()
    res = df.iloc[rows].reset_index(drop=True)
    res.columns = [f"Column_{i+1}" for i in range(res.shape[1])]
    return res

# ---------- UI ----------
st.title("🧱 Wall Detail Platform — انتخاب ستونی دستی برای پایداری")
st.write("این برنامه تعاملی است: اگر ستون‌ها در فایل اکسل ساختار متفاوتی دارند، از منوها ستون درست را انتخاب کن. این روش خطاهای تشخیص خودکار را حذف می‌کند.")

# load
try:
    sheets_h, sheets_raw = load_sheets(EXCEL_PATH)
except Exception as e:
    st.error("خطا در خواندن فایل اکسل: " + str(e))
    st.stop()

sheet_names = list(sheets_h.keys())
st.info("شیت‌های موجود: " + ", ".join(sheet_names))

# pick sheet names (defaults to sheet0, sheet1, sheet3 if present)
sheet0 = next((k for k in sheet_names if k.lower().strip()=="sheet0"), sheet_names[0])
sheet1 = next((k for k in sheet_names if k.lower().strip()=="sheet1"), sheet_names[1] if len(sheet_names)>1 else sheet_names[0])
sheet3 = next((k for k in sheet_names if k.lower().strip()=="sheet3"), sheet_names[-1])

df0 = sheets_h[sheet0]
df1 = sheets_h[sheet1]
df3 = sheets_h[sheet3]

st.subheader("مرحله ۱ — انتخاب ستون‌های شیت استان‌ها (Sheet0)")
st.write("پیش‌نمایش (چند سطر اول):")
st.dataframe(sheets_raw[sheet0].head(8))
col_opts0 = get_column_options(df0)
prov_code_col = st.selectbox("ستون کد استان (مثلاً P-01):", col_opts0, index=0)
prov_name_col = st.selectbox("ستون نام استان (اگر نام ندارد، همان کد را انتخاب کن):", col_opts0, index=1 if len(col_opts0)>1 else 0)

st.markdown("---")
st.subheader("مرحله ۲ — انتخاب ستون‌های شیت شهرها (Sheet1)")
st.write("پیش‌نمایش (چند سطر اول):")
st.dataframe(sheets_raw[sheet1].head(8))
col_opts1 = get_column_options(df1)
prov_code_col_1 = st.selectbox("ستون کد استان در Sheet1:", col_opts1, index=0)
prov_name_col_1 = st.selectbox("ستون نام استان در Sheet1 (معمولاً کنار کد):", col_opts1, index=1 if len(col_opts1)>1 else 0)
city_code_col = st.selectbox("ستون کد شهر (مثلاً C-01-01):", col_opts1, index=2 if len(col_opts1)>2 else 0)
city_name_col = st.selectbox("ستون نام شهر (فارسی):", col_opts1, index=3 if len(col_opts1)>3 else (2 if len(col_opts1)>2 else 0))

# build provinces list using chosen columns
try:
    provinces = build_province_list(df0, prov_code_col, prov_name_col)
except Exception as e:
    st.error("خطا در ساخت لیست استان‌ها با ستون‌های انتخاب‌شده: " + str(e))
    st.stop()

if not provinces:
    st.error("هیچ استان/ردیفی با ستون‌های انتخاب‌شده یافت نشد. مطمئن شو ستون‌ها درست انتخاب شده‌اند.")
    st.stop()

prov_display = [f"{c} — {n}" for c,n in provinces]
sel_prov_idx = st.selectbox("انتخاب استان از لیست:", range(len(prov_display)), format_func=lambda i: prov_display[i])
selected_prov_code, selected_prov_name = provinces[sel_prov_idx]

st.markdown("---")
st.subheader("مرحله ۳ — استخراج و انتخاب شهرها برای استان انتخاب‌شده")
cities = build_city_list(df1, prov_code_col_1, city_code_col, city_name_col, selected_prov_code)
st.write(f"تعداد شهرهای استخراج‌شده: {len(cities)}")
if len(cities) == 0:
    st.warning("هیچ شهری استخراج نشد — لطفاً ستون‌های Sheet1 را بررسی و دوباره انتخاب کن.")
    st.dataframe(sheets_raw[sheet1].head(30))
    st.stop()

city_labels = [name for code,name in cities]
sel_city_idx = st.selectbox("انتخاب شهر (فقط نام شهر نمایش داده می‌شود):", range(len(city_labels)), format_func=lambda i: city_labels[i])
selected_city_code, selected_city_name = cities[sel_city_idx]

st.markdown("---")
st.subheader("مرحله ۴ — نمایش جزئیات از Sheet3 و دانلود")
st.write("پیش‌نمایش Sheet3 (چند سطر اول):")
st.dataframe(sheets_raw[sheet3].head(8))

if st.button("نمایش دیتیل دیوار برای استان و شهر انتخابی"):
    with st.spinner("در حال جستجو در Sheet3..."):
        res = extract_sheet3_rows(df3, selected_prov_code, selected_city_code or selected_city_name)
        if res.empty:
            st.warning("هیچ ردیفی در Sheet3 مطابق استان و شهر انتخابی پیدا نشد.")
            st.write("پیش‌نمایش چند سطر اول Sheet3 برای بررسی:")
            st.dataframe(sheets_raw[sheet3].head(50))
        else:
            st.success(f"{len(res)} ردیف یافت شد.")
            st.dataframe(res, use_container_width=True)
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                res.to_excel(writer, index=False, sheet_name="Wall_Details")
            buf.seek(0)
            st.download_button("📥 دانلود خروجی (Excel)", data=buf, file_name=f"Wall_Details_{selected_city_name}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
