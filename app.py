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
            st.error("❌ فایل materials.xlsx در مسیر برنامه پیدا نشد.")
            return {}
        xls = pd.read_excel(path, sheet_name=None, engine="openpyxl")
        for name, df in xls.items():
            df.columns = df.columns.map(lambda c: str(c).strip() if not pd.isna(c) else c)
        return xls
    except Exception as e:
        st.error("خطا در بارگذاری اکسل: " + str(e))
        return {}

def make_province_options(df_sheet0):
    """استخراج استان‌ها (کد یا نام)."""
    if df_sheet0.empty:
        return []
    
    # اگر فقط یک ستون دارد (کدها)
    if len(df_sheet0.columns) == 1:
        col = df_sheet0.columns[0]
        rows = df_sheet0[col].dropna().astype(str).str.strip().tolist()
        return [(r, r, r) for r in rows]  # code, name, label = همان مقدار
    
    # اگر دو ستون دارد (کد + نام)
    code_col, name_col = df_sheet0.columns[:2]
    rows = df_sheet0[[code_col, name_col]].dropna(how="all")
    opts = []
    for _, r in rows.iterrows():
        code = str(r[code_col]).strip()
        name = str(r[name_col]).strip()
        label = f"{code} — {name}" if name else code
        opts.append((code, name if name else code, label))
    return opts

def find_cities_for_province(df_sheet1, selected_province):
    """فیلتر شهرها بر اساس استان انتخابی (کد یا نام)."""
    if df_sheet1.empty:
        return []
    
    # بررسی تمام سلول‌ها برای تطبیق کد استان
    mask = df_sheet1.apply(lambda x: x.astype(str).str.contains(selected_province, case=False, na=False))
    matched_rows = df_sheet1[mask.any(axis=1)]
    
    # اگر چیزی پیدا شد، شهرها را از آخرین ستون استخراج می‌کنیم
    if not matched_rows.empty:
        last_col = matched_rows.columns[-1]
        cities = matched_rows[last_col].dropna().astype(str).str.strip().unique().tolist()
        return cities
    
    # در غیر اینصورت، تمام مقادیر را به عنوان fallback برمی‌گردانیم
    all_values = pd.Series(df_sheet1.values.ravel()).dropna().astype(str).str.strip().unique().tolist()
    return all_values[:200]

def extract_wall_details(df_sheet3, selected_city):
    """یافتن جزئیات دیوار از Sheet3 بر اساس شهر."""
    if not selected_city:
        return pd.DataFrame()
    df = df_sheet3.copy().astype(str)
    mask = df.apply(lambda x: x.str.contains(selected_city, case=False, na=False))
    if mask.any().any():
        return df[mask.any(axis=1)]
    return pd.DataFrame()

# ---------------------- رابط کاربری ----------------------
st.title("🧱 پلتفرم نمایش جزئیات دیوار ساختمان‌ها")
st.write("این برنامه از فایل ثابت `materials.xlsx` اطلاعات استان‌ها، شهرها و دیتیل دیوارها را می‌خواند.")

sheets = load_sheets(EXCEL_PATH)
if not sheets:
    st.stop()

# انتخاب شیت‌ها
sheet0_name = [n for n in sheets if "0" in n][-1] if any("0" in n for n in sheets) else list(sheets.keys())[0]
sheet1_name = [n for n in sheets if "1" in n][-1] if any("1" in n for n in sheets) else list(sheets.keys())[0]
sheet3_name = [n for n in sheets if "3" in n][-1] if any("3" in n for n in sheets) else list(sheets.keys())[-1]

df0 = sheets[sheet0_name]
df1 = sheets[sheet1_name]
df3 = sheets[sheet3_name]

st.markdown("---")
st.subheader("۱. انتخاب استان و شهر")

provinces = make_province_options(df0)
if not provinces:
    st.error("هیچ استان یا کدی در Sheet0 پیدا نشد.")
    st.stop()

province_labels = [p[2] for p in provinces]
selected_province = st.selectbox("انتخاب استان:", province_labels)
selected_province_code = [p[0] for p in provinces if p[2] == selected_province][0]

# نمایش شهرها بر اساس استان انتخابی
cities = find_cities_for_province(df1, selected_province_code)
if not cities:
    st.warning("برای این استان، هیچ شهری در Sheet1 یافت نشد.")
else:
    selected_city = st.selectbox("انتخاب شهر:", cities)

    st.markdown("---")
    st.subheader("۲. مشاهده جزئیات دیوار")

    if st.button("نمایش جزئیات دیوار برای شهر انتخابی"):
        with st.spinner("در حال استخراج داده‌ها..."):
            result_df = extract_wall_details(df3, selected_city)
            if result_df.empty:
                st.warning("هیچ دیتیلی برای این شهر پیدا نشد.")
            else:
                st.success("✅ دیتیل دیوار پیدا شد")
                st.dataframe(result_df, use_container_width=True)

                # خروجی برای دانلود
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                    result_df.to_excel(writer, index=False, sheet_name="Wall_Details")
                buf.seek(0)
                st.download_button(
                    label="📥 دانلود خروجی (Excel)",
                    data=buf,
                    file_name=f"Wall_Details_{selected_city}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

st.markdown("---")
st.caption("🔹 فایل materials.xlsx را در همین مسیر به‌روزرسانی کنید تا داده‌های جدید نمایش داده شوند.")
