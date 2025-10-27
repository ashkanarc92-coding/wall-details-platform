# app.py
import streamlit as st
import pandas as pd
import io
import os

# تنظیمات کلی صفحه
st.set_page_config(page_title="پلتفرم دیتیل دیوار", layout="wide")

# مسیر فایل اکسل (فایل باید کنار app.py باشد)
EXCEL_PATH = "materials.xlsx"

# ---------------------------------------------------------
# تابع برای بارگذاری شیت‌های اکسل
# ---------------------------------------------------------
@st.cache_data
def load_sheets(path):
    """بارگذاری همه شیت‌های فایل اکسل و بازگرداندن به صورت دیکشنری از DataFrameها"""
    try:
        if not os.path.exists(path):
            st.error("فایل اکسل materials.xlsx در مسیر برنامه پیدا نشد.")
            return {}
        xls = pd.read_excel(path, sheet_name=None, engine="openpyxl")
        # پاکسازی نام ستون‌ها از فاصله اضافی
        for name, df in xls.items():
            df.columns = df.columns.map(lambda c: str(c).strip() if not pd.isna(c) else c)
        return xls
    except Exception as e:
        st.error("خطا در بارگذاری اکسل: " + str(e))
        return {}

# ---------------------------------------------------------
# توابع کمکی برای پردازش داده‌ها
# ---------------------------------------------------------
def guess_code_name_columns(df):
    """حدس زدن ستون‌های کد و نام"""
    cols = list(df.columns)
    if len(cols) >= 2:
        return cols[0], cols[1]
    elif len(cols) == 1:
        return cols[0], cols[0]
    else:
        return None, None


def make_province_options(df_sheet0):
    """تولید لیست استان‌ها از Sheet0"""
    code_col, name_col = guess_code_name_columns(df_sheet0)
    if code_col is None:
        return []
    rows = df_sheet0[[code_col, name_col]].dropna(how="all")
    opts = []
    for _, r in rows.iterrows():
        code = str(r[code_col]).strip() if not pd.isna(r[code_col]) else ""
        name = str(r[name_col]).strip() if not pd.isna(r[name_col]) else ""
        label = code if name == "" else f"{code} — {name}"
        opts.append((code, name, label))
    return opts


def find_cities_for_province(df_sheet1, selected_province_code, selected_province_name):
    """یافتن شهرهای مرتبط با استان انتخابی از Sheet1"""
    df = df_sheet1.copy()
    cols = list(df.columns)
    # جستجوی بر اساس کد استان
    for col in cols:
        try:
            matches = df[col].astype(str).str.strip().str.lower() == selected_province_code.strip().lower()
            if matches.any():
                next_col = cols[cols.index(col) + 1] if cols.index(col) + 1 < len(cols) else cols[0]
                return df.loc[matches, next_col].dropna().astype(str).str.strip().unique().tolist()
        except Exception:
            continue
    # جستجو بر اساس نام استان
    for col in cols:
        try:
            matches = df[col].astype(str).str.strip().str.lower() == selected_province_name.strip().lower()
            if matches.any():
                next_col = cols[cols.index(col) + 1] if cols.index(col) + 1 < len(cols) else cols[0]
                return df.loc[matches, next_col].dropna().astype(str).str.strip().unique().tolist()
        except Exception:
            continue
    # در صورت عدم موفقیت
    return df.iloc[:, -1].dropna().astype(str).str.strip().unique().tolist()


def extract_wall_details(df_sheet3, selected_city):
    """یافتن جزئیات دیوار از Sheet3 برای شهر انتخابی"""
    if not selected_city:
        return pd.DataFrame()
    df = df_sheet3.copy().astype(str)
    mask = df.apply(lambda x: x.str.contains(selected_city, case=False, na=False))
    if mask.any().any():
        return df[mask.any(axis=1)]
    return pd.DataFrame()

# ---------------------------------------------------------
# رابط کاربری Streamlit
# ---------------------------------------------------------
st.title("🧱 پلتفرم نمایش جزئیات دیوار ساختمان‌ها در ایران")
st.write("این برنامه از فایل ثابت `materials.xlsx` اطلاعات را می‌خواند. ابتدا استان و سپس شهر را انتخاب کنید تا جزئیات دیوار نمایش داده شود.")

# بارگذاری شیت‌ها
sheets = load_sheets(EXCEL_PATH)
if not sheets:
    st.stop()

# انتخاب شیت‌ها بر اساس نام
sheet0_name = [n for n in sheets if "0" in n][-1] if any("0" in n for n in sheets) else list(sheets.keys())[0]
sheet1_name = [n for n in sheets if "1" in n][-1] if any("1" in n for n in sheets) else list(sheets.keys())[0]
sheet3_name = [n for n in sheets if "3" in n][-1] if any("3" in n for n in sheets) else list(sheets.keys())[-1]

df0 = sheets[sheet0_name]
df1 = sheets[sheet1_name]
df3 = sheets[sheet3_name]

st.markdown("---")
st.subheader("۱. انتخاب استان و شهر")

# انتخاب استان
province_opts = make_province_options(df0)
if not province_opts:
    st.error("ساختار شیت استان‌ها (Sheet0) نامعتبر است.")
    st.stop()

province_labels = [p[2] for p in province_opts]
selected_label = st.selectbox("انتخاب استان:", province_labels)
selected_index = province_labels.index(selected_label)
selected_province_code, selected_province_name, _ = province_opts[selected_index]

# انتخاب شهر
cities = find_cities_for_province(df1, selected_province_code, selected_province_name)
selected_city = st.selectbox("انتخاب شهر:", cities)

st.markdown("---")
st.subheader("۲. مشاهده جزئیات دیوار")

if st.button("نمایش جزئیات دیوار برای شهر انتخابی"):
    with st.spinner("در حال استخراج داده‌ها..."):
        result_df = extract_wall_details(df3, selected_city)
        if result_df.empty:
            st.warning("هیچ داده‌ای برای این شهر در Sheet3 پیدا نشد.")
        else:
            st.success("جزئیات دیوار پیدا شد ✅")
            st.dataframe(result_df, use_container_width=True)

            # ایجاد فایل خروجی برای دانلود
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
st.caption("🔹 نکته: هر زمان فایل materials.xlsx را به‌روزرسانی کنید، کافی است فایل جدید را جایگزین کنید. برنامه به‌صورت خودکار داده‌ها را از فایل جدید می‌خواند.")
