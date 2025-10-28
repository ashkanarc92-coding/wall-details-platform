# app.py
import streamlit as st
import pandas as pd
import io
import os
import re

st.set_page_config(page_title="Wall Detail Platform", layout="wide")

EXCEL_PATH = "materials.xlsx"

# ------------------- کمکی: بارگذاری امن شیت‌ها -------------------
@st.cache_data
def load_all_sheets(path):
    if not os.path.exists(path):
        raise FileNotFoundError("فایل '{}' پیدا نشد. لطفاً آن را در کنار app.py قرار دهید.".format(path))
    # read all sheets
    xls = pd.read_excel(path, sheet_name=None, engine="openpyxl", header=None, dtype=object)
    # تبدیل None به '' و str کردن محتوا (تا پردازش متن راحت‌تر شود)
    for k, df in xls.items():
        xls[k] = df.fillna("").astype(str)
    return xls

# ------------------- کمکی: پیدا کردن ستون کد/نام استان در Sheet0 -------------------
def detect_provinces(df_sheet0):
    """
    ورودی: df_sheet0 (بدون header؛ تمام سلول‌ها قبلا به str تبدیل شده)
    خروجی: لیستی از تاپل‌ها: [(province_code, province_name), ...]
    الگوریتم:
      - در تمام سلول‌ها به دنبال الگوی کد استان 'P-XX' باشیم.
      - برای هر خانه‌ای که کد استان دارد، در همان ردیف به دنبال نام (ستون بعدی یا ستون سمت راست) می‌گردیم.
      - اگر نام پیدا نشد، همان کد را هم اسم قرار می‌دهیم.
    """
    provinces = []
    pattern_p = re.compile(r"(?i)\bP-\d{2}\b")  # P-01, P-12 ...
    df = df_sheet0
    rows, cols = df.shape
    for i in range(rows):
        for j in range(cols):
            cell = df.iat[i, j].strip()
            if pattern_p.search(cell):
                code = pattern_p.search(cell).group(0).upper()
                # تلاش برای گرفتن نام: اول ستون بعدی در همان ردیف، سپس ستون بعدی و ... (تا 3 ستون)
                name = ""
                for k in range(1, 4):
                    if j + k < cols:
                        cand = df.iat[i, j + k].strip()
                        if cand and not pattern_p.search(cand):
                            name = cand
                            break
                if not name:
                    # fallback: نگاه به ستون سمت چپ
                    for k in range(1, 4):
                        if j - k >= 0:
                            cand = df.iat[i, j - k].strip()
                            if cand and not pattern_p.search(cand):
                                name = cand
                                break
                if not name:
                    name = code
                if (code, name) not in provinces:
                    provinces.append((code, name))
    # اگر هیچ کدی پیدا نشد، احتمالاً header سطر دارد: تلاش کنیم سطرها را اسکن برای ردیف‌هایی که شبیه "P-01" اند
    if not provinces:
        # فرض: ممکن است داده‌ها از ردیف 2 به بعد باشند؛ به دنبال سطرهایی باشیم که در هر ستون یکی از الگوها باشد
        flattened = df.values.ravel()
        for val in flattened:
            val = str(val).strip()
            if pattern_p.search(val):
                code = pattern_p.search(val).group(0).upper()
                provinces.append((code, code))
    return provinces

# ------------------- کمکی: پیدا کردن شهرها برای استان در Sheet1 -------------------
def detect_cities_for_province(df_sheet1, province_code):
    """
    جستجو برای یافتن شهرهای مربوط به استان انتخاب‌شده (بر اساس نام شهر واقعی)
    """
    pattern_city_code = re.compile(r"(?i)\bC-\d{2}-\d{2}\b")  # C-01-01
    pattern_province_code = re.compile(r"(?i)\bP-\d{2}\b")
    df = df_sheet1.copy()
    rows, cols = df.shape
    found = []

    for i in range(rows):
        row_vals = [str(df.iat[i, j]).strip() for j in range(cols)]
        joined = " | ".join(row_vals)

        # اگر کد استان در این ردیف وجود دارد
        if province_code.lower() in joined.lower():
            city_code = ""
            city_name = ""

            for j in range(cols):
                cell = row_vals[j]
                # پیدا کردن کد شهر
                if not city_code and pattern_city_code.search(cell):
                    city_code = pattern_city_code.search(cell).group(0).upper()
                # پیدا کردن نام شهر با حروف فارسی
                if not city_name and re.search(r"[\u0600-\u06FF]", cell):
                    # حذف مواردی که شامل "استان" یا شبیه "P-xx" هستند
                    if not pattern_province_code.search(cell) and "استان" not in cell:
                        city_name = cell.strip()

            # اگر نام شهر پیدا نشد ولی کد شهر هست
            if not city_name and city_code:
                city_name = city_code

            if city_name and (city_code, city_name) not in found:
                found.append((city_code if city_code else city_name, city_name))

    # اگر هنوز هیچ شهری پیدا نشد، کل فایل را برای کد شهر جستجو می‌کنیم
    if not found:
        for i in range(rows):
            for j in range(cols):
                cell = str(df.iat[i, j]).strip()
                if pattern_city_code.search(cell):
                    code = pattern_city_code.search(cell).group(0).upper()
                    found.append((code, code))

    # حذف تکراری‌ها (بر اساس نام شهر)
    unique = []
    seen = set()
    for code, name in found:
        if name not in seen:
            unique.append((code, name))
            seen.add(name)

    return unique
# ------------------- کمکی: استخراج دیتیل‌ها از Sheet3 -------------------
def extract_details_sheet3(df_sheet3, selected_province_code, selected_city_identifier):
    """
    جستجو در Sheet3 برای ردیف‌هایی که با selected_province_code و selected_city_identifier مطابقت دارند.
    الویت: در صورت وجود کد شهر (C-... ) با آن تطابق داده می‌شود؛ در غیر اینصورت نام شهر تطابق داده می‌شود.
    """
    df = df_sheet3.copy().astype(str)
    rows, cols = df.shape
    # اگر selected_city_identifier حاوی C- باشد، با آن جستجو کن
    city_is_code = bool(re.match(r"(?i)^C-\d{2}-\d{2}$", str(selected_city_identifier)))
    matched_rows = pd.Series([False]*rows)
    for i in range(rows):
        row_text = " | ".join([df.iat[i, j] for j in range(cols)])
        # نیاز است هر دو شرط استان و شهر را داشته باشیم؛ اما بعضی رکوردها شاید فقط شهر داشته باشند
        cond_city = (selected_city_identifier.strip().lower() in row_text.lower()) if selected_city_identifier else False
        cond_prov = (selected_province_code.strip().lower() in row_text.lower()) if selected_province_code else False
        # اگر city_is_code True، تشدید جستجو برای الگوی دقیق‌تر
        if city_is_code:
            if cond_city:
                matched_rows.iat[i] = True
        else:
            # اگر city name، سعی به ترکیب هر دو شرط (اگر استان نیز در ردیف باشد)
            if cond_city and cond_prov:
                matched_rows.iat[i] = True
            elif cond_city and not matched_rows.any():
                # اگر هیچ ردیفی پیدا نشده باشد، اجازه بده حداقل بر اساس شهر match شود
                matched_rows.iat[i] = True
    if matched_rows.any():
        res = df.loc[matched_rows.values, :].reset_index(drop=True)
        # نام ستون‌ها را به چیزی قابل نمایش تبدیل می‌کنیم (مثلاً Col1, Col2 یا اگر header اصلی در فایل وجود داشت استفاده می‌کنیم)
        # چون در read ما header=None گرفتیم، ستون‌ها اعداد هستند؛ برای نمایش بهتر از رشته "Column_#" استفاده می‌کنیم
        res.columns = [f"Column_{i+1}" for i in range(res.shape[1])]
        return res
    else:
        return pd.DataFrame()

# ------------------- UI -------------------
st.title("🧱 پلتفرم نمایش جزئیات دیوار — مبتنی بر materials.xlsx")
st.write("راهنما: ابتدا استان را از لیست انتخاب کنید، سپس شهرِ مربوطه را انتخاب کنید و در نهایت جزئیات دیوار نمایش داده می‌شود. در انتها می‌توانید خروجی Excel دانلود کنید.")

# try load
try:
    sheets = load_all_sheets(EXCEL_PATH)
except FileNotFoundError as fe:
    st.error(str(fe))
    st.stop()
except Exception as e:
    st.error("خطا در خواندن فایل اکسل: " + str(e))
    st.stop()

# انتخاب شیت‌ها (نام‌ها را نشان می‌دهیم تا اگر فایل شیت نام‌های متفاوتی داشت مشخص شود)
sheet_names = list(sheets.keys())
st.info("شیت‌های موجود در فایل: " + ", ".join(sheet_names))

# طبق چیزی که گفتی: sheet0, sheet1, sheet3 استفاده شوند؛ اگر وجود نداشتند، اولین‌ها را می‌گیریم
sheet0_key = next((k for k in sheet_names if k.lower().strip() == "sheet0"), sheet_names[0])
sheet1_key = next((k for k in sheet_names if k.lower().strip() == "sheet1"), sheet_names[1] if len(sheet_names)>1 else sheet_names[0])
sheet3_key = next((k for k in sheet_names if k.lower().strip() == "sheet3"), sheet_names[-1])

df0 = sheets[sheet0_key]
df1 = sheets[sheet1_key]
df3 = sheets[sheet3_key]

st.write("نمایش بخشی از Sheet0 (استان‌ها) — چند سطر اول:")
st.dataframe(df0.head(8))

# استخراج استان‌ها
provs = detect_provinces(df0)
if not provs:
    st.error("ناتوانی در استخراج لیست استان‌ها از Sheet0. لطفاً مطمئن شوید کدهای استان (P-xx) در شیت موجودند.")
    st.stop()

# نمایش لیست استان‌ها در selectbox (نمایش نام به کاربر)
prov_labels = [f"{c} — {n}" if c!=n else c for c,n in provs]
sel_idx = st.selectbox("انتخاب استان (از منوی کشویی):", range(len(prov_labels)), format_func=lambda i: prov_labels[i])
selected_province_code, selected_province_name = provs[sel_idx][0], provs[sel_idx][1]

st.write(f"استان انتخاب‌شده: {selected_province_code} — {selected_province_name}")

# حالا استخراج شهرها از Sheet1 بر اساس selected_province_code
st.write("نمایش بخشی از Sheet1 (شهرها) — چند سطر اول:")
st.dataframe(df1.head(8))

cities = detect_cities_for_province(df1, selected_province_code)
if not cities:
    st.warning("برای این استان، نتوانستم شهرها را از Sheet1 استخراج کنم. تمام مقادیر Sheet1 را نمایش می‌دهم تا بررسی کنی.")
    st.dataframe(df1.head(30))
    st.stop()

# cities is list of tuples (identifier, label)
city_labels = [lab for ident, lab in cities]
city_idx = st.selectbox("انتخاب شهر (از منوی کشویی):", range(len(city_labels)), format_func=lambda i: city_labels[i])
selected_city_identifier = cities[city_idx][0]
selected_city_label = cities[city_idx][1]

st.write(f"شهر انتخاب‌شده: {selected_city_label} (identifier: {selected_city_identifier})")

st.markdown("---")
if st.button("نمایش جزئیات دیوار برای شهر انتخابی"):
    with st.spinner("در حال جستجو در Sheet3 برای دیتیل‌های مربوطه..."):
        res = extract_details_sheet3(df3, selected_province_code, selected_city_identifier)
        if res.empty:
            st.warning("هیچ ردیفی در Sheet3 مطابق استان و شهر انتخابی پیدا نشد. لطفاً مطمئن شوید کدهای 'P-..' و 'C-..' یا نام شهر دقیقاً در Sheet3 موجود است.")
            # برای کمک، چند سطر از Sheet3 را نشان بده
            st.write("نمونه‌ای از چند سطر Sheet3 برای بررسی:")
            st.dataframe(df3.head(30))
        else:
            st.success(f"{len(res)} ردیف یافته شد. نمایش جدول زیر:")
            st.dataframe(res, use_container_width=True)

            # دکمه دانلود
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                res.to_excel(writer, index=False, sheet_name="Wall_Details")
            buf.seek(0)
            st.download_button("📥 دانلود خروجی (Excel)", data=buf,
                               file_name=f"Wall_Details_{selected_city_identifier}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

