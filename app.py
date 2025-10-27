# app.py
import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Wall Details Explorer", layout="wide")

EXCEL_PATH = "materials.xlsx"  # <-- فایل اکسل تو باید کنار همین app.py باشه

@st.cache_data
def load_sheets(path):
    """بارگزاری تمام شیت‌ها به عنوان DataFrame در یک دیکشنری."""
    try:
        xls = pd.read_excel(path, sheet_name=None, engine="openpyxl")
        # پاکسازی نام ستون‌ها (ترایم)
        for name, df in xls.items():
            df.columns = df.columns.map(lambda c: str(c).strip() if not pd.isna(c) else c)
        return xls
    except FileNotFoundError:
        st.error(f"فایل پیدا نشد: {path}. لطفاً مطمئن شوید فایل در همان پوشه‌ی پروژه قرار دارد.")
        return {}
    except Exception as e:
        st.error(f"خطا در بارگزاری اکسل: {e}")
        return {}

def guess_code_name_columns(df):
    """
    سعی می‌کنیم ستون‌های کد و نام را حدس بزنیم:
    - اگر دیتافریم دو ستون یا بیشتر داشت، اولین ستون را 'code' و دومین را 'name' می‌گیریم.
    - در غیر اینصورت ستون اول را هر دو می‌گیریم.
    """
    cols = list(df.columns)
    if len(cols) >= 2:
        return cols[0], cols[1]
    elif len(cols) == 1:
        return cols[0], cols[0]
    else:
        return None, None

def make_province_options(df_sheet0):
    """از Sheet0 لیست استان‌ها را استخراج می‌کند — خروجی: list of (code, name, label)"""
    code_col, name_col = guess_code_name_columns(df_sheet0)
    if code_col is None:
        return []
    # سطرهایی که خالی نیستند
    rows = df_sheet0[[code_col, name_col]].dropna(how="all")
    opts = []
    for _, r in rows.iterrows():
        code = str(r[code_col]).strip() if not pd.isna(r[code_col]) else ""
        name = str(r[name_col]).strip() if not pd.isna(r[name_col]) else ""
        label = code if name == "" else f"{code} — {name}"
        if label not in [o[2] for o in opts]:
            opts.append((code, name, label))
    return opts

def find_cities_for_province(df_sheet1, selected_province_code, selected_province_name):
    """سعی می‌کند در Sheet1 ردیف‌های مربوط به استان انتخابی را پیدا کند و لیست شهرها را برگرداند."""
    # حدس ستون‌های کد استان و نام شهر
    cols = list(df_sheet1.columns)
    # اگر ستون‌هایی با 'province' ، 'code' یا 'P-' وجود دارد تلاش می‌کنیم آنها را پیدا کنیم
    df = df_sheet1.copy()
    # ساده‌ترین مسیر: بررسی هر ستون برای وجود selected_province_code
    if selected_province_code:
        for col in cols:
            try:
                matches = df[col].astype(str).str.strip().str.lower() == str(selected_province_code).strip().lower()
                if matches.any():
                    # فرض: ستون نام شهر، ستون بعدی یا یکی از ستون‌های دیگر است
                    # سعی می‌کنیم ستون نام شهر را حدس بزنیم:
                    # اگر ستون بعدی وجود دارد از آن استفاده کن، وگرنه پرش کن به اولین ستون غیرآن
                    candidate_name_col = None
                    col_idx = cols.index(col)
                    if col_idx + 1 < len(cols):
                        candidate_name_col = cols[col_idx + 1]
                    else:
                        # fallback: هر ستونی که غیرِ ستون کد باشد و حاوی نام‌ها باشد
                        for c in cols:
                            if c != col:
                                candidate_name_col = c
                                break
                    if candidate_name_col:
                        city_series = df.loc[matches, candidate_name_col].dropna().astype(str).str.strip().unique().tolist()
                        # city codes: اگر خودش هم کد شهر دارد (مثلاً ستون دیگر)، include codes
                        return city_series
            except Exception:
                continue

    # اگر با کد موفق نشدیم، تلاش با نام استان:
    if selected_province_name:
        for col in cols:
            try:
                matches = df[col].astype(str).str.strip().str.lower() == str(selected_province_name).strip().lower()
                if matches.any():
                    # candidate name col as above
                    col_idx = cols.index(col)
                    candidate_name_col = cols[col_idx + 1] if col_idx + 1 < len(cols) else None
                    if candidate_name_col is None:
                        for c in cols:
                            if c != col:
                                candidate_name_col = c
                                break
                    if candidate_name_col:
                        city_series = df.loc[matches, candidate_name_col].dropna().astype(str).str.strip().unique().tolist()
                        return city_series
            except Exception:
                continue

    # fallback عمومی: اگر ستونِ مشخصی برای 'city' یا 'name' وجود داشت از آن استفاده کن
    for keyword in ["city", "town", "name", "shahr", "شهر"]:
        for c in cols:
            if keyword in str(c).lower():
                return df[c].dropna().astype(str).str.strip().unique().tolist()

    # اگر هیچکدام نشد، لیست یکتا از تمام سلول‌های دیتافریم (compact) را برگردان
    flattened = pd.Series(df.values.ravel()).dropna().astype(str).str.strip().unique().tolist()
    return flattened[:200]  # محدود به 200 آیتم برای نمایش

def extract_wall_details(df_sheet3, selected_city):
    """در Sheet3 سعی می‌کنیم ردیف/ردیف‌هایی که مرتبط با شهر انتخابی‌اند را پیدا کنیم و آنها را برگردانیم."""
    if selected_city is None or selected_city == "":
        return pd.DataFrame()

    df = df_sheet3.copy().astype(object)
    # جستجو در کل دیتافریم برای مقدار city (case-insensitive)
    mask = df.applymap(lambda x: str(x).strip().lower() if not pd.isna(x) else "").applymap(lambda s: selected_city.strip().lower() in s if s else False)
    # هر ردیفی که هر ستونی True داشت انتخاب می‌شود
    rows_with_city = mask.any(axis=1)
    if rows_with_city.any():
        return df.loc[rows_with_city, :]

    # اگر نتیجه‌ای نبود، تلاش می‌کنیم ستون‌هایی که نامشان حاوی 'city' یا 'code' است را بررسی کنیم
    for col in df.columns:
        if any(k in str(col).lower() for k in ["city", "shahr", "code", "کد"]):
            matches = df[col].astype(str).str.strip().str.lower() == selected_city.strip().lower()
            if matches.any():
                return df.loc[matches, :]

    # fallback: هیچ پیدا نشد — کاربر را راهنمایی می‌کنیم
    return pd.DataFrame()

# ---------- UI ----------
st.title("پلتفرم نمایش دیتیل‌های دیوار — (Streamlit)")
st.write("این اپ از فایل `materials.xlsx` که کنار همین برنامه قرار دارد داده‌ها را می‌خواند. ابتدا استان را انتخاب کنید، سپس شهر و در نهایت جزئیات مربوطه نمایش داده می‌شود.")

sheets = load_sheets(EXCEL_PATH)
if not sheets:
    st.stop()

# بررسی وجود شیت‌های مورد نیاز
if not any(n.lower() in ("sheet0", "sheet 0", "sheet0".lower()) for n in sheets.keys()):
    # ولی از نام‌های دیگر هم ممکن است استفاده شود؛ بهتر است نام شیت‌ها را نشان دهیم
    st.write("شیت‌های موجود در فایل:")
    st.write(list(sheets.keys()))

# انتخاب شیت‌ها (اگر نام‌های دقیق متفاوت است کاربر می‌تواند انتخاب کند)
sheet0_name = None
sheet1_name = None
sheet3_name = None

# تلاش برای پیدا کردن شیت بر اساس اسم‌های متعارف
for name in sheets.keys():
    low = name.lower()
    if "sheet0" in low or "province" in low or "استان" in low:
        sheet0_name = name
    if "sheet1" in low or "city" in low or "شهر" in low:
        sheet1_name = name
    if "sheet3" in low or "detail" in low or "جدول" in low or "دیوار" in low:
        sheet3_name = name

# اگر هر کدام پیدا نشد، می‌گذاریم کاربر خودش انتخاب کند
col1, col2 = st.columns([1, 1])
with col1:
    sheet0_name = st.selectbox("شیت استان‌ها (Sheet0) — اگر تشخیص خودکار اشتباه بود انتخاب کنید:", options=list(sheets.keys()), index=list(sheets.keys()).index(sheet0_name) if sheet0_name in sheets else 0)
with col2:
    sheet1_name = st.selectbox("شیت شهرها (Sheet1) — اگر تشخیص خودکار اشتباه بود انتخاب کنید:", options=list(sheets.keys()), index=list(sheets.keys()).index(sheet1_name) if sheet1_name in sheets else 0)

# گزینه شیت جزئیات
sheet3_name = st.selectbox("شیت جزئیات دیوارها (Sheet3):", options=list(sheets.keys()), index=list(sheets.keys()).index(sheet3_name) if sheet3_name in sheets else 0)

df0 = sheets[sheet0_name]
df1 = sheets[sheet1_name]
df3 = sheets[sheet3_name]

st.markdown("---")
st.subheader("انتخاب استان و شهر")
# ساخت گزینه‌های استان
province_opts = make_province_options(df0)
if len(province_opts) == 0:
    st.error("نمی‌توانم استان‌ها را از Sheet0 استخراج کنم. لطفاً ساختار Sheet0 را بررسی کنید.")
    st.stop()

province_labels = [p[2] for p in province_opts]
selected_label = st.selectbox("استان:", province_labels)
selected_index = province_labels.index(selected_label)
selected_province_code, selected_province_name, _ = province_opts[selected_index]

# حالا لیست شهرها را بر اساس Sheet1 پیدا کن
city_list = find_cities_for_province(df1, selected_province_code, selected_province_name)
# اگر خیلی طولانی است کوتاهش کن
if len(city_list) > 400:
    city_list = city_list[:400]

selected_city = st.selectbox("شهر (یا کد شهر):", options=city_list)

st.markdown("---")
if st.button("نمایش جزئیات دیوار برای شهر انتخابی"):
    with st.spinner("در حال استخراج..."):
        result_df = extract_wall_details(df3, selected_city)
        if result_df.empty:
            st.warning("هیچ ردیفی در شیت جزئیات (Sheet3) مطابق شهر/کد انتخابی پیدا نشد.")
            st.write("لطفاً موارد زیر را بررسی کنید:")
            st.write("- آیا در Sheet3 یک ستون مخصوص کد یا نام شهر وجود دارد؟ (مثلاً C-01-01 یا نام شهر)")
            st.write("- آیا مقادیر دقیقاً مطابقت دارند؟ (حسّاسیت: این حالت موردی را نادیده می‌گیرد)")
            st.write("- در صورت نیاز می‌توانید ردیف/ستون موردنظر را دستی انتخاب کنید (امکان افزودن این قابلیت وجود دارد).")
        else:
            st.success(f"{len(result_df)} سطر پیدا شد. (نمایش جدول)")
            st.dataframe(result_df)

            # دکمه دانلود خروجی به اکسل
            to_save = result_df.copy()
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                to_save.to_excel(writer, index=False, sheet_name="Wall_Details")
                writer.save()
            buf.seek(0)
            st.download_button(label="دانلود خروجی (Excel)", data=buf, file_name=f"wall_details_{selected_city}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.markdown("---")
st.caption("راهنما: اگر ساختار فایل شما خیلی خاص است (ردیف‌های ثابت یا ستون‌های با نام‌های فارسی/خاص)، من می‌توانم بر اساس مثال واقعی شیت‌ها کد را دقیق‌تر تنظیم کنم تا فیلترها ۱۰۰٪ صحیح کار کنند.")
