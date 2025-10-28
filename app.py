# app.py
import streamlit as st
import pandas as pd
import io
import os
import re

st.set_page_config(page_title="Wall Detail Platform", layout="wide")
EXCEL_PATH = "materials.xlsx"

# ---------- helpers ----------
@st.cache_data
def load_all_sheets(path):
    if not os.path.exists(path):
        raise FileNotFoundError(f"فایل '{path}' پیدا نشد. لطفاً کنار app.py قرار دهید.")
    # header=None because your file may have irregular headers
    xls = pd.read_excel(path, sheet_name=None, engine="openpyxl", header=None, dtype=object)
    # fillna and cast to str for safe text ops
    for k, df in xls.items():
        xls[k] = df.fillna("").astype(str)
    return xls

def is_persian_text(s):
    return bool(re.search(r"[\u0600-\u06FF]", s))

def detect_provinces(df0):
    """Detect province (P-xx) codes and names (best-effort)."""
    pat_p = re.compile(r"(?i)\bP-\d{2}\b")
    provinces = []
    rows, cols = df0.shape
    for i in range(rows):
        for j in range(cols):
            cell = str(df0.iat[i,j]).strip()
            if cell and pat_p.search(cell):
                code = pat_p.search(cell).group(0).upper()
                # prefer right-adjacent cell as name, fallback left
                name = ""
                if j+1 < cols:
                    cand = str(df0.iat[i,j+1]).strip()
                    if cand and not pat_p.search(cand):
                        name = cand
                if not name and j-1 >= 0:
                    cand = str(df0.iat[i,j-1]).strip()
                    if cand and not pat_p.search(cand):
                        name = cand
                if not name:
                    name = code
                if (code, name) not in provinces:
                    provinces.append((code, name))
    # fallback: if nothing found, try scanning first column for non-empty unique values
    if not provinces:
        col0 = df0.iloc[:,0].astype(str).str.strip().replace("", pd.NA).dropna().unique().tolist()
        for v in col0:
            provinces.append((v, v))
    return provinces

def detect_city_columns(df1):
    """
    Return (prov_col_idx, city_code_col_idx, city_name_col_idx)
    Strategy:
      - examine first 6 columns (or all if fewer)
      - for each col compute:
         * count of values matching P- pattern
         * count of values matching C- pattern
         * count of Persian-containing values
         * proportion numeric-like values
      - choose prov_col as col with most P- matches (fallback col0)
      - choose city_code_col as col with most C- matches (fallback col2)
      - choose city_name_col as col with high Persian count and low numeric proportion and not energy-keyword
    """
    pat_p = re.compile(r"(?i)\bP-\d{2}\b")
    pat_c = re.compile(r"(?i)\bC-\d{2}-\d{2}\b")
    df = df1.copy()
    cols = list(df.columns)
    ncols = min(len(cols), 8)
    stats = []
    energy_keywords = ["نیاز", "انرژی", "سرمایش", "گرمایش", "kwh", "kW", "مصرف"]
    for idx in range(ncols):
        col = cols[idx]
        vals = df[col].astype(str).str.strip()
        total = len(vals)
        p_count = vals.str.contains(pat_p, case=False, na=False).sum()
        c_count = vals.str.contains(pat_c, case=False, na=False).sum()
        persian_count = vals.apply(lambda s: bool(re.search(r"[\u0600-\u06FF]", s))).sum()
        numeric_like = vals.apply(lambda s: bool(re.match(r"^[\d\.\,\- ]+$", s))).sum()
        header_text = str(col)
        has_energy_kw = any(kw in header_text for kw in energy_keywords) or vals.apply(lambda s: any(kw in s for kw in energy_keywords)).any()
        stats.append({
            "idx": idx,
            "col": col,
            "p_count": int(p_count),
            "c_count": int(c_count),
            "persian_count": int(persian_count),
            "numeric_count": int(numeric_like),
            "has_energy_kw": bool(has_energy_kw)
        })
    # pick prov_col: max p_count
    prov_col_idx = max(stats, key=lambda x: x["p_count"])["idx"] if any(s["p_count"]>0 for s in stats) else 0
    # pick city_code: max c_count
    city_code_idx = max(stats, key=lambda x: x["c_count"])["idx"] if any(s["c_count"]>0 for s in stats) else (2 if len(cols)>2 else 0)
    # pick city_name: prefer col with high persian_count and low numeric_count and not energy
    candidates = [s for s in stats if not s["has_energy_kw"]]
    if not candidates:
        candidates = stats
    # score = persian_count - numeric_count
    city_name_idx = max(candidates, key=lambda x: (x["persian_count"] - x["numeric_count"]))["idx"]
    # Final safety fallback
    if city_name_idx == prov_col_idx:
        city_name_idx = city_code_idx if city_code_idx != prov_col_idx else (prov_col_idx+1 if prov_col_idx+1 < len(cols) else prov_col_idx)
    return prov_col_idx, city_code_idx, city_name_idx, stats

def detect_cities_for_province(df1, province_code):
    # detect candidate columns
    prov_idx, ccode_idx, cname_idx, stats = detect_city_columns(df1)
    cols = list(df1.columns)
    prov_col, ccode_col, cname_col = cols[prov_idx], cols[ccode_idx], cols[cname_idx]

    # debug info
    st.write("DEBUG: تشخیص ستون‌ها در Sheet1:")
    st.write(pd.DataFrame(stats).set_index("idx"))

    st.write(f"DEBUG: انتخاب شده -> prov_col idx:{prov_idx} ({prov_col}), city_code idx:{ccode_idx} ({ccode_col}), city_name idx:{cname_idx} ({cname_col})")

    # filter rows where province_code matches either in province code col or province name col (case-insensitive)
    df = df1.copy()
    df = df.replace("", pd.NA).dropna(how="all")
    def row_has_prov(r):
        for c in [prov_col]:
            try:
                if str(r[c]).strip() and province_code.strip().lower() in str(r[c]).strip().lower():
                    return True
            except Exception:
                continue
        # also check entire row as fallback
        joined = " | ".join([str(r[c]).strip() for c in cols])
        if province_code.strip().lower() in joined.lower():
            return True
        return False

    filtered = df[df.apply(row_has_prov, axis=1)]
    st.write(f"DEBUG: تعداد ردیف‌های فیلتر شده بر اساس استان ({province_code}) = {len(filtered)}")
    if filtered.empty:
        # fallback: if province_code not found, try matching by province name if code could be name
        filtered = df[df[prov_col].astype(str).str.contains(province_code, case=False, na=False)]
    # Now extract cities
    cities = []
    for _, row in filtered.iterrows():
        code = str(row[ccode_col]).strip()
        name = str(row[cname_col]).strip()
        # ignore if name looks like energy header or numeric-only
        if not name or name.lower() in ["nan", "none"] or re.match(r"^[\d\.\,\- ]+$", name):
            name = code
        if not name:
            continue
        cities.append((code, name))
    # unique preserve order
    seen = set()
    uniq = []
    for code, name in cities:
        if name not in seen:
            seen.add(name)
            uniq.append((code, name))
    return uniq

def extract_details_sheet3(df3, prov_code, city_identifier):
    df = df3.copy().astype(str)
    rows, cols = df.shape
    pat_c = re.compile(r"(?i)^C-\d{2}-\d{2}$")
    city_is_code = bool(pat_c.match(str(city_identifier)))
    matched = []
    for i in range(rows):
        row_text = " | ".join([df.iat[i,j] for j in range(cols)])
        cond_city = city_identifier.strip().lower() in row_text.lower()
        cond_prov = prov_code.strip().lower() in row_text.lower()
        if city_is_code:
            if cond_city:
                matched.append(i)
        else:
            if cond_city and cond_prov:
                matched.append(i)
            elif cond_city and not matched:
                matched.append(i)
    if not matched:
        return pd.DataFrame()
    res = df.iloc[matched].reset_index(drop=True)
    res.columns = [f"Column_{i+1}" for i in range(res.shape[1])]
    return res

# ---------- UI ----------
st.title("🧱 Wall Detail Platform — Debug-enabled")
st.write("راهنما: ابتدا استان را انتخاب کنید؛ سپس شهرِ همان استان نمایش داده می‌شود. (اگر شهرها نمایش داده نشدند بخش DEBUG را ببین)")

# load
try:
    sheets = load_all_sheets(EXCEL_PATH)
except Exception as e:
    st.error("خطا در خواندن فایل اکسل: " + str(e))
    st.stop()

st.write("شیت‌های موجود: " + ", ".join(list(sheets.keys())))
# choose keys (prefer sheet0/1/3 names)
keys = list(sheets.keys())
sheet0 = next((k for k in keys if k.lower().strip()=="sheet0"), keys[0])
sheet1 = next((k for k in keys if k.lower().strip()=="sheet1"), keys[1] if len(keys)>1 else keys[0])
sheet3 = next((k for k in keys if k.lower().strip()=="sheet3"), keys[-1])
df0 = sheets[sheet0]
df1 = sheets[sheet1]
df3 = sheets[sheet3]

st.write("Preview Sheet0 (اولی):")
st.dataframe(df0.head(8))
st.write("Preview Sheet1 (شهرها):")
st.dataframe(df1.head(8))

# provinces
provs = detect_provinces(df0)
st.write(f"DEBUG: استخراج استان‌ها: {len(provs)} مورد")
st.write(provs)
if not provs:
    st.error("هیچ استانی پیدا نشد در Sheet0.")
    st.stop()

prov_labels = [name for code, name in provs]
prov_idx = st.selectbox("انتخاب استان:", range(len(prov_labels)), format_func=lambda i: prov_labels[i])
selected_prov_code, selected_prov_name = provs[prov_idx][0], provs[prov_idx][1]
st.write(f"استان انتخاب‌شده: {selected_prov_code} — {selected_prov_name}")

# cities
cities = detect_cities_for_province(df1, selected_prov_code)
st.write(f"DEBUG: تعداد شهرهای استخراج‌شده = {len(cities)}")
st.write("DEBUG: نمونه شهرها (code, name):")
st.write(cities[:40])
if not cities:
    st.warning("هیچ شهری برای این استان استخراج نشد. لطفاً خروجی DEBUG را اینجا پیست کن.")
    st.stop()

city_labels = [name for code, name in cities]
city_idx = st.selectbox("انتخاب شهر:", range(len(city_labels)), format_func=lambda i: city_labels[i])
selected_city_code, selected_city_name = cities[city_idx][0], cities[city_idx][1]
st.write(f"شهر انتخاب‌شده: {selected_city_code} — {selected_city_name}")

# show details
if st.button("نمایش دیتیل دیوار"):
    with st.spinner("در حال جستجو..."):
        res = extract_details_sheet3(df3, selected_prov_code, selected_city_code or selected_city_name)
        if res.empty:
            st.warning("هیچ ردیفی در Sheet3 مطابق استان/شهر انتخابی یافت نشد.")
            st.write("نمونه‌ای از چند سطر Sheet3 برای بررسی:")
            st.dataframe(df3.head(30))
        else:
            st.success(f"{len(res)} ردیف یافت شد.")
            st.dataframe(res, use_container_width=True)
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                res.to_excel(writer, index=False, sheet_name="Wall_Details")
            buf.seek(0)
            st.download_button("دانلود خروجی (Excel)", data=buf, file_name=f"Wall_Details_{selected_city_name}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
