import io
import re
import zipfile
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="CRS Gain Validator",
    page_icon="📊",
    layout="wide"
)

st.title("📊 CRS Gain Validator")
st.subheader("ระบบตรวจเช็คและแก้ไข CRS Gain, pdschTypeBGain อัตโนมัติ (รองรับดาวน์โหลดแยกไฟล์ 🚀)")

# ==================================================
# 1. SESSION STATE FOR DYNAMIC WEEKS
# ==================================================
if "num_weeks" not in st.session_state:
    st.session_state.num_weeks = 2

# ==================================================
# 2. VALIDATION RULES & CACHED HELPER FUNCTIONS
# ==================================================
VALID_PAIRS = {
    300: 1,
    477: 2,
    600: 3,
    -300: 1,
    177: 1,
    0: 0
}

def clean_col_name(name: str) -> str:
    if pd.isna(name):
        return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(name)).lower()

def parse_int_series(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors='coerce').astype('Int64')

# ช่วยแปลง DataFrame เป็น Bytes สำหรับปุ่มดาวน์โหลดไฟล์เดี่ยว
@st.cache_data(show_spinner=False)
def convert_df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, header=False)
    return output.getvalue()

@st.cache_data(show_spinner=False)
def process_excel_cached(file_bytes: bytes, filename: str):
    df_full = pd.read_excel(io.BytesIO(file_bytes), header=None)
    
    if len(df_full) <= 3:
        raise ValueError("ไฟล์มีข้อมูลไม่ถึง Row 4")

    header_row = df_full.iloc[3]
    crs_col_idx = None
    pdsch_col_idx = None

    for idx, col in enumerate(header_row):
        clean_name = clean_col_name(col)
        if clean_name == "crsgain":
            crs_col_idx = idx
        elif clean_name == "pdschtypebgain":
            pdsch_col_idx = idx

    if crs_col_idx is None or pdsch_col_idx is None:
        raise ValueError("ไม่พบ 'crsGain' หรือ 'pdschTypeBGain' ใน Row 4")

    crs_series = parse_int_series(df_full.iloc[4:, crs_col_idx])
    pdsch_series = parse_int_series(df_full.iloc[4:, pdsch_col_idx])

    expected_pdsch = crs_series.map(VALID_PAIRS)
    valid_mask = crs_series.isin(VALID_PAIRS) & (expected_pdsch == pdsch_series)
    non_null_mask = crs_series.notna() & pdsch_series.notna()
    
    results_series = pd.Series("", index=crs_series.index)
    results_series[non_null_mask] = valid_mask[non_null_mask].map({True: 't', False: 'f'})

    t_count = (results_series == "t").sum()
    f_count = (results_series == "f").sum()

    # สร้าง DataFrame Checked
    df_checked = df_full.copy()
    res_col = pd.Series("", index=df_checked.index)
    res_col.iloc[3] = "result"
    res_col.iloc[4:] = results_series.values
    df_checked[len(df_checked.columns)] = res_col

    # สร้าง DataFrame Fixed
    df_fixed = df_full.copy()
    fixed_pdsch = pdsch_series.copy()
    fix_mask = crs_series.isin(VALID_PAIRS) & crs_series.notna()
    fixed_pdsch[fix_mask] = expected_pdsch[fix_mask]
    df_fixed.iloc[4:, pdsch_col_idx] = fixed_pdsch.values

    return {
        "filename": filename,
        "checked": df_checked,
        "fixed": df_fixed,
        "t": t_count,
        "f": f_count
    }

def convert_dfs_to_zip(w_results) -> bytes:
    zip_b = io.BytesIO()
    with zipfile.ZipFile(zip_b, "w", zipfile.ZIP_DEFLATED) as z:
        for r in w_results:
            c_bytes = convert_df_to_excel_bytes(r["checked"])
            z.writestr(f"checked_{r['filename']}", c_bytes)

            f_bytes = convert_df_to_excel_bytes(r["fixed"])
            z.writestr(f"fixed_{r['filename']}", f_bytes)
            
    return zip_b.getvalue()

# ==================================================
# 3. CONTROLS FOR ADDING/REMOVING WEEKS
# ==================================================
st.markdown("### 🛠️ จัดการชุดไฟล์ (Week)")
ctrl_col1, ctrl_col2, _ = st.columns([1, 1, 3])

with ctrl_col1:
    if st.button("➕ เพิ่มชุดไฟล์ (Week)", use_container_width=True):
        if st.session_state.num_weeks < 50:
            st.session_state.num_weeks += 1
            st.rerun()
        else:
            st.warning("จำกัดสูงสุด 50 ชุด")

with ctrl_col2:
    if st.button("➖ ลดชุดไฟล์", use_container_width=True):
        if st.session_state.num_weeks > 1:
            st.session_state.num_weeks -= 1
            st.rerun()

st.caption(f"จำนวนชุดอัปโหลดปัจจุบัน: **{st.session_state.num_weeks}** ชุด")
st.divider()

# ==================================================
# 4. DYNAMIC UPLOAD SECTIONS
# ==================================================
weeks_data_list = []
cols_per_row = 3

for i in range(0, st.session_state.num_weeks, cols_per_row):
    cols = st.columns(cols_per_row)
    for j in range(cols_per_row):
        week_idx = i + j
        if week_idx < st.session_state.num_weeks:
            with cols[j]:
                st.subheader(f"📅 ชุดที่ {week_idx + 1}")
                w_name = st.text_input(
                    f"ชื่อ Week {week_idx + 1}:",
                    value=f"2026{16 + week_idx}",
                    key=f"w_name_{week_idx}"
                )
                up_files = st.file_uploader(
                    f"📂 ลากไฟล์ Excel (ชุดที่ {week_idx + 1})",
                    type=["xlsx"],
                    accept_multiple_files=True,
                    key=f"uploader_{week_idx}"
                )
                weeks_data_list.append({
                    "name": w_name,
                    "files": up_files if up_files else []
                })

# ==================================================
# 5. PROCESS ALL WEEKS
# ==================================================
chart_rows = []
all_weeks_results = {}

has_files = any(len(w["files"]) > 0 for w in weeks_data_list)

if has_files:
    with st.spinner("⏳ กำลังประมวลผลไฟล์..."):
        for w_data in weeks_data_list:
            w_name = w_data["name"]
            files = w_data["files"]

            if files:
                w_results = []
                for uploaded_file in files[:30]:
                    try:
                        file_bytes = uploaded_file.getvalue()
                        res = process_excel_cached(file_bytes, uploaded_file.name)
                        w_results.append(res)
                    except Exception as e:
                        st.error(f"❌ [{w_name}] {uploaded_file.name}: {str(e)}")

                if w_results:
                    all_weeks_results[w_name] = w_results
                    tot_t = sum(r["t"] for r in w_results)
                    tot_f = sum(r["f"] for r in w_results)
                    chart_rows.append({
                        "Week": w_name,
                        "f - Incorrect": tot_f,
                        "t - Correct": tot_t
                    })

# ==================================================
# 6. COMBINED CHART
# ==================================================
if chart_rows:
    st.divider()
    st.header("📈 กราฟเปรียบเทียบผลรวมทุก Week")

    combined_df = pd.DataFrame(chart_rows).set_index("Week")

    st.bar_chart(
        combined_df,
        color=["#E05252", "#2878D0"],
        height=500,
        stack=True
    )

    st.dataframe(combined_df.reset_index(), use_container_width=True, hide_index=True)

    # ==================================================
    # 7. DOWNLOAD SECTION (ZIP + SEPARATE FILES)
    # ==================================================
    st.divider()
    st.header("📥 ดาวน์โหลดไฟล์ Excel")

    tab_titles = [f"📦 {w_name}" for w_name in all_weeks_results.keys()]
    if tab_titles:
        tabs = st.tabs(tab_titles)
        for tab, (w_name, w_results) in zip(tabs, all_weeks_results.items()):
            with tab:
                # 1. ปุ่มดาวน์โหลด ZIP รวมของ Week นั้นๆ
                zip_data = convert_dfs_to_
