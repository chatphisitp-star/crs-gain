import io
import re
import zipfile
import pandas as pd
import streamlit as st
import altair as alt
from openpyxl.styles import PatternFill

st.set_page_config(
    page_title="CRS Gain Validator",
    page_icon="⚡",
    layout="wide"
)

# ==================================================
# 1. SIDEBAR CONFIGURATION
# ==================================================
st.sidebar.header("⚙️ การตั้งค่าระบบ")

with st.sidebar.expander("📝 ปรับแต่งข้อความหัวข้อ (Headers)", expanded=False):
    app_title = st.text_input("หัวข้อหลัก (Title):", value=" CRS Gain Validator ")
    app_subtitle = st.text_input(
        "หัวข้อย่อย (Subheader):", 
        value="ระบบตรวจเช็คและแก้ไข CRS Gain, pdschTypeBGain อัตโนมัติ"
    )
    section_manage_title = st.text_input("หัวข้อส่วนจัดการชุดไฟล์:", value="🛠️ จัดการชุดไฟล์ (Week)")
    section_chart_title = st.text_input("หัวข้อส่วนแสดงกราฟ:", value="CRS Gain, pdschTypeBGain")
    section_download_title = st.text_input("หัวข้อส่วนดาวน์โหลด:", value=" ดาวน์โหลดไฟล์ Excel")

with st.sidebar.expander("📄 ตั้งชื่อไฟล์ Export", expanded=False):
    prefix_checked = st.text_input("คำนำหน้าไฟล์ Checked:", value="checked_")
    prefix_fixed = st.text_input("คำนำหน้าไฟล์ Fixed:", value="fixed_")

with st.sidebar.expander("🎨 ปรับแต่งสีระบบ", expanded=True):
    st.subheader("1. สีในไฟล์ Excel")
    highlight_excel = st.checkbox("เปิดใช้งานไฮไลต์สีใน Excel", value=True)
    color_fail_excel = st.color_picker("สีไฮไลต์ Error 'f' (Checked File)", "#FFC7CE")
    color_fix_excel = st.color_picker("สีไฮไลต์จุดที่แก้ไข (Fixed File)", "#C6EFCE")

    st.subheader("2. สีกราฟแท่ง")
    color_incorrect_chart = st.color_picker("สีกราฟ Incorrect (f)", "#E05252")
    color_correct_chart = st.color_picker("สีกราฟ Correct (t)", "#2878D0")

st.title(app_title)
st.subheader(app_subtitle)

# ==================================================
# 2. STATE MANAGEMENT (บันทึกสถานะล่าสุดตลอด)
# ==================================================
if "num_weeks" not in st.session_state:
    st.session_state.num_weeks = 2

# ฟังก์ชันอัปเดตจำนวนชุดพร้อมจำค่าล่าสุด
def update_weeks(delta: int):
    new_val = st.session_state.num_weeks + delta
    st.session_state.num_weeks = max(1, min(50, new_val))

def set_weeks_direct():
    st.session_state.num_weeks = st.session_state.direct_num_input

# ==================================================
# 3. FAST VALIDATION ENGINE (DYNAMIC HEADER SEARCH)
# ==================================================
VALID_PAIRS = {
    300: 1,
    477: 2,
    600: 3,
    -300: 1,
    177: 1,
    0: 0
}

VALID_PAIRS_SERIES = pd.Series(VALID_PAIRS)

def clean_col_name(name: str) -> str:
    if pd.isna(name):
        return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(name)).lower()

@st.cache_data(show_spinner=False)
def process_excel_fast(file_bytes: bytes, filename: str):
    df_full = pd.read_excel(io.BytesIO(file_bytes), header=None)
    
    if len(df_full) == 0:
        raise ValueError("ไฟล์ไม่มีข้อมูล")

    header_row_idx = None
    crs_col_idx = None
    pdsch_col_idx = None

    # วนลูปสแกนหาแถวที่มีทั้ง crsGain และ pdschTypeBGain (สแกนแถวไหนก็ได้)
    for r_idx in range(len(df_full)):
        row_values = df_full.iloc[r_idx]
        found_crs = None
        found_pdsch = None
        
        for c_idx, val in enumerate(row_values):
            clean_name = clean_col_name(val)
            if clean_name == "crsgain":
                found_crs = c_idx
            elif clean_name == "pdschtypebgain":
                found_pdsch = c_idx
        
        # เมื่อพบทั้งสองคอลัมน์ในแถวเดียวกัน บันทึกตำแหน่งแถวและคอลัมน์ไว้ทันที
        if found_crs is not None and found_pdsch is not None:
            header_row_idx = r_idx
            crs_col_idx = found_crs
            pdsch_col_idx = found_pdsch
            break

    if header_row_idx is None:
        raise ValueError("ไม่พบหัวข้อ 'crsGain' และ 'pdschTypeBGain' ในไฟล์")

    # จุดเริ่มต้นของข้อมูลคือแถวถัดจาก Header
    data_start_idx = header_row_idx + 1

    crs_series = pd.to_numeric(df_full.iloc[data_start_idx:, crs_col_idx], errors='coerce')
    pdsch_series = pd.to_numeric(df_full.iloc[data_start_idx:, pdsch_col_idx], errors='coerce')

    expected_pdsch = crs_series.map(VALID_PAIRS_SERIES)
    valid_mask = crs_series.isin(VALID_PAIRS_SERIES.index) & (expected_pdsch == pdsch_series)
    non_null_mask = crs_series.notna() & pdsch_series.notna()

    results_series = pd.Series("", index=crs_series.index)
    results_series[non_null_mask] = valid_mask[non_null_mask].map({True: 't', False: 'f'})

    t_count = int((results_series == "t").sum())
    f_count = int((results_series == "f").sum())

    # สร้าง Checked DF
    df_checked = df_full.copy()
    res_col = pd.Series("", index=df_checked.index)
    res_col.iloc[header_row_idx] = "result"
    res_col.iloc[data_start_idx:] = results_series.values
    df_checked[len(df_checked.columns)] = res_col

    # สร้าง Fixed DF
    df_fixed = df_full.copy()
    fixed_pdsch = pdsch_series.copy()
    fix_mask = crs_series.isin(VALID_PAIRS_SERIES.index) & crs_series.notna()
    fixed_pdsch[fix_mask] = expected_pdsch[fix_mask]
    
    df_fixed.iloc[data_start_idx:, pdsch_col_idx] = fixed_pdsch.values
    df_fixed[len(df_fixed.columns)] = res_col

    return {
        "filename": filename,
        "checked_df": df_checked,
        "fixed_df": df_fixed,
        "pdsch_col_idx": pdsch_col_idx,
        "header_row_idx": header_row_idx,
        "t": t_count,
        "f": f_count
    }

def df_to_excel_bytes_styled(df: pd.DataFrame, header_row_idx: int = 3, highlight_type=None, target_col_idx=None, fail_color="#FFC7CE", fix_color="#C6EFCE") -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, header=False)
        
        if highlight_excel and highlight_type:
            worksheet = writer.sheets['Sheet1']
            hex_fail = fail_color.lstrip('#')
            hex_fix = fix_color.lstrip('#')
            
            fill_fail = PatternFill(start_color=hex_fail, end_color=hex_fail, fill_type="solid")
            fill_fix = PatternFill(start_color=hex_fix, end_color=hex_fix, fill_type="solid")

            last_col_idx = df.shape[1]
            data_start_idx = header_row_idx + 1
            res_series = df.iloc[data_start_idx:, last_col_idx - 1]
            fail_rows = res_series[res_series == "f"].index

            if highlight_type == "checked":
                for r_idx in fail_rows:
                    row_num = r_idx + 1  # openpyxl 1-based index
                    for col in range(1, last_col_idx + 1):
                        worksheet.cell(row=row_num, column=col).fill = fill_fail

            elif highlight_type == "fixed" and target_col_idx is not None:
                for r_idx in fail_rows:
                    row_num = r_idx + 1
                    worksheet.cell(row=row_num, column=target_col_idx + 1).fill = fill_fix

    output.seek(0)
    return output.getvalue()

# ==================================================
# 4. CONTROLS FOR ADDING/REMOVING WEEKS
# ==================================================
st.markdown(f"### {section_manage_title}")

c1, c2, c3, c4, c5, c6 = st.columns(6)

with c1:
    st.button("➕ เพิ่ม 1", use_container_width=True, on_click=update_weeks, args=(1,))
with c2:
    st.button("➕ เพิ่ม 5", use_container_width=True, on_click=update_weeks, args=(5,))
with c3:
    st.button("➕ เพิ่ม 10", use_container_width=True, on_click=update_weeks, args=(10,))
with c4:
    st.button("➖ ลด 1", use_container_width=True, on_click=update_weeks, args=(-1,))
with c5:
    st.button("➖ ลด 5", use_container_width=True, on_click=update_weeks, args=(-5,))
with c6:
    st.button("➖ ลด 10", use_container_width=True, on_click=update_weeks, args=(-10,))

# ซิงค์ค่าตัวเลขล่าสุดตลอดเวลา
st.number_input(
    "หรือระบุจำนวนชุดอัปโหลดโดยตรง (1 - 50):", 
    min_value=1, 
    max_value=50, 
    value=st.session_state.num_weeks,
    key="direct_num_input",
    on_change=set_weeks_direct
)

st.caption(f"จำนวนชุดอัปโหลดปัจจุบัน: **{st.session_state.num_weeks}** ชุด")
st.divider()

# ==================================================
# 5. DYNAMIC UPLOAD SECTIONS (จำชื่อค่าที่กรอกล่าสุด)
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
                
                # เช็ค Session State เพื่อจำชื่อที่เคยตั้งไว้ล่าสุด
                name_key = f"w_name_{week_idx}"
                if name_key not in st.session_state:
                    st.session_state[name_key] = f"2026{16 + week_idx}"

                w_name = st.text_input(
                    f"ชื่อ Week {week_idx + 1}:",
                    key=name_key
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
# 6. PROCESS ALL WEEKS
# ==================================================
chart_rows = []
all_weeks_results = {}

has_files = any(len(w["files"]) > 0 for w in weeks_data_list)

if has_files:
    with st.spinner("🚀 กำลังประมวลผลข้อมูลในความเร็วสูง..."):
        for w_data in weeks_data_list:
            w_name = w_data["name"]
            files = w_data["files"]

            if files:
                w_results = []
                for uploaded_file in files[:50]:
                    try:
                        file_bytes = uploaded_file.getvalue()
                        res = process_excel_fast(file_bytes, uploaded_file.name)
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
# 7. COMBINED CHART (เอียงข้อความแกน X 90 องศา)
# ==================================================
if chart_rows:
    st.divider()

    combined_df = pd.DataFrame(chart_rows)
    
    # แปลงข้อมูลเป็น Long Format สำหรับสร้าง Stacked Bar Chart
    df_melted = combined_df.melt(
        id_vars=["Week"], 
        value_vars=["f - Incorrect", "t - Correct"],
        var_name="Status", 
        value_name="Count"
    )

    chart = (
        alt.Chart(df_melted)
        .mark_bar()
        .encode(
            x=alt.X("Week:N", title=None, axis=alt.Axis(labelAngle=-90)),
