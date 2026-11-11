import io
import re
import zipfile
import pandas as pd
import streamlit as st
from openpyxl.styles import PatternFill

st.set_page_config(
    page_title="CRS Gain Validator",
    page_icon="📊",
    layout="wide"
)

# ==================================================
# 1. SIDEBAR CONFIGURATION (ปรับแต่งหัวข้อ สี และการจัดการชุด)
# ==================================================
st.sidebar.header("⚙️ การตั้งค่าระบบ")

# 1.1 ส่วนปรับแต่งข้อความหัวข้อ (Title Settings)
with st.sidebar.expander("📝 ปรับแต่งข้อความหัวข้อ (Headers)", expanded=False):
    app_title = st.text_input("หัวข้อหลัก (Title):", value="📊 CRS Gain Validator")
    app_subtitle = st.text_input(
        "หัวข้อย่อย (Subheader):", 
        value="ระบบตรวจเช็คและแก้ไข CRS Gain, pdschTypeBGain อัตโนมัติ"
    )
    section_manage_title = st.text_input("หัวข้อส่วนจัดการชุดไฟล์:", value="🛠️ จัดการชุดไฟล์ (Week)")
    section_chart_title = st.text_input("หัวข้อส่วนแสดงกราฟ:", value="📈 กราฟเปรียบเทียบผลรวมทุก Week")
    section_download_title = st.text_input("หัวข้อส่วนดาวน์โหลด:", value="📥 ดาวน์โหลดไฟล์ Excel")

# 1.2 ส่วนปรับแต่งรูปแบบชื่อไฟล์ (File Name Format)
with st.sidebar.expander("📄 ตั้งชื่อไฟล์ Export", expanded=False):
    prefix_checked = st.text_input("คำนำหน้าไฟล์ Checked:", value="checked_")
    prefix_fixed = st.text_input("คำนำหน้าไฟล์ Fixed:", value="fixed_")

# 1.3 ส่วนปรับแต่งสี (Color Picker)
with st.sidebar.expander("🎨 ปรับแต่งสีระบบ", expanded=True):
    st.subheader("1. สีในไฟล์ Excel")
    highlight_excel = st.checkbox("เปิดใช้งานไฮไลต์สีใน Excel", value=True)
    color_fail_excel = st.color_picker("สีไฮไลต์ Error 'f' (Checked File)", "#FFC7CE")
    color_fix_excel = st.color_picker("สีไฮไลต์จุดที่แก้ไข (Fixed File)", "#C6EFCE")

    st.subheader("2. สีกราฟแท่ง")
    color_incorrect_chart = st.color_picker("สีกราฟ Incorrect (f)", "#E05252")
    color_correct_chart = st.color_picker("สีกราฟ Correct (t)", "#2878D0")


# ==================================================
# 2. RENDER MAIN HEADERS
# ==================================================
st.title(app_title)
st.subheader(app_subtitle)

# ==================================================
# 3. SESSION STATE FOR DYNAMIC WEEKS
# ==================================================
if "num_weeks" not in st.session_state:
    st.session_state.num_weeks = 2

# ==================================================
# 4. VALIDATION RULES & HELPER FUNCTIONS
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

# ==================================================
# 5. EXCEL GENERATION WITH STYLES
# ==================================================
def df_to_excel_bytes_styled(df: pd.DataFrame, highlight_type=None, target_col_idx=None, fail_color="#FFC7CE", fix_color="#C6EFCE") -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, header=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        if highlight_excel and highlight_type:
            hex_fail = fail_color.lstrip('#')
            hex_fix = fix_color.lstrip('#')
            
            fill_fail = PatternFill(start_color=hex_fail, end_color=hex_fail, fill_type="solid")
            fill_fix = PatternFill(start_color=hex_fix, end_color=hex_fix, fill_type="solid")

            last_col_idx = df.shape[1]

            for row in range(5, df.shape[0] + 1):
                res_val = worksheet.cell(row=row, column=last_col_idx).value
                
                if highlight_type == "checked" and res_val == "f":
                    for col in range(1, last_col_idx + 1):
                        worksheet.cell(row=row, column=col).fill = fill_fail

                elif highlight_type == "fixed" and res_val == "f":
                    if target_col_idx is not None:
                        worksheet.cell(row=row, column=target_col_idx + 1).fill = fill_fix

    output.seek(0)
    return output.getvalue()

# ==================================================
# 6. PROCESS EXCEL FUNCTION
# ==================================================
def process_excel(uploaded_file):
    file_bytes = uploaded_file.getvalue()
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
    df_fixed[len(df_fixed.columns)] = res_col

    return {
        "filename": uploaded_file.name,
        "checked_df": df_checked,
        "fixed_df": df_fixed,
        "pdsch_col_idx": pdsch_col_idx,
        "t": t_count,
        "f": f_count
    }

# ==================================================
# 7. CONTROLS FOR ADDING/REMOVING WEEKS (ปรับแต่งการเพิ่มครั้งละ 5, 10)
# ==================================================
st.markdown(f"### {section_manage_title}")

# จัดกลุ่มปุ่มเพิ่ม/ลดทีละ 1, 5, 10
c1, c2, c3, c4, c5, c6 = st.columns(6)

with c1:
    if st.button("➕ เพิ่ม 1", use_container_width=True):
        st.session_state.num_weeks = min(50, st.session_state.num_weeks + 1)
        st.rerun()
with c2:
    if st.button("➕ เพิ่ม 5", use_container_width=True):
        st.session_state.num_weeks = min(50, st.session_state.num_weeks + 5)
        st.rerun()
with c3:
    if st.button("➕ เพิ่ม 10", use_container_width=True):
        st.session_state.num_weeks = min(50, st.session_state.num_weeks + 10)
        st.rerun()
with c4:
    if st.button("➖ ลด 1", use_container_width=True):
        st.session_state.num_weeks = max(1, st.session_state.num_weeks - 1)
        st.rerun()
with c5:
    if st.button("➖ ลด 5", use_container_width=True):
        st.session_state.num_weeks = max(1, st.session_state.num_weeks - 5)
        st.rerun()
with c6:
    if st.button("➖ ลด 10", use_container_width=True):
        st.session_state.num_weeks = max(1, st.session_state.num_weeks - 10)
        st.rerun()

# ช่องให้ระบุจำนวนชุดโดยตรง
direct_num = st.number_input(
    "หรือระบุจำนวนชุดอัปโหลดโดยตรง (1 - 50):", 
    min_value=1, 
    max_value=50, 
    value=st.session_state.num_weeks,
    key="direct_num_input"
)

if direct_num != st.session_state.num_weeks:
    st.session_state.num_weeks = direct_num
    st.rerun()

st.caption(f"จำนวนชุดอัปโหลดปัจจุบัน: **{st.session_state.num_weeks}** ชุด")
st.divider()

# ==================================================
# 8. DYNAMIC UPLOAD SECTIONS
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
# 9. PROCESS ALL WEEKS
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
                        res = process_excel(uploaded_file)
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
# 10. COMBINED CHART
# ==================================================
if chart_rows:
    st.divider()
    st.header(section_chart_title)

    combined_df = pd.DataFrame(chart_rows).set_index("Week")

    st.bar_chart(
        combined_df,
        color=[color_incorrect_chart, color_correct_chart],
        height=500,
        stack=True
    )

    st.dataframe(combined_df.reset_index(), use_container_width=True, hide_index=True)

    # ==================================================
    # 11. DOWNLOAD SECTION (ZIP + SEPARATE FILES)
    # ==================================================
    st.divider()
    st.header(section_download_title)

    tab_titles = [f"📦 {w_name}" for w_name in all_weeks_results.keys()]
    if tab_titles:
        tabs = st.tabs(tab_titles)
        for tab_idx, (tab, (w_name, w_results)) in enumerate(zip(tabs, all_weeks_results.items())):
            with tab:
                # 1. ดาวน์โหลด ZIP รวม
                zip_b = io.BytesIO()
                with zipfile.ZipFile(zip_b, "w", zipfile.ZIP_DEFLATED) as z:
                    for r in w_results:
                        chk_b = df_to_excel_bytes_styled(
                            r["checked_df"], highlight_type="checked", 
                            fail_color=color_fail_excel, fix_color=color_fix_excel
                        )
                        fix_b = df_to_excel_bytes_styled(
                            r["fixed_df"], highlight_type="fixed", target_col_idx=r["pdsch_col_idx"],
                            fail_color=color_fail_excel, fix_color=color_fix_excel
                        )
                        z.writestr(f"{prefix_checked}{r['filename']}", chk_b)
                        z.writestr(f"{prefix_fixed}{r['filename']}", fix_b)

                st.download_button(
                    f"📦 ดาวน์โหลดไฟล์ทั้งหมดของ {w_name} (ZIP รวม)",
                    data=zip_b.getvalue(),
                    file_name=f"crs_{w_name}_results.zip",
                    mime="application/zip",
                    type="primary",
                    use_container_width=True,
                    key=f"dl_zip_{w_name}_{tab_idx}"
                )

                st.markdown("---")
                st.subheader(f"📄 รายการไฟล์และปุ่มดาวน์โหลดแยก ({len(w_results)} ไฟล์)")

                # 2. ดาวน์โหลดแยกรายไฟล์
                for file_idx, r in enumerate(w_results):
                    with st.expander(f"📄 {r['filename']} | ✅ Correct (t): {r['t']} | ❌ Incorrect (f): {r['f']}", expanded=False):
                        btn_col1, btn_col2 = st.columns(2)
                        
                        chk_bytes = df_to_excel_bytes_styled(
                            r["checked_df"], highlight_type="checked", 
                            fail_color=color_fail_excel, fix_color=color_fix_excel
                        )
                        fix_bytes = df_to_excel_bytes_styled(
                            r["fixed_df"], highlight_type="fixed", target_col_idx=r["pdsch_col_idx"],
                            fail_color=color_fail_excel, fix_color=color_fix_excel
                        )

                        with btn_col1:
                            st.download_button(
                                label=f"⬇️ โหลด Checked ({r['filename']})",
                                data=chk_bytes,
                                file_name=f"{prefix_checked}{r['filename']}",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True,
                                key=f"btn_chk_{tab_idx}_{file_idx}_{r['filename']}"
                            )

                        with btn_col2:
                            st.download_button(
                                label=f"⬇️ โหลด Fixed ({r['filename']})",
                                data=fix_bytes,
                                file_name=f"{prefix_fixed}{r['filename']}",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True,
                                key=f"btn_fix_{tab_idx}_{file_idx}_{r['filename']}"
                            )

    st.success("🎉 ประมวลผลเรียบร้อยแล้ว!")
