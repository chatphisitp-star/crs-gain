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
st.subheader("ระบบตรวจเช็คและแก้ไข CRS Gain, pdschTypeBGain อัตโนมัติ (รองรับสูงสุด 50 ชุด/Week)")

# ==================================================
# 1. SESSION STATE FOR DYNAMIC WEEKS
# ==================================================

# กำหนดจำนวนชุดเริ่มต้นที่ 2 ชุด
if "num_weeks" not in st.session_state:
    st.session_state.num_weeks = 2

# ==================================================
# 2. VALIDATION RULES & HELPER FUNCTIONS
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
    """ทำความสะอาดชื่อคอลัมน์ ลบเว้นวรรคและอักขระพิเศษออก"""
    if pd.isna(name):
        return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(name)).lower()


def parse_int_safe(val):
    """แปลงค่าเป็น int อย่างปลอดภัย รองรับ string, float (เช่น 300.0)"""
    if pd.isna(val):
        return None
    try:
        return int(float(str(val).strip()))
    except (ValueError, TypeError):
        return None


def validate_row(crs, pdsch) -> str:
    crs_int = parse_int_safe(crs)
    pdsch_int = parse_int_safe(pdsch)

    if crs_int is None or pdsch_int is None:
        return ""

    return (
        't' if crs_int in VALID_PAIRS
        and VALID_PAIRS[crs_int] == pdsch_int
        else 'f'
    )


def fix_pdsch_val(crs, pdsch):
    crs_int = parse_int_safe(crs)

    if crs_int is not None and crs_int in VALID_PAIRS:
        return VALID_PAIRS[crs_int]

    return pdsch


def process_excel(uploaded_file):
    file_bytes = uploaded_file.getvalue()

    df_full = pd.read_excel(io.BytesIO(file_bytes), header=None)
    df_header = pd.read_excel(io.BytesIO(file_bytes), header=3)

    crs_col_idx = None
    pdsch_col_idx = None

    for idx, col in enumerate(df_header.columns):
        clean_name = clean_col_name(col)

        if clean_name == "crsgain":
            crs_col_idx = idx
        elif clean_name == "pdschtypebgain":
            pdsch_col_idx = idx

    if crs_col_idx is None or pdsch_col_idx is None:
        raise ValueError("ไม่พบ 'crsGain' หรือ 'pdschTypeBGain' ใน Row 4")

    crs_series = df_full.iloc[4:, crs_col_idx]
    pdsch_series = df_full.iloc[4:, pdsch_col_idx]

    results = [validate_row(c, p) for c, p in zip(crs_series, pdsch_series)]
    fixed_pdsch = [fix_pdsch_val(c, p) for c, p in zip(crs_series, pdsch_series)]

    t_count = results.count("t")
    f_count = results.count("f")

    # สร้าง DataFrame Checked
    df_checked = df_full.copy()
    result_column = [""] * len(df_checked)
    result_column[3] = "result"

    for i, res in enumerate(results):
        result_column[4 + i] = res

    df_checked[len(df_checked.columns)] = result_column

    # สร้าง DataFrame Fixed
    df_fixed = df_full.copy()
    for i, val in enumerate(fixed_pdsch):
        df_fixed.iloc[4 + i, pdsch_col_idx] = val

    checked_bytes = io.BytesIO()
    with pd.ExcelWriter(checked_bytes, engine="openpyxl") as writer:
        df_checked.to_excel(writer, index=False, header=False)

    fixed_bytes = io.BytesIO()
    with pd.ExcelWriter(fixed_bytes, engine="openpyxl") as writer:
        df_fixed.to_excel(writer, index=False, header=False)

    return {
        "filename": uploaded_file.name,
        "checked": df_checked,
        "fixed": df_fixed,
        "checked_bytes": checked_bytes.getvalue(),
        "fixed_bytes": fixed_bytes.getvalue(),
        "t": t_count,
        "f": f_count
    }


# ==================================================
# 3. CONTROLS FOR ADDING/REMOVING WEEKS (สูงสุด 50 ชุด)
# ==================================================

st.markdown("### 🛠️ จัดการชุดไฟล์ (Week)")
ctrl_col1, ctrl_col2, _ = st.columns([1, 1, 3])

with ctrl_col1:
    if st.button("➕ เพิ่มชุดไฟล์ (Week)", use_container_width=True):
        if st.session_state.num_weeks < 50:
            st.session_state.num_weeks += 1
        else:
            st.warning("จำกัดสูงสุด 50 ชุด")

with ctrl_col2:
    if st.button("➖ ลดชุดไฟล์", use_container_width=True):
        if st.session_state.num_weeks > 1:
            st.session_state.num_weeks -= 1

st.caption(f"จำนวนชุดอัปโหลดปัจจุบัน: **{st.session_state.num_weeks}** ชุด")
st.divider()

# ==================================================
# 4. DYNAMIC UPLOAD SECTIONS
# ==================================================

weeks_data_list = []

# จัดวางกล่องอัปโหลดในรูปแบบ Grid (3 ชุดต่อ 1 แถว)
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

for w_idx, w_data in enumerate(weeks_data_list):
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
# 6. COMBINED CHART (ทุก Week เรียงต่อกันบนแกน X)
# ==================================================

if chart_rows:
    st.divider()
    st.header("📈 กราฟเปรียบเทียบผลรวมทุก Week (เรียงต่อกันบนแกน X)")

    combined_df = pd.DataFrame(chart_rows).set_index("Week")

    # กราฟแท่งแสดงผลรวมของแต่ละ Week วางต่อกันตามแกน X
    st.bar_chart(
        combined_df,
        color=["#E05252", "#2878D0"],
        height=500,
        stack=True
    )

    st.dataframe(combined_df.reset_index(), use_container_width=True, hide_index=True)

    # ==================================================
    # 7. DOWNLOAD SECTION FOR ALL WEEKS
    # ==================================================

    st.divider()
    st.header("📥 ดาวน์โหลดไฟล์ Excel แยกตาม Week")

    tab_titles = [f"📦 {w_name}" for w_name in all_weeks_results.keys()]
    if tab_titles:
        tabs = st.tabs(tab_titles)
        for tab, (w_name, w_results) in zip(tabs, all_weeks_results.items()):
            with tab:
                zip_b = io.BytesIO()
                with zipfile.ZipFile(zip_b, "w", zipfile.ZIP_DEFLATED) as z:
                    for r in w_results:
                        z.writestr(f"checked_{r['filename']}", r["checked_bytes"])
                        z.writestr(f"fixed_{r['filename']}", r["fixed_bytes"])

                st.download_button(
                    f"📦 ดาวน์โหลดไฟล์ทั้งหมดของ {w_name} (ZIP)",
                    data=zip_b.getvalue(),
                    file_name=f"crs_{w_name}_results.zip",
                    mime="application/zip",
                    type="primary",
                    use_container_width=True
                )

    st.success("🎉 ประมวลผลเรียบร้อยแล้ว!")
