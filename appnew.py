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
st.subheader("ระบบตรวจเช็คและแก้ไข CRS Gain, pdschTypeBGain อัตโนมัติ (รองรับหลาย Week)")

# ใช้ session_state สำหรับเก็บประวัติกราฟแต่ละ Week
if "weeks_data" not in st.session_state:
    st.session_state.weeks_data = {}

# ==================================================
# 1. VALIDATION RULES & HELPER FUNCTIONS
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
# 2. UPLOAD SECTION (แยกเป็น 2 ชุดไฟล์)
# ==================================================

col_w1, col_w2 = st.columns(2)

with col_w1:
    st.subheader("📅 ชุดที่ 1 (Week 1)")
    week1_name = st.text_input("ชื่อแท่งกราฟชุดที่ 1:", value="202616", key="w1_name")
    uploaded_files_1 = st.file_uploader(
        "📂 ลากไฟล์ Excel (ชุดที่ 1) มาวางที่นี่",
        type=["xlsx"],
        accept_multiple_files=True,
        key="uploader_1"
    )

with col_w2:
    st.subheader("📅 ชุดที่ 2 (Week 2)")
    week2_name = st.text_input("ชื่อแท่งกราฟชุดที่ 2:", value="202617", key="w2_name")
    uploaded_files_2 = st.file_uploader(
        "📂 ลากไฟล์ Excel (ชุดที่ 2) มาวางที่นี่",
        type=["xlsx"],
        accept_multiple_files=True,
        key="uploader_2"
    )

# ==================================================
# 3. PROCESS FILES FOR BOTH WEEKS
# ==================================================

chart_rows = []

def process_file_list(files_list, week_label):
    all_results = []
    for uploaded_file in files_list[:30]:
        try:
            res = process_excel(uploaded_file)
            all_results.append(res)
        except Exception as e:
            st.error(f"❌ [{week_label}] {uploaded_file.name}: {str(e)}")
    return all_results

# ประมวลผลชุดที่ 1
results_w1 = process_file_list(uploaded_files_1, week1_name) if uploaded_files_1 else []
if results_w1:
    tot_t1 = sum(r["t"] for r in results_w1)
    tot_f1 = sum(r["f"] for r in results_w1)
    chart_rows.append({
        "Week": week1_name,
        "f - Incorrect": tot_f1,
        "t - Correct": tot_t1
    })

# ประมวลผลชุดที่ 2
results_w2 = process_file_list(uploaded_files_2, week2_name) if uploaded_files_2 else []
if results_w2:
    tot_t2 = sum(r["t"] for r in results_w2)
    tot_f2 = sum(r["f"] for r in results_w2)
    chart_rows.append({
        "Week": week2_name,
        "f - Incorrect": tot_f2,
        "t - Correct": tot_t2
    })

# ==================================================
# 4. SHOW COMBINED CHART (กราฟแท่งต่อด้านข้าง)
# ==================================================

if chart_rows:
    st.divider()
    st.header("📈 กราฟเปรียบเทียบระหว่าง Week (1 ชุดไฟล์ = 1 แท่ง)")

    combined_df = pd.DataFrame(chart_rows).set_index("Week")

    # แสดงแท่งกราฟของ Week 1 และ Week 2 เรียงต่อกันบนแกน X
    st.bar_chart(
        combined_df,
        color=["#E05252", "#2878D0"],
        height=450,
        stack=True
    )

    st.dataframe(combined_df.reset_index(), use_container_width=True, hide_index=True)

    # ==================================================
    # 5. DOWNLOAD SECTION FOR BOTH WEEKS
    # ==================================================

    st.divider()
    st.header("📥 ดาวน์โหลดไฟล์ Excel แยกตามชุด")

    tab1, tab2 = st.tabs([f"📦 {week1_name}", f"📦 {week2_name}"])

    with tab1:
        if results_w1:
            zip_b1 = io.BytesIO()
            with zipfile.ZipFile(zip_b1, "w", zipfile.ZIP_DEFLATED) as z:
                for r in results_w1:
                    z.writestr(f"checked_{r['filename']}", r["checked_bytes"])
                    z.writestr(f"fixed_{r['filename']}", r["fixed_bytes"])

            st.download_button(
                f"📦 ดาวน์โหลดไฟล์ทั้งหมดของ {week1_name} (ZIP)",
                data=zip_b1.getvalue(),
                file_name=f"crs_{week1_name}_results.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True
            )
        else:
            st.info("ยังไม่มีการอัปโหลดไฟล์ในชุดที่ 1")

    with tab2:
        if results_w2:
            zip_b2 = io.BytesIO()
            with zipfile.ZipFile(zip_b2, "w", zipfile.ZIP_DEFLATED) as z:
                for r in results_w2:
                    z.writestr(f"checked_{r['filename']}", r["checked_bytes"])
                    z.writestr(f"fixed_{r['filename']}", r["fixed_bytes"])

            st.download_button(
                f"📦 ดาวน์โหลดไฟล์ทั้งหมดของ {week2_name} (ZIP)",
                data=zip_b2.getvalue(),
                file_name=f"crs_{week2_name}_results.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True
            )
        else:
            st.info("ยังไม่มีการอัปโหลดไฟล์ในชุดที่ 2")

    st.success("🎉 ประมวลผลเรียบร้อยแล้ว!")
