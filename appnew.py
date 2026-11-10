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
st.subheader("ระบบตรวจเช็คและแก้ไข CRS Gain, pdschTypeBGain อัตโนมัติ")

# ==================================================
# 1. UPLOAD EXCEL
# ==================================================

uploaded_files = st.file_uploader(
    "📂 ลากไฟล์ Excel มาวางที่นี่ (สูงสุด 30 ไฟล์)",
    type=["xlsx"],
    accept_multiple_files=True
)

# ==================================================
# 2. VALIDATION RULES
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


# ==================================================
# 3. PROCESS EXCEL
# ==================================================

def process_excel(uploaded_file):
    file_bytes = uploaded_file.getvalue()

    # อ่านข้อมูลทั้งหมดโดยไม่ตัด Header
    df_full = pd.read_excel(io.BytesIO(file_bytes), header=None)

    # อ่าน Header ที่ Row 4 (index 3)
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

    # อ่านข้อมูลตั้งแต่ Row 5 (index 4)
    crs_series = df_full.iloc[4:, crs_col_idx]
    pdsch_series = df_full.iloc[4:, pdsch_col_idx]

    # ตรวจสอบและแก้ไข
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

    # แปลงเป็น Bytes ล่วงหน้าเพื่อประสิทธิภาพการดาวน์โหลด
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
# 4. PROCESS MULTIPLE FILES
# ==================================================

if uploaded_files:

    if len(uploaded_files) > 30:
        st.warning("⚠️ เพิ่มไฟล์มากกว่า 30 ไฟล์ ระบบจะประมวลผลเฉพาะ 30 ไฟล์แรกเท่านั้น")

    files = uploaded_files[:30]
    all_results = []

    st.divider()
    st.header("📋 ผลการตรวจสอบ")
    progress = st.progress(0)

    for index, uploaded_file in enumerate(files):
        try:
            result = process_excel(uploaded_file)
            all_results.append(result)
        except Exception as e:
            st.error(f"❌ {uploaded_file.name}: {str(e)}")

        progress.progress((index + 1) / len(files))

    # ==================================================
    # 5. SUMMARY & STACKED BAR CHART
    # ==================================================

    if all_results:
        total_t = sum(r["t"] for r in all_results)
        total_f = sum(r["f"] for r in all_results)

        st.divider()
        st.header("📊 ภาพรวมทั้งหมด")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("จำนวนไฟล์", len(all_results))
        c2.metric("t - ถูกต้อง", f"{total_t:,}")
        c3.metric("f - ผิดพลาด", f"{total_f:,}")
        c4.metric("รวมรายการ", f"{total_t + total_f:,}")

        # ==================================================
        # 6. COMPARISON CHART (1 ไฟล์ = 1 แท่ง / f บน / t ล่าง)
        # ==================================================

        st.divider()
        st.header("📈 กราฟเปรียบเทียบทุกไฟล์ Excel (1 ไฟล์ = 1 แท่ง)")

        # สลับเอา f ขึ้นก่อน t เพื่อให้ส่วนที่เป็น f แสดงซ้อนอยู่ด้านบนสุด
        comparison_df = pd.DataFrame([
            {
                "Excel File": r["filename"],
                "f - Incorrect": r["f"],
                "t - Correct": r["t"]
            }
            for r in all_results
        ]).set_index("Excel File")

        st.bar_chart(
            comparison_df,
            color=["#E05252", "#2878D0"],  # สีแดงสำหรับ f และสีฟ้าสำหรับ t
            height=500,
            stack=True
        )

        st.dataframe(comparison_df.reset_index(), use_container_width=True, hide_index=True)

        # ==================================================
        # 7. INDIVIDUAL DETAILS
        # ==================================================

        st.divider()
        st.header("📊 รายละเอียดแยกตามไฟล์")

        for index, result in enumerate(all_results):
            with st.expander(
                f"📁 {result['filename']} | t = {result['t']:,} | f = {result['f']:,}",
                expanded=False
            ):
                col1, col2, col3 = st.columns(3)
                col1.metric("t - ถูกต้อง", result["t"])
                col2.metric("f - ผิดพลาด", result["f"])
                col3.metric("รวมรายการ", result["t"] + result["f"])

                st.caption("ตัวอย่างข้อมูล Checked (15 แถวแรก)")
                st.dataframe(result["checked"].head(15), use_container_width=True)

        # ==================================================
        # 8. DOWNLOAD FILES (BULK & INDIVIDUAL)
        # ==================================================

        st.divider()
        st.header("📥 ดาวน์โหลดไฟล์ Excel")

        # ZIP Download All
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for r in all_results:
                zip_file.writestr(f"checked_{r['filename']}", r["checked_bytes"])
                zip_file.writestr(f"fixed_{r['filename']}", r["fixed_bytes"])

        st.download_button(
            "📦 ดาวน์โหลดไฟล์ทั้งหมด (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="crs_validation_results.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True
        )

        st.write("")

        # Individual Download
        for index, result in enumerate(all_results):
            st.markdown(f"**📁 {result['filename']}**")
            col1, col2 = st.columns(2)

            col1.download_button(
                "📥 ดาวน์โหลด Checked",
                data=result["checked_bytes"],
                file_name=f"checked_{result['filename']}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"checked_{index}"
            )

            col2.download_button(
                "🛠️ ดาวน์โหลด Fixed",
                data=result["fixed_bytes"],
                file_name=f"fixed_{result['filename']}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"fixed_{index}"
            )

        st.success("🎉 ประมวลผลเรียบร้อยแล้ว!")

    else:
        st.error("ไม่สามารถประมวลผลไฟล์ Excel ได้")
