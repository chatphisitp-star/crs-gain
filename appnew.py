
import streamlit as st
import pandas as pd
import io

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


def validate_row(crs, pdsch):

    if pd.isna(crs) or pd.isna(pdsch):
        return ""

    try:
        crs_int = int(float(str(crs).strip()))
        pdsch_int = int(float(str(pdsch).strip()))
    except (ValueError, TypeError):
        return ""

    return (
        't' if crs_int in VALID_PAIRS
        and VALID_PAIRS[crs_int] == pdsch_int
        else 'f'
    )


def fix_pdsch_val(crs, pdsch):

    if pd.isna(crs):
        return pdsch

    try:
        crs_int = int(float(str(crs).strip()))

        if crs_int in VALID_PAIRS:
            return VALID_PAIRS[crs_int]

    except (ValueError, TypeError):
        pass

    return pdsch


# ==================================================
# 3. PROCESS EXCEL
# ==================================================

def process_excel(uploaded_file):

    file_bytes = uploaded_file.getvalue()

    # อ่านข้อมูลทั้งหมดโดยไม่ตัด Header
    df_full = pd.read_excel(
        io.BytesIO(file_bytes),
        header=None
    )

    # อ่าน Header ที่ Row 4
    df_header = pd.read_excel(
        io.BytesIO(file_bytes),
        header=3
    )

    crs_col_idx = None
    pdsch_col_idx = None

    for idx, col in enumerate(df_header.columns):

        clean_name = str(col).strip().lower()

        if clean_name == "crsgain":
            crs_col_idx = idx

        elif clean_name == "pdschtypebgain":
            pdsch_col_idx = idx

    if crs_col_idx is None or pdsch_col_idx is None:
        raise ValueError(
            "ไม่พบ crsGain หรือ pdschTypeBGain ใน Row 4"
        )

    # อ่านข้อมูลตั้งแต่ Row 5
    crs_series = df_full.iloc[4:, crs_col_idx]
    pdsch_series = df_full.iloc[4:, pdsch_col_idx]

    # ตรวจสอบ
    results = [
        validate_row(c, p)
        for c, p in zip(crs_series, pdsch_series)
    ]

    # แก้ไข
    fixed_pdsch = [
        fix_pdsch_val(c, p)
        for c, p in zip(crs_series, pdsch_series)
    ]

    t_count = results.count("t")
    f_count = results.count("f")

    # สร้าง Checked
    df_checked = df_full.copy()

    result_column = [""] * len(df_checked)
    result_column[3] = "result"

    for i, res in enumerate(results):
        result_column[4 + i] = res

    df_checked[len(df_checked.columns)] = result_column

    # สร้าง Fixed
    df_fixed = df_full.copy()

    for i, val in enumerate(fixed_pdsch):
        df_fixed.iloc[4 + i, pdsch_col_idx] = val

    return {
        "filename": uploaded_file.name,
        "checked": df_checked,
        "fixed": df_fixed,
        "t": t_count,
        "f": f_count
    }


# ==================================================
# 4. PROCESS MULTIPLE FILES
# ==================================================

if uploaded_files:

    if len(uploaded_files) > 30:
        st.warning("รองรับสูงสุด 30 ไฟล์")

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
            st.error(
                f"❌ {uploaded_file.name}: {str(e)}"
            )

        progress.progress((index + 1) / len(files))

    # ==================================================
    # 5. SUMMARY
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
        # 6. COMPARISON CHART
        # ==================================================

        st.divider()
        st.header("📈 กราฟเปรียบเทียบทุกไฟล์ Excel")

        comparison_df = pd.DataFrame([
            {
                "Excel File": r["filename"],
                "t - Correct": r["t"],
                "f - Incorrect": r["f"]
            }
            for r in all_results
        ])

        chart_comparison = comparison_df.set_index("Excel File")

        st.bar_chart(
            chart_comparison,
            color=["#2878D0", "#E05252"],
            height=500
        )

        st.dataframe(
            comparison_df,
            use_container_width=True,
            hide_index=True
        )

        # ==================================================
        # 7. INDIVIDUAL CHART
        # ==================================================

        st.divider()
        st.header("📊 กราฟแยกตามไฟล์ Excel")

        for index, result in enumerate(all_results):

            with st.expander(
                f"📁 {result['filename']} "
                f"| t = {result['t']:,} "
                f"| f = {result['f']:,}",
                expanded=True
            ):

                col1, col2, col3 = st.columns(3)

                col1.metric("t - ถูกต้อง", result["t"])
                col2.metric("f - ผิดพลาด", result["f"])
                col3.metric(
                    "รวมรายการ",
                    result["t"] + result["f"]
                )

                chart_df = pd.DataFrame({
                    "จำนวนรายการ": [
                        result["t"],
                        result["f"]
                    ]
                }, index=["t - Correct", "f - Incorrect"])

                st.bar_chart(
                    chart_df,
                    color="#2878D0",
                    height=300
                )

                st.caption("ตัวอย่างข้อมูล Checked")

                st.dataframe(
                    result["checked"].head(15),
                    use_container_width=True
                )

        # ==================================================
        # 8. DOWNLOAD FILES
        # ==================================================

        st.divider()
        st.header("📥 ดาวน์โหลดไฟล์ Excel")

        for index, result in enumerate(all_results):

            st.markdown(f"**📁 {result['filename']}**")

            col1, col2 = st.columns(2)

            # Checked
            buffer1 = io.BytesIO()

            with pd.ExcelWriter(
                buffer1,
                engine="openpyxl"
            ) as writer:

                result["checked"].to_excel(
                    writer,
                    index=False,
                    header=False
                )

            col1.download_button(
                "📥 ดาวน์โหลด Checked",
                data=buffer1.getvalue(),
                file_name=f"checked_{result['filename']}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"checked_{index}"
            )

            # Fixed
            buffer2 = io.BytesIO()

            with pd.ExcelWriter(
                buffer2,
                engine="openpyxl"
            ) as writer:

                result["fixed"].to_excel(
                    writer,
                    index=False,
                    header=False
                )

            col2.download_button(
                "🛠️ ดาวน์โหลด Fixed",
                data=buffer2.getvalue(),
                file_name=f"fixed_{result['filename']}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"fixed_{index}"
            )

        st.success("🎉 ประมวลผลเรียบร้อยแล้ว!")

    else:
        st.error("ไม่สามารถประมวลผลไฟล์ Excel ได้")
