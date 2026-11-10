
import streamlit as st
import pandas as pd
import io
import plotly.graph_objects as go

# ==================================================
# 1. PAGE CONFIG
# ==================================================

st.set_page_config(
    page_title="CRS Gain Validator",
    page_icon="📊",
    layout="wide"
)

st.title("📊 CRS Gain Validator")
st.subheader("ระบบตรวจเช็คและแก้ไข CRS Gain, pdschTypeBGain อัตโนมัติ")

st.info("รองรับการอัปโหลดไฟล์ Excel สูงสุด 30 ไฟล์ในครั้งเดียว")

# ==================================================
# 2. UPLOAD FILES
# ==================================================

uploaded_files = st.file_uploader(
    "📂 ลากไฟล์ Excel (.xlsx) มาวางที่นี่",
    type=["xlsx"],
    accept_multiple_files=True
)

# ==================================================
# 3. VALIDATION RULES
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
        't'
        if crs_int in VALID_PAIRS
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
# 4. PROCESS EXCEL
# ==================================================

def process_excel(uploaded_file):

    file_bytes = uploaded_file.getvalue()

    # อ่านข้อมูลทั้งหมด ไม่ตัด Header
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
            "ไม่พบคอลัมน์ crsGain หรือ pdschTypeBGain ใน Row 4"
        )

    # อ่านข้อมูลตั้งแต่ Row 5
    crs_series = df_full.iloc[4:, crs_col_idx]
    pdsch_series = df_full.iloc[4:, pdsch_col_idx]

    # ตรวจสอบ t / f
    results = [
        validate_row(c, p)
        for c, p in zip(crs_series, pdsch_series)
    ]

    # แก้ไข pdschTypeBGain
    fixed_pdsch = [
        fix_pdsch_val(c, p)
        for c, p in zip(crs_series, pdsch_series)
    ]

    # นับผลลัพธ์
    t_count = results.count("t")
    f_count = results.count("f")

    # สร้างไฟล์ Checked
    df_checked = df_full.copy()

    result_column = [""] * len(df_checked)
    result_column[3] = "result"

    for i, res in enumerate(results):
        result_column[4 + i] = res

    df_checked[len(df_checked.columns)] = result_column

    # สร้างไฟล์ Fixed
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
# 5. PROCESS MULTIPLE FILES
# ==================================================

if uploaded_files:

    if len(uploaded_files) > 30:
        st.warning("รองรับสูงสุด 30 ไฟล์ จะประมวลผล 30 ไฟล์แรก")

    files_to_process = uploaded_files[:30]

    all_results = []

    st.divider()
    st.header("📋 กำลังประมวลผลไฟล์ Excel")

    progress = st.progress(0)

    for index, uploaded_file in enumerate(files_to_process):

        try:

            result = process_excel(uploaded_file)

            all_results.append(result)

            st.success(
                f"✅ {uploaded_file.name} "
                f"| t: {result['t']:,} "
                f"| f: {result['f']:,}"
            )

        except Exception as e:

            st.error(
                f"❌ {uploaded_file.name}: {str(e)}"
            )

        progress.progress(
            (index + 1) / len(files_to_process)
        )

    # ==================================================
    # 6. SUMMARY
    # ==================================================

    if all_results:

        total_t = sum(r["t"] for r in all_results)
        total_f = sum(r["f"] for r in all_results)

        st.divider()
        st.header("📊 ภาพรวมผลการตรวจสอบ")

        col1, col2, col3, col4 = st.columns(4)

        col1.metric(
            "จำนวนไฟล์",
            f"{len(all_results)}"
        )

        col2.metric(
            "t - ถูกต้อง",
            f"{total_t:,}"
        )

        col3.metric(
            "f - ผิดพลาด",
            f"{total_f:,}"
        )

        col4.metric(
            "รวมรายการ",
            f"{total_t + total_f:,}"
        )

        # ==================================================
        # 7. COMPARISON CHART
        # ==================================================

        st.divider()
        st.header("📈 กราฟเปรียบเทียบทุกไฟล์ Excel")

        comparison_df = pd.DataFrame([
            {
                "File": r["filename"],
                "t": r["t"],
                "f": r["f"]
            }
            for r in all_results
        ])

        fig_compare = go.Figure()

        fig_compare.add_trace(
            go.Bar(
                name="t - Correct",
                x=comparison_df["File"],
                y=comparison_df["t"],
                marker_color="#2878D0",
                text=comparison_df["t"],
                textposition="auto"
            )
        )

        fig_compare.add_trace(
            go.Bar(
                name="f - Incorrect",
                x=comparison_df["File"],
                y=comparison_df["f"],
                marker_color="#E05252",
                text=comparison_df["f"],
                textposition="auto"
            )
        )

        fig_compare.update_layout(
            title="CRS Gain Comparison - All Excel Files",
            xaxis_title="Excel File",
            yaxis_title="Number of Records",
            barmode="group",
            height=550,
            xaxis_tickangle=-45,
            template="plotly_white",
            legend_title="Result"
        )

        st.plotly_chart(
            fig_compare,
            use_container_width=True
        )

        # ==================================================
        # 8. INDIVIDUAL CHART
        # ==================================================

        st.divider()
        st.header("📊 กราฟแสดงผลแยกตามไฟล์ Excel")

        for index, result in enumerate(all_results):

            with st.expander(
                f"📁 {result['filename']} "
                f"| t = {result['t']:,} "
                f"| f = {result['f']:,}",
                expanded=True
            ):

                col1, col2, col3 = st.columns(3)

                col1.metric(
                    "t - ถูกต้อง",
                    f"{result['t']:,}"
                )

                col2.metric(
                    "f - ผิดพลาด",
                    f"{result['f']:,}"
                )

                col3.metric(
                    "รวมรายการ",
                    f"{result['t'] + result['f']:,}"
                )

                fig = go.Figure()

                fig.add_trace(
                    go.Bar(
                        x=["t", "f"],
                        y=[result["t"], result["f"]],
                        marker_color=[
                            "#2878D0",
                            "#E05252"
                        ],
                        text=[
                            result["t"],
                            result["f"]
                        ],
                        textposition="auto"
                    )
                )

                fig.update_layout(
                    title=f"CRS Gain Result: {result['filename']}",
                    xaxis_title="Validation Result",
                    yaxis_title="Number of Records",
                    height=350,
                    template="plotly_white",
                    showlegend=False
                )

                st.plotly_chart(
                    fig,
                    use_container_width=True,
                    key=f"chart_{index}"
                )

                st.caption("ตัวอย่างข้อมูล Checked")

                st.dataframe(
                    result["checked"].head(15),
                    use_container_width=True
                )

        # ==================================================
        # 9. DOWNLOAD EXCEL
        # ==================================================

        st.divider()
        st.header("📥 ดาวน์โหลดไฟล์ Excel")

        for index, result in enumerate(all_results):

            st.markdown(
                f"**📁 {result['filename']}**"
            )

            col1, col2 = st.columns(2)

            # Download Checked
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
                label="📥 ดาวน์โหลด Checked",
                data=buffer1.getvalue(),
                file_name=f"checked_{result['filename']}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"checked_{index}"
            )

            # Download Fixed
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
                label="🛠️ ดาวน์โหลด Fixed",
                data=buffer2.getvalue(),
                file_name=f"fixed_{result['filename']}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"fixed_{index}"
            )

        st.divider()

        st.success(
            "🎉 ประมวลผลไฟล์ Excel ทั้งหมดเรียบร้อยแล้ว!"
        )

    else:

        st.error("ไม่สามารถประมวลผลไฟล์ Excel ได้")
