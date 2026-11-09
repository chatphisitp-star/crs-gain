import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="CRS Gain Validator", layout="centered")
st.title("📊 ระบบตรวจเช็ค CRS Gain อัตโนมัติ")

# 1. ส่วนลากวางไฟล์บนหน้าเว็บ
uploaded_file = st.file_uploader("ลากไฟล์ Excel (.xlsx) มาวางที่นี่", type=["xlsx"])

def validate_row(crs, pdsch):
    if pd.isna(crs) or pd.isna(pdsch):
        return ""
    try:
        crs_int = int(float(str(crs).strip()))
        pdsch_int = int(float(str(pdsch).strip()))
    except (ValueError, TypeError):
        return ""
    
    valid_pairs = {300: 1, 477: 2, 600: 3, -300: 1, 177: 1, 0: 0}
    return 't' if crs_int in valid_pairs and valid_pairs[crs_int] == pdsch_int else 'f'

if uploaded_file is not None:
    # อ่านโครงสร้างทั้งหมดโดยไม่ตัด header เพื่อเก็บข้อมูลเดิมไว้ 100%
    df_full = pd.read_excel(uploaded_file, header=None)
    
    # อ่านแบบกำหนด header=3 (Row 4) เพื่อค้นหาตำแหน่งคอลัมน์
    df_header = pd.read_excel(uploaded_file, header=3)
    
    crs_col_idx = None
    pdsch_col_idx = None
    
    # ค้นหาดัชนีคอลัมน์ (Index) ของ crsGain และ pdschTypeBGain
    for idx, col in enumerate(df_header.columns):
        clean_name = str(col).strip().lower()
        if clean_name == 'crsgain':
            crs_col_idx = idx
        elif clean_name == 'pdschtypebgain':
            pdsch_col_idx = idx

    if crs_col_idx is not None and pdsch_col_idx is not None:
        # ดึงข้อมูลมาตรวจเช็คตั้งแต่แถวที่ 5 เป็นต้นไป (Index 4 ใน df_full)
        crs_series = df_full.iloc[4:, crs_col_idx]
        pdsch_series = df_full.iloc[4:, pdsch_col_idx]
        
        # คำนวณผลลัพธ์
        results = [validate_row(c, p) for c, p in zip(crs_series, pdsch_series)]
        
        # สร้างคอลัมน์ผลลัพธ์ต่อท้ายสุด โดยให้แถวที่ 4 (Row 4) เป็นหัวข้อ 'result'
        result_column = [''] * len(df_full)
        result_column[3] = 'result'  # ใส่ชื่อคอลัมน์ใน Row 4
        
        # ใส่ผลลัพธ์ตั้งแต่ Row 5 เป็นต้นไป
        for i, res in enumerate(results):
            result_column[4 + i] = res
            
        # เพิ่มคอลัมน์ผลลัพธ์เข้า DataFrame เดิม
        df_full[len(df_full.columns)] = result_column
        
        st.success("✅ ตรวจสอบข้อมูลเรียบร้อยแล้ว! (คงข้อมูลเดิมไว้ครบถ้วน)")
        st.dataframe(df_full.head(15))
        
        output_name = f"checked_{uploaded_file.name}"
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_full.to_excel(writer, index=False, header=False)
            
        st.download_button(
            label="📥 กดเพื่อดาวน์โหลดไฟล์ผลลัพธ์ (.xlsx)",
            data=buffer.getvalue(),
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("❌ ไม่พบคอลัมน์ 'crsGain' หรือ 'pdschTypeBGain' ในแถวที่ 4 กรุณาตรวจสอบไฟล์อีกครั้ง")
