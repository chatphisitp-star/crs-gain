import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="CRS Gain Validator", layout="centered")
st.title("📊 ระบบตรวจเช็คและแก้ไข CRS Gain, pdschTypeBGain อัตโนมัติ")

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

def fix_pdsch_val(crs, pdsch):
    if pd.isna(crs):
        return pdsch
    try:
        crs_int = int(float(str(crs).strip()))
        valid_pairs = {300: 1, 477: 2, 600: 3, -300: 1, 177: 1, 0: 0}
        if crs_int in valid_pairs:
            return valid_pairs[crs_int]
    except (ValueError, TypeError):
        pass
    return pdsch

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
        
        # 1. คำนวณผลลัพธ์ t / f
        results = [validate_row(c, p) for c, p in zip(crs_series, pdsch_series)]
        
        # 2. คำนวณค่า pdschTypeBGain ที่แก้ไขถูกต้องแล้ว
        fixed_pdsch = [fix_pdsch_val(c, p) for c, p in zip(crs_series, pdsch_series)]
        
        # --- สร้าง DF แบบที่ 1: เพิ่มคอลัมน์ result ต่อท้าย ---
        df_checked = df_full.copy()
        result_column = [''] * len(df_checked)
        result_column[3] = 'result'  # ใส่ชื่อคอลัมน์ใน Row 4
        for i, res in enumerate(results):
            result_column[4 + i] = res
        df_checked[len(df_checked.columns)] = result_column
        
        # --- สร้าง DF แบบที่ 2: แก้ไขค่าในคอลัมน์ pdschTypeBGain โดยตรง ---
        df_fixed = df_full.copy()
        for i, val in enumerate(fixed_pdsch):
            df_fixed.iloc[4 + i, pdsch_col_idx] = val

        st.success("✅ ประมวลผลข้อมูลเรียบร้อยแล้ว!")
        
        error_count = results.count('f')
        if error_count > 0:
            st.warning(f"⚠️ พบข้อมูลที่ไม่ถูกต้องจำนวน {error_count} แถว")
        else:
            st.info("🎉 ข้อมูลถูกต้องทั้งหมดแล้ว")

        st.dataframe(df_checked.head(15))
        
        col1, col2 = st.columns(2)
        
        # ปุ่มที่ 1: ดาวน์โหลดไฟล์ตรวจเช็ค (มีคอลัมน์ result เพิ่มขึ้นมา)
        buffer1 = io.BytesIO()
        with pd.ExcelWriter(buffer1, engine='openpyxl') as writer:
            df_checked.to_excel(writer, index=False, header=False)
            
        col1.download_button(
            label="📥 ดาวน์โหลดไฟล์ตรวจเช็ค (.xlsx)",
            data=buffer1.getvalue(),
            file_name=f"checked_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # ปุ่มที่ 2: ดาวน์โหลดไฟล์ที่แก้ไขค่า pdschTypeBGain ให้ถูกต้องแล้ว
        buffer2 = io.BytesIO()
        with pd.ExcelWriter(buffer2, engine='openpyxl') as writer:
            df_fixed.to_excel(writer, index=False, header=False)
            
        col2.download_button(
            label="🛠️ ดาวน์โหลดไฟล์ที่แก้ไขค่าถูกต้องแล้ว (.xlsx)",
            data=buffer2.getvalue(),
            file_name=f"fixed_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    else:
        st.error("❌ ไม่พบคอลัมน์ 'crsGain' หรือ 'pdschTypeBGain' ในแถวที่ 4 กรุณาตรวจสอบไฟล์อีกครั้ง")
