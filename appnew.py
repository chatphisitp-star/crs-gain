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
    # อ่านไฟล์ Excel โดยใช้แถวที่ 4 (header=3) เป็นชื่อคอลัมน์
    df = pd.read_excel(uploaded_file, header=3)
    
    # ตัดช่องว่างข้างหน้า/ข้างหลังของชื่อคอลัมน์ออกทั้งหมดเพื่อป้องกันพิมพ์เกิน
    df.columns = df.columns.astype(str).str.strip()
    
    # ระบุชื่อคอลัมน์เป้าหมาย
    crs_col = 'crsGain'
    pdsch_col = 'pdschTypeBGain'
    
    # ตรวจสอบว่าพบคอลัมน์ที่ต้องการหรือไม่
    if crs_col in df.columns and pdsch_col in df.columns:
        # ประมวลผลสร้างคอลัมน์ result
        df['result'] = [validate_row(c, p) for c, p in zip(df[crs_col], df[pdsch_col])]
        
        # แสดงตัวอย่างตารางผลลัพธ์บนหน้าเว็บ
        st.success(f"✅ ตรวจสอบข้อมูลเรียบร้อยแล้ว! (ดึงข้อมูลจากคอลัมน์ '{crs_col}' และ '{pdsch_col}')")
        st.dataframe(df.head(10))
        
        # ปุ่มกดดาวน์โหลดไฟล์ผลลัพธ์
        output_name = f"checked_{uploaded_file.name}"
        
        # แปลง dataframe เป็น excel สำหรับดาวน์โหลด
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
            
        st.download_button(
            label="📥 กดเพื่อดาวน์โหลดไฟล์ผลลัพธ์ (.xlsx)",
            data=buffer.getvalue(),
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error(f"❌ ไม่พบคอลัมน์ '{crs_col}' หรือ '{pdsch_col}' ในแถวที่ 4 ของไฟล์ Excel กรุณาตรวจสอบชื่อคอลัมน์อีกครั้ง")
