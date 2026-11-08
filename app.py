import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel SKU Generator", layout="wide")
st.title("📦 ระบบสร้างรหัส SKU จากไฟล์ Excel")

channel_map = {
    'Eat in': {'code': '01', 'abbr': 'EI'},
    'Take away': {'code': '02', 'abbr': 'TA'},
    'Delivery': {'code': '03', 'abbr': 'DL'}
}

st.subheader("📥 1. นำเข้าไฟล์ข้อมูล (Import Excel)")
uploaded_file = st.file_uploader("เลือกไฟล์ Excel ของคุณ (รองรับ .xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.write("📋 **ตัวอย่างข้อมูลที่นำเข้า:**")
        st.dataframe(df.head())
        columns = df.columns.tolist()
        
        st.subheader("⚙️ 2. ตั้งค่าคอลัมน์ข้อมูล")
        col1, col2, col3 = st.columns(3)
        with col1:
            sku_col = st.selectbox("เลือกคอลัมน์ที่เป็น [SKUเดิม]", options=columns)
        with col2:
            name_col = st.selectbox("เลือกคอลัมน์ที่เป็น [ชื่อสินค้า]", options=columns)
        with col3:
            channel_col = st.selectbox("เลือกคอลัมน์ที่เป็น [ช่องทางการขาย]", options=columns)
            
        if st.button("⚡ เริ่มสร้างรหัส SKU ใหม่ทั้งหมด"):
            df_process = df.copy()
            
            def generate_new_sku(row):
                channel = str(row[channel_col]).strip()
                sku_old = str(row[sku_col]).strip()
                if channel in channel_map:
                    code = channel_map[channel]['code']
                    abbr = channel_map[channel]['abbr']
                    return f"{sku_old}{code}{abbr}"
                else:
                    return f"{sku_old}-ERROR"

            def get_channel_code(channel):
                channel = str(channel).strip()
                return channel_map[channel]['code'] if channel in channel_map else '99'

            df_process['New_Product_Code'] = df_process.apply(generate_new_sku, axis=1)
            df_process['_Sort_Code'] = df_process[channel_col].apply(get_channel_code)
            
            df_result = df_process.sort_values(by='_Sort_Code').reset_index(drop=True)
            df_result = df_result.drop(columns=['_Sort_Code'])
            
            st.success("🎉 ประมวลผลและเรียงลำดับข้อมูลเสร็จเรียบร้อย!")
            st.write("📋 **ผลลัพธ์ข้อมูลใหม่:**")
            st.dataframe(df_result)
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_result.to_excel(writer, index=False, sheet_name='New_SKU')
            
            st.subheader("📤 3. ส่งออกไฟล์ข้อมูล (Export Excel)")
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ Excel ผลลัพธ์",
                data=buffer.getvalue(),
                file_name="processed_sku_products.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการอ่านไฟล์: {e}")
else:
    st.info("💡 กรุณาอัปโหลดไฟล์ Excel ที่มีคอลัมน์สำหรับ SKU เดิม, ชื่อสินค้า และช่องทางการขาย (เช่น Eat in, Take away, Delivery)")
