import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="SKU Generator Pro", layout="wide")
st.title("📦 ระบบสร้างรหัส SKU สินค้า")

channel_map = {
    'Eat in': {'code': '01', 'abbr': 'EI'},
    'Take away': {'code': '02', 'abbr': 'TA'},
    'Delivery': {'code': '03', 'abbr': 'DL'}
}

# สร้างแท็บแยกการทำงานระหว่าง กรอกเอง กับ อัปโหลดไฟล์
tab1, tab2 = st.tabs(["📝 กรอกข้อมูลเอง (Manual)", "📥 อัปโหลดไฟล์ Excel (Import)"])

# ==========================================
# แท็บที่ 1: กรอกข้อมูลเองทีละรายการ
# ==========================================
with tab1:
    st.subheader("กรอกข้อมูลสินค้าเพื่อสร้างรหัสทันที")
    
    with st.form("manual_form", clear_on_submit=False):
        col_sku, col_name, col_channel = st.columns(3)
        
        with col_sku:
            sku_input = st.text_input("รหัส SKU เดิม (เช่น PROD001)", placeholder="กรอกรหัส SKU").strip()
        with col_name:
            product_name = st.text_input("ชื่อสินค้า (เช่น ชานม)", placeholder="กรอกชื่อสินค้า").strip()
        with col_channel:
            channel_input = st.selectbox("ช่องทางการขาย", options=list(channel_map.keys()), key="manual_channel")
            
        submit_button = st.form_submit_button(label="⚡ สร้างรหัส SKU ใหม่")

    if submit_button:
        if not sku_input or not product_name:
            st.error("⚠️ กรุณากรอกรหัส SKU และชื่อสินค้าให้ครบถ้วน")
        else:
            code = channel_map[channel_input]['code']
            abbr = channel_map[channel_input]['abbr']
            new_sku = f"{sku_input}{code}{abbr}"
            
            st.success("🎉 สร้างรหัสสินค้าสำเร็จ!")
            st.metric(label="รหัสสินค้าใหม่ (New SKU Code)", value=new_sku)
            
            # แสดงตารางสรุป
            result_data = {
                "ข้อมูล": ["SKU เดิม", "ชื่อสินค้า", "ช่องทางการขาย", "รหัสสินค้าใหม่"],
                "รายละเอียด": [sku_input, product_name, channel_input, new_sku]
            }
            st.table(pd.DataFrame(result_data))

# ==========================================
# แท็บที่ 2: อัปโหลดไฟล์ Excel
# ==========================================
with tab2:
    st.subheader("นำเข้าไฟล์ข้อมูลเพื่อประมวลผลทีละหลายรายการ")
    uploaded_file = st.file_uploader("เลือกไฟล์ Excel ของคุณ (รองรับ .xlsx)", type=["xlsx"])

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            st.write("📋 **ตัวอย่างข้อมูลที่นำเข้า:**")
            st.dataframe(df.head())
            columns = df.columns.tolist()
            
            st.subheader("⚙️ ตั้งค่าคอลัมน์ข้อมูล")
            col1, col2, col3 = st.columns(3)
            with col1:
                sku_col = st.selectbox("เลือกคอลัมน์ที่เป็น [SKUเดิม]", options=columns, key="import_sku")
            with col2:
                name_col = st.selectbox("เลือกคอลัมน์ที่เป็น [ชื่อสินค้า]", options=columns, key="import_name")
            with col3:
                channel_col = st.selectbox("เลือกคอลัมน์ที่เป็น [ช่องทางการขาย]", options=columns, key="import_channel")
                
            if st.button("⚡ เริ่มสร้างรหัส SKU จากไฟล์ ทั้งหมด"):
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
                
                st.subheader("📤 ส่งออกไฟล์ข้อมูล (Export Excel)")
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
