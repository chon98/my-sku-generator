import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="SKU Generator Pro", layout="wide")
st.title("📦 ระบบสร้างรหัส SKU สินค้าตามช่องทางการขาย")

channel_map = {
    'Eat in': {'code': '01', 'abbr': 'EI'},
    'Take away': {'code': '02', 'abbr': 'TA'},
    'Delivery': {'code': '03', 'abbr': 'DL'}
}

tab1, tab2 = st.tabs(["📝 กรอกข้อมูลเอง (Manual)", "📥 อัปโหลดไฟล์ Excel (Import)"])

with tab1:
    st.subheader("กรอกข้อมูลสินค้าเพื่อสร้างรหัส SKU")
    
    with st.form("manual_form", clear_on_submit=False):
        col_sku, col_name, col_channel = st.columns(3)
        
        with col_sku:
            sku_input = st.text_input("รหัส SKU เดิม (เช่น PROD001)", placeholder="กรอกรหัส SKU").strip()
        with col_name:
            product_name = st.text_input("ชื่อสินค้า (เช่น ชานม)", placeholder="กรอกชื่อสินค้า").strip()
        with col_channel:
            # เพิ่มตัวเลือก "ทุกช่องทางการขาย" เข้าไปใน List
            channel_options = ["ทุกช่องทางการขาย"] + list(channel_map.keys())
            channel_input = st.selectbox("เลือกช่องทางการขาย", options=channel_options, key="manual_channel")
            
        submit_button = st.form_submit_button(label="⚡ สร้างรหัส SKU")

    if submit_button:
        if not sku_input or not product_name:
            st.error("⚠️ กรุณากรอกรหัส SKU และชื่อสินค้าให้ครบถ้วน")
        else:
            manual_results = []
            
            # กรณีเลือกทำทุกช่องทาง
            if channel_input == "ทุกช่องทางการขาย":
                for channel, info in channel_map.items():
                    new_sku = f"{sku_input}{info['code']}{info['abbr']}"
                    manual_results.append({
                        "SKU เดิม": sku_input,
                        "ชื่อสินค้า": product_name,
                        "ช่องทางการขาย": channel,
                        "รหัสสินค้าใหม่": new_sku
                    })
                st.success("🎉 สร้างรหัสสินค้าครบทุกช่องทางการขายสำเร็จ!")
            
            # กรณีเลือกช่องทางใดช่องทางหนึ่งเจาะจง
            else:
                info = channel_map[channel_input]
                new_sku = f"{sku_input}{info['code']}{info['abbr']}"
                manual_results.append({
                    "SKU เดิม": sku_input,
                    "ชื่อสินค้า": product_name,
                    "ช่องทางการขาย": channel_input,
                    "รหัสสินค้าใหม่": new_sku
                })
                st.success(f"🎉 สร้างรหัสสินค้าสำหรับ {channel_input} สำเร็จ!")
            
            df_manual_res = pd.DataFrame(manual_results)
            st.table(df_manual_res)

with tab2:
    st.subheader("นำเข้าไฟล์ข้อมูลเพื่อแตกรหัสสินค้าเป็นทุกช่องทางการขาย")
    uploaded_file = st.file_uploader("เลือกไฟล์ Excel ของคุณ (รองรับ .xlsx)", type=["xlsx"])

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            st.write("📋 **ตัวอย่างข้อมูลที่นำเข้า:**")
            st.dataframe(df.head())
            columns = df.columns.tolist()
            
            st.subheader("⚙️ ตั้งค่าคอลัมน์ข้อมูล")
            col1, col2 = st.columns(2)
            with col1:
                sku_col = st.selectbox("เลือกคอลัมน์ที่เป็น [SKUเดิม]", options=columns, key="import_sku")
            with col2:
                name_col = st.selectbox("เลือกคอลัมน์ที่เป็น [ชื่อสินค้า]", options=columns, key="import_name")
                
            if st.button("⚡ เริ่มแตกไฟล์และสร้างรหัส SKU ทั้งหมด"):
                expanded_rows = []
                
                for _, row in df.iterrows():
                    sku_old = str(row[sku_col]).strip()
                    p_name = str(row[name_col]).strip()
                    
                    other_data = row.drop([sku_col, name_col]).to_dict()
                    
                    for channel, info in channel_map.items():
                        new_sku = f"{sku_old}{info['code']}{info['abbr']}"
                        
                        new_row = {
                            sku_col: sku_old,
                            name_col: p_name,
                            'Channel': channel,
                            'New_Product_Code': new_sku,
                            '_Sort_Code': info['code']
                        }
                        
                        new_row.update(other_data)
                        expanded_rows.append(new_row)
                
                df_result = pd.DataFrame(expanded_rows)
                df_result = df_result.sort_values(by=['_Sort_Code', sku_col]).reset_index(drop=True)
                df_result = df_result.drop(columns=['_Sort_Code'])
                
                st.success("🎉 ระบบแตกไฟล์ข้อมูลและสร้างรหัสทุกช่องทางการขายเสร็จเรียบร้อย!")
                st.write("📋 **ผลลัพธ์ข้อมูลใหม่:**")
                st.dataframe(df_result)
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_result.to_excel(writer, index=False, sheet_name='All_Channels_SKU')
                
                st.subheader("📤 ส่งออกไฟล์ข้อมูล (Export Excel)")
                st.download_button(
                    label="📥 ดาวน์โหลดไฟล์ Excel ผลลัพธ์ (ครบทุกช่องทาง)",
                    data=buffer.getvalue(),
                    file_name="multi_channel_sku_products.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการอ่านไฟล์: {e}")
    else:
        st.info("💡 กรุณาอัปโหลดไฟล์ Excel ที่มีคอลัมน์สำหรับ SKU เดิม และ ชื่อสินค้า (ระบบจะนำแต่ละแถวไปสร้างครบทั้ง 3 ช่องทางให้เองอัตโนมัติ)")
