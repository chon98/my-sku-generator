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

tab1, tab2 = st.tabs(["📝 กรอกข้อมูลเองแบบหลายช่องทาง (Grid Manual)", "📥 อัปโหลดไฟล์ Excel (Import)"])

with tab1:
    st.subheader("📝 ตารางกรอกข้อมูลสินค้า (ระบุช่องทางการขายแยกรายแถวได้)")
    st.caption("💡 วิธีใช้งาน: ดับเบิ้ลคลิกเพื่อพิมพ์ข้อมูล / เลือกช่องทางการขายจากเมนู Dropdown / กดปุ่มด้านล่างสุดของตารางเพื่อเพิ่มแถวใหม่ (+ Add row)")
    
    init_data = [
        {"SKU เดิม": "", "ชื่อสินค้า": "", "ช่องทางการขาย": "Eat in"}
    ]
    df_init = pd.DataFrame(init_data)
    
    edited_df = st.data_editor(
        df_init,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "SKU เดิม": st.column_config.TextColumn("รหัส SKU เดิม (เช่น PROD001)", required=True),
            "ชื่อสินค้า": st.column_config.TextColumn("ชื่อสินค้า (เช่น ชานม)", required=True),
            "ช่องทางการขาย": st.column_config.SelectboxColumn(
                "เลือกช่องทางการขาย",
                options=["ทุกช่องทางการขาย"] + list(channel_map.keys()),
                required=True
            )
        }
    )
    
    if st.button("⚡ เริ่มสร้างรหัส SKU จากตารางด้านบน", key="btn_manual_grid"):
        manual_results = []
        has_error = False
        
        for index, row in edited_df.iterrows():
            sku_old = str(row["SKU เดิม"]).strip() if pd.notna(row["SKU เดิม"]) else ""
            p_name = str(row["ชื่อสินค้า"]).strip() if pd.notna(row["ชื่อสินค้า"]) else ""
            channel_input = row["ช่องทางการขาย"]
            
            if not sku_old and not p_name:
                continue
                
            if not sku_old or not p_name:
                st.error(f"⚠️ แถวที่ {index + 1} กรอกข้อมูลไม่ครบ กรุณาระบุทั้ง SKU เดิม และ ชื่อสินค้า")
                has_error = True
                break
            
            if channel_input == "ทุกช่องทางการขาย":
                for channel, info in channel_map.items():
                    new_sku = f"{sku_old}{info['code']}{info['abbr']}"
                    manual_results.append({
                        "New_Product_Code": new_sku,
                        "ชื่อสินค้า": p_name,
                        "ช่องทางการขาย": channel,
                        "_Sort_Code": info['code'],
                        "SKU เดิม": sku_old
                    })
            else:
                info = channel_map[channel_input]
                new_sku = f"{sku_old}{info['code']}{info['abbr']}"
                manual_results.append({
                    "New_Product_Code": new_sku,
                    "ชื่อสินค้า": p_name,
                    "ช่องทางการขาย": channel_input,
                    "_Sort_Code": info['code'],
                    "SKU เดิม": sku_old
                })
        
        if not has_error and len(manual_results) > 0:
            df_manual_res = pd.DataFrame(manual_results)
            # เรียงลำดับตามช่องทางการขาย และ SKU เดิม จากนั้นจัดคอลัมน์ตามลำดับที่ต้องการ
            df_manual_res = df_manual_res.sort_values(by=['_Sort_Code', 'SKU เดิม']).reset_index(drop=True)
            output_cols = ["New_Product_Code", "ชื่อสินค้า", "ช่องทางการขาย"]
            df_manual_final = df_manual_res[output_cols]
            
            st.success(f"🎉 สร้างรหัสสินค้าสำเร็จทั้งหมด {len(df_manual_final)} รายการ!")
            
            st.write("📋 **รหัสสินค้าใหม่ทั้งหมด (คลิกไอคอนสี่เหลี่ยมซ้อนกันที่มุมขวาบนของกล่องเพื่อคัดลอกทั้งหมด):**")
            all_new_skus_text = "\n".join(df_manual_final["New_Product_Code"].tolist())
            st.code(all_new_skus_text, language="text")
            
            st.write("📊 **ตารางสรุปรายละเอียด (เรียงลำดับคอลัมน์เรียบร้อย):**")
            st.dataframe(df_manual_final, use_container_width=True)
            
            # เพิ่มปุ่ม Export Excel สำหรับฝั่ง Manual ให้ด้วย
            buffer_manual = io.BytesIO()
            with pd.ExcelWriter(buffer_manual, engine='openpyxl') as writer:
                df_manual_final.to_excel(writer, index=False, sheet_name='Manual_SKU')
            
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ Excel ผลลัพธ์ (จากตารางกรอกเอง)",
                data=buffer_manual.getvalue(),
                file_name="manual_sku_products.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        elif not has_error and len(manual_results) == 0:
            st.warning("⚠️ ไม่พบข้อมูลในตาราง กรุณากรอกรหัสสินค้าก่อนกดปุ่ม")

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
                    
                    for channel, info in channel_map.items():
                        new_sku = f"{sku_old}{info['code']}{info['abbr']}"
                        
                        new_row = {
                            'New_Product_Code': new_sku,
                            'ชื่อสินค้า': p_name,
                            'ช่องทางการขาย': channel,
                            '_Sort_Code': info['code'],
                            '_Sku_Old': sku_old
                        }
                        expanded_rows.append(new_row)
                
                df_result = pd.DataFrame(expanded_rows)
                # เรียงลำดับตามช่องทางการขาย และ SKU เดิม จากนั้นจัดคอลัมน์ตามลำดับที่ต้องการ
                df_result = df_result.sort_values(by=['_Sort_Code', '_Sku_Old']).reset_index(drop=True)
                
                output_cols = ["New_Product_Code", "ชื่อสินค้า", "ช่องทางการขาย"]
                df_final_export = df_result[output_cols]
                
                st.success("🎉 ระบบแตกไฟล์ข้อมูลและสร้างรหัสทุกช่องทางการขายเสร็จเรียบร้อย!")
                st.write("📋 **ผลลัพธ์ข้อมูลใหม่ (เรียงลำดับคอลัมน์เรียบร้อย):**")
                st.dataframe(df_final_export, use_container_width=True)
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_final_export.to_excel(writer, index=False, sheet_name='All_Channels_SKU')
                
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
