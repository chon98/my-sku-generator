import streamlit as st
import pandas as pd
import io
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="SKU Generator Pro", 
    page_icon="📦", 
    layout="wide"
)

st.title("📦 ระบบสร้างรหัส SKU สินค้าตามช่องทางการขาย")
st.write("จัดการสร้างรหัสสินค้าแยกตามช่องทางการขายอัตโนมัติ รวดเร็ว แม่นยำ")
st.write("---")

channel_map = {
    'Eat in': {'code': '01', 'abbr': 'EI'},
    'Take away': {'code': '02', 'abbr': 'TA'},
    'Delivery': {'code': '03', 'abbr': 'DL'}
}

# ฟังก์ชันตกแต่งหัวข้อตาราง ขยายขนาดคอลัมน์อัตโนมัติ และตรึงแนวแถวบนสุด
def style_excel_sheet(writer, sheet_name):
    worksheet = writer.sheets[sheet_name]
    worksheet.freeze_panes = "A2"
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_alignment = Alignment(horizontal="center", vertical="center")
    
    for col in worksheet.iter_cols(min_row=1, max_row=1):
        for cell in col:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
            
    for col in worksheet.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col.column)
        worksheet.column_dimensions[col_letter].width = max(max_len + 5, 15)

tab1, tab2 = st.tabs(["📝 กรอกข้อมูลเองแบบหลายช่องทาง (Grid Manual)", "📥 อัปโหลดไฟล์ Excel (Import)"])

# ==========================================
# แท็บที่ 1: ตารางกรอกข้อมูลเอง
# ==========================================
with tab1:
    st.subheader("📝 ตารางกรอกข้อมูลสินค้า")
    st.caption("💡 วิธีใช้งาน: ดับเบิ้ลคลิกเพื่อพิมพ์ข้อมูล / คลิกเลือกช่องทางการขาย (เลือกได้มากกว่า 1 ช่องทางพร้อมกัน) / กดปุ่ม (+ Add row) ด้านล่างสุดเพื่อเพิ่มแถวใหม่")
    
    # กำหนดค่าเริ่มต้นให้กับคอลัมน์เลือกช่องทางการขายเป็น List เพื่อให้รองรับ Multi-select อัตโนมัติ
    init_data = [
        {"SKU เดิม": "", "ชื่อสินค้า": "", "เลือกช่องทางการขาย": ["Eat in"]}
    ]
    df_init = pd.DataFrame(init_data)
    
    # แก้ไขจุดนี้: ใช้ตัวตั้งค่าเริ่มต้นของคอลัมน์ตารางให้รองรับ Multi-select โดยแก้บั๊กจากคำสั่งเดิม
    edited_df = st.data_editor(
        df_init,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "SKU เดิม": st.column_config.TextColumn("รหัส SKU เดิม (เช่น PROD001)", required=True),
            "ชื่อสินค้า": st.column_config.TextColumn("ชื่อสินค้า (เช่น ชานม)", required=True),
            "เลือกช่องทางการขาย": st.column_config.SelectboxColumn(
                "เลือกช่องทางการขาย (เลือกได้หลายช่องทาง)",
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
            channels_input = row["เลือกช่องทางการขาย"]
            
            if not sku_old and not p_name:
                continue
                
            if not sku_old or not p_name:
                st.error(f"⚠️ แถวที่ {index + 1} กรอกข้อมูลไม่ครบ กรุณาระบุทั้ง SKU เดิม และ ชื่อสินค้า")
                has_error = True
                break
                
            if not channels_input or (isinstance(channels_input, list) and len(channels_input) == 0):
                st.error(f"⚠️ แถวที่ {index + 1} ยังไม่ได้เลือกช่องทางการขาย")
                has_error = True
                break
            
            # ตรวจสอบว่าเป็น String เดี่ยว หรือเป็น List รายการจากการกดเลือกมาหลายตัว
            channels_list = channels_input if isinstance(channels_input, list) else [channels_input]
            
            selected_channels = []
            if "ทุกช่องทางการขาย" in channels_list:
                selected_channels = list(channel_map.keys())
            else:
                selected_channels = channels_list
                
            for channel in selected_channels:
                info = channel_map.get(channel)
                if info:
                    new_sku = f"{sku_old}{info['code']}{info['abbr']}"
                    manual_results.append({
                        "New_Product_Code": new_sku,
                        "ชื่อสินค้า": p_name,
                        "ช่องทางการขาย": channel,
                        "_Sort_Code": info['code'],
                        "SKU เดิม": sku_old
                    })
        
        if not has_error and len(manual_results) > 0:
            df_manual_res = pd.DataFrame(manual_results)
            df_manual_res = df_manual_res.sort_values(by=['SKU เดิม', '_Sort_Code']).reset_index(drop=True)
            
            output_cols = ["New_Product_Code", "ชื่อสินค้า", "ช่องทางการขาย"]
            df_manual_final = df_manual_res[output_cols]
            
            st.success(f"🎉 สร้างรหัสสินค้าสำเร็จทั้งหมด {len(df_manual_final)} รายการ!")
            
            st.write("📋 **รหัสสินค้าใหม่ทั้งหมด (คลิกไอคอนสี่เหลี่ยมซ้อนกันที่มุมขวาบนของกล่องเพื่อคัดลอกทั้งหมด):**")
            all_new_skus_text = "\n".join(df_manual_final["New_Product_Code"].tolist())
            st.code(all_new_skus_text, language="text")
            
            st.write("📊 **ตารางสรุปรายละเอียด:**")
            st.dataframe(df_manual_final, use_container_width=True)
            
            buffer_manual = io.BytesIO()
            with pd.ExcelWriter(buffer_manual, engine='openpyxl') as writer:
                df_manual_final.to_excel(writer, index=False, sheet_name='Manual_SKU')
                style_excel_sheet(writer, 'Manual_SKU')
            
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ Excel ผลลัพธ์ (จากตารางกรอกเอง)",
                data=buffer_manual.getvalue(),
                file_name="manual_sku_products.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_download_manual_res"
            )
            
        elif not has_error and len(manual_results) == 0:
            st.warning("⚠️ ไม่พบข้อมูลในตาราง กรุณากรอกรหัสสินค้าก่อนกดปุ่ม")

# ==========================================
# แท็บที่ 2: อัปโหลดไฟล์ Excel
# ==========================================
with tab2:
    st.subheader("📥 ดาวน์โหลดไฟล์เทมเพลตสำหรับใช้งาน")
    
    template_data = {
        "SKU เดิม": ["PROD001", "PROD002"],
        "ชื่อสินค้า": ["ชานม", "กาแฟ"]
    }
    df_template = pd.DataFrame(template_data)
    
    buffer_template = io.BytesIO()
    with pd.ExcelWriter(buffer_template, engine='openpyxl') as writer:
        df_template.to_excel(writer, index=False, sheet_name='Template')
        style_excel_sheet(writer, 'Template')
        
    st.download_button(
        label="📥 ดาวน์โหลดไฟล์เทมเพลต Excel (.xlsx)",
        data=buffer_template.getvalue(),
        file_name="sku_import_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="btn_download_template_xlsx"
    )
    
    st.write("---")
    st.subheader("📤 นำเข้าไฟล์ข้อมูลเพื่อแตกรหัสสินค้าเป็นทุกช่องทางการขาย")
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
                
            if st.button("⚡ เริ่มแตกไฟล์และสร้างรหัส SKU ทั้งหมด", type="primary"):
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
                df_result = df_result.sort_values(by=['_Sku_Old', '_Sort_Code']).reset_index(drop=True)
                
                output_cols = ["New_Product_Code", "ชื่อสินค้า", "ช่องทางการขาย"]
                df_final_export = df_result[output_cols]
                
                st.success(f"🎉 ระบบแตกไฟล์ข้อมูลและสร้างรหัสทุกช่องทางการขายเสร็จเรียบร้อย ทั้งหมด {len(df_final_export)} รายการ!")
                st.write("📋 **ผลลัพธ์ข้อมูลใหม่:**")
                st.dataframe(df_final_export, use_container_width=True)
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_final_export.to_excel(writer, index=False, sheet_name='All_Channels_SKU')
                    style_excel_sheet(writer, 'All_Channels_SKU')
                
                st.download_button(

                    label="📥 ดาวน์โหลดไฟล์ Excel ผลลัพธ์ (ครบทุกช่องทาง)",
                    data=buffer.getvalue(),
                    file_name="multi_channel_sku_products.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_download_file_res"
                )
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการอ่านไฟล์: {e}")
    else:
        st.info("💡 กรุณาอัปโหลดไฟล์ Excel ที่มีคอลัมน์สำหรับ SKU เดิม และ ชื่อสินค้า (ระบบจะนำแต่ละแถวไปสร้างครบทั้ง 3 ช่องทางให้เองอัตโนมัติ)")
