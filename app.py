import streamlit as st
import pandas as pd
import io
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# 1. ตั้งค่าหน้าเว็บให้รองรับธีมของระบบโดยสมบูรณ์
st.set_page_config(
    page_title="SKU Generator Pro", 
    page_icon="📦", 
    layout="wide"
)

# คลังคำแปลระบบ 2 ภาษา (ไทย / อังกฤษ)
translation = {
    'TH': {
        'title': "📦 ระบบบริหารจัดการและสร้างรหัส SKU สินค้า",
        'subtitle': "ปรับปรุงรหัสสินค้าแยกตามช่องทางการขายอัตโนมัติ รวดเร็ว แม่นยำ",
        'tab1': "📝 กรอกข้อมูลรายรายการ (Grid Manual)",
        'tab2': "📥 ประมวลผลผ่านไฟล์ (Import Excel)",
        'grid_title': "📝 ตารางกรอกข้อมูลสินค้า",
        'grid_caption': "💡 วิธีใช้งาน: ดับเบิ้ลคลิกเพื่อพิมพ์ข้อมูล / เลือกช่องทางการขายจากเมนู Dropdown / กดปุ่ม (+ Add row) ด้านล่างสุดของตารางเพื่อเพิ่มแถวใหม่",
        'col_old_sku': "รหัส SKU เดิม (เช่น PROD001)",
        'col_prod_name': "ชื่อสินค้า (เช่น ชานม)",
        'col_channel': "เลือกช่องทางการขาย",
        'all_channels': "ทุกช่องทางการขาย",
        'btn_generate': "⚡ เริ่มสร้างรหัส SKU ใหม่",
        'err_incomplete': "⚠️ แถวที่ {} กรอกข้อมูลไม่ครบ กรุณาระบุทั้ง SKU เดิม และ ชื่อสินค้า",
        'err_empty': "⚠️ ไม่พบข้อมูลในตาราง กรุณากรอกรหัสสินค้าก่อนกดปุ่ม",
        'success_gen': "🎉 สร้างรหัสสินค้าสำเร็จทั้งหมด {} รายการ!",
        'copy_title': "📋 รหัสสินค้าใหม่ทั้งหมด (คลิกไอคอนสี่เหลี่ยมซ้อนกันที่มุมขวาบนของกล่องเพื่อคัดลอกทั้งหมด):",
        'summary_title': "📊 ตารางสรุปรายละเอียด:",
        'btn_download_res': "📥 ดาวน์โหลดไฟล์ Excel ผลลัพธ์",
        'dl_template_title': "📥 ดาวน์โหลดไฟล์เทมเพลตสำหรับใช้งาน",
        'btn_download_tpl': "📥 ดาวน์โหลดไฟล์เทมเพลต Excel (.xlsx)",
        'import_title': "📤 นำเข้าไฟล์ข้อมูลเพื่อแตกรหัสสินค้าเป็นทุกช่องทางการขาย",
        'import_uploader': "เลือกไฟล์ Excel ของคุณ (รองรับ .xlsx)",
        'preview_import': "📋 ตัวอย่างข้อมูลที่นำเข้าจากไฟล์:",
        'setting_import': "⚙️ ตั้งค่าจับคู่คอลัมน์ข้อมูล",
        'select_old_sku': "เลือกคอลัมน์ที่เป็น [SKUเดิม]",
        'select_prod_name': "เลือกคอลัมน์ที่เป็น [ชื่อสินค้า]",
        'btn_process_file': "⚡ เริ่มแตกไฟล์และสร้างรหัส SKU ทั้งหมด",
        'info_upload': "💡 กรุณาอัปโหลดไฟล์ Excel ที่มีคอลัมน์สำหรับ SKU เดิม และ ชื่อสินค้า (ระบบจะนำแต่ละแถวไปสร้างครบทั้ง 3 ช่องทางให้เองอัตโนมัติ)",
        'col_header_sku': "New_Product_Code",
        'col_header_name': "ชื่อสินค้า",
        'col_header_chan': "ช่องทางการขาย",
        'tpl_sku': "SKU เดิม",
        'tpl_name': "ชื่อสินค้า"
    },
    'EN': {
        'title': "📦 Multi-Channel SKU Generator Pro",
        'subtitle': "Transform and optimize your product SKUs by sales channels seamlessly.",
        'tab1': "📝 Manual Entry Grid",
        'tab2': "📥 Bulk Import Excel",
        'grid_title': "📝 Product Information Grid",
        'grid_caption': "💡 Instructions: Double-click to edit cell / Select channel from dropdown / Click (+ Add row) at the bottom to insert a new row.",
        'col_old_sku': "Original SKU (e.g., PROD001)",
        'col_prod_name': "Product Name (e.g., Milk Tea)",
        'col_channel': "Select Channel",
        'all_channels': "All Channels",
        'btn_generate': "⚡ Generate New SKUs",
        'err_incomplete': "⚠️ Row {} is incomplete. Please provide both Original SKU and Product Name.",
        'err_empty': "⚠️ Grid is empty. Please enter product data before generating.",
        'success_gen': "🎉 Successfully generated {} SKU items!",
        'copy_title': "📋 All New SKUs (Click the copy icon on the top-right of the block to copy all):",
        'summary_title': "📊 Summary Details Table:",
        'btn_download_res': "📥 Download Excel Result",
        'dl_template_title': "📥 Download Standard Template",
        'btn_download_tpl': "📥 Download Excel Template (.xlsx)",
        'import_title': "📤 Import Excel File to Expand All Channels",
        'import_uploader': "Choose your Excel file (Supports .xlsx)",
        'preview_import': "📋 Imported Data Preview:",
        'setting_import': "⚙️ Column Mapping Settings",
        'select_old_sku': "Map column for [Original SKU]",
        'select_prod_name': "Map column for [Product Name]",
        'btn_process_file': "⚡ Process File & Generate All SKUs",
        'info_upload': "💡 Please upload an Excel file containing Original SKU and Product Name columns (System will auto-expand 3 channels for every item).",
        'col_header_sku': "New_Product_Code",
        'col_header_name': "Product_Name",
        'col_header_chan': "Sales_Channel",
        'tpl_sku': "Original SKU",
        'tpl_name': "Product Name"
    }
}

# 2. จัดวางส่วนหัวของเว็บไซต์และเมนูสลับภาษาอย่างเป็นระเบียบ
header_left, header_right = st.columns([3, 1])
with header_right:
    lang = st.selectbox("🌐 Language / ภาษา", options=["TH", "EN"], index=0)

lang_dict = translation[lang]

with header_left:
    st.title(lang_dict['title'])
    st.write(lang_dict['subtitle'])

st.write("---")

channel_map = {
    'Eat in': {'code': '01', 'abbr': 'EI'},
    'Take away': {'code': '02', 'abbr': 'TA'},
    'Delivery': {'code': '03', 'abbr': 'DL'}
}

# ฟังก์ชันตกแต่งไฟล์ Excel
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

# 3. แท็บสำหรับการประมวลผลข้อมูล
tab1, tab2 = st.tabs([lang_dict['tab1'], lang_dict['tab2']])

# ==========================================
# แท็บที่ 1: ตารางกรอกข้อมูลเอง (Grid Manual)
# ==========================================
with tab1:
    st.subheader(lang_dict['grid_title'])
    st.caption(lang_dict['grid_caption'])
    
    init_data = [
        {lang_dict['tpl_sku']: "", lang_dict['tpl_name']: "", lang_dict['col_channel']: "Eat in"}
    ]
    df_init = pd.DataFrame(init_data)
    
    edited_df = st.data_editor(
        df_init,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            lang_dict['tpl_sku']: st.column_config.TextColumn(lang_dict['col_old_sku'], required=True),
            lang_dict['tpl_name']: st.column_config.TextColumn(lang_dict['col_prod_name'], required=True),
            lang_dict['col_channel']: st.column_config.SelectboxColumn(
                lang_dict['col_channel'],
                options=[lang_dict['all_channels']] + list(channel_map.keys()),
                required=True
            )
        }
    )
    
    if st.button(lang_dict['btn_generate'], key="btn_manual_grid", type="primary"):
        manual_results = []
        has_error = False
        
        for index, row in edited_df.iterrows():
            sku_old = str(row[lang_dict['tpl_sku']]).strip() if pd.notna(row[lang_dict['tpl_sku']]) else ""
            p_name = str(row[lang_dict['tpl_name']]).strip() if pd.notna(row[lang_dict['tpl_name']]) else ""
            channel_input = row[lang_dict['col_channel']]
            
            if not sku_old and not p_name:
                continue
                
            if not sku_old or not p_name:
                st.error(lang_dict['err_incomplete'].format(index + 1))
                has_error = True
                break
            
            if channel_input == lang_dict['all_channels']:
                for channel, info in channel_map.items():
                    new_sku = f"{sku_old}{info['code']}{info['abbr']}"
                    manual_results.append({
                        lang_dict['col_header_sku']: new_sku,
                        lang_dict['col_header_name']: p_name,
                        lang_dict['col_header_chan']: channel,
                        "_Sort_Code": info['code'],
                        "_Sku_Old": sku_old
                    })
            else:
                actual_channel = channel_input
                info = channel_map.get(actual_channel, channel_map['Eat in'])
                new_sku = f"{sku_old}{info['code']}{info['abbr']}"
                manual_results.append({
                    lang_dict['col_header_sku']: new_sku,
                    lang_dict['col_header_name']: p_name,
                    lang_dict['col_header_chan']: actual_channel,
                    "_Sort_Code": info['code'],
                    "_Sku_Old": sku_old
                })
        
        if not has_error and len(manual_results) > 0:
            df_manual_res = pd.DataFrame(manual_results)
            df_manual_res = df_manual_res.sort_values(by=['_Sku_Old', '_Sort_Code']).reset_index(drop=True)
            output_cols = [lang_dict['col_header_sku'], lang_dict['col_header_name'], lang_dict['col_header_chan']]
            df_manual_final = df_manual_res[output_cols]
            
            st.success(lang_dict['success_gen'].format(len(df_manual_final)))
            
            st.write(lang_dict['copy_title'])
            all_new_skus_text = "\n".join(df_manual_final[lang_dict['col_header_sku']].tolist())
            st.code(all_new_skus_text, language="text")
            
            st.write(lang_dict['summary_title'])
            st.dataframe(df_manual_final, use_container_width=True)
            
            buffer_manual = io.BytesIO()
            with pd.ExcelWriter(buffer_manual, engine='openpyxl') as writer:
                df_manual_final.to_excel(writer, index=False, sheet_name='Manual_SKU')
