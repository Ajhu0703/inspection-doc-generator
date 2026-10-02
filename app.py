import os
import io
import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

st.set_page_config(page_title="材料驗收文件產生器", page_icon="📋", layout="wide")

st.title("📋 課程材料驗收文件自動產生器")
st.markdown("將材料 Excel 明細與照片自動整合填入 Word 範本中，產出標準驗收紀錄表。")

# 1. 基礎資訊輸入區
st.header("1. 填寫基本資訊")
col1, col2 = st.columns(2)
with col1:
    class_name = st.text_input("課程班級名稱與編號", value="115W0149-飲調輕食暨餐飲創業(桃園)-第 1 期")
with col2:
    delivery_date = st.text_input("進貨日期 (YYY/MM/DD)", value="115/02/03")

# 2. 檔案上傳區
st.header("2. 上傳驗收資料")
col_a, col_b = st.columns(2)

with col_a:
    excel_file = st.file_uploader("上傳材料明細 Excel (.xlsx)", type=["xlsx"])
    template_file = st.file_uploader("上傳空白 Word 範本 (.docx)", type=["docx"])

with col_b:
    image_files = st.file_uploader("上傳品項照片 (可複選多張照片)", type=["jpg", "png", "jpeg"], accept_multiple_files=True)

# 替換文字的輔助函式（可同時替換一般段落與表格內的文字）
def replace_text_in_paragraph(p, new_class, new_date):
    if "115W0149" in p.text or "115W0150" in p.text or "課程班級" in p.text or "飲調輕食" in p.text or "烘焙輕食" in p.text:
        # 若段落中包含舊的班級關鍵字，更新為新輸入的班級名稱
        for run in p.runs:
            if "115W" in run.text or "飲調" in run.text or "烘焙" in run.text or "第" in run.text:
                run.text = ""
        p.text = new_class
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if "進貨日" in p.text or "進貨日期" in p.text:
        p.text = f"進貨日 {new_date}" if "進貨日 " in p.text else f"進貨日期：{new_date}"
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

# 3. 核心處理邏輯
if st.button("🚀 開始生成驗收文件", type="primary"):
    if not excel_file:
        st.error("請上傳材料明細 Excel 檔案！")
    elif not template_file:
        st.error("請上傳 Word 空白範本！")
    else:
        try:
            df = pd.read_excel(excel_file)
            
            # 整理照片列表
            img_list = []
            if image_files:
                for img in image_files:
                    base_name = os.path.splitext(img.name)[0].strip()
                    img_list.append({
                        'name': base_name,
                        'bytes': img.getvalue()
                    })

            doc = Document(template_file)

            # A. 替換一般段落中的班級與日期
            for p in doc.paragraphs:
                replace_text_in_paragraph(p, class_name, delivery_date)

            # B. 替換所有表格（包含頂部頁首表格）中的班級與日期
            for tbl in doc.tables:
                for row in tbl.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            replace_text_in_paragraph(p, class_name, delivery_date)

            # C. 尋找主要資料表格（包含「品名」或「照片」標題）
            target_table = None
            for tbl in doc.tables:
                header_text = "".join([c.text for c in tbl.rows[0].cells])
                if "品名" in header_text or "照片" in header_text:
                    target_table = tbl
                    break

            if target_table:
                target_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                
                # 自動判斷標題列各欄位位置
                hdr_cells = [cell.text.strip() for cell in target_table.rows[0].cells]
                
                idx_no = 0
                idx_name = 1
                idx_spec = 2
                idx_qty = 3
                idx_photo = 4
                idx_note = len(hdr_cells) - 1
                
                for c_i, c_text in enumerate(hdr_cells):
                    if "項" in c_text: idx_no = c_i
                    elif "品名" in c_text: idx_name = c_i
                    elif "規格" in c_text: idx_spec = c_i
                    elif "數量" in c_text: idx_qty = c_i
                    elif "照片" in c_text: idx_photo = c_i
                    elif "備註" in c_text: idx_note = c_i

                # 清理標題列底下的預設空白列
                while len(target_table.rows) > 1:
                    row_text = "".join([c.text.strip() for c in target_table.rows[1].cells])
                    if row_text == "" or "項次" not in row_text:
                        target_table._tbl.remove(target_table.rows[1]._tr)
                    else:
                        break

                # 遍歷 Excel 的每一行資料並寫入表格
                for idx, row in df.iterrows():
                    row_cells = target_table.add_row().cells
                    
                    item_no = str(row.get('品號', row.get('項次', idx + 1)))
                    item_name = str(row.get('品名', '')).strip()
                    item_spec = str(row.get('規格', '')).strip()
                    item_qty = f"{row.get('數量', '')} {row.get('單位', '')}".strip()
                    
                    # 寫入文字
                    if idx_no < len(row_cells):
                        row_cells[idx_no].text = str(item_no)
                        row_cells[idx_no].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    if idx_name < len(row_cells):
                        row_cells[idx_name].text = item_name
                    if idx_spec < len(row_cells):
                        row_cells[idx_spec].text = item_spec
                    if idx_qty < len(row_cells):
                        row_cells[idx_qty].text = item_qty
                        row_cells[idx_qty].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    if idx_note < len(row_cells):
                        row_cells[idx_note].text = ""

                    # 處理照片
                    if idx_photo < len(row_cells):
                        cell_photo = row_cells[idx_photo]
                        p_photo = cell_photo.paragraphs[0]
                        p_photo.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        
                        matching_img = None
                        idx_str = str(idx + 1)
                        no_str = str(item_no)
                        
                        for img_obj in img_list:
                            img_name = img_obj['name']
                            if (item_name and item_name in img_name) or \
                               (no_str and no_str in img_name) or \
                               (idx_str == img_name or f"照片{idx_str}" in img_name or f"序號{idx_str}" in img_name):
                                matching_img = img_obj['bytes']
                                break
                        
                        if not matching_img and idx < len(img_list):
                            matching_img = img_list[idx]['bytes']

                        if matching_img:
                            image_stream = io.BytesIO(matching_img)
                            p_photo.add_run().add_picture(image_stream, width=Cm(4.0), height=Cm(3.0))
                        else:
                            p_photo.text = "（待補照片）"
            else:
                st.error("Word 範本中找不到包含『品名』或『照片』的表格，請檢查範本格式！")

            # 導出文件
            doc_io = io.BytesIO()
            doc.save(doc_io)
            doc_io.seek(0)

            st.success("🎉 驗收文件產製完成！")
            
            filename = f"驗收資料_{class_name}_{delivery_date}.docx".replace("/", "")
            st.download_button(
                label="📥 點此下載驗收文件 (.docx)",
                data=doc_io,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        except Exception as e:
            st.error(f"處理檔案時發生錯誤：{str(e)}")
