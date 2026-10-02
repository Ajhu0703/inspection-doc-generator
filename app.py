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
    class_name = st.text_input("課程班級名稱與編號", value="115W0149-飲調輕食暨餐飲創業(桃園)-第1期")
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

# 3. 核心處理邏輯
if st.button("🚀 開始生成驗收文件", type="primary"):
    if not excel_file:
        st.error("請上傳材料明細 Excel 檔案！")
    elif not template_file:
        st.error("請上傳 Word 空白範本！")
    else:
        try:
            df = pd.read_excel(excel_file)
            
            # 整理照片列表（將檔名與照片資料存入列表）
            img_list = []
            if image_files:
                for img in image_files:
                    base_name = os.path.splitext(img.name)[0].strip()
                    img_list.append({
                        'name': base_name,
                        'bytes': img.getvalue()
                    })

            doc = Document(template_file)

            # 替換段落標籤
            for p in doc.paragraphs:
                if "課程班級" in p.text:
                    p.text = f"課程班級：{class_name}"
                if "進貨日期" in p.text:
                    p.text = f"進貨日期：{delivery_date}"

            if len(doc.tables) > 0:
                table = doc.tables[0]
                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                
                # 若表格除了標題列外已有空白列，將其清理（避免產生第一頁全空白的問題）
                while len(table.rows) > 1:
                    # 檢查第二列是否為預留空白列，若是則刪除
                    row_text = "".join([c.text.strip() for c in table.rows[1].cells])
                    if row_text == "" or "項次" not in row_text:
                        # 移除第二列
                        table._tbl.remove(table.rows[1]._tr)
                    else:
                        break

                num_cols = len(table.columns)
                
                # 遍歷 Excel 的每一行資料並寫入表格
                for idx, row in df.iterrows():
                    row_cells = table.add_row().cells
                    
                    item_no = str(row.get('品號', row.get('項次', idx + 1)))
                    item_name = str(row.get('品名', '')).strip()
                    item_spec = str(row.get('規格', '')).strip()
                    item_qty = f"{row.get('數量', '')} {row.get('單位', '')}".strip()
                    
                    # 寫入各文字欄位
                    if num_cols >= 1:
                        row_cells[0].text = str(item_no)
                        row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    if num_cols >= 2:
                        row_cells[1].text = item_name
                    if num_cols >= 3:
                        row_cells[2].text = item_spec
                    if num_cols >= 4:
                        row_cells[3].text = item_qty
                        row_cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    if num_cols >= 6:
                        row_cells[5].text = ""  # 備註欄保持空白

                    # 插入照片（通常在第 5 欄，即 index 4）
                    photo_idx = 4 if num_cols >= 5 else (num_cols - 1)
                    cell_photo = row_cells[photo_idx]
                    p_photo = cell_photo.paragraphs[0]
                    p_photo.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    # 增強版照片匹配邏輯（支援：1, 01, 照片1, 序號1, 品名匹配）
                    matching_img = None
                    idx_str = str(idx + 1)
                    no_str = str(item_no)
                    
                    for img_obj in img_list:
                        img_name = img_obj['name']
                        # 匹配條件：檔名包含品名、包含項次數字、或包含品號
                        if (item_name and item_name in img_name) or \
                           (no_str and no_str in img_name) or \
                           (idx_str == img_name or f"照片{idx_str}" in img_name or f"序號{idx_str}" in img_name):
                            matching_img = img_obj['bytes']
                            break
                    
                    # 若依然沒配對成功，若照片數量與行數相同，退回按順序配對
                    if not matching_img and idx < len(img_list):
                        matching_img = img_list[idx]['bytes']

                    if matching_img:
                        image_stream = io.BytesIO(matching_img)
                        # 限制寬度 3.2cm、高度 2.5cm，防止照片撐爆表格
                        p_photo.add_run().add_picture(image_stream, width=Cm(3.2), height=Cm(2.5))
                    else:
                        p_photo.text = "（待補照片）"

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
