import streamlit as st
import openpyxl
import pdfplumber
import fitz
import io
import os
from datetime import datetime

st.set_page_config(page_title="通行費對帳工具", layout="wide")
st.title("通行費自動對帳與標註工具")

if 'processed_excel' not in st.session_state: st.session_state.processed_excel = None
if 'processed_pdf' not in st.session_state: st.session_state.processed_pdf = None

uploaded_pdf = st.file_uploader("1. 上傳遠通電收 PDF", type="pdf")
uploaded_excel = st.file_uploader("2. 上傳完整的 T_E 申請表 (含多個月份)", type="xlsx")

def format_date(date_val):
    try:
        if isinstance(date_val, str): return datetime.strptime(date_val, '%d-%b-%y').strftime('%Y/%m/%d')
        if hasattr(date_val, 'strftime'): return date_val.strftime('%Y/%m/%d')
        return str(date_val)
    except: return None

# --- 新增：工作表選擇邏輯 ---
selected_sheet = None
if uploaded_excel:
    # 預先讀取 workbook 取得工作表清單
    temp_wb = openpyxl.load_workbook(uploaded_excel, read_only=True)
    sheet_names = temp_wb.sheetnames
    
    # 智慧預設值：抓取當前月份，例如 "4月"
    current_month_str = f"{datetime.now().month}月"
    default_idx = sheet_names.index(current_month_str) if current_month_str in sheet_names else 0
    
    selected_sheet = st.selectbox("3. 選擇要處理的月份工作表", sheet_names, index=default_idx)
    st.info(f"目前選擇處理：**{selected_sheet}**")

if uploaded_pdf and uploaded_excel and selected_sheet:
    if st.button("🚀 開始處理"):
        try:
            # --- 終極字體偵測 ---
            font_path = None
            for file in os.listdir("."):
                if file.lower().endswith((".ttc", ".ttf")):
                    font_path = file
                    break
            if not font_path:
                for root, dirs, files in os.walk("."):
                    for file in files:
                        if file.lower().endswith((".ttc", ".ttf")):
                            font_path = os.path.join(root, file)
                            break
                    if font_path: break
            
            prefix = "項目 " if font_path else "No. "
            pdf_bytes = uploaded_pdf.getvalue()
            
            # --- 修改處：讀取指定的 sheet ---
            wb = openpyxl.load_workbook(uploaded_excel)
            ws = wb[selected_sheet] 
            
            # 1. 處理 Excel
            header_row = 7 
            date_col_idx, toll_col_idx, item_col_idx = 4, 11, 1
            
            toll_map, serial_map = {}, {}
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        for line in text.split('\n'):
                            parts = line.split()
                            if len(parts) >= 3 and '/' in parts[0]:
                                if parts[2].replace('元', '').isdigit():
                                    toll_map[parts[0]] = int(parts[2].replace('元', ''))
            
            for row in range(header_row + 1, ws.max_row + 1):
                raw_date = ws.cell(row=row, column=date_col_idx).value
                if not raw_date: continue
                
                d_str = format_date(raw_date)
                
                if d_str in toll_map and toll_map[d_str] is not None:
                    ws.cell(row=row, column=toll_col_idx).value = toll_map[d_str]
                    item_val = ws.cell(row=row, column=item_col_idx).value
                    if item_val is not None:
                        try:
                            clean_item = int(float(item_val))
                            serial_map[d_str] = f"{prefix}{clean_item:02d}"
                        except:
                            serial_map[d_str] = f"{prefix}{item_val}"
                    toll_map[d_str] = None 
            
            out_excel = io.BytesIO()
            wb.save(out_excel)
            st.session_state.processed_excel = out_excel.getvalue()

            # 2. PDF 標註
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            for page in doc:
                words = page.get_text("words")
                if font_path:
                    page.insert_font(fontname="custom_font", fontfile=font_path)

                for w in words:
                    if w[4] in serial_map:
                        date_w = w
                        line_words =[lw for lw in words if abs(lw[1] - date_w[1]) < 5]
                        line_words.sort(key=lambda x: x[0])
                        
                        km_w, toll_w = None, None
                        for idx, lw in enumerate(line_words):
                            if "公里" in lw[4]:
                                km_w = lw
                                if idx + 1 < len(line_words):
                                    toll_w = line_words[idx + 1] 
                                break
                        
                        if km_w and toll_w:
                            mid_x = (km_w[2] + toll_w[0]) / 2 
                        else:
                            mid_x = date_w[2] + 140
                        
                        text_to_insert = serial_map[date_w[4]]
                        
                        page.insert_text((mid_x - 18, date_w[3] - 2), 
                                         text_to_insert, 
                                         fontsize=12, 
                                         fontname="custom_font" if font_path else None, 
                                         color=(0,0,0))
            
            out_pdf = io.BytesIO()
            doc.save(out_pdf)
            st.session_state.processed_pdf = out_pdf.getvalue()
            
            if font_path:
                st.success(f"🎉 處理成功！工作表：{selected_sheet}")
            else:
                st.warning("⚠️ 處理成功，但未找到字體檔。")
                
        except Exception as e:
            st.error(f"錯誤: {e}")

# --- 下載區 ---
if st.session_state.processed_excel:
    # 下載檔名自動加上月份，方便管理
    fn = f"{selected_sheet}_更新後_{uploaded_excel.name}"
    st.download_button("💾 下載更新後的 Excel", st.session_state.processed_excel, fn)
if st.session_state.processed_pdf:
    st.download_button("💾 下載標註後的 PDF", st.session_state.processed_pdf, f"標註_{selected_sheet}_" + uploaded_pdf.name)
