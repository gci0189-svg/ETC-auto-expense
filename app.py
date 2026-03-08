import streamlit as st
import openpyxl
import pdfplumber
import fitz
import io
from datetime import datetime

st.set_page_config(page_title="通行費對帳工具", layout="wide")
st.title("通行費自動對帳與標註工具")

if 'processed_excel' not in st.session_state: st.session_state.processed_excel = None
if 'processed_pdf' not in st.session_state: st.session_state.processed_pdf = None

uploaded_pdf = st.file_uploader("上傳遠通電收 PDF", type="pdf")
uploaded_excel = st.file_uploader("上傳 T_E 申請表 Excel", type="xlsx")

def format_date(date_val):
    try:
        if isinstance(date_val, str): return datetime.strptime(date_val, '%d-%b-%y').strftime('%Y/%m/%d')
        if hasattr(date_val, 'strftime'): return date_val.strftime('%Y/%m/%d')
        return str(date_val)
    except: return None

if uploaded_pdf and uploaded_excel:
    if st.button("開始處理"):
        try:
            pdf_bytes = uploaded_pdf.getvalue()
            wb = openpyxl.load_workbook(uploaded_excel)
            ws = wb.active
            
            # 1. 處理 Excel
            header_row = 7 
            date_col_idx, toll_col_idx, item_col_idx = 4, 11, 1
            
            toll_map, serial_map = {}, {}
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                for page in pdf.pages:
                    for line in page.extract_text().split('\n'):
                        parts = line.split()
                        if len(parts) >= 3 and '/' in parts[0]:
                            if parts[2].replace('元', '').isdigit():
                                toll_map[parts[0]] = int(parts[2].replace('元', ''))
            
            for row in range(header_row + 1, ws.max_row + 1):
                d_str = format_date(ws.cell(row=row, column=date_col_idx).value)
                if d_str in toll_map:
                    ws.cell(row=row, column=toll_col_idx).value = toll_map[d_str]
                    item_val = ws.cell(row=row, column=item_col_idx).value
                    if item_val: serial_map[d_str] = f"項目{item_val}"
                    toll_map[d_str] = None 
            
            out_excel = io.BytesIO()
            wb.save(out_excel)
            st.session_state.processed_excel = out_excel.getvalue()

            # 2. PDF 標註 (精確對齊)
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            for page in doc:
                words = page.get_text("words") # words 格式: (x0, y0, x1, y1, text, ...)
                for i in range(len(words) - 2):
                    # 邏輯：找到日期 (words[i])，里程 (words[i+1])，通行費 (words[i+2])
                    if words[i][4] in serial_map:
                        # 計算里程結束位置 (x1 of words[i+1]) 與 通行費開始位置 (x0 of words[i+2])
                        x_start = words[i+1][2]
                        x_end = words[i+2][0]
                        # 標註放在中間
                        center_x = (x_start + x_end) / 2
                        page.insert_text((center_x - 20, words[i][1] + 1), serial_map[words[i][4]], fontsize=8, color=(0,0,0))
            
            out_pdf = io.BytesIO()
            doc.save(out_pdf)
            st.session_state.processed_pdf = out_pdf.getvalue()
            st.success("處理成功！")
        except Exception as e:
            st.error(f"錯誤: {e}")

if st.session_state.processed_excel:
    st.download_button("下載更新後的 Excel", st.session_state.processed_excel, uploaded_excel.name)
if st.session_state.processed_pdf:
    st.download_button("下載標註後的 PDF", st.session_state.processed_pdf, "標註_" + uploaded_pdf.name)
