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

            # 2. PDF 標註 (以 y 座標歸類每一行)
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            for page in doc:
                words = page.get_text("words")
                # 建立一個 y 座標到文字的映射，誤差 5 以內視為同一行
                rows = {}
                for w in words:
                    y_level = round(w[1] / 5) * 5
                    if y_level not in rows: rows[y_level] = []
                    rows[y_level].append(w)
                
                # 遍歷每一行找日期
                for y, row_words in rows.items():
                    # 依 x 座標排序該行文字
                    row_words.sort(key=lambda x: x[0])
                    for w in row_words:
                        if w[4] in serial_map:
                            # 找到日期字串後，往右找里程與通行費區域 (假設第 2 個詞是里程，第 3 個是通行費)
                            # 如果該行沒有這麼多詞，直接定在日期右邊固定距離
                            x_pos = w[2] + 40 
                            # 插入文字：改為 12pt, 黑色, 顯示 "項目XX"
                            page.insert_text((x_pos, w[3]-2), serial_map[w[4]], fontsize=12, color=(0,0,0))
            
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
