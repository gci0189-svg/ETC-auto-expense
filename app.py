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
    """將所有日期統一轉為 YYYY/MM/DD 字串"""
    try:
        if isinstance(date_val, str):
            # 處理 1-Dec-25 格式
            return datetime.strptime(date_val, '%d-%b-%y').strftime('%Y/%m/%d')
        return date_val.strftime('%Y/%m/%d')
    except: return str(date_val)

if uploaded_pdf and uploaded_excel:
    if st.button("開始處理"):
        try:
            pdf_bytes = uploaded_pdf.getvalue()
            wb = openpyxl.load_workbook(uploaded_excel)
            ws = wb.active
            
            # 1. 找到表頭行 (找 "服務日期" 關鍵字)
            header_row = 7 # 根據您的圖，大約在第 7 行
            date_col_idx, toll_col_idx, item_col_idx = 4, 11, 1 # D, K, A
            
            # 2. 解析 PDF
            toll_map = {}
            serial_map = {} # {日期: 項目號}
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                for page in pdf.pages:
                    for line in page.extract_text().split('\n'):
                        parts = line.split()
                        if len(parts) >= 3 and '/' in parts[0]:
                            toll_map[parts[0]] = parts[2].replace('元', '')
            
            # 3. 更新 Excel 並建立標註映射
            for row in range(header_row + 1, ws.max_row + 1):
                raw_d = ws.cell(row=row, column=date_col_idx).value
                if raw_d:
                    d_str = format_date(raw_d)
                    if d_str in toll_map:
                        ws.cell(row=row, column=toll_col_idx).value = int(toll_map[d_str])
                        serial_map[d_str] = str(ws.cell(row=row, column=item_col_idx).value)
                        # 清空 toll_map 該日期的值，確保只填入第一筆
                        toll_map[d_str] = None 
            
            out_excel = io.BytesIO()
            wb.save(out_excel)
            st.session_state.processed_excel = out_excel.getvalue()

            # 4. PDF 標註 (精確對齊)
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            for page in doc:
                words = page.get_text("words")
                for i in range(len(words)):
                    if words[i][4] in serial_map:
                        # 放置在里程與通行費中間 (利用座標偏移)
                        x_pos = words[i][2] + 40 
                        page.insert_text((x_pos, words[i][1] + 1), f"項目{serial_map[words[i][4]]}", fontsize=9, color=(0,0,0))
            
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
