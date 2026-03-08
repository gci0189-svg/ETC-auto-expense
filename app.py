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
                    if item_val is not None:
                        try:
                            # 確保項目序號是整數 (去除 71.0 的 .0)
                            clean_item = int(float(item_val))
                            serial_map[d_str] = f"項目 {clean_item}"
                        except:
                            serial_map[d_str] = f"項目 {item_val}"
                    toll_map[d_str] = None 
            
            out_excel = io.BytesIO()
            wb.save(out_excel)
            st.session_state.processed_excel = out_excel.getvalue()

            # 2. PDF 標註 (精準置中 + 中文字體支援)
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            for page in doc:
                words = page.get_text("words")
                
                # 篩選有日期的字詞
                for w in words:
                    if w[4] in serial_map:
                        date_w = w
                        # 抓取跟這個日期「同一行」的所有文字 (y 座標誤差小於 5)
                        line_words = [lw for lw in words if abs(lw[1] - date_w[1]) < 5]
                        line_words.sort(key=lambda x: x[0]) # 照 x 座標排序由左至右
                        
                        km_w, toll_w = None, None
                        for idx, lw in enumerate(line_words):
                            if "公里" in lw[4]:
                                km_w = lw # 找到里程字詞
                                if idx + 1 < len(line_words):
                                    toll_w = line_words[idx + 1] # 緊接在里程後的就是通行費
                                break
                        
                        # 計算中心點
                        if km_w and toll_w:
                            # (里程的右邊界 + 通行費的左邊界) / 2
                            mid_x = (km_w[2] + toll_w[0]) / 2 
                        else:
                            # 防呆機制：如果找不到，就固定放在日期右邊 140 的位置
                            mid_x = date_w[2] + 140
                        
                        # 要寫入的文字
                        text_to_insert = serial_map[date_w[4]]
                        
                        # 插入文字：指定字體為 cjk (支援中文)，大小 12pt，置中偏移約 20 單位
                        page.insert_text((mid_x - 20, date_w[3] - 2), 
                                         text_to_insert, 
                                         fontsize=12, 
                                         fontname="cjk", 
                                         color=(0,0,0))
            
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
