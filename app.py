import streamlit as st
import pandas as pd
import pdfplumber
import fitz
import io

st.set_page_config(page_title="通行費對帳工具", layout="wide")
st.title("通行費自動對帳與標註工具")

# 初始化 session_state 以防止處理後按鈕消失
if 'processed_excel' not in st.session_state:
    st.session_state.processed_excel = None
if 'processed_pdf' not in st.session_state:
    st.session_state.processed_pdf = None

uploaded_pdf = st.file_uploader("上傳遠通電收 PDF", type="pdf")
uploaded_excel = st.file_uploader("上傳 T_E 申請表 Excel", type="xlsx")

if uploaded_pdf and uploaded_excel:
    if st.button("開始處理"):
        try:
            pdf_bytes = uploaded_pdf.getvalue()
            
            # 1. 讀取與處理 Excel
            raw_df = pd.read_excel(uploaded_excel, header=None)
            header_idx = next((i for i in range(min(30, len(raw_df))) if "項目" in str(raw_df.iloc[i].values) and "服務日期" in str(raw_df.iloc[i].values)), -1)
            df = pd.read_excel(uploaded_excel, header=header_idx)
            
            date_col = next((c for c in df.columns if '服務日期' in str(c)), None)
            toll_col = next((c for c in df.columns if '過路費' in str(c)), None)
            item_col = next((c for c in df.columns if '項目' in str(c)), None)
            
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce').dt.strftime('%Y/%m/%d')
            
            # 2. 解析 PDF
            toll_map = {}
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                for page in pdf.pages:
                    for line in page.extract_text().split('\n'):
                        parts = line.split()
                        if len(parts) >= 3 and '/' in parts[0]:
                            try:
                                amount = int(parts[2].replace('元', ''))
                                toll_map[parts[0]] = amount
                            except: continue

            # 3. 填入 Excel
            for date, amount in toll_map.items():
                mask = (df[date_col] == date)
                if mask.any():
                    df.at[df[mask].index[0], toll_col] = amount
            
            out_excel = io.BytesIO()
            df.to_excel(out_excel, index=False)
            st.session_state.processed_excel = out_excel.getvalue()

            # 4. PDF 標註 (放置於「里程」與「通行費」中間)
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            serial_map = {row[date_col]: str(row[item_col]) for _, row in df.iterrows() if pd.notna(row[toll_col]) and row[toll_col] != 0}
            
            for page in doc:
                words = page.get_text("words")
                for i in range(len(words) - 1):
                    # 邏輯：找到日期字串後，往右搜尋里程與通行費的中間坐標
                    if words[i][4] in serial_map:
                        x_mid = (words[i+1][2] + words[i+2][0]) / 2 # 概算兩欄間距中心
                        page.insert_text((x_mid - 15, words[i][1] + 1), f"[{serial_map[words[i][4]]}]", fontsize=9, color=(1, 0, 0))
            
            out_pdf = io.BytesIO()
            doc.save(out_pdf)
            st.session_state.processed_pdf = out_pdf.getvalue()
            st.success("檔案處理完畢！")

        except Exception as e:
            st.error(f"錯誤: {e}")

# 顯示下載按鈕 (從 session_state 讀取)
if st.session_state.processed_excel:
    st.download_button("下載處理後的 Excel", st.session_state.processed_excel, "T_E_Processed.xlsx")
if st.session_state.processed_pdf:
    st.download_button("下載標註序號的 PDF", st.session_state.processed_pdf, "AXE-5073_Signed.pdf")
