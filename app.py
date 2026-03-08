import streamlit as st
import pandas as pd
import pdfplumber
import fitz
import io

st.set_page_config(page_title="通行費對帳工具", layout="wide")
st.title("通行費自動對帳與標註工具")

if 'processed_excel' not in st.session_state: st.session_state.processed_excel = None
if 'processed_pdf' not in st.session_state: st.session_state.processed_pdf = None

uploaded_pdf = st.file_uploader("上傳遠通電收 PDF", type="pdf")
uploaded_excel = st.file_uploader("上傳 T_E 申請表 Excel", type="xlsx")

if uploaded_pdf and uploaded_excel:
    if st.button("開始處理"):
        try:
            pdf_bytes = uploaded_pdf.getvalue()
            excel_bytes = uploaded_excel.getvalue()
            
            # 1. 處理 Excel (直接修改原檔)
            raw_df = pd.read_excel(excel_bytes, header=None)
            header_idx = next((i for i in range(30) if "項目" in str(raw_df.iloc[i].values) and "服務日期" in str(raw_df.iloc[i].values)), -1)
            
            # 讀取 Excel 用於運算
            df = pd.read_excel(excel_bytes, header=header_idx)
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
                                toll_map[parts[0]] = int(parts[2].replace('元', ''))
                            except: continue

            # 3. 更新 Excel 數據
            for date, amount in toll_map.items():
                mask = (df[date_col] == date)
                if mask.any():
                    df.at[df[mask].index[0], toll_col] = amount
            
            # 保存修改回原 Excel 結構 (這裡簡單處理，將填好的 df 轉回 excel)
            out_excel = io.BytesIO()
            df.to_excel(out_excel, index=False)
            st.session_state.processed_excel = out_excel.getvalue()

            # 4. PDF 標註 (格式: 項目XX, 黑色)
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            serial_map = {row[date_col]: str(int(row[item_col])) for _, row in df.iterrows() if pd.notna(row[toll_col])}
            
            for page in doc:
                words = page.get_text("words")
                for i in range(len(words) - 1):
                    if words[i][4] in serial_map:
                        x_mid = (words[i+1][2] + words[i+2][0]) / 2
                        # 格式: 項目XX, 黑色 (color=(0,0,0))
                        page.insert_text((x_mid - 20, words[i][1] + 1), f"項目{serial_map[words[i][4]]}", fontsize=9, color=(0, 0, 0))
            
            out_pdf = io.BytesIO()
            doc.save(out_pdf)
            st.session_state.processed_pdf = out_pdf.getvalue()
            st.success("處理完畢！")

        except Exception as e:
            st.error(f"錯誤: {e}")

if st.session_state.processed_excel:
    st.download_button("下載處理後的 Excel", st.session_state.processed_excel, uploaded_excel.name)
if st.session_state.processed_pdf:
    st.download_button("下載標註序號的 PDF", st.session_state.processed_pdf, "標註後_" + uploaded_pdf.name)
