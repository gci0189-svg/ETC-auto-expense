import streamlit as st
import pandas as pd
import pdfplumber
import fitz
import io

st.title("通行費自動對帳與標註工具")

uploaded_pdf = st.file_uploader("上傳遠通電收 PDF", type="pdf")
uploaded_excel = st.file_uploader("上傳 T_E 申請表 Excel", type="xlsx")

if uploaded_pdf and uploaded_excel:
    if st.button("開始處理"):
        try:
            # 1. 讀取 Excel
            raw_df = pd.read_excel(uploaded_excel, header=None)
            header_row_index = -1
            for i in range(min(30, len(raw_df))):
                row_values = str(raw_df.iloc[i].values)
                if "項目" in row_values and "服務日期" in row_values:
                    header_row_index = i
                    break
            
            if header_row_index == -1:
                st.error("找不到標題列！請確認 Excel 是否包含 '項目' 與 '服務日期' 欄位。")
                st.stop()
            
            df = pd.read_excel(uploaded_excel, header=header_row_index)
            
            # 欄位偵測
            date_col = next((c for c in df.columns if '服務日期' in str(c)), None)
            toll_col = next((c for c in df.columns if '過路費' in str(c)), None)
            item_col = next((c for c in df.columns if '項目' in str(c)), None)
            
            # 格式化日期以供後續比對
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce').dt.strftime('%Y/%m/%d')
            
            # 2. 解析 PDF (為了避免 file pointer 問題，我們讀取 bytes)
            pdf_bytes = uploaded_pdf.getvalue()
            toll_map = {}
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        for line in text.split('\n'):
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
            
            output_excel = io.BytesIO()
            df.to_excel(output_excel, index=False)
            st.success("Excel 處理完成！")
            st.download_button("下載處理後的 Excel", output_excel.getvalue(), "T_E_Processed.xlsx")

            # 4. 在 PDF 上標註序號 (使用 bytes 重新開啟)
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            serial_map = {}
            for idx, row in df.iterrows():
                if pd.notna(row[toll_col]) and row[toll_col] != 0:
                    serial_map[row[date_col]] = str(row[item_col])

            for page in doc:
                words = page.get_text("words")
                for word in words:
                    # word[4] 是文字內容
                    if word[4] in serial_map:
                        page.insert_text((word[0] - 25, word[1] + 8), f"[{serial_map[word[4]]}]", 
                                         fontsize=9, color=(1, 0, 0))
            
            output_pdf = io.BytesIO()
            doc.save(output_pdf)
            st.download_button("下載標註序號的 PDF", output_pdf.getvalue(), "AXE-5073_Signed.pdf")
            
        except Exception as e:
            st.error(f"發生系統錯誤: {str(e)}")
