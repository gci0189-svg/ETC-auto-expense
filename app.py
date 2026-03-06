import streamlit as st
import pandas as pd
import pdfplumber
import fitz
import io

st.title("通行費自動對帳與標註工具")

# --- 上傳區 ---
uploaded_pdf = st.file_uploader("上傳遠通電收 PDF", type="pdf")
uploaded_excel = st.file_uploader("上傳 T_E 申請表 Excel", type="xlsx")

if uploaded_pdf and uploaded_excel:
    if st.button("開始處理"):
        try:
            # 1. 讀取 Excel (先不指定 header，讀取全資料以搜尋表頭)
            raw_df = pd.read_excel(uploaded_excel, header=None)
            
            # 搜尋包含關鍵字 '項目' 與 '服務日期' 的列作為標題
            header_row_index = -1
            for i in range(min(30, len(raw_df))): # 掃描前 30 行
                row_values = str(raw_df.iloc[i].values)
                if "項目" in row_values and "服務日期" in row_values:
                    header_row_index = i
                    break
            
            if header_row_index == -1:
                st.error("找不到標題列！請確認 Excel 是否包含 '項目' 與 '服務日期' 欄位。")
                st.stop()
            
            # 以正確的標題列重新讀取 Excel
            df = pd.read_excel(uploaded_excel, header=header_row_index)
            
            # 自動偵測列名
            date_col = next((c for c in df.columns if '服務日期' in str(c)), None)
            toll_col = next((c for c in df.columns if '過路費' in str(c)), None)
            item_col = next((c for c in df.columns if '項目' in str(c)), None)
            
            if not date_col or not toll_col or not item_col:
                st.error(f"偵測到標題但無法對應欄位！請確認欄位名稱。目前的欄位有: {list(df.columns)}")
                st.stop()

            # 轉換日期格式以便比對
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce').dt.strftime('%Y/%m/%d')
            
            # 2. 解析 PDF
            toll_map = {}
            with pdfplumber.open(uploaded_pdf) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        for line in text.split('\n'):
                            parts = line.split()
                            # 格式: 2025/12/31 44.3公里 29元
                            if len(parts) >= 3 and '/' in parts[0]:
                                date_str = parts[0]
                                try:
                                    amount = int(parts[2].replace('元', ''))
                                    toll_map[date_str] = amount
                                except:
                                    continue

            # 3. 填入 Excel (當日第一筆)
            for date, amount in toll_map.items():
                mask = (df[date_col] == date)
                if mask.any():
                    first_idx = df[mask].index[0]
                    df.at[first_idx, toll_col] = amount
            
            output_excel = io.BytesIO()
            df.to_excel(output_excel, index=False)
            st.success("Excel 處理完成！")
            st.download_button("下載處理後的 Excel", output_excel.getvalue(), "T_E_Processed.xlsx")

            # 4. 在 PDF 上標註序號
            doc = fitz.open(stream=uploaded_pdf.read(), filetype="pdf")
            serial_map = {}
            for idx, row in df.iterrows():
                if pd.notna(row[toll_col]) and row[toll_col] != 0:
                    serial_map[row[date_col]] = str(row[item_col])

            for page in doc:
                words = page.get_text("words")
                for word in words:
                    word_text = word[4]
                    if word_text in serial_map:
                        page.insert_text((word[0] - 25, word[1] + 8), f"[{serial_map[word_text]}]", 
                                         fontsize=9, color=(1, 0, 0))
            
            output_pdf = io.BytesIO()
            doc.save(output_pdf)
            st.download_button("下載標註序號的 PDF", output_pdf.getvalue(), "AXE-5073_Signed.pdf")
            
        except Exception as e:
            st.error(f"發生系統錯誤: {str(e)}")
