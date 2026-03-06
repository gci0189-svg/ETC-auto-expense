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
            # 1. 讀取 Excel
            df = pd.read_excel(uploaded_excel)
            
            # 自動偵測日期與過路費欄位
            # 根據您提供的資料，服務日期對應的是 '服務日期'，過路費對應的是 '過路費'
            date_col = next((c for c in df.columns if '服務日期' in str(c)), None)
            toll_col = next((c for c in df.columns if '過路費' in str(c)), None)
            item_col = next((c for c in df.columns if '項目' in str(c)), None)
            
            if not date_col or not toll_col:
                st.error(f"找不到必要欄位！Excel 偵測到的欄位有: {list(df.columns)}")
                st.stop()

            # 確保日期格式一致
            df[date_col] = pd.to_datetime(df[date_col]).dt.strftime('%Y/%m/%d')
            
            # 2. 解析 PDF 並建立日期與費用的對應表
            toll_map = {}
            with pdfplumber.open(uploaded_pdf) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        for line in text.split('\n'):
                            # 針對 PDF 日期格式進行解析
                            parts = line.split()
                            # 假設格式: 2025/12/31 44.3公里 29元
                            if len(parts) >= 3 and '/' in parts[0]:
                                date_str = parts[0]
                                try:
                                    amount = int(parts[2].replace('元', ''))
                                    toll_map[date_str] = amount
                                except:
                                    continue

            # 3. 填入 Excel (將 PDF 金額填入對應日期當天的第一筆項目)
            # 這裡採用分組邏輯，只在當天第一筆填入
            for date in toll_map:
                mask = (df[date_col] == date)
                if mask.any():
                    first_idx = df[mask].index[0]
                    df.at[first_idx, toll_col] = toll_map[date]
            
            # 產生新的 Excel
            output_excel = io.BytesIO()
            df.to_excel(output_excel, index=False)
            st.success("Excel 處理完成！")
            st.download_button("下載處理後的 Excel", output_excel.getvalue(), "T_E_Processed.xlsx")

            # 4. 在 PDF 上標註序號
            doc = fitz.open(stream=uploaded_pdf.read(), filetype="pdf")
            
            # 建立日期對應項目序號的表
            serial_map = {}
            for idx, row in df.iterrows():
                # 若該欄位有過路費數據，記錄其項目序號
                if pd.notna(row[toll_col]) and row[toll_col] != 0:
                    serial_map[row[date_col]] = str(row[item_col])

            # 遍歷頁面標註
            for page in doc:
                words = page.get_text("words")
                for word in words:
                    word_text = word[4]
                    if word_text in serial_map:
                        # 在日期文字左側寫入紅色的項目序號
                        page.insert_text((word[0] - 25, word[1] + 8), f"[{serial_map[word_text]}]", 
                                         fontsize=9, color=(1, 0, 0))
            
            output_pdf = io.BytesIO()
            doc.save(output_pdf)
            st.download_button("下載標註序號的 PDF", output_pdf.getvalue(), "AXE-5073_Signed.pdf")
            
        except Exception as e:
            st.error(f"發生錯誤: {str(e)}")
