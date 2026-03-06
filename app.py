import streamlit as st
import pandas as pd
import pdfplumber
import fitz
import io
from datetime import datetime

st.title("通行費自動對帳與標註工具")

# --- 上傳區 ---
uploaded_pdf = st.file_uploader("上傳遠通電收 PDF", type="pdf")
uploaded_excel = st.file_uploader("上傳 T_E 申請表 Excel", type="xlsx")

if uploaded_pdf and uploaded_excel:
    if st.button("開始處理"):
        # 1. 讀取 Excel
        df = pd.read_excel(uploaded_excel)
        # 確保日期格式一致 (轉換為 YYYY/MM/DD)
        df['服務日期'] = pd.to_datetime(df['服務日期']).dt.strftime('%Y/%m/%d')
        
        # 2. 解析 PDF 並建立日期與費用的對應表
        toll_map = {}
        with pdfplumber.open(uploaded_pdf) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                for line in text.split('\n'):
                    # 簡單邏輯：搜尋符合 YYYY/MM/DD 的行並提取金額
                    # 實際請依據您的 PDF 結構調整邏輯
                    parts = line.split()
                    if len(parts) >= 3 and '/' in parts[0]:
                        date_str = parts[0]
                        try:
                            # 假設金額在第三個區塊 (例如: 2025/12/31 44.3公里 29元)
                            amount = int(parts[2].replace('元', ''))
                            toll_map[date_str] = amount
                        except:
                            continue

        # 3. 填入 Excel (新增新欄位或更新)
        df['過路費'] = df['服務日期'].map(toll_map)
        
        # 產生新的 Excel
        output_excel = io.BytesIO()
        df.to_excel(output_excel, index=False)
        st.success("Excel 處理完成！")
        st.download_button("下載處理後的 Excel", output_excel.getvalue(), "T_E_Processed.xlsx")

        # 4. 在 PDF 上標註序號
        doc = fitz.open(stream=uploaded_pdf.read(), filetype="pdf")
        
        # 建立 Excel 序號對應表
        serial_map = {}
        for idx, row in df.iterrows():
            if pd.notna(row['過路費']): # 假設這行有過路費，則記錄此項目序號
                serial_map[row['服務日期']] = str(int(row['項目']))

        # 遍歷頁面標註
        for page in doc:
            words = page.get_text("words") # 獲取所有文字座標
            for word in words:
                word_text = word[4]
                if word_text in serial_map:
                    # 在日期文字左側寫入項目序號 (紅字)
                    page.insert_text((word[0] - 30, word[1] + 8), serial_map[word_text], 
                                     fontsize=10, color=(1, 0, 0))
        
        # 儲存 PDF
        output_pdf = io.BytesIO()
        doc.save(output_pdf)
        st.download_button("下載標註序號的 PDF", output_pdf.getvalue(), "AXE-5073_Signed.pdf")
