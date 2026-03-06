import streamlit as st
import pandas as pd
import pdfplumber
from datetime import datetime
import io

st.title("遠通通行費自動填寫工具")

pdf_file = st.file_uploader("上傳通行費PDF", type="pdf")
excel_file = st.file_uploader("上傳里程明細Excel", type="xlsx")

def parse_pdf(file):
    records = {}
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text: continue
            lines = text.split("\n")
            for line in lines:
                parts = line.split()
                if len(parts) >= 3:
                    try:
                        date_raw = parts[0] # 假設格式如 2025/12/01
                        fee = parts[2].replace("元","")
                        dt = datetime.strptime(date_raw, "%Y/%m/%d")
                        # 轉換為 1-Dec-25 格式與 Excel 比對
                        key = dt.strftime("%-d-%b-%y")
                        records[key] = int(fee)
                    except:
                        pass
    return records

if st.button("開始處理"):
    if pdf_file and excel_file:
        try:
            # 1. 讀取 Excel，跳過前 6 行標題，直接從第 7 行(index 6)開始讀取欄位名
            df = pd.read_excel(excel_file, header=6)
            
            # 2. 解析 PDF
            records = parse_pdf(pdf_file)
            
            # 3. 比對與填寫
            # 我們尋找「服務日期」這一欄，並將結果填入「過路費」這一欄
            if "服務日期" in df.columns and "過路費" in df.columns:
                count = 0
                for i, row in df.iterrows():
                    # 安全轉換 Excel 日期
                    raw_date = row["服務日期"]
                    date_val = pd.to_datetime(raw_date, errors='coerce')
                    
                    if pd.isna(date_val):
                        continue
                        
                    # 轉成與 PDF 一致的格式 (例如 1-Dec-25)
                    date_str = date_val.strftime("%-d-%b-%y")
                    
                    if date_str in records:
                        df.at[i, "過路費"] = records[date_str]
                        count += 1
                
                st.success(f"處理完成！共成功匹配 {count} 筆日期。")
                
                # 4. 提供下載 (Excel 格式)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                
                st.download_button(
                    label="點我下載完成的 Excel",
                    data=output.getvalue(),
                    file_name="里程明細_已填寫.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("找不到『服務日期』或『過路費』欄位，請檢查 Excel 格式。")
        except Exception as e:
            st.error(f"執行出錯: {e}")
    else:
        st.warning("請先上傳兩個檔案再開始。")
