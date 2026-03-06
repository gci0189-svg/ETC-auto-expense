import streamlit as st
import pandas as pd
import pdfplumber
from datetime import datetime

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
                # 假設 PDF 格式為: 2026/03/01 ... 費用
                if len(parts) >= 3:
                    try:
                        date = parts[0]
                        fee = parts[2].replace("元","")
                        
                        dt = datetime.strptime(date, "%Y/%m/%d")
                        # Linux 環境(Streamlit) 使用 %-d 去除前導零
                        key = dt.strftime("%-d-%b-%y")
                        records[key] = int(fee)
                    except:
                        pass
    return records

if st.button("開始處理"):
    if pdf_file and excel_file:
        # 讀取 Excel
        df = pd.read_excel(excel_file)
        records = parse_pdf(pdf_file)
        seen = set()

        for i, row in df.iterrows():
            # 強制將第一欄轉為日期格式，若轉換失敗會變成 NaT
            date_cell = pd.to_datetime(row[0], errors='coerce')

            # 如果這格不是日期（或是標題列），就跳過
            if pd.isna(date_cell):
                continue

            # 轉換為與 PDF 比對的格式
            date_str = date_cell.strftime("%-d-%b-%y")

            if date_str not in seen:
                if date_str in records:
                    # 在 K 欄填入費用 (K 是第 11 欄，若 Excel 沒這麼多欄請注意)
                    df.loc[i, "K"] = records[date_str]
                seen.add(date_str)

        st.success("完成！請點擊下方按鈕下載")

        # 輸出轉成 CSV 或 Excel
        st.download_button(
            "下載處理後的檔案",
            df.to_csv(index=False).encode('utf-8-sig'), # 使用 utf-8-sig 解決 Excel 開啟亂碼
            "里程明細_已完成.csv",
            "text/csv"
        )
    else:
        st.error("請先上傳 PDF 與 Excel 檔案")
