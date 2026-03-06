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
            lines = text.split("\n")

            for line in lines:
                parts = line.split()

                if len(parts) >= 3:
                    try:
                        date = parts[0]
                        fee = parts[2].replace("元","")

                        dt = datetime.strptime(date,"%Y/%m/%d")
                        key = dt.strftime("%-d-%b-%y")

                        records[key] = int(fee)

                    except:
                        pass

    return records


if st.button("開始處理"):

    df = pd.read_excel(excel_file)

    records = parse_pdf(pdf_file)

    seen = set()

    for i,row in df.iterrows():

        date_cell = row[0]

        if pd.isna(date_cell):
            continue

        date_str = date_cell.strftime("%-d-%b-%y")

        if date_str not in seen:

            if date_str in records:
                df.at[i,"K"] = records[date_str]

            seen.add(date_str)

    st.success("完成")

    st.download_button(
        "下載Excel",
        df.to_csv(index=False).encode(),
        "里程明細_完成.csv",
        "text/csv"
    )
