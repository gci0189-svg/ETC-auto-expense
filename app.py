import streamlit as st
import openpyxl
import pdfplumber
import io

st.set_page_config(page_title="通行費對帳工具", layout="wide")
st.title("通行費自動對帳工具")

uploaded_pdf = st.file_uploader("上傳遠通電收 PDF", type="pdf")
uploaded_excel = st.file_uploader("上傳 T_E 申請表 Excel", type="xlsx")

if uploaded_pdf and uploaded_excel:
    if st.button("開始處理"):
        try:
            # 1. 載入 Excel 原檔
            wb = openpyxl.load_workbook(uploaded_excel)
            ws = wb.active
            
            # 2. 定位標題列 (假設標題列在第 7 行)
            # 掃描前 15 行找到 '項目' 與 '服務日期'
            header_row = 0
            for row in range(1, 16):
                row_values = [str(ws.cell(row=row, column=col).value) for col in range(1, 10)]
                if "項目" in row_values and "服務日期" in row_values:
                    header_row = row
                    break
            
            # 找出 '服務日期' 和 '過路費' 的欄位索引
            date_col_idx, toll_col_idx = 0, 0
            for col in range(1, 20):
                val = str(ws.cell(row=header_row, column=col).value)
                if "服務日期" in val: date_col_idx = col
                if "過路費" in val: toll_col_idx = col
            
            # 3. 解析 PDF 並建立映射
            pdf_bytes = uploaded_pdf.getvalue()
            toll_map = {}
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                for page in pdf.pages:
                    for line in page.extract_text().split('\n'):
                        parts = line.split()
                        if len(parts) >= 3 and '/' in parts[0]:
                            toll_map[parts[0]] = int(parts[2].replace('元', ''))

            # 4. 填入 Excel (僅修改數據)
            # 從 header_row + 1 開始讀資料
            processed_dates = set()
            for row in range(header_row + 1, ws.max_row + 1):
                cell_date = ws.cell(row=row, column=date_col_idx).value
                if cell_date:
                    # 簡單處理 Excel 日期格式 (支援字串或 datetime)
                    date_str = cell_date.strftime('%Y/%m/%d') if hasattr(cell_date, 'strftime') else str(cell_date)
                    if date_str in toll_map and date_str not in processed_dates:
                        ws.cell(row=row, column=toll_col_idx).value = toll_map[date_str]
                        processed_dates.add(date_str)

            # 儲存
            out_excel = io.BytesIO()
            wb.save(out_excel)
            st.session_state.processed_excel = out_excel.getvalue()
            st.success("檔案已保留原始格式並更新數據！")

        except Exception as e:
            st.error(f"錯誤: {e}")

if st.session_state.get('processed_excel'):
    st.download_button("下載更新後的 Excel", st.session_state.processed_excel, "Updated_" + uploaded_excel.name)
