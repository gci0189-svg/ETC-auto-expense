"""
Diebold Nixdorf 費用申報整合工具
=====================================
功能：
  Tab 1 - 加油費計算：上傳加油發票PDF → 自動統計發票總額＆稅額 → 計算Personal Car公里數
  Tab 2 - 通行費對帳：上傳T_E申請表＋遠通電收PDF → 填入過路費＋標註PDF
  
  兩個Tab共用「里程津貼小計」：Tab1 自動從Tab2的申請表讀取，省去手動輸入。

安裝與執行：
  pip install streamlit openpyxl pdfplumber pymupdf
  streamlit run app.py
"""

import streamlit as st
import openpyxl
import pdfplumber
import fitz          # pymupdf
import io
import os
import re
import math
from datetime import datetime

# ─────────────────────────────────────────
# 頁面設定
# ─────────────────────────────────────────
st.set_page_config(page_title="DN 費用申報整合工具", layout="wide", page_icon="🚗")

st.markdown("""
<style>
    .main-title { font-size:1.8rem; font-weight:700; color:#1F4E79; margin-bottom:0; }
    .sub-title  { font-size:0.95rem; color:#666; margin-bottom:1.5rem; }
    .result-box { background:#E8F4FD; border-left:4px solid #1F4E79;
                  padding:1rem 1.2rem; border-radius:6px; margin:0.8rem 0; }
    .warn-box   { background:#FFF8E1; border-left:4px solid #F59E0B;
                  padding:0.8rem 1rem; border-radius:6px; margin:0.5rem 0; }
    .success-box{ background:#E8F5E9; border-left:4px solid #2E7D32;
                  padding:0.8rem 1rem; border-radius:6px; margin:0.5rem 0; }
    .metric-big { font-size:2rem; font-weight:700; color:#C00000; }
    .metric-lbl { font-size:0.85rem; color:#555; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">🚗 DN 費用申報整合工具</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">加油費計算 ＋ 通行費對帳，一站搞定</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────
# Session State 初始化
# ─────────────────────────────────────────
for key in ['processed_excel', 'processed_pdf', 'mileage_allowance',
            'fuel_total', 'fuel_tax', 'personal_car_km', 'selected_sheet']:
    if key not in st.session_state:
        st.session_state[key] = None

# ─────────────────────────────────────────
# 工具函式
# ─────────────────────────────────────────

def format_date_slash(date_val):
    """各種格式 → 2026/03/31"""
    try:
        if isinstance(date_val, str):
            return datetime.strptime(date_val.strip(), '%d-%b-%y').strftime('%Y/%m/%d')
        if hasattr(date_val, 'strftime'):
            return date_val.strftime('%Y/%m/%d')
    except Exception:
        pass
    return None


def read_mileage_allowance(excel_bytes, sheet_name):
    """從T_E申請表讀取指定月份的「里程津貼小計」（第10欄，小計行）"""
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
    if sheet_name not in wb.sheetnames:
        return None
    ws = wb[sheet_name]
    for row in ws.iter_rows():
        vals = [c.value for c in row]
        # 小計行特徵：第1欄=None, 第2欄='小計'
        if vals[0] is None and vals[1] == '小計':
            return vals[9]   # col10 = 里程津貼
    return None


def parse_fuel_invoices_from_pdf(pdf_bytes):
    """
    從加油發票PDF解析所有發票的總額與稅額。
    台灣電子發票格式：掃描多張，每張有「銷售額合計」、「稅額」或類似欄位。
    策略：用 pdfplumber 抓全部文字，用正則找金額。
    回傳 list of (total, tax)
    """
    invoices = []
    
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        full_text = ""
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                full_text += t + "\n"
    
    # ── 策略1：找「銷售額合計 XXXX 稅額 XX」格式 ──
    # 台灣電子發票明細常見格式
    pattern_sales_tax = re.findall(
        r'銷售額合計[：:\s]*(\d+)[\s\S]{0,30}?稅額[：:\s]*(\d+)', full_text
    )
    if pattern_sales_tax:
        for s, t in pattern_sales_tax:
            invoices.append((int(s), int(t)))
        return invoices
    
    # ── 策略2：找「含稅金額」或「總金額」後面的數字 ──
    # 格式：含稅金額 1235  稅額 59
    pattern_total_tax = re.findall(
        r'(?:含稅金額|發票金額|總金額|合計金額)[：:\s]*(\d+)[\s\S]{0,50}?(?:稅額|營業稅)[：:\s]*(\d+)', 
        full_text
    )
    if pattern_total_tax:
        for total, tax in pattern_total_tax:
            invoices.append((int(total), int(tax)))
        return invoices

    # ── 策略3：找連續出現的「XXXX 元/稅 XX」行 ──
    # 處理掃描式發票：每行可能只有金額
    # 嘗試找發票號碼後的金額段落
    # 台灣發票號碼格式：2碼英文 + 8碼數字
    inv_sections = re.split(r'[A-Z]{2}-?\d{8}', full_text)
    if len(inv_sections) > 1:
        for section in inv_sections[1:]:   # 跳過第一段（發票號前）
            lines = section.strip().split('\n')
            # 嘗試在段落中找最大數字（總額）和稅額
            numbers = re.findall(r'\b(\d{3,6})\b', section[:300])
            if len(numbers) >= 2:
                amounts = sorted([int(n) for n in numbers], reverse=True)
                # 最大的是總額，找接近5%的是稅額
                total = amounts[0]
                # 找最接近 total/21 的數字作為稅額（台灣5%稅率，含稅則除21）
                expected_tax = total / 21
                tax_candidates = [n for n in amounts[1:] if abs(n - expected_tax) < expected_tax * 0.3]
                if tax_candidates:
                    tax = min(tax_candidates, key=lambda x: abs(x - expected_tax))
                    invoices.append((total, tax))
    
    if invoices:
        return invoices
    
    # ── 策略4：退回純數字行解析（最後手段）──
    # 找所有行中出現「元」字旁邊的數字（加油站常見格式）
    lines = full_text.split('\n')
    amount_lines = []
    for line in lines:
        m = re.search(r'(\d{3,5})\s*元', line)
        if m:
            amount_lines.append(int(m.group(1)))
    
    # 以每組(總額, 稅額)的方式配對：稅額約為總額/21
    i = 0
    while i < len(amount_lines) - 1:
        total = amount_lines[i]
        tax_expected = total / 21
        if abs(amount_lines[i+1] - tax_expected) < tax_expected * 0.4:
            invoices.append((total, amount_lines[i+1]))
            i += 2
        else:
            i += 1
    
    return invoices


def parse_toll_from_pdf(pdf_bytes):
    """從遠通電收PDF解析通行費，回傳 {日期字串: 金額} dict"""
    toll_map = {}
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split('\n'):
                parts = line.split()
                if len(parts) >= 3 and re.match(r'\d{4}/\d{2}/\d{2}', parts[0]):
                    amount_str = parts[2].replace('元', '')
                    if amount_str.isdigit():
                        date_str = parts[0]
                        if date_str not in toll_map:
                            toll_map[date_str] = int(amount_str)
    return toll_map


def find_font():
    """搜尋可用的中文字體"""
    for root, dirs, files in os.walk("."):
        for f in files:
            if f.lower().endswith(('.ttc', '.ttf')):
                return os.path.join(root, f)
    # 嘗試系統字體路徑
    system_fonts = [
        '/usr/share/fonts/truetype/wqy/wqy-microhei.ttc',
        '/System/Library/Fonts/PingFang.ttc',
        'C:/Windows/Fonts/msjh.ttc',
    ]
    for fp in system_fonts:
        if os.path.exists(fp):
            return fp
    return None


# ─────────────────────────────────────────
# Tab 佈局
# ─────────────────────────────────────────
tab1, tab2 = st.tabs(["⛽ 加油費計算", "🛣️ 通行費對帳"])


# ╔══════════════════════════════════════════╗
# ║  Tab 1：加油費計算                       ║
# ╚══════════════════════════════════════════╝
with tab1:
    st.subheader("⛽ 加油費計算器")
    st.markdown("上傳加油發票PDF，自動統計總額與稅額，並計算申報公里數。")

    col_left, col_right = st.columns([1, 1], gap="large")

    with col_left:
        st.markdown("#### 📄 上傳加油發票 PDF")
        fuel_pdf = st.file_uploader(
            "支援多張發票合併掃描的PDF", type="pdf", key="fuel_pdf"
        )

        st.markdown("#### 📊 T_E 申請表（自動帶入里程津貼）")
        te_excel_tab1 = st.file_uploader(
            "上傳T_E申請表可自動帶入里程津貼小計", type="xlsx", key="te_tab1"
        )

        selected_sheet_tab1 = None
        if te_excel_tab1:
            wb_tmp = openpyxl.load_workbook(te_excel_tab1, read_only=True)
            sheets = wb_tmp.sheetnames
            current_month = f"{datetime.now().month}月"
            default_idx = sheets.index(current_month) if current_month in sheets else 0
            selected_sheet_tab1 = st.selectbox(
                "選擇月份工作表", sheets, index=default_idx, key="sheet_tab1"
            )

        # 里程津貼輸入：自動帶入或手動輸入
        auto_allowance = None
        if te_excel_tab1 and selected_sheet_tab1:
            te_excel_tab1.seek(0)
            auto_allowance = read_mileage_allowance(te_excel_tab1.read(), selected_sheet_tab1)

        if auto_allowance:
            st.markdown(f"""
            <div class="success-box">
            ✅ 已從申請表自動讀取 <b>{selected_sheet_tab1}</b> 里程津貼小計：
            <span style="font-size:1.3rem;font-weight:700;color:#1F4E79"> NT$ {int(auto_allowance):,}</span>
            </div>
            """, unsafe_allow_html=True)
            mileage_allowance_input = auto_allowance
        else:
            mileage_allowance_input = st.number_input(
                "💰 總里程津貼（手動輸入，若未上傳申請表）",
                min_value=0, value=0, step=100,
                help="從T_E申請表月份工作表的「里程津貼」小計欄取得"
            )

        run_fuel = st.button("🚀 開始計算", type="primary", key="run_fuel")

    with col_right:
        st.markdown("#### 📊 計算結果")

        if run_fuel:
            if not fuel_pdf:
                st.error("請先上傳加油發票PDF！")
            else:
                with st.spinner("解析發票PDF中..."):
                    fuel_pdf.seek(0)
                    invoices = parse_fuel_invoices_from_pdf(fuel_pdf.read())

                if not invoices:
                    st.markdown("""
                    <div class="warn-box">
                    ⚠️ <b>自動解析失敗</b>：PDF格式可能為掃描圖片或特殊排版。<br>
                    請手動輸入下方的發票金額。
                    </div>
                    """, unsafe_allow_html=True)
                    fuel_total_val = st.number_input("發票總額合計", min_value=0, step=1, key="manual_total")
                    fuel_tax_val = st.number_input("發票稅額合計", min_value=0, step=1, key="manual_tax")
                else:
                    st.markdown("**解析到的發票明細：**")
                    fuel_total_val = 0
                    fuel_tax_val = 0
                    for i, (total, tax) in enumerate(invoices, 1):
                        st.markdown(f"發票 {i}：總額 **NT$ {total:,}**，稅額 NT$ {tax:,}")
                        fuel_total_val += total
                        fuel_tax_val += tax

                    st.markdown("---")
                    st.markdown(f"**發票合計：NT$ {fuel_total_val:,}　稅額合計：NT$ {fuel_tax_val:,}**")

                # 計算申報數字
                if mileage_allowance_input > 0 and fuel_total_val > 0:
                    diff = mileage_allowance_input - fuel_total_val
                    personal_car_km = math.ceil(diff / 7)
                    personal_car_amt = personal_car_km * 7

                    st.session_state.fuel_total = fuel_total_val
                    st.session_state.fuel_tax = fuel_tax_val
                    st.session_state.personal_car_km = personal_car_km

                    st.markdown(f"""
                    <div class="result-box">
                    <div class="metric-lbl">Concur 申報金額</div>
                    <br>
                    <table width="100%">
                    <tr>
                      <td><div class="metric-lbl">🚗 Personal Car Mileage</div>
                          <div class="metric-big">{personal_car_km:,} 公里</div>
                          <div class="metric-lbl">金額 NT$ {personal_car_amt:,}</div>
                      </td>
                      <td><div class="metric-lbl">⛽ Fuel（油資補助）</div>
                          <div class="metric-big">NT$ {fuel_total_val:,}</div>
                          <div class="metric-lbl">稅額 NT$ {fuel_tax_val:,}</div>
                      </td>
                    </tr>
                    </table>
                    <br>
                    <div class="metric-lbl">
                    計算公式：⌈({mileage_allowance_input:,} - {fuel_total_val:,}) ÷ 7⌉ = {personal_car_km:,} 公里
                    </div>
                    </div>
                    """, unsafe_allow_html=True)

                    st.markdown("""
                    <div class="success-box">
                    ✅ <b>Concur填寫方式</b><br>
                    1. Personal Car Mileage → Distance 填上方公里數<br>
                    2. Fuel → Amount 填發票總額合計<br>
                    3. Fuel → Tax Amount 填發票稅額合計
                    </div>
                    """, unsafe_allow_html=True)

                elif fuel_total_val > 0:
                    st.markdown(f"""
                    <div class="result-box">
                    <b>發票統計結果</b><br>
                    總額：NT$ {fuel_total_val:,}　稅額：NT$ {fuel_tax_val:,}<br>
                    <span style="color:#999">請輸入或上傳申請表以計算公里數</span>
                    </div>
                    """, unsafe_allow_html=True)

        else:
            st.markdown("""
            <div style="color:#aaa; padding:2rem; text-align:center;">
            上傳發票PDF後按「開始計算」
            </div>
            """, unsafe_allow_html=True)

        # 手動輸入區（備用）
        with st.expander("📝 手動輸入發票金額（PDF解析失敗時使用）"):
            st.markdown("直接在下方輸入最多10筆發票：")
            manual_totals = []
            manual_taxes = []
            for i in range(1, 11):
                c1, c2 = st.columns(2)
                with c1:
                    t = st.number_input(f"發票{i} 總額", min_value=0, step=1, key=f"mt_{i}")
                with c2:
                    x = st.number_input(f"發票{i} 稅額", min_value=0, step=1, key=f"mx_{i}")
                if t > 0:
                    manual_totals.append(t)
                    manual_taxes.append(x)

            if manual_totals:
                m_total = sum(manual_totals)
                m_tax = sum(manual_taxes)
                st.markdown(f"**手動合計：總額 NT$ {m_total:,}，稅額 NT$ {m_tax:,}**")
                if mileage_allowance_input > 0:
                    km = math.ceil((mileage_allowance_input - m_total) / 7)
                    st.markdown(f"**Personal Car 公里數：{km:,} 公里**")


# ╔══════════════════════════════════════════╗
# ║  Tab 2：通行費對帳                       ║
# ╚══════════════════════════════════════════╝
with tab2:
    st.subheader("🛣️ 通行費對帳與標註")
    st.markdown("上傳遠通電收PDF＋T_E申請表，自動填入過路費並標註PDF。")

    col_a, col_b = st.columns([1, 1], gap="large")

    with col_a:
        st.markdown("#### 📁 上傳檔案")

        toll_pdf = st.file_uploader("1. 遠通電收通行費 PDF", type="pdf", key="toll_pdf")
        te_excel = st.file_uploader("2. T_E 申請表 (.xlsx)", type="xlsx", key="te_main")

        selected_sheet = None
        if te_excel:
            wb_tmp = openpyxl.load_workbook(te_excel, read_only=True)
            sheet_names = wb_tmp.sheetnames
            current_month = f"{datetime.now().month}月"
            default_idx = sheet_names.index(current_month) if current_month in sheet_names else 0
            selected_sheet = st.selectbox(
                "3. 選擇月份工作表", sheet_names, index=default_idx, key="sheet_main"
            )
            st.session_state.selected_sheet = selected_sheet

            if selected_sheet:
                te_excel.seek(0)
                allowance = read_mileage_allowance(te_excel.read(), selected_sheet)
                if allowance:
                    st.session_state.mileage_allowance = allowance
                    st.markdown(f"""
                    <div class="success-box">
                    ✅ {selected_sheet} 里程津貼小計：<b>NT$ {int(allowance):,}</b><br>
                    <span style="font-size:0.85rem;color:#555">（Tab1 加油費計算將自動使用此數值）</span>
                    </div>
                    """, unsafe_allow_html=True)

        if toll_pdf and te_excel and selected_sheet:
            if st.button("🚀 開始對帳與標註", type="primary", key="run_toll"):
                with st.spinner("處理中..."):
                    try:
                        # ── 解析遠通電收 PDF ──
                        toll_pdf.seek(0)
                        toll_map = parse_toll_from_pdf(toll_pdf.read())

                        if not toll_map:
                            st.error("無法從PDF解析通行費資料，請確認格式正確。")
                            st.stop()

                        # ── 更新 Excel ──
                        te_excel.seek(0)
                        wb = openpyxl.load_workbook(te_excel)
                        ws = wb[selected_sheet]

                        HEADER_ROW = 7
                        DATE_COL   = 4    # D：服務日期
                        TOLL_COL   = 11   # K：過路費
                        ITEM_COL   = 1    # A：項目

                        serial_map = {}
                        matched_dates = set()

                        for row in range(HEADER_ROW + 1, ws.max_row + 1):
                            raw_date = ws.cell(row=row, column=DATE_COL).value
                            if not raw_date:
                                continue
                            d_str = format_date_slash(raw_date)
                            if not d_str:
                                continue
                            if d_str in toll_map and d_str not in matched_dates:
                                ws.cell(row=row, column=TOLL_COL).value = toll_map[d_str]
                                item_val = ws.cell(row=row, column=ITEM_COL).value
                                if item_val is not None:
                                    try:
                                        clean_item = int(float(item_val))
                                        serial_map[d_str] = f"項目 {clean_item:02d}"
                                    except Exception:
                                        serial_map[d_str] = f"項目 {item_val}"
                                matched_dates.add(d_str)

                        out_excel = io.BytesIO()
                        wb.save(out_excel)
                        st.session_state.processed_excel = out_excel.getvalue()

                        # ── 標註 PDF ──
                        font_path = find_font()
                        toll_pdf.seek(0)
                        doc = fitz.open(stream=toll_pdf.read(), filetype="pdf")

                        for page in doc:
                            words = page.get_text("words")
                            if font_path:
                                try:
                                    page.insert_font(fontname="custom_font", fontfile=font_path)
                                except Exception:
                                    font_path = None

                            for w in words:
                                word_text = w[4]
                                if word_text in serial_map:
                                    date_w = w
                                    line_words = [lw for lw in words if abs(lw[1] - date_w[1]) < 5]
                                    line_words.sort(key=lambda x: x[0])

                                    km_w, toll_w = None, None
                                    for idx, lw in enumerate(line_words):
                                        if "公里" in lw[4]:
                                            km_w = lw
                                            if idx + 1 < len(line_words):
                                                toll_w = line_words[idx + 1]
                                            break

                                    mid_x = (km_w[2] + toll_w[0]) / 2 if (km_w and toll_w) else date_w[2] + 140
                                    label = serial_map[word_text]

                                    page.insert_text(
                                        (mid_x - 18, date_w[3] - 2),
                                        label,
                                        fontsize=11,
                                        fontname="custom_font" if font_path else "helv",
                                        color=(0, 0, 0.7)
                                    )

                        out_pdf = io.BytesIO()
                        doc.save(out_pdf)
                        st.session_state.processed_pdf = out_pdf.getvalue()

                        st.success(f"✅ 處理完成！共比對 {len(matched_dates)} 筆通行費資料。")

                        if len(toll_map) > len(matched_dates):
                            unmatched = set(toll_map.keys()) - matched_dates
                            st.markdown(f"""
                            <div class="warn-box">
                            ⚠️ 以下日期在PDF中有通行費，但申請表中未找到對應日期（可能為非工作日或未登錄）：<br>
                            {', '.join(sorted(unmatched))}
                            </div>
                            """, unsafe_allow_html=True)

                    except Exception as e:
                        st.error(f"處理錯誤：{e}")
                        import traceback
                        st.code(traceback.format_exc())

    with col_b:
        st.markdown("#### 📥 下載結果")

        if st.session_state.processed_excel and selected_sheet:
            te_name = te_excel.name if te_excel else "T_E申請表.xlsx"
            fn_excel = f"{selected_sheet}_通行費更新_{te_name}"
            st.download_button(
                "💾 下載更新後的 Excel（含過路費）",
                data=st.session_state.processed_excel,
                file_name=fn_excel,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        if st.session_state.processed_pdf and toll_pdf:
            fn_pdf = f"標註_{selected_sheet}_{toll_pdf.name}" if selected_sheet else f"標註_{toll_pdf.name}"
            st.download_button(
                "💾 下載標註後的通行費 PDF",
                data=st.session_state.processed_pdf,
                file_name=fn_pdf,
                mime="application/pdf"
            )

        if not st.session_state.processed_excel and not st.session_state.processed_pdf:
            st.markdown("""
            <div style="color:#aaa; padding:2rem; text-align:center;">
            上傳檔案後按「開始對帳與標註」，<br>結果將在此下載
            </div>
            """, unsafe_allow_html=True)

        # 通行費預覽
        if toll_pdf:
            with st.expander("🔍 預覽遠通電收PDF解析結果"):
                toll_pdf.seek(0)
                preview_map = parse_toll_from_pdf(toll_pdf.read())
                if preview_map:
                    total_toll = sum(preview_map.values())
                    st.markdown(f"**共解析 {len(preview_map)} 筆，總計 NT$ {total_toll:,} 元**")
                    for date, amt in sorted(preview_map.items()):
                        st.markdown(f"- {date}：**{amt} 元**")
                else:
                    st.warning("未解析到通行費資料")


# ─────────────────────────────────────────
# 底部說明
# ─────────────────────────────────────────
st.markdown("---")
st.markdown("""
<div style="font-size:0.8rem; color:#999; text-align:center;">
🔒 所有資料僅在本機處理，不上傳任何伺服器 ｜ 
欄位對應：日期=D欄、過路費=K欄、項目=A欄（第8行為表頭）
</div>
""", unsafe_allow_html=True)
