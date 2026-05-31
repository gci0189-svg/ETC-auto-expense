"""
DN 費用申報整合工具 v4 (物理分欄與數據鎖定加固版)
============================================================
佈局：單頁寬版
  上方：st.columns([3, 2])
    左 3/5 → 通行費對帳（上傳T_E申請表＋遠通電收PDF，自動生成標註PDF、比對明細、並在Excel內附稽核報告頁與橫向PDF）
    右 2/5 → 加油費計算（自動解析最多10張發票，按日期物理分欄排序，即時同步至最上方 Concur 快速填寫對照表）
  下方：橫線分隔 → 電信費處理（移除密碼＋擷取第一頁）

安裝：
  pip install streamlit openpyxl pdfplumber pymupdf pypdf pyzbar pillow opencv-python-headless pandas
  streamlit run app.py
"""

import streamlit as st
import streamlit.components.v1 as components
import openpyxl
import pdfplumber
import fitz  # PyMuPDF
import io, os, re, math
import pandas as pd
import shutil
import tempfile
import subprocess
from datetime import datetime
from PIL import Image

# ─────────────────────────────────────────
# 安全防禦性與系統環境檢測
# ─────────────────────────────────────────
try:
    from pyzbar.pyzbar import decode as decode_qrcode
    PYZBAR_AVAILABLE = True
except ImportError:
    PYZBAR_AVAILABLE = False

try:
    from pypdf import PdfReader, PdfWriter
    PYPDF_AVAILABLE = True
except ImportError:
    PYPDF_AVAILABLE = False

# 檢測雲端主機中是否安裝有 LibreOffice 執行檔
SOFFICE_PATH = shutil.which('soffice') or shutil.which('libreoffice')
LIBREOFFICE_AVAILABLE = SOFFICE_PATH is not None

# ─────────────────────────────────────────
# 頁面設定
# ─────────────────────────────────────────
st.set_page_config(page_title="DN 費用申報整合工具", layout="wide", page_icon="🚗")

st.markdown("""
<style>
  .block-container{padding-top:1rem;padding-bottom:1rem;padding-left:1rem;padding-right:1rem;max-width:100%}
  section[data-testid="stMain"] > div {padding-left:1rem}
  h1,h2,h3{margin-top:0}
  h2{font-size:1.15rem!important;color:#1F4E79;border-bottom:2px solid #1F4E79;padding-bottom:4px}
  h3{font-size:1rem!important;color:#333}
  .success-box{background:#E8F5E9;border-left:4px solid #2E7D32;padding:.6rem 1rem;border-radius:4px;margin:.4rem 0;font-size:.9rem}
  .warn-box{background:#FFF8E1;border-left:4px solid #F59E0B;padding:.6rem 1rem;border-radius:4px;margin:.4rem 0;font-size:.9rem}
  .info-box{background:#E8F4FD;border-left:4px solid #1F4E79;padding:.6rem 1rem;border-radius:4px;margin:.4rem 0;font-size:.9rem}
  .section-title{font-size:1.05rem;font-weight:700;color:#1F4E79;
                 border-bottom:2px solid #1F4E79;padding-bottom:4px;margin-bottom:.8rem}
  div[data-testid="stVerticalBlock"] > div[data-testid="stVerticalBlock"]{
    border-right: 1px solid #e0e0e0;
  }
</style>
""", unsafe_allow_html=True)

st.markdown(
    '<p style="font-size:1rem;font-weight:700;color:#1F4E79;margin:0 0 0.8rem 0;">🚗 DN 費用申報整合工具</p>',
    unsafe_allow_html=True
)

# Session State 初始化 (擴充至 10 張發票容量)
for k in ['toll_excel','toll_pdf_out','telecom_pdf','mileage_allowance',
          'selected_sheet','mileage_manual','merged_pdf','audit_df','mileage_pdf_out',
          'tolls_parking_amount', 'mileage_distance', 'fuel_amount', 'fuel_tax']:
    if k not in st.session_state:
        st.session_state[k] = None if k not in ['mileage_manual', 'tolls_parking_amount', 'mileage_distance', 'fuel_amount', 'fuel_tax'] else 0

# 加油發票金額與稅額狀態初始化 (10組)
for i in range(1, 11):
    if f"inv_t{i}" not in st.session_state:
        st.session_state[f"inv_t{i}"] = 0
    if f"inv_x{i}" not in st.session_state:
        st.session_state[f"inv_x{i}"] = 0

# 全域作用域變數安全初始化
invoice_rows = []
mileage_input = 0

# ═══════════════════════════════════════════
# 工具與回呼（Callback）全域函式
# ═══════════════════════════════════════════

def auto_tax(i):
    """
    [全域回呼]：當加油總額改變時，自動計算稅額，100% 避免 Scope 錯誤。
    """
    total = st.session_state[f"inv_t{i}"]
    if total > 0:
        sales = round(total / 1.05)
        st.session_state[f"inv_x{i}"] = round(sales * 0.05)
    else:
        st.session_state[f"inv_x{i}"] = 0


def format_date_slash(v):
    try:
        if isinstance(v, str):
            return pd.to_datetime(v.strip()).strftime('%Y/%m/%d')
        if hasattr(v, 'strftime'):
            return v.strftime('%Y/%m/%d')
    except Exception:
        pass
    return None


def read_mileage_allowance(excel_bytes, sheet_name):
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
    if sheet_name not in wb.sheetnames:
        return None
    ws = wb[sheet_name]
    for row in ws.iter_rows():
        vals = [c.value for c in row]
        if vals[0] is None and vals[1] == '小計':
            return vals[9]   # 欄J = 里程津貼小計
    return None


def parse_fuel_pdf_totals(pdf_bytes):
    """
    [物理空間投影對齊演算法 - 分欄中點邊界版]：
    1. 採用選用括號容錯正則，完美捕獲 Formosa 聯的 '1578 (TX)E' 格式。
    2. 子字串安全對齊定位，解決 '1578元' 或 '金額:1578' 的 X 軸坐標抓取問題。
    3. 利用直欄中點（Column Midpoints）進行物理投影分欄，徹底解決多直列並排發票的位移問題。
    4. 按發票交易日期由舊到新排序。
    """
    pairs = []
    
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            
            if len(text.strip()) < 30:
                try:
                    import pytesseract
                    img = page.to_image(resolution=300).original
                    text = pytesseract.image_to_string(img, lang='chi_tra+eng', config='--psm 6 --oem 3')
                except:
                    continue
            
            # 1. 提取高精準度發票金額 (500~5000)
            valid_amounts = []
            for line in text.split('\n'):
                # 允許金額與 TX 之間夾帶選用括號字元
                PAT_TX = re.compile(r'(\d{3,5})\s*[\(（]?\s*(?:TX|T[X×Xx]|1Ⅸ|Ⅸ)\b')
                tx_vals = [int(v) for v in PAT_TX.findall(line) if 500 <= int(v) <= 5000]
                if tx_vals:
                    valid_amounts.extend(tx_vals)
                    continue
                
                # 備援匹配總計字樣
                PAT_TOTAL = re.compile(r'(?:總\s*[計計十十訁]|額\s*[計計])\s*[\$＄]?\s*(\d{3,5})')
                PAT_CONCAT = re.compile(r'總計(\d{3,5})')
                for pat in [PAT_TOTAL, PAT_CONCAT]:
                    vals = [int(v) for v in pat.findall(line) if 500 <= int(v) <= 5000]
                    if vals:
                        valid_amounts.extend(vals)
                        break
            
            if not valid_amounts:
                continue
                
            # 2. 抓取頁面中所有的單字，並識別日期與其橫向 X 座標
            words = page.extract_words()
            dates = []  # 儲存 (x_center, date_str)
            for w in words:
                txt = w['text'].strip()
                date_match = re.search(r'(\d{4}[-/]\d{2}[-/]\d{2})', txt)
                if date_match:
                    std_d = date_match.group(1).replace('-', '/')
                    x_center = (w['x0'] + w['x1']) / 2
                    dates.append((x_center, std_d))
            
            # 將日期由左至右水平排序
            dates.sort(key=lambda d: d[0])
            
            # 如果沒有找到日期，直接當作無日期綁定
            if not dates:
                for amt in valid_amounts:
                    pairs.append(("9999/12/31", amt))
                continue
                
            # 3. 建立水平邊界線，進行完美的直欄分欄
            # 日期的數量代表總共有幾個直欄 (Column)
            num_cols = len(dates)
            midpoints = []
            for idx in range(num_cols - 1):
                midpoints.append((dates[idx][0] + dates[idx + 1][0]) / 2)
            
            # 定義分欄檢索函數
            def get_column_date(x_coord):
                # 尋找落在哪個區間
                for col_idx, limit in enumerate(midpoints):
                    if x_coord < limit:
                        return dates[col_idx][1]
                return dates[-1][1]  # 落在最後一個區間
            
            # 4. 對每個高精準度提取的發票金額，找到它在頁面中的 word 物件，進行特徵座標匹配
            for amt in valid_amounts:
                amt_str = str(amt)
                matching_words = [w for w in words if amt_str in w['text']]
                if matching_words:
                    best_word = None
                    for mw in matching_words:
                        line_words = [w['text'].upper() for w in words if abs(w['top'] - mw['top']) < 4]
                        line_text = " ".join(line_words)
                        if any(x in line_text for x in ["TX", "Tㄨ", "元", "合計", "金額"]):
                            best_word = mw
                            break
                    if not best_word:
                        best_word = matching_words[0]
                        
                    ax = (best_word['x0'] + best_word['x1']) / 2
                    assigned_date = get_column_date(ax)
                    pairs.append((assigned_date, amt))
                else:
                    pairs.append((dates[-1][1], amt))
                        
    # 依日期由舊到新排序並去重
    seen_amounts = set()
    unique_sorted = []
    
    pairs.sort(key=lambda x: x[0])
    for dt, amt in pairs:
        if amt not in seen_amounts:
            seen_amounts.add(amt)
            unique_sorted.append(amt)
            
    return unique_sorted


def parse_toll_from_pdf(pdf_bytes):
    """
    從遠通 PDF 解析通行費，並進行每日加總
    [Surgical Bugfix]: 採用高精確度的「通行交易行」正則規則，
    排除非交易日期(如:查詢時間、已於"xxxx/xx/xx"扣款等文字)。
    """
    toll_map = {}
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            # 日期 ＋ 里程數(公里) ＋ 通行費(元) 的嚴格交易格式
            rows = re.findall(r'(\d{4}/\d{2}/\d{2})\s+([\d\.]+)(?:公里)?\s+(\d+)(?:元)?', text)
            for date_str, mileage, amt in rows:
                std_date = format_date_slash(date_str)
                if std_date:
                    toll_map[std_date] = toll_map.get(std_date, 0) + int(amt)
    return toll_map


def find_font():
    for root, _, files in os.walk("."):
        for f in files:
            if f.lower().endswith(('.ttc', '.ttf')):
                return os.path.join(root, f)
    for fp in ['/usr/share/fonts/truetype/wqy/wqy-microhei.ttc',
               '/System/Library/Fonts/PingFang.ttc',
               'C:/Windows/Fonts/msjh.ttc']:
        if os.path.exists(fp):
            return fp
    return None


def install_local_fonts():
    """
    [Linux 補丁]：搜尋專案目錄下的所有 .ttf 與 .ttc 字型檔，
    自動安裝至 Linux 使用者系統字型目錄中，並刷新 OS 字型快取。
    """
    try:
        user_font_dir = os.path.expanduser('~/.fonts')
        if not os.path.exists(user_font_dir):
            os.makedirs(user_font_dir)
        
        fonts_copied = False
        for f in os.listdir('.'):
            if f.lower().endswith(('.ttf', '.ttc')):
                src_path = f
                dest_path = os.path.join(user_font_dir, f)
                if not os.path.exists(dest_path):
                    shutil.copy(src_path, dest_path)
                    fonts_copied = True
        
        if fonts_copied:
            subprocess.run(['fc-cache', '-f'], stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)
    except Exception as e:
        pass


def convert_excel_to_pdf(excel_bytes, sheet_name):
    """
    [高精準度 PDF 生成]：使用無頭 LibreOffice 將原始 Excel 工作表直接轉換成符合官方排版規範的 PDF。
    """
    if not LIBREOFFICE_AVAILABLE:
        return None
    try:
        install_local_fonts()

        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
        for name in list(wb.sheetnames):
            if name != sheet_name:
                del wb[name]
                
        ws = wb[sheet_name]
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = '9'  # A4 代碼
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        wb.active = 0
        
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_xlsx:
            xlsx_path = tmp_xlsx.name
            wb.save(xlsx_path)
            
        output_dir = tempfile.gettempdir()
        cmd = [
            SOFFICE_PATH,
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', output_dir,
            xlsx_path
        ]
        subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)
        
        pdf_filename = os.path.basename(xlsx_path).replace('.xlsx', '.pdf')
        pdf_path = os.path.join(output_dir, pdf_filename)
        
        with open(pdf_path, 'rb') as f:
            pdf_bytes = f.read()
            
        try:
            os.remove(xlsx_path)
            os.remove(pdf_path)
        except:
            pass
            
        return pdf_bytes
    except Exception as e:
        st.error(f"PDF 轉換失敗: {e}")
        return None


def build_results_html(invoice_rows, mileage_allowance):
    """
    invoice_rows: list of (total, tax) 每張發票
    回傳仿試算表的 HTML 字串 (動態支援至 10 列明細)
    """
    total_amount = sum(r[0] for r in invoice_rows)
    total_tax    = sum(r[1] for r in invoice_rows)
    km = math.ceil(max(0, mileage_allowance - total_amount) / 7) if mileage_allowance > 0 else 0
    amt = km * 7

    TD    = "border:1px solid #bbb;padding:6px 10px;font-size:13px;font-family:Arial,sans-serif;"
    TDNUM = TD + "text-align:right;"
    HDR   = TD + "background:#1F4E79;color:#fff;font-weight:700;text-align:center;font-size:13px;"
    SUB   = TD + "background:#D6E4F0;font-weight:700;text-align:center;"
    TOT   = TD + "background:#BDD7EE;font-weight:700;"
    BLK   = "border:none;background:transparent;width:16px;"

    right_rows = [
        ("總里程津貼",          f"{int(mileage_allowance):,}" if mileage_allowance else "—",
         "#FFF2CC", "#1F4E79", True),
        ("加油發票合計",         f"{total_amount:,}",  "#FFFFFF", "#333", False),
        ("發票稅額合計",         f"{total_tax:,}",     "#FCE4D6", "#C00000", False),
        ("Personal Car 公里數",  f"{km:,}",            "#E2EFDA", "#C00000", True),
        ("Personal Car 金額",    f"{amt:,}",           "#E2EFDA", "#333", False),
        ("Fuel（油資補助）",     f"{total_amount:,}",  "#FFFFFF", "#333", False),
    ]

    rows_html = ""
    # 支援 10 組發票列印
    for i in range(10):
        l1 = datetime.now().strftime('%Y/%m') if i < len(invoice_rows) else ""
        l2 = f"{invoice_rows[i][0]:,}"        if i < len(invoice_rows) else ""
        l3 = f"{invoice_rows[i][1]:,}"        if i < len(invoice_rows) else ""
        if i < len(right_rows):
            rl, rv, rbg, rc, rb = right_rows[i]
            fw = "700" if rb else "400"
            fs = "14px" if rb else "13px"
            r1 = f'<td style="{TD}background:{rbg};">{rl}</td>'
            r2 = f'<td style="{TDNUM}background:{rbg};color:{rc};font-weight:{fw};font-size:{fs};">{rv}</td>'
        else:
            r1 = f'<td style="{TD}"></td>'
            r2 = f'<td style="{TD}"></td>'
        rows_html += (
            f'<tr>'
            f'<td style="{TD}">{l1}</td>'
            f'<td style="{TDNUM}">{l2}</td>'
            f'<td style="{TDNUM}">{l3}</td>'
            f'<td style="{BLK}"></td>'
            f'{r1}{r2}</tr>'
        )

    formula = (
        f"⌈({int(mileage_allowance):,} − {total_amount:,}) ÷ 7⌉ = {km:,} 公里"
        if mileage_allowance else ""
    )

    html = (
        '<div style="overflow-x:auto;">'
        '<table style="border-collapse:collapse;width:100%;font-family:Arial,sans-serif;">'
        '<colgroup>'
        '<col style="width:13%"><col style="width:14%"><col style="width:12%">'
        '<col style="width:2%">'
        '<col style="width:35%"><col style="width:18%">'
        '</colgroup>'
        f'<tr><td colspan="3" style="{HDR}">加油發票登記</td>'
        f'<td style="{BLK}"></td>'
        f'<td colspan="2" style="{HDR}">申報金額計算</td></tr>'
        f'<tr><td style="{SUB}">日期</td><td style="{SUB}">發票總額</td>'
        f'<td style="{SUB}">發票稅額</td><td style="{BLK}"></td>'
        f'<td style="{SUB}">項目</td><td style="{SUB}">金額 (TWD)</td></tr>'
        + rows_html +
        f'<tr><td style="{TOT}">合計</td>'
        f'<td style="{TOT}text-align:right;">{total_amount:,}</td>'
        f'<td style="{TOT}text-align:right;">{total_tax:,}</td>'
        f'<td style="{BLK}"></td>'
        f'<td colspan="2" style="{TD}font-size:11px;color:#666;">{formula}</td></tr>'
        '</table></div>'
    )
    return html, total_amount, total_tax, km, amt


# ═══════════════════════════════════════════
# 主要佈局：左 3/5（通行費）｜ 右 2/5（加油費）
# ═══════════════════════════════════════════
col_toll, col_fuel = st.columns([3, 2], gap="large")

# ╔══════════════════════════════════════════╗
# ║  左側：通行費對帳                        ║
# ╚══════════════════════════════════════════╝
with col_toll:
    st.markdown('<div class="section-title">🛣️ 通行費對帳</div>', unsafe_allow_html=True)

    parking_pdf = st.file_uploader("① 停車費 PDF", type="pdf", key="parking_pdf")
    toll_pdf    = st.file_uploader("② 遠通電收 PDF", type="pdf", key="toll_pdf")
    te_excel    = st.file_uploader("③ T_E 申請表 (.xlsx)", type="xlsx", key="te_main")

    selected_sheet = None
    if te_excel:
        wb_tmp = openpyxl.load_workbook(te_excel, read_only=True)
        sheets = wb_tmp.sheetnames
        cm = f"{datetime.now().month}月"
        default_idx = sheets.index(cm) if cm in sheets else 0
        selected_sheet = st.selectbox("④ 選擇月份工作表", sheets, index=default_idx, key="s_main")
        st.session_state.selected_sheet = selected_sheet

        if selected_sheet:
            te_excel.seek(0)
            excel_bytes = te_excel.read()
            allowance = read_mileage_allowance(excel_bytes, selected_sheet)
            if allowance:
                st.session_state.mileage_allowance = allowance
                st.markdown(f"""
                <div class="success-box">
                ✅ <b>{selected_sheet}</b> 里程津貼小計：<b>NT$ {int(allowance):,}</b>
                （已同步至右側加油費計算）
                </div>""", unsafe_allow_html=True)

            # ── 🚀 [即時自動解析機制：物理列加總版] ──
            # 優先讀取對帳標註後的 Excel 二進位檔案，防止上傳加油發票時金額洗回原始的 4,090
            active_bytes = st.session_state.toll_excel if st.session_state.toll_excel is not None else excel_bytes
            try:
                wb_auto = openpyxl.load_workbook(io.BytesIO(active_bytes), data_only=True)
                if selected_sheet in wb_auto.sheetnames:
                    ws_auto = wb_auto[selected_sheet]
                    t_total = 0
                    p_total = 0
                    for r in range(8, ws_auto.max_row + 1):
                        is_sub = False
                        for col_idx in [1, 2, 3]:
                            val = ws_auto.cell(row=r, column=col_idx).value
                            if val and str(val).strip() == "小計":
                                is_sub = True
                                break
                        if is_sub:
                            continue  # 忽略小計儲存格本身，防止公式未刷新的 0 或舊快取干擾
                        
                        val_k = ws_auto.cell(row=r, column=11).value  # 過路費 (K欄)
                        val_l = ws_auto.cell(row=r, column=12).value  # 停車費 (L欄)
                        try:
                            if val_k is not None: t_total += int(float(val_k))
                        except: pass
                        try:
                            if val_l is not None: p_total += int(float(val_l))
                        except: pass
                    # 即時更新全域 state，實現零按鍵一上傳立刻顯示通行停車費總和
                    st.session_state.tolls_parking_amount = t_total + p_total
            except:
                pass

    if toll_pdf and te_excel and selected_sheet:
        if st.button("🚀 開始對帳與標註", type="primary", key="run_toll"):
            with st.spinner("對帳比對、標註中以及高精確度 PDF 轉換中..."):
                try:
                    # 1. 解析 PDF
                    toll_pdf.seek(0)
                    toll_map = parse_toll_from_pdf(toll_pdf.read())
                    if not toll_map:
                        st.error("無法解析通行費PDF，請確認格式")
                        st.stop()

                    # 2. 開啟並寫入 T_E 申報明細
                    te_excel.seek(0)
                    wb = openpyxl.load_workbook(te_excel)
                    ws = wb[selected_sheet]
                    DATE_COL, TOLL_COL, ITEM_COL = 4, 11, 1
                    serial_map, matched = {}, set()

                    # 注入頁面列印設定，保證 Excel 檔案下載後直接另存 PDF 也 100% 是一頁寬
                    ws.sheet_properties.pageSetUpPr.fitToPage = True
                    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
                    ws.page_setup.paperSize = '9'  # A4
                    ws.page_setup.fitToWidth = 1
                    ws.page_setup.fitToHeight = 0

                    # 將 PDF 的通行費匹配回 Excel 中
                    for row in range(8, ws.max_row + 1):
                        raw_date = ws.cell(row=row, column=DATE_COL).value
                        if not raw_date: continue
                        d_str = format_date_slash(raw_date)
                        if not d_str: continue
                        if d_str in toll_map and d_str not in matched:
                            ws.cell(row=row, column=TOLL_COL).value = toll_map[d_str]
                            item_val = ws.cell(row=row, column=ITEM_COL).value
                            if item_val is not None:
                                try:    serial_map[d_str] = f"項目 {int(float(item_val)):02d}"
                                except: serial_map[d_str] = f"項目 {item_val}"
                            matched.add(d_str)

                    # 3. 雙向對帳稽核計算 (Excel 日常加總 vs PDF 日常加總)
                    excel_daily = {}
                    for row in range(8, ws.max_row + 1):
                        raw_date = ws.cell(row=row, column=DATE_COL).value
                        if not raw_date: continue
                        d_str = format_date_slash(raw_date)
                        if not d_str: continue
                        
                        val = ws.cell(row=row, column=TOLL_COL).value
                        val_num = 0
                        if val is not None:
                            try:    val_num = int(float(val))
                            except: pass
                        excel_daily[d_str] = excel_daily.get(d_str, 0) + val_num

                    # 合併並彙總
                    all_dates = sorted(list(set(excel_daily.keys()) | set(toll_map.keys())))
                    audit_rows = []
                    for d in all_dates:
                        ex_val = excel_daily.get(d, 0)
                        pdf_val = toll_map.get(d, 0)
                        diff = ex_val - pdf_val
                        status = "✅ 匹配" if diff == 0 else "❌ 金額不符"
                        audit_rows.append({
                            "日期": d,
                            "Excel金額": ex_val,
                            "PDF金額": pdf_val,
                            "差異": diff,
                            "狀態": status
                        })

                    st.session_state.audit_df = pd.DataFrame(audit_rows)

                    # 3.5 [雙向物理累加計算]：繞過 Excel 公式限制，手動計算 K 與 L 欄之實體數值總和
                    tolls_total = 0
                    parking_total = 0
                    for r in range(8, ws.max_row + 1):
                        is_sub_col = False
                        for col_idx in [1, 2, 3]:
                            val = ws.cell(row=r, column=col_idx).value
                            if val and str(val).strip() == "小計":
                                is_sub_col = True
                                break
                        if is_sub_col:
                            continue  # 忽略小計行本身
                            
                        val_k = ws.cell(row=r, column=11).value  # 過路費 (K欄)
                        val_l = ws.cell(row=r, column=12).value  # 停車費 (L欄)
                        try:
                            if val_k is not None: tolls_total += int(float(val_k))
                        except: pass
                        try:
                            if val_l is not None: parking_total += int(float(val_l))
                        except: pass
                        
                    # 完美填入對照表 State 中
                    st.session_state.tolls_parking_amount = tolls_total + parking_total

                    # 4. 將稽核報告寫入 Excel 中（新增一個稽核頁籤）
                    audit_sheet_name = f"對帳稽核_{selected_sheet}"
                    if audit_sheet_name in wb.sheetnames:
                        del wb[audit_sheet_name]
                    audit_ws = wb.create_sheet(title=audit_sheet_name)
                    
                    headers = ["日期", "Excel金額", "PDF金額", "差異", "狀態"]
                    audit_ws.append(headers)
                    # 美化稽核頁籤表頭
                    for col_num, header in enumerate(headers, 1):
                        cell = audit_ws.cell(row=1, column=col_num)
                        cell.font = openpyxl.styles.Font(bold=True, color="FFFFFF")
                        cell.fill = openpyxl.styles.PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
                        cell.alignment = openpyxl.styles.Alignment(horizontal="center")

                    for r in audit_rows:
                        audit_ws.append([r["日期"], r["Excel金額"], r["PDF金額"], r["差異"], r["狀態"]])

                    # 自動調整欄寬
                    for col in audit_ws.columns:
                        max_len = max(len(str(cell.value or '')) for cell in col)
                        col_letter = openpyxl.utils.get_column_letter(col[0].column)
                        audit_ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

                    # 存檔明細檔
                    out_excel = io.BytesIO()
                    wb.save(out_excel)
                    excel_saved_bytes = out_excel.getvalue()
                    st.session_state.toll_excel = excel_saved_bytes

                    # 5. 直接將 Excel 轉換成 format 與 Logo 100% 相同、列印寬自適應為一頁的 PDF
                    if LIBREOFFICE_AVAILABLE:
                        st.session_state.mileage_pdf_out = convert_excel_to_pdf(excel_saved_bytes, selected_sheet)

                    # ── 標註遠通電收 PDF ──
                    font_path = find_font()
                    toll_pdf.seek(0)
                    doc = fitz.open(stream=toll_pdf.read(), filetype="pdf")
                    for page in doc:
                        words = page.get_text("words")
                        if font_path:
                            try:   page.insert_font(fontname="cf", fontfile=font_path)
                            except: font_path = None
                        for w in words:
                            if w[4] not in serial_map: continue
                            dw = w
                            lw = sorted([x for x in words if abs(x[1]-dw[1]) < 5], key=lambda x: x[0])
                            km_w  = next((x for x in lw if "公里" in x[4]), None)
                            toll_w = lw[lw.index(km_w)+1] if km_w and lw.index(km_w)+1 < len(lw) else None
                            mx = (km_w[2]+toll_w[0])/2 if (km_w and toll_w) else dw[2]+140
                            page.insert_text(
                                (mx-18, dw[3]-2), serial_map[w[4]], fontsize=11,
                                fontname="cf" if font_path else "helv", color=(0, 0, 0.7)
                            )

                    # ── 標註後儲存 ──
                    out_toll_only = io.BytesIO()
                    doc.save(out_toll_only)
                    st.session_state.toll_pdf_out = out_toll_only.getvalue()

                    # ── 合併：停車費 PDF + 標註後的遠通電收 ──
                    SIZE_LIMIT = 15 * 1024 * 1024   # 15MB

                    if parking_pdf:
                        parking_pdf.seek(0)
                        parking_doc = fitz.open(stream=parking_pdf.read(), filetype="pdf")
                        merged_doc  = fitz.open()
                        merged_doc.insert_pdf(parking_doc)   # 停車費優先
                        merged_doc.insert_pdf(doc)            # 標註遠通
                        parking_doc.close()

                        out_merged = io.BytesIO()
                        merged_doc.save(out_merged, garbage=4, deflate=True)
                        merged_bytes = out_merged.getvalue()
                        merged_size  = len(merged_bytes)

                        if merged_size > SIZE_LIMIT:
                            st.info(f"合併後 {merged_size/1024/1024:.1f}MB，開始降階壓縮...")
                            compressed = None
                            for quality in [85, 75, 60, 45]:
                                buf = io.BytesIO()
                                merged_doc.save(buf, garbage=4, deflate=True, deflate_images=True, deflate_fonts=True)
                                comp_doc = fitz.open(stream=buf.getvalue(), filetype="pdf")
                                out_comp = io.BytesIO()

                                writer_doc = fitz.open()
                                scale = 1.0
                                if quality <= 75: scale = 0.85
                                if quality <= 60: scale = 0.70
                                if quality <= 45: scale = 0.55

                                for pg in comp_doc:
                                    mat = fitz.Matrix(scale, scale)
                                    pix = pg.get_pixmap(matrix=mat, alpha=False)
                                    img_pdf = fitz.open()
                                    img_page = img_pdf.new_page(width=pg.rect.width, height=pg.rect.height)
                                    img_page.insert_image(img_page.rect, pixmap=pix)
                                    writer_doc.insert_pdf(img_pdf)

                                writer_doc.save(out_comp, garbage=4, deflate=True)
                                result = out_comp.getvalue()
                                result_size = len(result)

                                comp_doc.close()
                                writer_doc.close()

                                if result_size <= SIZE_LIMIT:
                                    compressed = result
                                    final_size = result_size
                                    used_quality = quality
                                    break

                            if compressed:
                                st.session_state['merged_pdf'] = compressed
                                st.session_state['merged_compressed'] = True
                                st.session_state['merged_size'] = final_size
                                st.session_state['merged_quality'] = used_quality
                            else:
                                st.session_state['merged_pdf'] = result
                                st.session_state['merged_compressed'] = True
                                st.session_state['merged_size'] = len(result)
                                st.session_state['merged_quality'] = 45
                        else:
                            st.session_state['merged_pdf'] = merged_bytes
                            st.session_state['merged_compressed'] = False
                            st.session_state['merged_size'] = merged_size

                        merged_doc.close()

                    doc.close()
                    st.success(f"✅ 完成！共比對 **{len(matched)}** 筆通行費")
                    unmatched = set(toll_map.keys()) - matched
                    if unmatched:
                        st.markdown(f"""<div class="warn-box">
                        ⚠️ PDF有記錄但申請表未找到的日期：{', '.join(sorted(unmatched))}
                        </div>""", unsafe_allow_html=True)

                except Exception as e:
                    st.error(f"錯誤：{e}")
                    import traceback; st.code(traceback.format_exc())

    # ── 檔案下載區 (3 欄式橫向並排) ──
    if st.session_state.toll_excel or st.session_state.toll_pdf_out or st.session_state.mileage_pdf_out:
        dl1, dl2, dl3 = st.columns(3, gap="small")
        with dl1:
            if st.session_state.toll_excel:
                te_name = te_excel.name if te_excel else "T_E申請表.xlsx"
                st.download_button(
                    "💾 下載更新後的 Excel（含稽核頁籤）",
                    st.session_state.toll_excel,
                    f"{selected_sheet}_對帳稽核_{te_name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        with dl2:
            if st.session_state.toll_pdf_out and toll_pdf:
                st.download_button(
                    "💾 下載標註後的遠通電收",
                    st.session_state.toll_pdf_out,
                    f"標註_{selected_sheet}_{toll_pdf.name}",
                    mime="application/pdf"
                )
        with dl3:
            if LIBREOFFICE_AVAILABLE:
                if st.session_state.mileage_pdf_out:
                    st.download_button(
                        "💾 下載原版格式里程 PDF (已修正亂碼)",
                        st.session_state.mileage_pdf_out,
                        f"{selected_sheet}_里程明細.pdf",
                        mime="application/pdf"
                    )
            else:
                st.info("💡 雲端尚未啟動 LibreOffice，但已為您的 Excel 預先植入「一頁寬」設定。請下載 Excel 並直接在電腦另存 PDF 即可，格式與字型完全正確。")

    # 合併 PDF 下載（停車費 + 標註遠通電收）
    if st.session_state.get('merged_pdf'):
        size_mb  = st.session_state['merged_size'] / 1024 / 1024
        was_comp = st.session_state.get('merged_compressed', False)
        quality  = st.session_state.get('merged_quality', '-')
        month_str = selected_sheet or datetime.now().strftime("%Y%m")

        if was_comp:
            st.markdown(f"""<div class="success-box">
            ✅ 壓縮完成：{size_mb:.1f}MB（低於15MB限制）
            </div>""", unsafe_allow_html=True)
        else:
            st.markdown(f"""<div class="success-box">
            ✅ 停車費＋遠通電收合併完成：{size_mb:.1f}MB
            </div>""", unsafe_allow_html=True)

        st.download_button(
            f"💾 下載合併PDF（{size_mb:.1f}MB）",
            data=st.session_state['merged_pdf'],
            file_name=f"{month_str}_停車費＋通行費.pdf",
            mime="application/pdf",
            type="primary"
        )

    # 顯示自動生成的對帳稽核報告表
    if st.session_state.audit_df is not None:
        with st.expander("🔍 檢視通行費對帳稽核報告 (即時驗證)"):
            st.markdown("**每日明細金額雙向稽核明細**")
            
            def highlight_diff(row):
                if row['狀態'] == '❌ 金額不符':
                    return ['background-color: #ffcccc'] * len(row)
                return ['background-color: #e6ffed'] * len(row)
            
            st.dataframe(
                st.session_state.audit_df.style.apply(highlight_diff, axis=1), 
                use_container_width=True
            )
            
            # 指標卡統計
            c1, c2 = st.columns(2)
            total_excel = int(st.session_state.audit_df['Excel金額'].sum())
            total_pdf = int(st.session_state.audit_df['PDF金額'].sum())
            c1.metric("Excel 總金額", f"{total_excel:,} 元")
            c2.metric("遠通 PDF 總金額", f"{total_pdf:,} 元")

    # 通行費預覽
    if toll_pdf:
        with st.expander("🔍 預覽遠通電收原始解析結果"):
            toll_pdf.seek(0)
            pm = parse_toll_from_pdf(toll_pdf.read())
            if pm:
                st.markdown(f"**共 {len(pm)} 筆，合計 NT$ {sum(pm.values()):,} 元**")
                for d, a in sorted(pm.items()):
                    st.markdown(f"- {d}：{a} 元")
            else:
                st.warning("未解析到通行費資料")

    # ── 電信費處理（T_E 申請表下方，間隔留白）──
    st.markdown("<div style='margin-top:2rem'></div>", unsafe_allow_html=True)
    st.markdown('<div class="section-title">📱 電信費 PDF 處理</div>', unsafe_allow_html=True)

    telecom_file = st.file_uploader("上傳電信費 PDF", type="pdf", key="telecom_up")

    tc_col1, tc_col2 = st.columns([3, 2], gap="medium")
    with tc_col1:
        password = st.text_input(
            "PDF 密碼（無密碼請留空）",
            type="password",
            placeholder="身分證末4碼 / 生日 MMDD",
            key="telecom_pwd"
        )
        st.markdown("""
        <div class="warn-box" style="font-size:.82rem">
        💡 台灣大哥大／遠傳：身分證末4碼<br>
        中華電信：生日 MMDD 或身分證末4碼<br>
        亞太電信：出生年月日 YYYYMMDD
        </div>""", unsafe_allow_html=True)

    with tc_col2:
        st.markdown("<div style='margin-top:1.8rem'></div>", unsafe_allow_html=True)
        if telecom_file:
            if st.button("🔓 移除密碼並擷取第一頁", type="primary", key="run_telecom"):
                telecom_file.seek(0)
                raw = telecom_file.read()
                passwords_to_try = list(dict.fromkeys([password, "", "0000"]))
                success = False
                for pwd in passwords_to_try:
                    ok, result_bytes, msg = remove_pdf_password_and_extract_page1(raw, pwd)
                    if ok:
                        st.session_state.telecom_pdf = result_bytes
                        st.success(f"✅ {msg}")
                        if pwd != password:
                            st.info(f"使用密碼「{pwd}」成功解密")
                        success = True
                        break
                if not success:
                    st.error("❌ 密碼錯誤，請確認後重試")
                    st.caption("其他格式：電話後4碼、生日6碼（YYYYMM）")

        if st.session_state.telecom_pdf:
            orig = telecom_file.name.replace('.pdf', '') if telecom_file else "電信費"
            st.download_button(
                "💾 下載（已解密，僅第一頁）",
                data=st.session_state.telecom_pdf,
                file_name=f"{orig}_第一頁.pdf",
                mime="application/pdf"
            )
            st.markdown("""
            <div class="success-box" style="font-size:.82rem">
            ✅ 下載後直接上傳至 Concur 作為電信費附件
            </div>""", unsafe_allow_html=True)


# ╔══════════════════════════════════════════╗
# ║  右側：加油費計算（手動輸入）            ║
# ╚══════════════════════════════════════════╝
with col_fuel:
    
    # ── [安全加固：提前預運算] ──
    # 在渲染最上方對照表前，提前運算 current 狀態，杜絕「慢一拍」顯示為 0 的現象
    current_invoice_rows = []
    for i in range(1, 11):
        t_val = st.session_state.get(f"inv_t{i}", 0)
        x_val = st.session_state.get(f"inv_x{i}", 0)
        if t_val > 0:
            current_invoice_rows.append((t_val, x_val))
            
    # 即時在全域計算最新合計
    if current_invoice_rows:
        temp_total_amount = sum(r[0] for r in current_invoice_rows)
        temp_total_tax = sum(r[1] for r in current_invoice_rows)
        m_allowance = st.session_state.get("mileage_manual", 0)
        temp_km = math.ceil(max(0, m_allowance - temp_total_amount) / 7) if m_allowance > 0 else 0
        
        st.session_state.fuel_amount = temp_total_amount
        st.session_state.fuel_tax = temp_total_tax
        st.session_state.mileage_distance = temp_km
    else:
        st.session_state.fuel_amount = 0
        st.session_state.fuel_tax = 0
        st.session_state.mileage_distance = 0

    # ── 右側：Concur 快速填寫對照面板 ──
    st.markdown('<div class="section-title">📋 Concur 快速填寫對照表</div>', unsafe_allow_html=True)
    
    st.markdown(f"""
    <div style="background:#F2F6FA; border-left:5px solid #1F4E79; padding:15px; border-radius:6px; margin-bottom:1.5rem;">
      <p style="margin:0 0 10px 0; font-weight:700; color:#1F4E79; font-size:0.95rem;">💡 複製上月申報單後，僅需更新以下 3 筆動態欄位：</p>
      <table style="width:100%; border-collapse:collapse; font-size:0.9rem;">
        <tr style="border-bottom:1px solid #e0e0e0;">
          <td style="padding:8px 0; font-weight:600; color:#333;">1. Personal Car Mileage</td>
          <td style="text-align:right; padding:8px 0; font-weight:700; color:#C00000; font-size:1.15rem;">
            {st.session_state.mileage_distance:,} <span style="font-size:0.8rem; font-weight:400; color:#555;">公里 (Distance)</span>
          </td>
        </tr>
        <tr style="border-bottom:1px solid #e0e0e0;">
          <td style="padding:8px 0; font-weight:600; color:#333;">2. Tolls/Road Charges/ Parking</td>
          <td style="text-align:right; padding:8px 0; font-weight:700; color:#1F4E79; font-size:1.15rem;">
            TWD {st.session_state.tolls_parking_amount:,} <span style="font-size:0.8rem; font-weight:400; color:#555;">(Amount)</span>
          </td>
        </tr>
        <tr>
          <td style="padding:8px 0; font-weight:600; color:#333;">3. Fuel</td>
          <td style="text-align:right; padding:8px 0;">
            <span style="font-weight:700; color:#333; font-size:1.15rem;">TWD {st.session_state.fuel_amount:,}</span> <span style="font-size:0.8rem; color:#555;">(Amount)</span><br>
            <span style="font-weight:700; color:#D35400; font-size:1.1rem;">TWD {st.session_state.fuel_tax:,}</span> <span style="font-size:0.8rem; color:#555;">(Tax Amount)</span>
          </td>
        </tr>
      </table>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="section-title">⛽ 加油費計算</div>', unsafe_allow_html=True)

    # 里程津貼：優先從左側申請表自動帶入，直接寫入 session_state 確保同步
    if st.session_state.mileage_allowance:
        mileage_val = int(st.session_state.mileage_allowance)
        if st.session_state.get("mileage_manual", 0) != mileage_val:
            st.session_state["mileage_manual"] = mileage_val
        st.markdown(f"""
        <div class="info-box">
        📊 里程津貼自動帶入：<b>NT$ {mileage_val:,}</b>
        </div>""", unsafe_allow_html=True)

    mileage_input = st.number_input(
        "💰 總里程津貼（可手動修改）",
        min_value=0,
        step=100,
        key="mileage_manual"
    )

    st.markdown("**🧾 加油發票**")

    fuel_pdf_file = st.file_uploader(
        "上傳加油發票PDF（自動解析總計金額）",
        type="pdf", key="fuel_pdf_upload"
    )

    # 上傳PDF後自動解析、排序並填入
    if fuel_pdf_file:
        if st.button("🔍 自動解析發票金額", key="parse_fuel"):
            with st.spinner("解析與日期自動排序中..."):
                fuel_pdf_file.seek(0)
                parsed = parse_fuel_pdf_totals(fuel_pdf_file.read())

            if parsed:
                # 填入前 10 張，多的截掉
                for i, total in enumerate(parsed[:10], 1):
                    st.session_state[f"inv_t{i}"] = total
                    sales = round(total / 1.05)
                    st.session_state[f"inv_x{i}"] = round(sales * 0.05)
                # 剩餘欄位清空
                for i in range(len(parsed[:10]) + 1, 11):
                    st.session_state[f"inv_t{i}"] = 0
                    st.session_state[f"inv_x{i}"] = 0

                st.markdown(f"""
                <div class="success-box">
                ✅ 解析到 <b>{len(parsed)}</b> 筆發票（已依日期自動排序）：{parsed[:10]}
                {"（超過10張，請分批上傳）" if len(parsed) > 10 else ""}
                </div>""", unsafe_allow_html=True)
                
                # 強制觸發頁面刷新，使對照表立刻渲染
                st.rerun()
            else:
                st.markdown("""
                <div class="warn-box">
                ⚠️ 未自動解析到金額，請手動輸入
                </div>""", unsafe_allow_html=True)

    hc1, hc2 = st.columns([3, 2])
    with hc1: st.markdown("<div style='font-size:.8rem;color:#888;padding:2px 0'>發票總額</div>", unsafe_allow_html=True)
    with hc2: st.markdown("<div style='font-size:.8rem;color:#888;padding:2px 0'>稅額（可修改）</div>", unsafe_allow_html=True)

    invoice_rows = []
    for i in range(1, 11):
        ic1, ic2 = st.columns([3, 2])
        with ic1:
            total = st.number_input(
                f"總額{i}", min_value=0, step=1,
                key=f"inv_t{i}",
                on_change=auto_tax, args=(i,),
                label_visibility="collapsed"
            )
        with ic2:
            tax = st.number_input(
                f"稅額{i}", min_value=0, step=1,
                key=f"inv_x{i}",
                label_visibility="collapsed"
            )
        if total > 0:
            invoice_rows.append((total, tax))

    # 有資料就即時顯示結算表
    if invoice_rows:
        st.markdown("---")
        html_table, total_amount, total_tax, km, amt = build_results_html(
            invoice_rows, mileage_input
        )
        # 即時更新全域 state 以渲染最上方的對照表
        st.session_state.fuel_amount = total_amount
        st.session_state.fuel_tax = total_tax
        st.session_state.mileage_distance = km
        
        components.html(html_table, height=400, scrolling=False)

        # 快速摘要（方便複製數字）
        st.markdown(f"""
        <div class="info-box" style="margin-top:.5rem">
        <b>Concur 填寫摘要</b><br>
        Fuel → Amount：<b>{total_amount:,}</b>　Tax Amount：<b>{total_tax:,}</b><br>
        Personal Car → Distance：<b>{km:,} 公里</b>（金額 {amt:,}）
        </div>""", unsafe_allow_html=True)
    else:
        st.markdown("""
        <div style="color:#aaa;text-align:center;padding:2rem 0;font-size:.9rem;">
        輸入發票金額後即時顯示結算表
        </div>""", unsafe_allow_html=True)


st.markdown("""
<div style="font-size:.75rem;color:#bbb;text-align:center;margin-top:2rem">
🔒 所有資料僅在本機處理，不上傳任何伺服器
</div>""", unsafe_allow_html=True)
