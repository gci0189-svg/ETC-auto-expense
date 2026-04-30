"""
DN 費用申報整合工具 v4
=====================
佈局：單頁寬版
  上方：st.columns([2, 1])
    左 2/3 → 通行費對帳（上傳T_E申請表＋遠通電收PDF）
    右 1/3 → 加油費計算（手動輸入發票金額，顯示結算表）
  下方：橫線分隔 → 電信費處理（移除密碼＋擷取第一頁）

安裝：
  pip install streamlit openpyxl pdfplumber pymupdf pypdf
  streamlit run app.py
"""

import streamlit as st
import streamlit.components.v1 as components
import openpyxl
import pdfplumber
import fitz
import io, os, re, math
from datetime import datetime

try:
    from pypdf import PdfReader, PdfWriter
    PYPDF_AVAILABLE = True
except ImportError:
    PYPDF_AVAILABLE = False

# ─────────────────────────────────────────
# 頁面設定
# ─────────────────────────────────────────
st.set_page_config(page_title="DN 費用申報整合工具", layout="wide", page_icon="🚗")

st.markdown("""
<style>
  .block-container{padding-top:1.2rem;padding-bottom:1rem}
  h1{font-size:1.5rem!important;color:#1F4E79}
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

st.markdown("# 🚗 DN 費用申報整合工具")

# Session State
for k in ['toll_excel','toll_pdf_out','telecom_pdf','mileage_allowance','selected_sheet']:
    if k not in st.session_state:
        st.session_state[k] = None

# ═══════════════════════════════════════════
# 工具函式
# ═══════════════════════════════════════════

def format_date_slash(v):
    try:
        if isinstance(v, str):
            return datetime.strptime(v.strip(), '%d-%b-%y').strftime('%Y/%m/%d')
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


def parse_toll_from_pdf(pdf_bytes):
    toll_map = {}
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split('\n'):
                parts = line.split()
                if len(parts) >= 3 and re.match(r'\d{4}/\d{2}/\d{2}', parts[0]):
                    amt = parts[2].replace('元', '')
                    if amt.isdigit() and parts[0] not in toll_map:
                        toll_map[parts[0]] = int(amt)
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


def remove_pdf_password_and_extract_page1(pdf_bytes, password=""):
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        total_pages = len(doc)
        if doc.is_encrypted:
            if not doc.authenticate(password):
                doc.close()
                return False, None, f"密碼錯誤（嘗試：「{password}」）"
        new_doc = fitz.open()
        new_doc.insert_pdf(doc, from_page=0, to_page=0)
        out = io.BytesIO()
        new_doc.save(out, encryption=fitz.PDF_ENCRYPT_NONE)
        new_doc.close(); doc.close()
        return True, out.getvalue(), f"成功！已移除密碼並擷取第1頁（共 {total_pages} 頁）"
    except Exception:
        pass
    if PYPDF_AVAILABLE:
        try:
            reader = PdfReader(io.BytesIO(pdf_bytes))
            if reader.is_encrypted:
                if reader.decrypt(password) == 0:
                    return False, None, f"密碼錯誤（嘗試：「{password}」）"
            writer = PdfWriter()
            writer.add_page(reader.pages[0])
            out = io.BytesIO()
            writer.write(out)
            return True, out.getvalue(), f"成功（備援）！擷取第1頁（共 {len(reader.pages)} 頁）"
        except Exception as e:
            return False, None, f"處理失敗：{e}"
    return False, None, "解密失敗，請確認密碼"


def build_results_html(invoice_rows, mileage_allowance):
    """
    invoice_rows: list of (total, tax)  每張發票
    回傳仿試算表的 HTML 字串
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
# 主要佈局：左 2/3（通行費）｜ 右 1/3（加油費）
# ═══════════════════════════════════════════
col_toll, col_fuel = st.columns([3, 2], gap="large")

# ╔══════════════════════════════════════════╗
# ║  左側：通行費對帳                        ║
# ╚══════════════════════════════════════════╝
with col_toll:
    st.markdown('<div class="section-title">🛣️ 通行費對帳</div>', unsafe_allow_html=True)

    toll_pdf = st.file_uploader("① 遠通電收 PDF", type="pdf", key="toll_pdf")
    te_excel = st.file_uploader("② T_E 申請表 (.xlsx)", type="xlsx", key="te_main")

    selected_sheet = None
    if te_excel:
        wb_tmp = openpyxl.load_workbook(te_excel, read_only=True)
        sheets = wb_tmp.sheetnames
        cm = f"{datetime.now().month}月"
        default_idx = sheets.index(cm) if cm in sheets else 0
        selected_sheet = st.selectbox("③ 選擇月份工作表", sheets, index=default_idx, key="s_main")
        st.session_state.selected_sheet = selected_sheet

        if selected_sheet:
            te_excel.seek(0)
            allowance = read_mileage_allowance(te_excel.read(), selected_sheet)
            if allowance:
                st.session_state.mileage_allowance = allowance
                st.markdown(f"""
                <div class="success-box">
                ✅ <b>{selected_sheet}</b> 里程津貼小計：<b>NT$ {int(allowance):,}</b>
                （已同步至右側加油費計算）
                </div>""", unsafe_allow_html=True)

    if toll_pdf and te_excel and selected_sheet:
        if st.button("🚀 開始對帳與標註", type="primary", key="run_toll"):
            with st.spinner("處理中..."):
                try:
                    toll_pdf.seek(0)
                    toll_map = parse_toll_from_pdf(toll_pdf.read())
                    if not toll_map:
                        st.error("無法解析通行費PDF，請確認格式")
                        st.stop()

                    te_excel.seek(0)
                    wb = openpyxl.load_workbook(te_excel)
                    ws = wb[selected_sheet]
                    DATE_COL, TOLL_COL, ITEM_COL = 4, 11, 1
                    serial_map, matched = {}, set()

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

                    out_excel = io.BytesIO()
                    wb.save(out_excel)
                    st.session_state.toll_excel = out_excel.getvalue()

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
                    out_pdf = io.BytesIO()
                    doc.save(out_pdf)
                    st.session_state.toll_pdf_out = out_pdf.getvalue()

                    st.success(f"✅ 完成！共比對 **{len(matched)}** 筆通行費")
                    unmatched = set(toll_map.keys()) - matched
                    if unmatched:
                        st.markdown(f"""<div class="warn-box">
                        ⚠️ PDF有記錄但申請表未找到的日期：{', '.join(sorted(unmatched))}
                        </div>""", unsafe_allow_html=True)

                except Exception as e:
                    st.error(f"錯誤：{e}")
                    import traceback; st.code(traceback.format_exc())

    # 下載區
    dl1, dl2 = st.columns(2, gap="small")
    with dl1:
        if st.session_state.toll_excel:
            te_name = te_excel.name if te_excel else "T_E申請表.xlsx"
            st.download_button(
                "💾 下載更新後的 Excel",
                st.session_state.toll_excel,
                f"{selected_sheet}_通行費_{te_name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    with dl2:
        if st.session_state.toll_pdf_out and toll_pdf:
            st.download_button(
                "💾 下載標註後的 PDF",
                st.session_state.toll_pdf_out,
                f"標註_{selected_sheet}_{toll_pdf.name}",
                mime="application/pdf"
            )

    # 通行費預覽
    if toll_pdf:
        with st.expander("🔍 預覽遠通電收解析結果"):
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
    st.markdown('<div class="section-title">⛽ 加油費計算</div>', unsafe_allow_html=True)

    # 里程津貼：優先從左側申請表自動帶入
    if st.session_state.mileage_allowance:
        mileage_val = int(st.session_state.mileage_allowance)
        st.markdown(f"""
        <div class="info-box">
        📊 里程津貼自動帶入：<b>NT$ {mileage_val:,}</b>
        </div>""", unsafe_allow_html=True)
    else:
        mileage_val = 0

    mileage_input = st.number_input(
        "💰 總里程津貼（可手動修改）",
        min_value=0,
        value=mileage_val,
        step=100,
        key="mileage_manual"
    )

    st.markdown("**🧾 輸入發票金額（每張填一列）**")
    st.caption("只需填「發票總額」，稅額自動計算 ceil(總額 ÷ 1.05 × 5%)")

    # 最多5張發票
    invoice_rows = []
    for i in range(1, 6):
        v = st.number_input(
            f"發票 {i}",
            min_value=0, value=0, step=1,
            key=f"inv_{i}",
            label_visibility="collapsed" if i > 1 else "visible"
        )
        if v > 0:
            tax = math.ceil(v / 1.05 * 0.05)
            invoice_rows.append((v, tax))

    # 有資料就即時顯示結算表
    if invoice_rows:
        st.markdown("---")
        html_table, total_amount, total_tax, km, amt = build_results_html(
            invoice_rows, mileage_input
        )
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
