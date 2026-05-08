"""
PDF → Excel 結案表填充器
功能:上傳銀河專案 PDF,自動填入結案表模板對應位置
"""
import streamlit as st
import pdfplumber
import re
import unicodedata
import subprocess
import tempfile
import os
import shutil
from pathlib import Path
from io import BytesIO
from openpyxl import load_workbook
from datetime import datetime

st.set_page_config(page_title="PDF → 結案表填充器", page_icon="📊", layout="wide")

# ==================== PDF 抽取邏輯 ====================

def normalize(s):
    """去空白、Unicode 正規化(把 ⽂→文 等異體字統一)"""
    if s is None:
        return ""
    s = str(s)
    s = unicodedata.normalize('NFKC', s)
    return re.sub(r'\s+', '', s)


def is_main_data_row(row):
    """判斷是否為主表資料列"""
    if not row or len(row) < 13:
        return False
    站版 = normalize(row[0])
    if not 站版 or 站版 == "合計" or "站版" in 站版:
        return False
    try:
        int(row[4])
        return True
    except (ValueError, TypeError):
        return False


# 主軸縮寫對照(PDF 中主軸欄位常被切成單字)
主軸縮寫對照 = {
    "自": "自品置入",
    "需": "需求置入",
    "圖": "圖書/閱讀",
    "會": "會員/品牌",
    "百": "百貨",
    "競": "競品置入",
    "全": "全域置入",
    "社": "社團置入",
}

# 站版前綴關鍵字 — 用來辨識一行是否以「站版」開頭
站版前綴 = ["Threads", "PTT", "Facebook", "Dcard", "Mobile01", "eyny", "Plurk",
             "BabyHome", "FG", "伊莉", "狄卡", "批踢踢", "FashionGuide"]

# 行尾「主軸縮寫 + 日期前半 + 9 個數字」的特徵
ROW_PATTERN = re.compile(
    r'(自|需|圖|會|百|競|全|社)\s+'                                # 主軸縮寫
    r'\d{4}-\s+'                                                   # 日期 yyyy-
    r'(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+'                            # 發文 回應 聲量 正向
    r'(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)'                       # 正面 負面 議題 產品 複合
    r'\s*$', re.MULTILINE
)


def extract_main_table_via_text(pdf_path):
    """用「行尾 9 個數字」特徵從純文字中抓主表資料(比 extract_tables 穩健)"""
    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            full_text += unicodedata.normalize('NFKC', page_text) + "\n"
    
    rows = []
    for m in ROW_PATTERN.finditer(full_text):
        主軸縮寫 = m.group(1)
        nums = [int(m.group(i)) for i in range(2, 11)]
        
        # 取該行從行首到主軸縮寫之前的文字 = 站版 + 標題
        line_start = full_text.rfind('\n', 0, m.start()) + 1
        line_part = full_text[line_start:m.start()].strip()
        
        # 確認該行包含某個站版前綴
        has_prefix = any(kw in line_part for kw in 站版前綴)
        if not has_prefix:
            continue
        
        rows.append({
            "站版": line_part,
            "標題": "",  # 文字解析時站版+標題混在一起,標題可選
            "主軸": 主軸縮寫對照[主軸縮寫],
            "專案發文量": nums[0],
            "網友回應量": nums[1],
            "討論聲量總數": nums[2],
            "正向聲量總數": nums[3],
            "正面討論": nums[4],
            "負面討論": nums[5],
            "議題討論": nums[6],
            "產品討論": nums[7],
            "複合討論": nums[8],
        })
    return rows


def extract_main_table(pdf_path):
    """抽取主表(優先用文字解析,備援用表格解析)"""
    # 文字解析法:更穩健,能避免 pdfplumber 跨頁切錯
    rows = extract_main_table_via_text(pdf_path)
    if rows:
        return rows
    
    # 備援:表格解析(萬一文字解析失效)
    fallback_rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                if not table:
                    continue
                for row in table:
                    if is_main_data_row(row):
                        fallback_rows.append({
                            "站版": normalize(row[0]),
                            "標題": (row[1] or "").strip().replace('\n', ' '),
                            "主軸": normalize(row[2]),
                            "專案發文量": int(row[4]),
                            "網友回應量": int(row[5]),
                            "討論聲量總數": int(row[6]),
                            "正向聲量總數": int(row[7]),
                            "正面討論": int(row[8]),
                            "負面討論": int(row[9]),
                            "議題討論": int(row[10]),
                            "產品討論": int(row[11]),
                            "複合討論": int(row[12]),
                        })
    return fallback_rows


def extract_post_types(pdf_path):
    """抽取「文案類型發文篇數」表格,回傳 dict {類型名: 數量}"""
    types = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                if not table:
                    continue
                head = [normalize(c) for c in table[0]] if table[0] else []
                if "文案類型" in head and "累計" in head:
                    for row in table[1:]:
                        if len(row) >= 2 and row[0] and row[1] is not None:
                            name = normalize(row[0])
                            if name in ("合計", ""):
                                continue
                            try:
                                types[name] = int(str(row[1]).strip())
                            except ValueError:
                                pass
    return types


# ==================== 計算邏輯 ====================

def categorize_station(stb):
    """把站版名稱對應到模板分類"""
    s = normalize(stb)
    if "PTT" in s.upper():
        return "PTT"
    if "Dcard" in s or "狄卡" in s:
        return "Dcard"
    if "Threads" in s:
        return "Threads"
    if "Facebook" in s or "FB" in s:
        return "FB社團"
    return "其他版面"


def calculate_all(main_data, post_types):
    """計算所有要填入模板的數字"""
    result = {}
    
    # ===== 工作表 1:篇數 =====
    # 模板欄位 → PDF 對應名稱
    type_mapping = {
        "2~3圖+分享文(素人拍+過稿)": ["2~3圖+分享文(提供照片+過稿)", "2~3圖+分享文(素人拍+過稿)", "2~3圖+分享文"],
        "1圖+分享文(素人拍+過稿)": ["1圖+分享文(提供照片+過稿)", "1圖+分享文(素人拍+過稿)", "1圖+分享文"],
        "分享文": ["分享文"],
        "詳文": ["詳文"],
        "主文": ["主文"],
        "推文": ["推文"],
        "FB-主文": ["FB-主文"],
        "FB-推文": ["FB-推文"],
        "置入主文": ["置入主文"],
        "置入推文": ["置入推文"],
    }
    
    篇數 = {}
    for template_name, pdf_names in type_mapping.items():
        value = 0
        for pdf_name in pdf_names:
            for k, v in post_types.items():
                if normalize(pdf_name) == normalize(k):
                    value = v
                    break
            if value > 0:
                break
        篇數[template_name] = value
    result["篇數"] = 篇數
    
    # ===== 工作表 2:KPI =====
    kpi = {}
    # 網友回應數 = 網友回應量整列加總
    kpi["網友回應數"] = sum(r["網友回應量"] for r in main_data)
    # 議題曝光數 = 主軸不含「置入」的列數
    kpi["議題曝光數"] = sum(1 for r in main_data if "置入" not in r["主軸"])
    # 內文指名度 = 跳過(使用者填)
    kpi["內文指名度"] = None
    # 好評增加數 = 各列「正向聲量總數 - 正面討論」先逐列相減再加總
    kpi["好評增加數"] = sum(r["正向聲量總數"] - r["正面討論"] for r in main_data)
    result["KPI"] = kpi
    
    # ===== 工作表 3:網友回應分布 =====
    # 模板順序:PTT、Dcard、Threads、其他版面、FB社團
    分布 = {}
    for cat in ["PTT", "Dcard", "Threads", "其他版面", "FB社團"]:
        items = [r for r in main_data if categorize_station(r["站版"]) == cat]
        分布[cat] = {
            "實際溝通": sum(r["專案發文量"] for r in items),
            "正面討論": sum(r["正面討論"] for r in items),
            "負面討論": sum(r["負面討論"] for r in items),
            "議題討論": sum(r["議題討論"] for r in items),
            "產品討論": sum(r["產品討論"] for r in items),
            "複合討論": sum(r["複合討論"] for r in items),
        }
    result["分布"] = 分布
    result["分布來源"] = {
        cat: [r["站版"] for r in main_data if categorize_station(r["站版"]) == cat]
        for cat in ["PTT", "Dcard", "Threads", "其他版面", "FB社團"]
    }
    
    return result


# ==================== 填充模板 ====================

def fill_template(template_path, calc_result):
    """把計算結果填入模板,回傳填好的 xlsx bytes"""
    wb = load_workbook(template_path)
    
    # ----- 工作表 1:篇數 -----
    ws1 = wb["篇數"]
    篇數欄位映射 = {
        "2~3圖+分享文(素人拍+過稿)": "C8",
        "1圖+分享文(素人拍+過稿)": "D8",
        "分享文": "E8",
        "詳文": "F8",
        "主文": "G8",
        "推文": "H8",
        "FB-主文": "I8",
        "FB-推文": "J8",
        "置入主文": "K8",
        "置入推文": "L8",
    }
    for name, cell in 篇數欄位映射.items():
        ws1[cell] = calc_result["篇數"].get(name, 0)
    
    # 區塊 A 上方小表(第 3-5 列):欄位順序跟第 7-8 列相同
    # 第 4 列「篇數」= 與第 8 列同樣的數字
    # 第 5 列「總計」= 橫列加總
    上方小表映射 = {
        "2~3圖+分享文(素人拍+過稿)": "C4",  # 對應 C3「2~3圖分享文」
        "1圖+分享文(素人拍+過稿)": "D4",   # 對應 D3「1圖分享文」
        "分享文": "E4",
        "詳文": "F4",
        "主文": "G4",
        "推文": "H4",
        "FB-主文": "I4",
        "FB-推文": "J4",
        "置入主文": "K4",
        "置入推文": "L4",
    }
    for name, cell in 上方小表映射.items():
        ws1[cell] = calc_result["篇數"].get(name, 0)
    # 第 5 列「總計」= 橫列加總
    ws1["C5"] = "=SUM(C4:L4)"
    
    # 區塊 B (直式清單) - 第 14~24 列
    區塊B映射 = {
        "2~3圖+分享文(素人拍+過稿)": "C15",
        "1圖+分享文(素人拍+過稿)": "C16",
        "分享文": "C17",
        "詳文": "C18",
        "主文": "C19",
        "推文": "C20",
        "FB-主文": "C21",
        "FB-推文": "C22",
        "置入主文": "C23",
        "置入推文": "C24",
    }
    for name, cell in 區塊B映射.items():
        ws1[cell] = calc_result["篇數"].get(name, 0)
    # C25 是合計,用公式
    ws1["C25"] = "=SUM(C15:C24)"
    
    # ----- 工作表 2:KPI -----
    ws2 = wb["KPI"]
    ws2["D4"] = f"{calc_result['KPI']['網友回應數']}篇"
    ws2["D5"] = f"{calc_result['KPI']['議題曝光數']}串"
    # D6 內文指名度 - 跳過(留原值或清空)
    ws2["D6"] = ""  # 清空,讓使用者自己填
    ws2["D7"] = f"{calc_result['KPI']['好評增加數']}篇"
    
    # ----- 工作表 3:網友回應分布 -----
    ws3 = wb["網友回應分布"]
    # 列對應(已知模板:PTT=5列, Dcard=6列, Threads=7列, 其他=8列, FB=9列)
    版面列對應 = {
        "PTT": 5,
        "Dcard": 6,
        "Threads": 7,
        "其他版面": 8,
        "FB社團": 9,
    }
    for cat, row_num in 版面列對應.items():
        d = calc_result["分布"][cat]
        ws3.cell(row=row_num, column=3).value = d["實際溝通"]   # C
        ws3.cell(row=row_num, column=4).value = d["正面討論"]   # D
        ws3.cell(row=row_num, column=5).value = d["負面討論"]   # E
        ws3.cell(row=row_num, column=6).value = d["議題討論"]   # F
        ws3.cell(row=row_num, column=7).value = d["產品討論"]   # G
        ws3.cell(row=row_num, column=8).value = d["複合討論"]   # H
        # I 欄(回應小計)= 正+負+議+產+複,用公式自動計算
        ws3.cell(row=row_num, column=9).value = f"=SUM(D{row_num}:H{row_num})"
    # 第 14 列「發文篇數總計」模板已內建公式 =SUM(C5:C13) 等,不需另外處理
    
    # 輸出到 bytes
    output = BytesIO()
    wb.save(output)
    return output.getvalue()


def convert_xls_to_xlsx(xls_bytes):
    """把 .xls 轉成 .xlsx (因為 openpyxl 無法處理舊版 .xls)"""
    with tempfile.TemporaryDirectory() as tmpdir:
        in_path = os.path.join(tmpdir, "input.xls")
        with open(in_path, "wb") as f:
            f.write(xls_bytes)
        
        result = subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'xlsx', '--outdir', tmpdir, in_path],
            capture_output=True, text=True, timeout=120
        )
        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice 轉檔失敗: {result.stderr}")
        
        out_path = os.path.join(tmpdir, "input.xlsx")
        if not os.path.exists(out_path):
            raise RuntimeError("LibreOffice 沒有產生 xlsx 檔")
        
        with open(out_path, "rb") as f:
            return f.read()


# ==================== Streamlit UI ====================

st.title("📊 PDF → 結案表填充器")
st.caption("上傳銀河 PDF + Excel 模板,自動算出對應數字並填入")

with st.sidebar:
    st.header("📋 填充規則")
    st.markdown("""
    **工作表 1 - 篇數**
    對應 PDF「文案類型發文篇數」
    
    **工作表 2 - KPI**
    - 網友回應數 = 網友回應量加總
    - 議題曝光數 = 主軸非「置入」的列數
    - 內文指名度 = 留空(手動填)
    - 好評增加數 = 正向聲量總和 − 正面討論總和
    
    **工作表 3 - 網友回應分布**
    - 實際溝通 = 專案發文量
    - 各討論欄 = 各自加總
    - 回應小計 = 公式自動計算
    - 未列出的版面 → 併入「其他版面」
    """)

col_l, col_r = st.columns(2)

with col_l:
    st.subheader("1️⃣ 上傳 Excel 模板")
    template_file = st.file_uploader(
        "結案表格(.xls 或 .xlsx)",
        type=["xls", "xlsx"],
        key="template",
    )

with col_r:
    st.subheader("2️⃣ 上傳 PDF 報表")
    pdf_file = st.file_uploader(
        "銀河專案 PDF",
        type=["pdf"],
        key="pdf",
    )

if template_file and pdf_file:
    st.divider()
    
    # 處理模板:.xls 要先轉成 .xlsx
    template_bytes = template_file.getvalue()
    if template_file.name.lower().endswith(".xls"):
        with st.spinner("🔄 轉換 .xls 模板格式中..."):
            try:
                template_bytes = convert_xls_to_xlsx(template_bytes)
            except Exception as e:
                st.error(f"❌ 模板轉檔失敗: {e}")
                st.info("💡 提示:請先用 Excel 把 .xls 另存為 .xlsx 再上傳")
                st.stop()
    
    # 把 PDF 存到暫存檔(pdfplumber 需要 path 或 fileobj)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
        tmp_pdf.write(pdf_file.getvalue())
        pdf_path = tmp_pdf.name
    
    try:
        with st.spinner("🔍 解析 PDF 並計算數據..."):
            main_data = extract_main_table(pdf_path)
            post_types = extract_post_types(pdf_path)
        
        if not main_data:
            st.error("❌ 無法從 PDF 抽取主表資料,請確認上傳的是正確格式的銀河專案 PDF")
            st.stop()
        
        result = calculate_all(main_data, post_types)
        
        st.success(f"✅ 解析成功!共抓到 **{len(main_data)}** 篇文章資料")
        
        # ===== 預覽計算結果 =====
        st.subheader("👀 預覽計算結果")
        
        tab1, tab2, tab3, tab4 = st.tabs(["📑 篇數", "📈 KPI", "📊 網友回應分布", "📂 原始資料"])
        
        with tab1:
            st.write("**將填入工作表「篇數」**")
            篇數 = result["篇數"]
            cols = st.columns(5)
            for i, (k, v) in enumerate(篇數.items()):
                cols[i % 5].metric(k, v)
            st.caption(f"📍 PDF 中抓到的文案類型:{', '.join(post_types.keys()) if post_types else '(無)'}")
        
        with tab2:
            st.write("**將填入工作表「KPI」**")
            kpi = result["KPI"]
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("網友回應數", f"{kpi['網友回應數']} 篇")
            c2.metric("議題曝光數", f"{kpi['議題曝光數']} 串")
            c3.metric("內文指名度", "(手動填)", delta="留空")
            c4.metric("好評增加數", f"{kpi['好評增加數']} 篇")
            
            with st.expander("📐 計算明細"):
                st.write(f"- **網友回應數** = 網友回應量加總 = **{kpi['網友回應數']}**")
                non_p = [r['主軸'] for r in main_data if "置入" not in r['主軸']]
                st.write(f"- **議題曝光數** = 主軸非「置入」的列數 = **{kpi['議題曝光數']}**")
                st.write(f"  非置入主軸明細: {non_p}")
                # 好評增加數明細:各列(正向聲量 - 正面討論)
                row_diffs = [(r['站版'][:20], r['正向聲量總數'], r['正面討論'], r['正向聲量總數']-r['正面討論']) for r in main_data]
                st.write(f"- **好評增加數** = 各列(正向聲量−正面討論)先計算再加總 = **{kpi['好評增加數']}**")
                with st.expander("查看每列計算"):
                    import pandas as pd
                    df_diff = pd.DataFrame(row_diffs, columns=["站版", "正向聲量", "正面討論", "差值"])
                    st.dataframe(df_diff, use_container_width=True, hide_index=True)
        
        with tab3:
            st.write("**將填入工作表「網友回應分布」**")
            import pandas as pd
            df_dist = pd.DataFrame(result["分布"]).T
            df_dist["回應小計"] = df_dist[["正面討論", "負面討論", "議題討論", "產品討論", "複合討論"]].sum(axis=1)
            df_dist = df_dist[["實際溝通", "正面討論", "負面討論", "議題討論", "產品討論", "複合討論", "回應小計"]]
            st.dataframe(df_dist, use_container_width=True)
            
            with st.expander("🗂️ 各分類包含的站版"):
                for cat, sources in result["分布來源"].items():
                    if sources:
                        st.write(f"**{cat}**: {', '.join(set(sources))}")
                    else:
                        st.write(f"**{cat}**: (無)")
        
        with tab4:
            st.write("**從 PDF 抽出的原始明細**")
            import pandas as pd
            df_raw = pd.DataFrame(main_data)
            st.dataframe(df_raw, use_container_width=True, hide_index=True)
        
        st.divider()
        
        # ===== 產生並下載填好的 Excel =====
        with st.spinner("📝 填入模板..."):
            try:
                filled_bytes = fill_template(BytesIO(template_bytes), result)
            except Exception as e:
                st.error(f"❌ 填入模板失敗: {e}")
                import traceback
                st.code(traceback.format_exc())
                st.stop()
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_name = f"結案表格_已填_{timestamp}.xlsx"
        
        st.download_button(
            label="📥 下載填好的 Excel",
            data=filled_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )
        
        st.info("💡 **提醒**:\n"
                "- 工作表 KPI 的「內文指名度」需要您手動填入\n"
                "- 原模板已保留(只是用上傳檔生成新檔)\n"
                "- 「網友回應分布」第一個區塊(第 4-13 列)未填,因模板註記為「系統抓取後手動處理」")
    
    finally:
        # 清理暫存檔
        if os.path.exists(pdf_path):
            os.unlink(pdf_path)

else:
    st.info("👆 請上傳 Excel 模板 + PDF 報表後開始")
