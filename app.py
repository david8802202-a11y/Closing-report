"""
PDF → 報表填充器(主入口)
負責:
- 共用 PDF 解析
- Streamlit UI 入口與功能路由
- 把 PDF defaults 交給對應模組(closure_report / monthly_report / indexing_search)
"""
import os
import re
import tempfile
import unicodedata
from datetime import datetime
from io import BytesIO

import pdfplumber
import streamlit as st

import closure_report
import monthly_report
import indexing_search


st.set_page_config(page_title="PDF → 報表填充器", page_icon="📊", layout="wide")


# ==================== 共用 PDF 解析 ====================
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
# 由於不同專案可能有不同主軸類型,擴充常見對照;未列出的縮寫會用「文字解析」自動補
主軸縮寫對照 = {
    "自": "自品置入",
    "需": "需求置入",
    "圖": "圖書/閱讀",
    "會": "會員/品牌",
    "百": "百貨",
    "競": "競品置入",  # 或「競品議題置入」
    "全": "全域置入",
    "社": "社團置入",
    "消": "消費者需求置入",
    "情": "情境話題",
    "產": "產品話題",
    "話": "話題",
}

# 站版前綴關鍵字 — 用來辨識一行是否以「站版」開頭
站版前綴 = ["Threads", "PTT", "Facebook", "Dcard", "Mobile01", "eyny", "Plurk",
             "BabyHome", "FG", "伊莉", "狄卡", "批踢踢", "FashionGuide"]

# 行尾「主軸 + 日期前半 + 9 個數字」的通用特徵
# 主軸允許:任何非數字、非空白的字符組合 — 不限定中文開頭
# 涵蓋「自/需/圖」「全院-」「miradry」「ABC-123」等任何寫法
ROW_PATTERN = re.compile(
    r'([^\s\d][^\s]*)\s+'                                          # 主軸縮寫(非數字、非空白開頭)
    r'\d{4}-\s+'                                                   # 日期 yyyy-
    r'(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+'                            # 發文 回應 聲量 正向
    r'(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)'                       # 正面 負面 議題 產品 複合
    r'\s*$', re.MULTILINE
)


# 主軸縮寫對照(PDF 中主軸欄位常被切成單字)
# 由於不同專案可能有不同主軸類型,擴充常見對照;未列出的縮寫會用「文字解析」自動補
主軸縮寫對照 = {
    "自": "自品置入",
    "需": "需求置入",
    "圖": "圖書/閱讀",
    "會": "會員/品牌",
    "百": "百貨",
    "競": "競品置入",  # 或「競品議題置入」
    "全": "全域置入",
    "社": "社團置入",
    "消": "消費者需求置入",
    "情": "情境話題",
    "產": "產品話題",
    "話": "話題",
}

# 站版前綴關鍵字 — 用來辨識一行是否以「站版」開頭
站版前綴 = ["Threads", "PTT", "Facebook", "Dcard", "Mobile01", "eyny", "Plurk",
             "BabyHome", "FG", "伊莉", "狄卡", "批踢踢", "FashionGuide"]

# 行尾「主軸 + 日期前半 + 9 個數字」的通用特徵
# 主軸允許:任何非數字、非空白的字符組合 — 不限定中文開頭
# 涵蓋「自/需/圖」「全院-」「miradry」「ABC-123」等任何寫法
ROW_PATTERN = re.compile(
    r'([^\s\d][^\s]*)\s+'                                          # 主軸縮寫(非數字、非空白開頭)
    r'\d{4}-\s+'                                                   # 日期 yyyy-
    r'(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+'                            # 發文 回應 聲量 正向
    r'(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)'                       # 正面 負面 議題 產品 複合
    r'\s*$', re.MULTILINE
)




def get_main_category_from_pdf(pdf_path):
    """從 PDF「操作主軸」區塊抽出實際主軸全名清單(用文字解析,避免跨頁切表問題)"""
    return [name for name, _ in get_main_category_with_counts(pdf_path)]

def get_main_category_with_counts(pdf_path):
    """從 PDF「操作主軸」區塊抽出主軸全名 + 累計串數
    
    回傳 [(主軸名, 累計串數), ...] 依 PDF 中的順序
    
    注意:PDF 上「操作主軸表」的「累計」欄是權威來源,直接使用它,
    避免因主表跨頁破碎而漏抓主軸串數。
    """
    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for page in pdf.pages:
            full_text += unicodedata.normalize('NFKC', page.extract_text() or "") + "\n"
    
    extracted = []
    
    # 找「操作主軸」區段
    start_idx = full_text.find("操作主軸")
    if start_idx < 0:
        return extracted
    
    # 找區段結尾(碰到下個區段就停)
    end_markers = ["專案聲量分佈", "網友關注度", "網友回應概況"]
    end_idx = len(full_text)
    for marker in end_markers:
        idx = full_text.find(marker, start_idx + 10)
        if 0 < idx < end_idx:
            end_idx = idx
    
    section = full_text[start_idx:end_idx]
    
    # 匹配「主軸名 + 空白 + 純數字」
    for line in section.split('\n'):
        line = line.strip()
        if not line:
            continue
        if line in ("操作主軸", "類型 累計", "類型", "累計"):
            continue
        m = re.match(r'^(.+?)\s+(\d+)\s*$', line)
        if m:
            name = m.group(1).strip()
            count = int(m.group(2))
            if name == "合計":
                continue
            # 排除站版列、含百分比的列
            if ":" in name or "%" in name:
                continue
            if re.search(r'\d+\.\d+', line):
                continue
            extracted.append((name, count))
    
    return extracted

def extract_guaranteed_reply(pdf_path):
    """從 PDF「專案執行進度摘要」抽出「網友回應數」的預計完成數
    
    PDF 格式:
      專案執行進度摘要
      專案累計 預計完成 完成率
      銀河發文量 X Y Z%
      網友回應數 X Y Z%    ← 抽這一列的「預計完成」= Y
      ...
    
    回傳 int 或 None(找不到)
    """
    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for page in pdf.pages:
            full_text += unicodedata.normalize('NFKC', page.extract_text() or "") + "\n"
    
    # 找「網友回應數」那一列(在「專案執行進度摘要」區塊裡)
    # 格式:「網友回應數 <累計> <預計完成> <完成率%>」
    # 用 re 抓出「網友回應數」後面第二個數字
    section_start = full_text.find("專案執行進度摘要")
    if section_start < 0:
        return None
    # 限制搜尋範圍到下一個區段
    section_end = len(full_text)
    for marker in ["操作主軸", "專案聲量分佈", "網友關注度", "網友回應概況"]:
        idx = full_text.find(marker, section_start + 10)
        if 0 < idx < section_end:
            section_end = idx
    section = full_text[section_start:section_end]
    
    # 抓「網友回應數 <數字> <數字>」— 前面數字是累計,第二個是預計完成
    m = re.search(r'網友回應數\s+(\d+)\s+(\d+)', section)
    if m:
        return int(m.group(2))
    return None

def resolve_main_category(prefix, full_names):
    """把截斷的主軸縮寫對應到完整名稱
    
    優先級:
    1. 完全相同
    2. 完整名稱開頭就是縮寫(prefix match)
    3. 無法對應 → 回傳原縮寫
    """
    if not prefix:
        return prefix
    # 1. 完全相同
    for name in full_names:
        if name == prefix:
            return name
    # 2. 開頭匹配(全名以縮寫開頭)
    for name in full_names:
        if name.startswith(prefix):
            return name
    # 3. 反向匹配(縮寫以全名開頭,例如「全院-逆時針」會出現「全院-逆」被截斷)
    for name in full_names:
        if prefix.startswith(name) and len(name) >= 2:
            return name
    return prefix

def extract_main_table_via_text(pdf_path):
    """用「行尾 9 個數字」特徵從純文字中抓主表資料(比 extract_tables 穩健)
    
    判斷邏輯:任何一行只要結尾是「中文字 + yyyy- + 9個數字」就視為資料列。
    這個結構足以唯一識別資料列(合計列、表頭、空白行都不會符合)。
    不依賴站版名稱字典,適用於任何站版。
    """
    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            full_text += unicodedata.normalize('NFKC', page_text) + "\n"
    
    # 從 PDF「操作主軸」表抓主軸全名清單(動態,不寫死)
    full_names = get_main_category_from_pdf(pdf_path)
    
    rows = []
    for m in ROW_PATTERN.finditer(full_text):
        主軸縮寫 = m.group(1)
        nums = [int(m.group(i)) for i in range(2, 11)]
        
        # 取該行從行首到主軸縮寫之前的文字 = 站版 + 標題
        line_start = full_text.rfind('\n', 0, m.start()) + 1
        line_part = full_text[line_start:m.start()].strip()
        
        # 過濾:行內容太短就跳過(避免抓到合計列等)
        if len(line_part) < 3:
            continue
        
        # 把可能截斷的主軸縮寫對應到完整名稱
        # 優先使用 PDF「操作主軸」表的全名清單做 prefix match
        # 查不到再用內建對照表;都查不到就用原縮寫
        主軸全名 = resolve_main_category(主軸縮寫, full_names)
        if 主軸全名 == 主軸縮寫 and 主軸縮寫 in 主軸縮寫對照:
            主軸全名 = 主軸縮寫對照[主軸縮寫]
        
        rows.append({
            "站版": line_part,
            "標題": "",
            "主軸": 主軸全名,
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
    """主表抽取的對外接口(目前用文字解析法,最穩定)"""
    return extract_main_table_via_text(pdf_path)

def extract_post_types(pdf_path):
    """抽取「文案類型發文篇數」表格,回傳 dict {類型名: 數量}
    
    用「純文字 + 區段定位」抓取,避免 pdfplumber 跨頁切表導致漏抓。
    """
    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for page in pdf.pages:
            full_text += unicodedata.normalize('NFKC', page.extract_text() or "") + "\n"
    
    types = {}
    
    # 找「文案類型發文篇數」這個區段
    start_idx = full_text.find("文案類型發文篇數")
    if start_idx < 0:
        start_idx = full_text.find("文案類型")
    if start_idx < 0:
        return types
    
    # 區段結尾的關鍵字(碰到就停)
    end_markers = [
        "專案執行進度摘要",
        "操作主軸",
        "專案聲量分佈",
        "網友關注度",
        "網友回應概況",
    ]
    
    # 找最近的結束標記
    end_idx = len(full_text)
    for marker in end_markers:
        idx = full_text.find(marker, start_idx + 10)
        if 0 < idx < end_idx:
            end_idx = idx
    
    section = full_text[start_idx:end_idx]
    
    # 匹配「文案類型名 + 空白 + 第一個數字」
    # 新版 PDF 可能有額外欄位:「文案類型 累計 預計完成 完成率」
    # 例如:「分享文 9 9 100.00%」— 只需抓第一個數字 9
    for line in section.split('\n'):
        line = line.strip()
        if not line:
            continue
        # 排除標題列(各種變體)
        if line in ("文案類型 累計", "文案類型", "累計",
                    "文案類型 累計 預計完成 完成率",
                    "文案類型 累計 預計 完成率"):
            continue
        # 排除百分比結尾的標題類字串(避免誤抓「文案類型 累計 預計完成 完成率」的位置)
        if line.startswith("文案類型 "):
            continue
        # 匹配「<類型名> <第一個數字> [後面可能還有其他欄位]」
        m = re.match(r'^(.+?)\s+(\d+)(?:\s+.*)?$', line)
        if m:
            name = m.group(1).strip()
            count = int(m.group(2))
            # 排除「合計」
            if name == "合計":
                continue
            # 排除站版分布的列(含「:」表示是站版名稱,不是文案類型)
            if ":" in name:
                continue
            # 排除類型名本身就含百分比的(不太可能,但保險)
            if "%" in name:
                continue
            types[name] = count
    
    return types

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


# ==================== 內建模板選擇(由各模組提供)====================
def get_builtin_template(report_type):
    """從對應模組取得 Excel 模板的 bytes"""
    if report_type == "結案表":
        return closure_report.get_excel_template_bytes()
    if report_type == "月報表":
        return monthly_report.get_excel_template_bytes()
    raise ValueError(f"未知的報表類型:{report_type}")


# ==================== Streamlit UI ====================
st.title("📊 PDF → 報表填充器")
st.caption("上傳銀河 PDF,自動算出對應數字並產出 Excel / Word / 簡報指令詞")

# === Step 1:選擇功能 ===
st.subheader("🎯 選擇要使用的功能")
report_type = st.radio(
    "功能類型",
    ["結案表", "月報表", "內文指名度查詢"],
    horizontal=True,
    key="report_type",
    help="前兩個會產出 Excel + Word + 簡報指令詞,第三個是查詢工具",
)

# === Sidebar:填充規則 ===
with st.sidebar:
    st.header("📋 填充規則")
    if report_type == "結案表":
        st.markdown("""
        **【結案表】填充規則**

        **工作表 1 - 篇數**
        對應 PDF「文案類型發文篇數」

        **工作表 2 - KPI**
        - 網友回應數 = 網友回應量加總
        - 議題曝光數 = 勾選主軸裡「不重複標題」的計數(同一標題只算 1 串)
        - 內文指名度 = 留空(手動填)
        - 好評增加數 = 各列(正向聲量−正面討論)加總

        **工作表 3 - 網友回應分布**
        - 實際溝通 = 專案發文量
        - 各討論欄 = 各自加總
        - 回應小計 = 公式自動計算
        - 未列出的版面 → 併入「其他版面」

        **延伸輸出**
        - Word 結案資訊整理(需補填部分欄位)
        - 結案指令詞 TXT(給 Gamma / Claude 做簡報)
        """)
    elif report_type == "月報表":
        st.markdown("""
        **【月報表】填充規則**

        **頁簽 1 - 網友回應摘要+操作數據**
        - 專案總發文量 = 發文量合計 + 文案類型明細
        - 網友總回應量 / 討論聲量(總數)
        - 各討論比率/篇數(分母 = 5 類加總)

        **頁簽 2 - 版面佔比**
        - PTT / DCARD / THREAD / 其他版面 的「網友回應量」加總

        **頁簽 3 - 總覽整理**
        - 不會被修改(保留模板原樣)

        **延伸輸出**
        - Word 月報資訊整理(需補填部分欄位)
        - 月報簡報指令詞 TXT
        """)
    else:
        st.markdown("""
        **【內文指名度查詢】使用說明**

        1. 上傳「專案發文總覽 PDF」
        2. 輸入關鍵字
        3. 選擇要搜尋「內文」或「標題」
        4. 顯示符合篇數 + 命中列表

        **關鍵字語法**
        - 單一詞: `Friday`
        - AND(都要含): `Friday+稻草人`
        - OR(任一含): `Friday,稻草人`
        """)

st.divider()

# === Step 2:上傳 PDF ===
st.subheader("📤 上傳 PDF 報表")
if report_type == "內文指名度查詢":
    st.caption("✨ 請上傳「專案發文總覽 PDF」(含 cnt.內文 欄位的版本)")
else:
    st.caption("✨ 模板已內建,直接上傳 PDF 即可")

pdf_file = st.file_uploader("銀河專案 PDF", type=["pdf"], key="pdf")

if not pdf_file:
    st.info("👆 請選擇功能,並上傳 PDF 報表後開始")
    st.stop()

st.divider()

# 把 PDF 存到暫存檔
with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
    tmp_pdf.write(pdf_file.getvalue())
    pdf_path = tmp_pdf.name

try:
    # ---------------------------------------------------
    # 分支 1:內文指名度查詢
    # ---------------------------------------------------
    if report_type == "內文指名度查詢":
        indexing_search.render_ui(pdf_path)

    # ---------------------------------------------------
    # 分支 2:結案表 / 月報表
    # ---------------------------------------------------
    else:
        try:
            template_bytes = get_builtin_template(report_type)
        except Exception as e:
            st.error(f"❌ 載入內建模板失敗:{e}")
            st.stop()

        with st.spinner("🔍 解析 PDF 並計算數據..."):
            main_data = extract_main_table(pdf_path)
            post_types = extract_post_types(pdf_path)

        if not main_data:
            st.error("❌ 無法從 PDF 抽取主表資料,請確認上傳的是正確格式的銀河專案 PDF")
            st.stop()

        st.success(f"✅ 解析成功!共抓到 **{len(main_data)}** 篇文章資料")

        st.subheader("👀 預覽計算結果")

        if report_type == "結案表":
            # ----- 主軸勾選 + 計算 -----
            主軸_累計 = get_main_category_with_counts(pdf_path)
            if not 主軸_累計:
                主軸統計 = {}
                for r in main_data:
                    主軸統計[r["主軸"]] = 主軸統計.get(r["主軸"], 0) + 1
                主軸_累計 = list(主軸統計.items())
            主軸_counts = dict(主軸_累計)
            主軸清單_排序 = sorted(主軸_累計, key=lambda x: (-x[1], x[0]))
            預設勾選 = [axis for axis, _ in 主軸清單_排序 if "置入" not in axis]
            主軸_實算串數 = {}
            for axis in {r["主軸"] for r in main_data}:
                unique_titles = {r["站版"] for r in main_data if r["主軸"] == axis}
                主軸_實算串數[axis] = len(unique_titles)

            with st.expander("🎯 選擇要計入「議題曝光數」的主軸", expanded=True):
                st.caption("💡 議題曝光數 = 勾選主軸中「不重複標題」的計數(同一標題不管幾則發文、不管幾個站版,都算 1 串)。預設勾選「非置入」類主軸。")
                cols = st.columns(3)
                selected_axes = []
                for i, (axis, _count) in enumerate(主軸清單_排序):
                    with cols[i % 3]:
                        default = axis in 預設勾選
                        串數 = 主軸_實算串數.get(axis, 0)
                        label = f"{axis} ({串數} 串)"
                        if st.checkbox(label, value=default, key=f"axis_{axis}"):
                            selected_axes.append(axis)
                if not selected_axes:
                    st.warning("⚠️ 未勾選任何主軸,議題曝光數會是 0")

            保證回應數 = extract_guaranteed_reply(pdf_path)
            result = closure_report.calculate_all(
                main_data, post_types,
                active_axes=selected_axes,
                axis_counts=主軸_counts,
                guaranteed_reply=保證回應數,
            )
            closure_report._render_closure_preview(
                result, main_data, post_types,
                selected_axes=selected_axes,
                axis_counts=主軸_counts,
            )
        else:
            result = None
            monthly_report._render_monthly_preview(main_data, post_types)

        st.divider()

        # ----- 填 Excel -----
        with st.spinner("📝 填入模板..."):
            try:
                if report_type == "結案表":
                    filled_bytes = closure_report.fill_template(BytesIO(template_bytes), result)
                else:
                    filled_bytes = monthly_report.fill_monthly_template(BytesIO(template_bytes), main_data, post_types)
            except Exception as e:
                st.error(f"❌ 填入模板失敗:{e}")
                import traceback
                st.code(traceback.format_exc())
                st.stop()

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_name = f"{report_type}_已填_{ts}.xlsx"
        st.download_button(
            label=f"📥 下載填好的 {report_type}",
            data=filled_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )

        # ----- 結案表:額外 Word + 指令詞 -----
        if report_type == "結案表":
            st.info("💡 **提醒**:\n"
                    "- 工作表 KPI 的「內文指名度」需要您手動填入(用第 3 個功能查詢)\n"
                    "- 「網友回應分布」第一個區塊未填,因模板註記為「系統抓取後手動處理」")

            st.divider()
            with st.expander("📄 生成 Word 結案資訊整理 + 結案指令詞(需補填部分欄位)", expanded=False):
                st.markdown("Word 包含 Excel 沒有的內容(聲量比較、正負評摘要、操作觀察等),需補填。指令詞 TXT 可貼給 Gamma/Claude 做簡報。")

                # 準備 pdf_defaults
                kpi_dict = result.get("KPI", {})
                篇數_dict = result.get("篇數", {})
                分布_dict = result.get("分布", {})

                細項_parts = []
                for name, cnt in 篇數_dict.items():
                    if cnt and cnt > 0:
                        short_name = name.split('(')[0].strip()
                        細項_parts.append(f"{short_name} {cnt} 篇")
                總篇數 = sum(v for v in 篇數_dict.values() if isinstance(v, (int, float)))
                細項_desc = "、".join(細項_parts) + f",共計 {總篇數} 篇" if 細項_parts else f"共 {總篇數} 篇"

                保證數 = kpi_dict.get("保證網友回應數") or 52
                word_kpi = [
                    {"項目": "網友回應數", "說明": f"開題與置入真實回應(原保證{保證數}篇)", "達成": f"{kpi_dict.get('網友回應數', 0)}篇"},
                    {"項目": "議題曝光數", "說明": "主動開題經營串數", "達成": f"{kpi_dict.get('議題曝光數', 0)}串"},
                    {"項目": "內文指名度", "說明": "提及品牌篇數", "達成": "(手動填)"},
                    {"項目": "好評增加數", "說明": "品牌正面好評篇數(含PTT/Dcard/Threads/社團)", "達成": f"{kpi_dict.get('好評增加數', 0)}篇"},
                ]

                word_分布 = []
                totals = {'實際': 0, '正': 0, '負': 0, '議': 0, '產': 0, '複': 0, '小計': 0}
                for cat_name in ["PTT", "Dcard", "Threads", "其他版面", "FB社團"]:
                    d = 分布_dict.get(cat_name, {})
                    if not d: continue
                    實際 = d.get("實際溝通", 0)
                    正 = d.get("正面討論", 0); 負 = d.get("負面討論", 0)
                    議 = d.get("議題討論", 0); 產 = d.get("產品討論", 0); 複 = d.get("複合討論", 0)
                    小計 = 正 + 負 + 議 + 產 + 複
                    if 實際 == 0 and 小計 == 0:
                        continue
                    word_分布.append({
                        "版面": cat_name, "實際溝通": 實際, "正面": 正, "負面": 負,
                        "議題": 議, "產品": 產, "複合": 複, "小計": 小計,
                    })
                    totals['實際'] += 實際; totals['正'] += 正; totals['負'] += 負
                    totals['議'] += 議; totals['產'] += 產; totals['複'] += 複; totals['小計'] += 小計
                word_分布.append({
                    "版面": "總計", "實際溝通": totals['實際'], "正面": totals['正'], "負面": totals['負'],
                    "議題": totals['議'], "產品": totals['產'], "複合": totals['複'], "小計": totals['小計'],
                })

                pdf_defaults = {
                    '執行篇數細項': 細項_desc,
                    '預計執行篇數': 總篇數, '實際執行篇數': 總篇數, '達成率': "100%",
                    'KPI': word_kpi,
                    '實際溝通量': 總篇數, '保證網友回應數': 保證數,
                    '實際回應數': kpi_dict.get('網友回應數', 0),
                    '分布表': word_分布,
                }

                word_data = closure_report.render_word_form(pdf_defaults)
                if word_data:
                    ts2 = datetime.now().strftime('%Y%m%d_%H%M%S')
                    with st.spinner("📝 生成 Word 報告 + 結案指令詞中..."):
                        try:
                            word_bytes = closure_report.build_word_report(word_data)
                            st.session_state["generated_word"] = {
                                "bytes": word_bytes,
                                "name": f"結案資訊整理_{ts2}.docx",
                            }
                        except Exception as e:
                            st.session_state["generated_word"] = None
                            st.error(f"❌ Word 生成失敗:{e}")
                            import traceback
                            st.code(traceback.format_exc())
                        try:
                            prompt_text = closure_report.build_prompt_text(word_data)
                            st.session_state["generated_prompt"] = {
                                "text": prompt_text,
                                "name": f"結案指令詞_{ts2}.txt",
                            }
                        except Exception as e:
                            st.session_state["generated_prompt"] = None
                            st.error(f"❌ 結案指令詞生成失敗:{e}")

                gw = st.session_state.get("generated_word")
                gp = st.session_state.get("generated_prompt")
                if gw or gp:
                    st.markdown("---")
                    st.markdown("### 📦 可下載檔案")
                if gw:
                    st.success(f"✅ Word 報告已生成:{gw['name']}")
                    st.download_button(
                        label=f"📥 下載 {gw['name']}",
                        data=gw['bytes'], file_name=gw['name'],
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True, type="primary", key="word_download",
                    )
                if gp:
                    st.success(f"✅ 結案指令詞已生成:{gp['name']}(給 Gamma / Claude 做簡報用)")
                    st.download_button(
                        label=f"📝 下載 {gp['name']}",
                        data=gp['text'].encode('utf-8'), file_name=gp['name'],
                        mime="text/plain", use_container_width=True, key="prompt_download",
                    )
                    with st.expander("👀 預覽結案指令詞內容"):
                        st.text(gp['text'])
                if gw or gp:
                    if st.button("🗑 清除上次生成的檔案", key="clear_generated"):
                        st.session_state["generated_word"] = None
                        st.session_state["generated_prompt"] = None
                        st.rerun()

        # ----- 月報表:額外 Word + 指令詞 -----
        else:
            st.info("💡 **提醒**:\n"
                    "- 「整體口碑操作議題方向」、「網友正評/負評/複合討論」需要您手動填入\n"
                    "- 「總覽整理」頁簽不會被修改(保留模板原樣)")

            st.divider()
            with st.expander("📄 生成 Word 月報資訊整理 + 簡報指令詞(需補填部分欄位)", expanded=False):
                st.markdown("Word 包含 Excel 沒有的內容(網路聲量比較、正負面討論主題、操作觀察等),需補填。簡報指令詞可直接貼給 Gamma/Claude 做簡報。")

                # 從主表自動算的 pdf_defaults
                細項_parts2 = []
                for name, cnt in (post_types or {}).items():
                    if cnt and cnt > 0:
                        short_name = name.split('(')[0].strip() if '(' in name else name
                        細項_parts2.append(f"{short_name} {cnt} 篇")
                monthly_總篇數 = sum(r["專案發文量"] for r in main_data)
                monthly_細項_desc = "、".join(細項_parts2) if 細項_parts2 else f"共 {monthly_總篇數} 篇"

                monthly_總回應 = sum(r["網友回應量"] for r in main_data)
                monthly_正 = sum(r["正面討論"] for r in main_data)
                monthly_負 = sum(r["負面討論"] for r in main_data)
                monthly_議 = sum(r["議題討論"] for r in main_data)
                monthly_產 = sum(r["產品討論"] for r in main_data)
                monthly_複 = sum(r["複合討論"] for r in main_data)

                cat_sum = {"PTT": 0, "DCARD": 0, "THREAD": 0, "其他版面": 0}
                for r in main_data:
                    cat_sum[monthly_report.categorize_station_monthly(r["站版"])] += r["網友回應量"]
                monthly_版面分佈 = [
                    {"站版": "PTT", "網友回應數": cat_sum["PTT"]},
                    {"站版": "DCARD", "網友回應數": cat_sum["DCARD"]},
                    {"站版": "THREAD", "網友回應數": cat_sum["THREAD"]},
                    {"站版": "其他版面", "網友回應數": cat_sum["其他版面"]},
                    {"站版": "合計", "網友回應數": sum(cat_sum.values())},
                ]

                monthly_pdf_defaults = {
                    '專案發文量_細項': monthly_細項_desc,
                    '專案發文量_總': monthly_總篇數,
                    '網友回應量': monthly_總回應,
                    '回應_正面': monthly_正, '回應_負面': monthly_負, '回應_議題': monthly_議,
                    '回應_產品': monthly_產, '回應_複合': monthly_複,
                    '各版面回應分佈': monthly_版面分佈,
                }

                monthly_data = monthly_report.render_monthly_form(monthly_pdf_defaults)
                if monthly_data:
                    ts3 = datetime.now().strftime('%Y%m%d_%H%M%S')
                    with st.spinner("📝 生成月報 Word + 簡報指令詞中..."):
                        try:
                            m_word_bytes = monthly_report.build_monthly_report(monthly_data)
                            st.session_state["monthly_generated_word"] = {
                                "bytes": m_word_bytes,
                                "name": f"月報資訊整理_{ts3}.docx",
                            }
                        except Exception as e:
                            st.session_state["monthly_generated_word"] = None
                            st.error(f"❌ Word 生成失敗:{e}")
                            import traceback
                            st.code(traceback.format_exc())
                        try:
                            m_prompt_text = monthly_report.build_monthly_prompt_text(monthly_data)
                            st.session_state["monthly_generated_prompt"] = {
                                "text": m_prompt_text,
                                "name": f"月報簡報指令詞_{ts3}.txt",
                            }
                        except Exception as e:
                            st.session_state["monthly_generated_prompt"] = None
                            st.error(f"❌ 簡報指令詞生成失敗:{e}")

                mgw = st.session_state.get("monthly_generated_word")
                mgp = st.session_state.get("monthly_generated_prompt")
                if mgw or mgp:
                    st.markdown("---")
                    st.markdown("### 📦 可下載檔案")
                if mgw:
                    st.success(f"✅ 月報 Word 已生成:{mgw['name']}")
                    st.download_button(
                        label=f"📥 下載 {mgw['name']}",
                        data=mgw['bytes'], file_name=mgw['name'],
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True, type="primary", key="monthly_word_download",
                    )
                if mgp:
                    st.success(f"✅ 簡報指令詞已生成:{mgp['name']}(給 Gamma / Claude 做簡報用)")
                    st.download_button(
                        label=f"📝 下載 {mgp['name']}",
                        data=mgp['text'].encode('utf-8'), file_name=mgp['name'],
                        mime="text/plain", use_container_width=True, key="monthly_prompt_download",
                    )
                    with st.expander("👀 預覽簡報指令詞內容"):
                        st.text(mgp['text'])
                if mgw or mgp:
                    if st.button("🗑 清除上次生成的月報檔案", key="clear_monthly_generated"):
                        st.session_state["monthly_generated_word"] = None
                        st.session_state["monthly_generated_prompt"] = None
                        st.rerun()

finally:
    if os.path.exists(pdf_path):
        os.unlink(pdf_path)
