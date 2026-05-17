"""
PDF → Excel 轉換器
功能:上傳電子原生 PDF,自動抽取所有表格,轉換成 Excel 下載
"""
import streamlit as st
import pdfplumber
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
from datetime import datetime
import re

st.set_page_config(page_title="PDF → Excel 轉換器", page_icon="📄", layout="wide")

# ==================== 工具函式 ====================

def clean_cell(value):
    """清理單一儲存格:去除多餘空白、換行符"""
    if value is None:
        return ""
    text = str(value).strip()
    # 把表格內換行統一成空格
    text = re.sub(r'\s+', ' ', text)
    return text


def clean_table(raw_table):
    """清理整個表格,並處理空欄位名稱"""
    if not raw_table or len(raw_table) < 1:
        return None
    
    # 清理每一格
    cleaned = [[clean_cell(c) for c in row] for row in raw_table]
    
    # 第一列當表頭
    headers = cleaned[0]
    rows = cleaned[1:] if len(cleaned) > 1 else []
    
    # 處理空表頭、重複表頭
    final_headers = []
    seen = {}
    for i, h in enumerate(headers):
        name = h if h else f"欄位{i+1}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 1
        final_headers.append(name)
    
    # 確保每一列長度與表頭一致
    fixed_rows = []
    for row in rows:
        if len(row) < len(final_headers):
            row = row + [""] * (len(final_headers) - len(row))
        elif len(row) > len(final_headers):
            row = row[:len(final_headers)]
        fixed_rows.append(row)
    
    # 移除完全空白的列
    fixed_rows = [r for r in fixed_rows if any(cell for cell in r)]
    
    if not fixed_rows:
        return None
    
    return pd.DataFrame(fixed_rows, columns=final_headers)


def extract_tables_from_pdf(pdf_file):
    """從 PDF 抽取所有表格,回傳 [(來源描述, DataFrame), ...]"""
    results = []
    with pdfplumber.open(pdf_file) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()
            for table_idx, raw in enumerate(tables, start=1):
                df = clean_table(raw)
                if df is not None and not df.empty:
                    label = f"第{page_idx}頁_表{table_idx}"
                    results.append((label, df))
    return results


def extract_text_from_pdf(pdf_file):
    """抽取所有純文字 (萬一找不到表格時備用)"""
    text_pages = []
    with pdfplumber.open(pdf_file) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            text_pages.append((f"第{i}頁", text))
    return text_pages


def style_worksheet(ws, df):
    """套用表頭樣式、自動欄寬"""
    # 表頭樣式
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", start_color="4472C4")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(border_style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    
    # 設定表頭
    for col_idx, _ in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = border
    
    # 內容樣式 + 框線
    body_font = Font(name="Arial", size=10)
    body_align = Alignment(vertical="center", wrap_text=True)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.font = body_font
            cell.alignment = body_align
            cell.border = border
    
    # 自動欄寬 (依據內容最長字元)
    for col_idx, col_name in enumerate(df.columns, start=1):
        col_letter = get_column_letter(col_idx)
        max_len = len(str(col_name))
        for value in df.iloc[:, col_idx - 1].astype(str):
            max_len = max(max_len, len(value))
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 50)
    
    # 凍結首列
    ws.freeze_panes = "A2"


def build_excel(tables, mode="multi"):
    """
    把表格資料建成 Excel
    mode: "multi" = 每個表一個工作表; "single" = 全部合併到一個工作表
    """
    wb = Workbook()
    wb.remove(wb.active)  # 移除預設空白工作表
    
    if mode == "multi":
        used_names = set()
        for label, df in tables:
            # Excel 工作表名稱限制 31 字、不能含特殊字元
            safe_name = re.sub(r'[\\/*?:\[\]]', '_', label)[:31]
            # 避免重名
            base = safe_name
            counter = 1
            while safe_name in used_names:
                safe_name = f"{base[:28]}_{counter}"
                counter += 1
            used_names.add(safe_name)
            
            ws = wb.create_sheet(safe_name)
            # 寫入表頭
            ws.append(list(df.columns))
            # 寫入資料
            for _, row in df.iterrows():
                ws.append(list(row.values))
            style_worksheet(ws, df)
    
    else:  # single 模式:合併
        ws = wb.create_sheet("合併資料")
        first = True
        for label, df in tables:
            # 加上來源標記列
            ws.append([f"=== {label} ==="])
            source_cell = ws.cell(row=ws.max_row, column=1)
            source_cell.font = Font(bold=True, color="C00000", size=11)
            
            ws.append(list(df.columns))
            header_row_idx = ws.max_row
            for _, row in df.iterrows():
                ws.append(list(row.values))
            
            # 簡單套用表頭樣式
            for col_idx in range(1, len(df.columns) + 1):
                c = ws.cell(row=header_row_idx, column=col_idx)
                c.font = Font(name="Arial", bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", start_color="4472C4")
                c.alignment = Alignment(horizontal="center")
            
            ws.append([])  # 空白行分隔
        
        # 自動欄寬
        for col_idx in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col_idx)
            max_len = 10
            for row in ws.iter_rows(min_col=col_idx, max_col=col_idx):
                for cell in row:
                    if cell.value:
                        max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = min(max_len + 2, 50)
    
    output = BytesIO()
    wb.save(output)
    return output.getvalue()


# ==================== UI ====================

st.title("📄 PDF → Excel 轉換器")
st.caption("上傳電子原生 PDF,自動抽取表格並轉換成 Excel")

with st.sidebar:
    st.header("⚙️ 設定")
    output_mode = st.radio(
        "輸出模式",
        ["每個表格獨立工作表", "全部合併到單一工作表"],
        help="多個表格時的處理方式"
    )
    show_preview = st.checkbox("顯示預覽", value=True)
    st.divider()
    st.markdown("""
    **支援類型**
    - ✅ 電子原生 PDF (文字可選取)
    - ✅ 表格清楚有框線
    - ❌ 掃描的圖片 PDF
    - ❌ 複雜合併儲存格表格
    """)

uploaded_file = st.file_uploader("📤 拖曳或點選上傳 PDF", type=["pdf"])

if uploaded_file:
    file_size = len(uploaded_file.getvalue()) / 1024
    st.info(f"📎 已上傳: **{uploaded_file.name}** ({file_size:.1f} KB)")
    
    with st.spinner("🔍 正在分析 PDF 並抽取表格..."):
        try:
            tables = extract_tables_from_pdf(uploaded_file)
        except Exception as e:
            st.error(f"❌ PDF 解析失敗: {e}")
            st.stop()
    
    if not tables:
        st.warning("⚠️ 沒有偵測到任何表格")
        with st.expander("📃 查看 PDF 純文字內容(用於除錯)"):
            uploaded_file.seek(0)
            text_pages = extract_text_from_pdf(uploaded_file)
            for label, text in text_pages:
                st.markdown(f"**{label}**")
                st.text(text[:2000] if text else "(無文字)")
        st.stop()
    
    st.success(f"✅ 成功偵測到 **{len(tables)}** 個表格")
    
    # 預覽
    if show_preview:
        st.subheader("👀 表格預覽")
        for label, df in tables:
            with st.expander(f"📊 {label} ({len(df)} 列 × {len(df.columns)} 欄)", expanded=True):
                st.dataframe(df, use_container_width=True, hide_index=True)
    
    # 統計資訊
    col1, col2, col3 = st.columns(3)
    col1.metric("📊 表格總數", len(tables))
    col2.metric("📝 總列數", sum(len(df) for _, df in tables))
    col3.metric("📋 總欄數", sum(len(df.columns) for _, df in tables))
    
    st.divider()
    
    # 產生 Excel
    mode = "multi" if "獨立" in output_mode else "single"
    excel_bytes = build_excel(tables, mode=mode)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_name = uploaded_file.name.rsplit(".", 1)[0]
    output_filename = f"{base_name}_{timestamp}.xlsx"
    
    st.download_button(
        label="📥 下載 Excel 檔案",
        data=excel_bytes,
        file_name=output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )

else:
    st.info("👆 請上傳一個 PDF 檔案開始")
    
    with st.expander("💡 使用說明 / 常見問題"):
        st.markdown("""
        **這個工具能做什麼?**
        - 自動抽取 PDF 中的所有表格
        - 把每個表格轉成 Excel 工作表
        - 套用專業樣式(表頭、框線、自動欄寬)
        - 凍結首列方便檢視
        
        **效果不好怎麼辦?**
        1. 確認 PDF 是「文字可選取」的版本(不是掃描檔)
        2. 表格如果沒有清楚的格線,辨識率會降低
        3. 跨頁表格可能會被切成兩個表(可選擇「合併」模式)
        4. 複雜的合併儲存格可能解析錯誤
        
        **完全沒抽到表格?**
        - 點開「查看 PDF 純文字內容」確認 PDF 是否真的有文字
        - 如果是掃描檔,需要 OCR(本工具暫不支援)
        """)
